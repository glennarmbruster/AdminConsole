#Requires -Version 5.1
<#
    Common.ps1
    Shared helper library for the Admin Console launcher and every
    standalone tool script (MailboxHealth.ps1, MessageTrace.ps1,
    InboxRules.ps1, etc.).

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    2026-09-16  Get-UserToolGroupMap: this operator's tool grouping.
    2026-09-16  The group ring on flat buttons is painted, not set as a
                FlatAppearance border: one border colour could not carry the
                keyline, so the ring was invisible on pale themes.
                Write-JsonFileAtomic cleans up its temp file in a finally.
    2026-09-16  Get-TenantConfig re-reads when the tenant file changes, the same
                staleness as the profile: Tenant Manager, Theme Manager and
                Park In Place all write it from their own processes.
    2026-09-16  Get-UserProfile notices the profile file changing underneath its
                cache. A theme saved as 'just me' from Theme Manager -- a
                separate process -- was invisible to a console re-run in the
                same PowerShell window, which still held the pre-save copy.
    2026-09-16  Group buttons -- one console button standing for several tools --
                are ringed in white, in both the raised and flat button
                styles, and keep the ring when greyed out.
    2026-09-16  Write-JsonFileAtomic: config writes are temp-file-and-rename, so a
                crash cannot leave a truncated file that reads as 'no settings'.
                Get-UserLayoutOwn for read-modify-write, so a 'just me' save no
                longer freezes a copy of the every-tenant block into one tenant.
                Resetting one tenant no longer clears the every-tenant block.
#>

#region Tenant configuration ----------------------------------------------------
<#
    EVERY tenant-specific fact enters the toolkit here and nowhere else.

    What used to sit at the top of this file:

        $script:MailDomain            = 'a real domain'
        $script:FranchiseDCs          = @( eight real server names )
        $script:HQDomain              = 'another real domain'
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

#region Multi-tenant ------------------------------------------------------------
<#
    ONE TOOLKIT FOLDER, SEVERAL TENANTS.

    The JSON files were the only thing separating one tenant's console from
    another's, so a second tenant meant a second copy of a hundred scripts --
    and every fix then had to be made twice, which is how the two copies drift.

    Instead: Config\Tenants\<name>.json, one per tenant, and a switch that says
    which to load.

        AdminConsole.ps1                     the default tenant
        AdminConsole.ps1 -Tenant wxyz        that file
        AdminConsole.ps1 -Tenant abcd        a nickname works too

    ABSENT MEANS OFF, AND THAT IS THE WHOLE REVERT PLAN. With no
    Config\Tenants folder, Get-TenantConfigPath returns Config\Tenant.json
    exactly as it always did and every function below is a no-op. Delete the
    folder and the toolkit is what it was this morning; nothing else has to be
    put back. That is deliberate -- a change to the file every tool depends on
    should be undoable by deleting something, not by restoring something.

    THE ACTIVE TENANT IS A GLOBAL, and that is not laziness. The console
    launches every tool with '&' INTO ITS OWN PROCESS (there are no
    NewProcess entries in Tools.json), so a global set once at startup is
    inherited by every tool without a single one of them having to declare a
    -Tenant parameter or remember to pass it on. A script-scoped variable in
    Common.ps1 would NOT be: a tool dot-sources its own copy of this file and
    gets its own script scope.
#>
$script:TenantsFolderName = 'Tenants'
$script:TenantIndexName   = 'Tenants.json'

function Get-TenantsFolder {
    <# Config\Tenants, NOT created on demand -- its absence is the off switch. #>
    return (Join-Path (Join-Path $PSScriptRoot $script:ConfigFolderName) $script:TenantsFolderName)
}

function Test-MultiTenantMode {
    <#
        Is this install multi-tenant? True when Config\Tenants exists AND holds
        at least one .json. A folder somebody created and left empty is not a
        configuration, and treating it as one would break a working console
        with no tenant files to fall back on.
    #>
    $dir = Get-TenantsFolder
    if (-not (Test-Path -LiteralPath $dir)) { return $false }
    return (@(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -ne $script:TenantIndexName }).Count -gt 0)
}

function Get-TenantRegistry {
    <#
        Every tenant this install knows about: Key, Nickname, File, Default.

        THE FILES ARE THE TRUTH, THE INDEX IS DECORATION. The list comes from
        what is actually in Config\Tenants; Tenants.json only adds nicknames and
        marks the default. So dropping a .json into the folder is enough to make
        a tenant appear, and a stale index cannot hide a tenant that exists or
        invent one that does not -- which is the failure mode of an index that
        is allowed to be authoritative.
    #>
    $dir = Get-TenantsFolder
    if (-not (Test-Path -LiteralPath $dir)) { return @() }

    $index = $null
    $ip = Join-Path $dir $script:TenantIndexName
    if (Test-Path -LiteralPath $ip) {
        try { $index = Get-Content -LiteralPath $ip -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { }
    }

    $meta = @{}
    $defaultKey = ''
    if ($index) {
        if ($index.PSObject.Properties['Default']) { $defaultKey = "$($index.Default)".Trim().ToLowerInvariant() }
        foreach ($e in @($index.Tenants)) {
            if (-not $e -or -not $e.PSObject.Properties['Key']) { continue }
            $k = "$($e.Key)".Trim().ToLowerInvariant()
            if ($k) { $meta[$k] = $e }
        }
    }

    # A PLAIN ARRAY, NOT List[object], AND THAT IS A BUG FIX NOT A STYLE CHOICE.
    #
    #     New-Object System.Collections.Generic.List[object]
    #
    # unquoted does NOT build a List of object. PowerShell parses the trailing
    # [object] as an ARGUMENT, so it calls New-Object -TypeName ...List
    # -ArgumentList ([object]) and looks for a List constructor taking a Type.
    # There is none, so it throws 'Argument types do not match' -- and it throws
    # at the first line that TOUCHES $out, which sent the error to
    # `$list = @($out)` and hid the real cause completely.
    #
    # Quoting the type name fixes it too. An array is used instead because this
    # is a handful of tenants, the performance difference is nothing, and @()
    # cannot be got wrong.
    $out = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        if ($f.Name -eq $script:TenantIndexName) { continue }
        $key  = $f.BaseName
        $k    = $key.ToLowerInvariant()
        $nick = ''
        if ($meta.ContainsKey($k) -and $meta[$k].PSObject.Properties['Nickname']) {
            $nick = "$($meta[$k].Nickname)".Trim()
        }
        $out += [pscustomobject]@{
            Key       = $key
            Nickname  = $nick
            File      = $f.FullName
            IsDefault = ($k -eq $defaultKey)
        }
    }

    # NOTHING MARKED DEFAULT? THE FIRST ONE IS IT, and the console says so. An
    # install with tenants but no default must still start: refusing to open
    # because an index key is missing would be a self-inflicted outage.
    $list = @($out)
    if ($list.Count -gt 0 -and -not (@($list | Where-Object { $_.IsDefault }).Count)) {
        $list[0].IsDefault = $true
    }
    return $list
}

function Resolve-TenantKey {
    <#
        Turn what somebody typed into a tenant Key. Matches the file name, the
        nickname, or either with the primary domain's suffix dropped -- so
        'wxyz', 'WXYZ.com' and a nickname all find the same file.

        Returns '' when nothing matches. The CALLER decides what that means;
        this function does not guess a tenant, because guessing which
        organisation to run against is the one thing it must never do.
    #>
    param([string]$Name)

    $want = "$Name".Trim()
    if (-not $want) { return '' }
    $want = ($want -replace '\.json$', '').Trim().ToLowerInvariant()

    $reg = Get-TenantRegistry
    foreach ($t in $reg) { if ($t.Key.ToLowerInvariant() -eq $want) { return $t.Key } }
    foreach ($t in $reg) { if ($t.Nickname -and $t.Nickname.ToLowerInvariant() -eq $want) { return $t.Key } }
    # Last: match on the leading label, so 'wxyz' finds 'wxyz.com'.
    $lead = ($want -split '\.')[0]
    foreach ($t in $reg) {
        if (($t.Key -split '\.')[0].ToLowerInvariant() -eq $lead) { return $t.Key }
        if ($t.Nickname -and ($t.Nickname -split '\.')[0].ToLowerInvariant() -eq $lead) { return $t.Key }
    }
    return ''
}

function Set-ActiveTenant {
    <#
        Choose the tenant for THIS PROCESS. Called once, by the console, before
        anything reads config.

        Returns the resolved key, or '' when multi-tenant mode is off. Throws on
        a name that does not match -- see the comment in Resolve-TenantKey: a
        typo must stop the console, not quietly open somebody else's directory.
    #>
    param([string]$Name)

    if (-not (Test-MultiTenantMode)) {
        $global:ActiveTenantKey = ''
        return ''
    }

    $reg = Get-TenantRegistry
    $key = ''
    if ("$Name".Trim()) {
        $key = Resolve-TenantKey -Name $Name
        if (-not $key) {
            $known = (@($reg | ForEach-Object {
                if ($_.Nickname) { "$($_.Key) ($($_.Nickname))" } else { $_.Key }
            }) -join ', ')
            throw "No tenant matches '$Name'. Known tenants: $known"
        }
    } else {
        $d = @($reg | Where-Object { $_.IsDefault })
        if ($d.Count -gt 0) { $key = $d[0].Key }
    }

    # A SECOND TENANT IN A CONNECTED PROCESS IS REFUSED.
    #
    # Swapping the config file does not swap the Exchange Online session, the
    # on-premises PSSession, SPO, Graph, or whatever PIM activated -- those
    # belong to this powershell.exe and outlive the console window that made
    # them. Running the console for a second tenant from the same prompt gives
    # you tenant B's configuration driving tenant A's live connections, and
    # nothing on screen would say so.
    #
    # THE FIRST TENANT IN A PROCESS IS ALWAYS ALLOWED, connected or not: the
    # connections at that point are whatever the operator made deliberately,
    # and this is not the place to second-guess them. Only a CHANGE is blocked.
    #
    # Disconnecting instead of refusing was considered and rejected: tokens
    # cache, modules keep state, and Import-PSSession leaves proxy functions
    # that still resolve after their session is gone. 'Probably clean' is not
    # a basis for pointing a disable tool at a directory.
    $prev = "$($global:ActiveTenantKey)"
    if ($prev -and $key -and $prev -ne $key) {
        $live = @(Get-ToolkitLiveConnection)
        if ($live.Count -gt 0) {
            throw ("This PowerShell window is signed in to $($live -join ', ') for tenant '$prev'. " +
                   "Those connections belong to the window, not to the console, so opening '$key' here " +
                   "would run $key's tools against $prev's sessions.`r`n`r`n" +
                   "Open a NEW PowerShell window and run:`r`n" +
                   "    .\AdminConsole.ps1 -Tenant $key")
        }
    }

    $global:ActiveTenantKey = $key
    $global:TenantConfigStamp = $null

    # INTO THE ENVIRONMENT AS WELL, AND THIS IS THE LINE THAT MAKES THE
    # DASHBOARD TOOLS WORK.
    #
    # A $global: is inherited by tools the console launches with '&' -- the
    # ones on the grid. It is NOT inherited by a NEW powershell.exe, and that is
    # how the Dashboard opens Theme Manager, Tenant Manager, Park In Place, Tool
    # Manager, the tester: Start-ToolDetached. In those processes the global is
    # empty, Get-TenantConfigPath falls back to Config\Tenant.json, and every
    # one of them silently edits a file the console is not reading. Reported
    # 2026-09-14 as "I change the theme and nothing happens" -- Theme Manager
    # had written the theme, correctly, to the wrong tenant file.
    #
    # Environment variables ARE inherited by child processes, whichever way they
    # are started. Get-ActiveTenantKey reads this when the global is empty.
    #
    # Always written, never only-when-set, so a stale value from an earlier
    # console in the same window is overwritten rather than trusted -- the same
    # lesson -SafeMode taught about $env: being per-window, not per-script.
    $env:ADMINCONSOLE_TENANT = $key

    # The config cache is keyed to no tenant at all, so it MUST be dropped here
    # -- otherwise the first Get-TenantConfig after switching returns whoever
    # was loaded last. That is the wrong-tenant bug in its purest form.
    $global:TenantConfig = $null
    return $key
}

function Get-ActiveTenantKey {
    <#
        The tenant in force, or '' for single-tenant installs.

        THE GLOBAL FIRST, THEN THE ENVIRONMENT. The global is what the console
        set in this process; the environment is what a PARENT console set before
        launching this process detached. A Dashboard tool has no global and
        needs the second. See the note in Set-ActiveTenant.

        The environment value is only honoured when a matching tenant FILE
        exists. A leftover variable from a tenant that has since been deleted
        must read as 'no tenant', not as a path to a file that is not there.
    #>
    if ($global:ActiveTenantKey) { return "$($global:ActiveTenantKey)" }
    $e = "$($env:ADMINCONSOLE_TENANT)".Trim()
    if ($e) {
        try {
            if (Test-Path -LiteralPath (Join-Path (Get-TenantsFolder) "$e.json")) { return $e }
        } catch { }
    }
    return ''
}

function Get-ToolkitLiveConnection {
    <#
        Which services this POWERSHELL PROCESS is signed in to, by name.

        WHY THIS MATTERS MORE THAN THE CONFIG FILE.
        -------------------------------------------
        Switching tenants swaps a JSON file. It does NOT swap:

            Connect-ExchangeOnline  -- a live session and an access token
            the on-premises PSSession and its imported cmdlets
            Connect-SPOService
            Connect-MgGraph
            whatever PIM activated for the roles you hold THERE

        Those belong to the PowerShell process, not to the console, and closing
        the console window does not end them. So running

            .\AdminConsole.ps1 -Tenant abcd

        from a prompt that earlier ran the WXYZ console gives you WXYZ's
        Exchange Online session with ABCD's config on top -- a console that is
        confidently labelled one tenant while every cmdlet it runs reaches the
        other. That is worse than either tenant alone, and nothing on screen
        would say so.

        A NEW POWERSHELL WINDOW IS THE ONLY CLEAN ANSWER. Disconnecting looks
        tempting and is not equivalent: tokens get cached, modules keep state,
        and Import-PSSession leaves proxy functions behind that still resolve
        after the session they proxied is gone. A fresh process has none of
        that by construction.

        Returns a list of service names, empty when the process is clean. Each
        probe is wrapped because they are allowed to fail -- a probe that
        throws must read as 'not connected', not take the console down.
    #>
    $live = @()
    try { if ((Test-ExoConnectionState).Connected) { $live += 'Exchange Online' } } catch { }
    try { if ((Test-EmsAvailable).Available)       { $live += 'Exchange on-premises' } } catch { }
    try { if ((Test-SpoConnectionState).Connected) { $live += 'SharePoint Online' } } catch { }
    try { if ((Test-GraphConnectionState).Connected) { $live += 'Microsoft Graph' } } catch { }
    return @($live)
}

function Start-TenantConsole {
    <#
        Open a tenant in a NEW PowerShell process.

        This is the supported way to change tenant, and it is a new process
        rather than a new window of this one for the reason in
        Get-ToolkitLiveConnection: the connections are the process's.

        -NoExit is deliberate. The console is a WinForms window and the shell
        behind it is where errors land -- a startup failure on a shell that
        closed itself takes its own explanation with it, which is exactly the
        case somebody needs it for.
    #>
    param([Parameter(Mandatory)][string]$Tenant)

    $key = Resolve-TenantKey -Name $Tenant
    if (-not $key) { throw "No tenant matches '$Tenant'." }

    $console = Join-Path $PSScriptRoot 'AdminConsole.ps1'
    if (-not (Test-Path -LiteralPath $console)) { throw "AdminConsole.ps1 not found beside Common.ps1." }

    # powershell.exe, NOT pwsh: this toolkit is 5.1 (#Requires at the top of
    # every script) and the on-premises Exchange tools do not load in 7.
    Start-Process -FilePath 'powershell.exe' `
        -ArgumentList @('-NoExit', '-ExecutionPolicy', 'Bypass',
                        '-File', "`"$console`"", '-Tenant', "`"$key`"")
    return $key
}

function Get-ActiveTenantLabel {
    <#
        What to show a human: the nickname when there is one, else the key.
        Empty on a single-tenant install, so callers can print it unconditionally
        and get nothing rather than "(none)".
    #>
    $k = Get-ActiveTenantKey
    if (-not $k) { return '' }
    foreach ($t in Get-TenantRegistry) {
        if ($t.Key -eq $k) { return $(if ($t.Nickname) { $t.Nickname } else { $t.Key }) }
    }
    return $k
}
#endregion

#region Park in place -----------------------------------------------------------
<#
    PARKED IN PLACE: the script stays in the folder, the button does not appear,
    and the additions scanner stops offering it.

    NOT THE SAME AS DisabledTools\. That moves the FILE, for the whole install,
    for everybody -- right when a tool is retired. This is per TENANT and moves
    nothing: the on-premises Exchange tools are real tools that this toolkit
    should keep, they are simply meaningless in a cloud-only tenant. One folder
    serving three tenants cannot say that by moving files, because the file
    is shared and the answer is not.

    So parking is a LIST, per tenant, of scripts the console should behave as
    though it had never been given. Un-park by removing the entry; the file was
    never touched.
#>
$script:ParkedFileName = 'Parked.json'

function Get-ParkedListPath {
    <#
        Where this tenant's parked list lives: IN THE TENANT FILE, under a
        'Parked' block. Everything about a tenant is in its one file --
        capabilities, directories, theme, band, and what it does not show.

        The first version wrote a separate <key>.parked.json beside the tenant
        file. That was a fourth place per tenant to look, so it was folded in
        (2026-09-14). Get-ParkedLegacyPath still names the old file: it is READ
        when the tenant file has no Parked block, and renamed away on the
        first save, so nothing already parked is lost by upgrading.
    #>
    return (Get-TenantConfigPath)
}

function Get-ParkedLegacyPath {
    $k = Get-ActiveTenantKey
    if ($k) { return (Join-Path (Get-TenantsFolder) "$k.parked.json") }
    return (Join-Path (Join-Path $PSScriptRoot $script:ConfigFolderName) $script:ParkedFileName)
}

function Get-ParkedStore {
    <#
        The raw Parked block for the active tenant, from wherever it lives.

        Returns an object with .Parked (entries) and .Sections (names), or
        $null when nothing is parked anywhere. Tenant file first; the legacy
        side file only when the tenant file has no block at all, so once a
        save has landed in the tenant file the old file is never consulted
        again even if it is still on disk.
    #>
    # THROUGH THE TENANT CACHE, NOT OFF THE DISK. Get-TenantConfig already
    # holds this exact file, parsed, for the life of the process, and every
    # writer (Set-TenantFileKey, Tenant Manager) clears that cache. Reading
    # the file here again meant four reads and four JSON parses of the same
    # bytes on every console start -- once each for scripts and sections,
    # twice each because the console asks twice.
    $tj = $null
    try { $tj = Get-TenantConfig } catch { $tj = $null }
    if ($tj -and $tj.PSObject.Properties['Loaded'] -and $tj.Loaded) {
        if ($tj.PSObject.Properties['Parked'] -and $tj.Parked) {
            $blk = $tj.Parked
            # The block is { Scripts, Sections }. A bare array is tolerated as
            # "scripts only", because that is what a hand edit will produce.
            if ($blk -is [System.Array]) { return [pscustomobject]@{ Parked = @($blk); Sections = @() } }
            return [pscustomobject]@{
                Parked   = $(if ($blk.PSObject.Properties['Scripts'])  { @($blk.Scripts) }  else { @() })
                Sections = $(if ($blk.PSObject.Properties['Sections']) { @($blk.Sections) } else { @() })
            }
        }
    }
    $lp = Get-ParkedLegacyPath
    if (Test-Path -LiteralPath $lp) {
        $j = $null
        try { $j = Get-Content -LiteralPath $lp -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { return $null }
        if (-not $j) { return $null }
        return [pscustomobject]@{
            Parked   = $(if ($j.PSObject.Properties['Parked'])   { @($j.Parked) }   else { @() })
            Sections = $(if ($j.PSObject.Properties['Sections']) { @($j.Sections) } else { @() })
        }
    }
    return $null
}

function Get-ParkedScript {
    <#
        The parked entries for the active tenant.

        Returns objects with Script / Reason / When. No block anywhere, empty
        list -- nothing is parked, which is the right answer for a fresh
        install and means no tenant has to have one before the console starts.
    #>
    $j = Get-ParkedStore
    if (-not $j) { return @() }

    $out = @()
    foreach ($e in @($j.Parked)) {
        if ($null -eq $e) { continue }
        if ($e.PSObject.Properties['Script']) {
            $s = "$($e.Script)".Trim()
            if (-not $s) { continue }
            $out += [pscustomobject]@{
                Script = $s
                Reason = $(if ($e.PSObject.Properties['Reason']) { "$($e.Reason)" } else { '' })
                When   = $(if ($e.PSObject.Properties['When'])   { "$($e.When)"   } else { '' })
            }
        } else {
            # A bare string is accepted, because it is what somebody hand-editing
            # this file will write.
            $s = "$e".Trim()
            if ($s) { $out += [pscustomobject]@{ Script = $s; Reason = ''; When = '' } }
        }
    }
    return @($out)
}

function Get-ParkedKey {
    <#
        The one comparison key for a parked script: its path RELATIVE TO THE
        TOOLKIT ROOT, lowercased, backslashes, no leading '.\'.

        Root tools key as 'foo.ps1'. Project tools key as
        'projects\<project>\foo.ps1'. It is NOT the bare file name, because the
        projects reuse names -- '1-ONPREM-Export.ps1' exists in more than one
        project folder, and a leaf-only match would park every one of them
        when one was ticked. Root keys are unchanged from the leaf-keyed
        version, so an existing parked list still reads the same.

        Accepts any form: '.\Foo.ps1', './Projects/X/Foo.ps1', 'Foo.ps1', or a
        full path under the toolkit folder.
    #>
    param([Parameter()][string]$Script)
    $s = "$Script".Trim() -replace '/', '\'
    if (-not $s) { return '' }
    $root = "$PSScriptRoot".TrimEnd('\')
    if ($root -and $s.StartsWith($root + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        $s = $s.Substring($root.Length + 1)
    }
    while ($s.StartsWith('.\')) { $s = $s.Substring(2) }
    return $s.TrimStart('\').ToLowerInvariant()
}

function Get-ParkedScriptIndex {
    <#
        Parked scripts as a lookup, keyed by Get-ParkedKey -- the lowercased
        path relative to the toolkit root.

        Tools.json says '.\Foo.ps1', the scanner says 'Foo.ps1', somebody
        hand-editing says './foo.ps1', and all three mean one file. One
        normalising key is the only comparison right for all of them -- and
        getting this wrong shows up as a park that silently does nothing, which
        is the failure mode hardest to notice.
    #>
    $h = @{}
    foreach ($e in @(Get-ParkedScript)) {
        $k = Get-ParkedKey -Script "$($e.Script)"
        if ($k) { $h[$k] = $e }
    }
    return $h
}

function Get-ParkedSection {
    <#
        Section NAMES parked for the active tenant. The whole section -- its
        heading and every button in it -- is left out of that tenant's console.

        WHY A SECTION AND NOT JUST ITS SCRIPTS. Parking the eight Projects
        scripts hides Projects today; the ninth project added next month comes
        straight back onto a tenant that never wanted any of them. Parking the
        SECTION says what was meant: this tenant does not do these.

        Tools.json stays one file for every tenant. This is the per-tenant
        expression of "not here", and it is a list, not a second Tools.json --
        see ARCHITECTURE 53 for why a Tools.json per tenant was rejected.
    #>
    $j = Get-ParkedStore
    if (-not $j) { return @() }
    return @(@($j.Sections) | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
}

function Test-SectionParked {
    param([Parameter(Mandatory)][string]$Name)
    $n = "$Name".Trim().ToLowerInvariant()
    foreach ($s in Get-ParkedSection) { if ($s.ToLowerInvariant() -eq $n) { return $true } }
    return $false
}

function Test-ScriptParked {
    <# Is this script parked for the active tenant? Takes any path form. #>
    param([Parameter(Mandatory)][string]$Script)
    $k = Get-ParkedKey -Script $Script
    if (-not $k) { return $false }
    return (Get-ParkedScriptIndex).ContainsKey($k)
}

function Save-ParkedScript {
    <#
        Write the parked list for the active tenant.

        REPLACES THE WHOLE LIST. The caller passes what the list should BE, not
        a delta -- the UI already holds the full picture, and an add/remove API
        here would be a second place for the two to disagree.
    #>
    param(
        [Parameter()][object[]]$Entries = @(),
        # Section names to park alongside. Passed explicitly, not read back
        # from the file, so a caller that only knows about scripts cannot
        # accidentally wipe the sections -- the UI sends both every time.
        [Parameter()][string[]]$Sections = @()
    )

    $path = Get-ParkedListPath

    $clean = @()
    $seen  = @{}
    foreach ($e in @($Entries)) {
        if ($null -eq $e) { continue }
        $s = "$($e.Script)".Trim()
        if (-not $s) { continue }
        $k = Get-ParkedKey -Script $s
        if (-not $k -or $seen[$k]) { continue }
        $seen[$k] = $true
        $clean += [PSCustomObject][ordered]@{
            Script = $s
            Reason = "$($e.Reason)"
            When   = $(if ("$($e.When)".Trim()) { "$($e.When)" } else { (Get-Date -Format 'yyyy-MM-dd') })
        }
    }

    $block = [PSCustomObject][ordered]@{
        _comment = @(
            'PARKED IN PLACE for this tenant: the script stays in the folder,',
            'the button does not appear, Add-ConsoleAdditions stops offering it.',
            'Sections: whole sections left off this console, tools included.',
            'Edited by ParkInPlace.ps1 on the Dashboard. Un-park by removing',
            'the entry -- nothing was moved.'
        )
        Scripts  = @($clean)
        Sections = @(@($Sections) | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    }
    Set-TenantFileKey -Name 'Parked' -Value $block

    # THE OLD SIDE FILE IS RETIRED, NOT DELETED. Renamed so it stops being read
    # (Get-ParkedStore only reads it when the tenant file has no block, and it
    # now does) and stays on disk for anyone who wants to check what it said.
    try {
        $lp = Get-ParkedLegacyPath
        if (Test-Path -LiteralPath $lp) { Move-Item -LiteralPath $lp -Destination "$lp.migrated" -Force }
    } catch { }
    return $path
}

function Write-JsonFileAtomic {
    <#
        Write a config file so a crash cannot leave a truncated one.

        WriteAllText opens the real file and truncates it before the first
        byte lands. Every config file this toolkit owns degrades SILENTLY when
        it comes back empty -- a half-written Tenant.json reads as 'every
        capability off' and a half-written profile as 'no customisation' --
        so the failure looks like a settings change nobody made rather than
        like a crash. Writing a sibling temp file and renaming makes the
        switch atomic on NTFS: the file is either the old one or the new one.

        Replace, not Move, when the target exists: Move-Item onto an existing
        path fails, and deleting first reopens the window this closes.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Json
    )
    $text = ($Json -replace "`r?`n", "`r`n")
    $enc  = New-Object System.Text.UTF8Encoding($true)
    $tmp  = "$Path.$PID.tmp"
    # THE CLEANUP IS IN A finally, NOT AFTER THE FALLBACK. The realistic reason
    # Replace throws is the destination being locked -- by antivirus, by an SMB
    # server, by another process holding the config -- and that same lock makes
    # the fallback write throw too. With the cleanup sitting after it, the line
    # never ran and every failed save left a permanent <file>.<pid>.tmp beside
    # the real one. Writing the temp inside the try covers a failure there as
    # well, which used to leak a half-written file.
    try {
        [System.IO.File]::WriteAllText($tmp, $text, $enc)
        if (Test-Path -LiteralPath $Path) {
            [System.IO.File]::Replace($tmp, $Path, $null)
        } else {
            [System.IO.File]::Move($tmp, $Path)
        }
    } catch {
        # A share that refuses Replace (some SMB servers do) still has to end
        # up with the content, so the direct write stands as the fallback --
        # and is allowed to throw to the caller, because a save that did not
        # happen must not report success.
        [System.IO.File]::WriteAllText($Path, $text, $enc)
    } finally {
        if (Test-Path -LiteralPath $tmp) {
            try { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } catch { }
        }
    }
    return $Path
}

function Set-TenantFileKey {
    <#
        Write ONE top-level key into the active tenant file, leaving every
        other key exactly as it was.

        This is how anything that is "about the tenant" gets saved by a tool
        that is not Tenant Manager -- the theme, the parked block, the band.
        Read-modify-write on the whole file, so a key this toolkit has never
        heard of survives too. UTF-8 with BOM and CRLF like every config file.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter()]$Value
    )
    $tp = Get-TenantConfigPath
    $tj = $null
    if (Test-Path -LiteralPath $tp) {
        try { $tj = Get-Content -LiteralPath $tp -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { $tj = $null }
    }
    if (-not $tj) { $tj = [PSCustomObject]@{} }
    Add-Member -InputObject $tj -NotePropertyName $Name -NotePropertyValue $Value -Force
    $dir = Split-Path -Parent $tp
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    $null = Write-JsonFileAtomic -Path $tp -Json ($tj | ConvertTo-Json -Depth 10)
    # Every cached copy of the tenant file is now stale.
    $global:TenantConfig = $null
    $global:TenantConfigStamp = $null
    return $tp
}

function Get-ScriptVariant {
    <#
        The TOOLKIT-VARIANT a script declares, or '' when it declares none.

        Read from the file rather than kept in a table, because the marker is
        already in every script and a table would be a second copy that goes
        stale the first time somebody adds a tool without updating it.
    #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    try {
        # Only the head of the file: the marker is in the help block at the top,
        # and reading a 5000-line tool to find line 17 is a hundred tools' worth
        # of pointless IO every time this window opens.
        $head = Get-Content -LiteralPath $Path -TotalCount 60 -ErrorAction Stop
        foreach ($line in $head) {
            if ($line -match '^#\s*TOOLKIT-VARIANT:\s*(\S+)') { return $Matches[1].Trim() }
        }
    } catch { }
    return ''
}

function Get-ScriptsForVariant {
    <#
        Every .ps1 in the toolkit root declaring one of the given variants.

        THIS IS WHAT MAKES 'PARK ON-PREM EXCHANGE' ONE CLICK. Parking the
        Connect button alone leaves twenty tools on the grid that cannot work
        without it -- each one greying out with a different explanation at the
        moment somebody presses it. Parking the whole variant is the honest
        expression of 'this tenant does not have that'.
    #>
    param([Parameter(Mandatory)][string[]]$Variant)

    $want = @{}
    foreach ($v in $Variant) { $want["$v".Trim().ToLowerInvariant()] = $true }

    $out = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        $v = Get-ScriptVariant -Path $f.FullName
        if (-not $v) { continue }
        if ($want[$v.ToLowerInvariant()]) { $out += $f.Name }
    }
    return @($out)
}
#endregion

#region User profile ------------------------------------------------------------
<#
    ONE TOOLKIT, ONE Tools.json, AND A VIEW OF IT PER PERSON.

    Sections.json and Tools.json describe what the toolkit HAS. They are shared
    and they stay shared -- adding a tool is still one edit in one place. What
    is per person is how that toolkit is ARRANGED on screen: which sections
    exist, how many columns, how many rows each section gets, which tools sit
    in which section, which tools are grouped behind one button, which are
    hidden, and which theme.

    Before this, Arrange wrote straight back to Sections.json and Tools.json.
    Two operators on one folder meant the second one to drag a button silently
    rearranged the first one's console. That is the whole reason this exists.

    PER PERSON AND PER TENANT. The profile holds a Default block and a block
    per tenant key; the tenant block wins where it says anything, the Default
    fills the rest. So one person can run a dense grid in the tenant they live
    in and a short one in the tenant they touch twice a year, and two people
    on the same machine share nothing but the tools themselves.

    ABSENT MEANS UNCHANGED. No profile file, no overlay, and the console
    behaves exactly as it did before any of this -- which is also what
    'Reset to default' restores by deleting the file.
#>
$script:UsersFolderName = 'Users'

function Get-UserProfileKey {
    <#
        Who this profile belongs to: the WINDOWS ACCOUNT running the console.

        Not the admin UPN. The UPN needs a live connection, is not known at
        startup, and changes with which admin account signed in -- while the
        layout belongs to the person at the keyboard whether or not they have
        connected to anything yet.

        Sanitised to what is safe in a file name, so a domain-qualified or
        odd account name cannot produce a path.
    #>
    $u = "$env:USERNAME".Trim()
    if (-not $u) { $u = 'default' }
    $u = ($u -replace '[^A-Za-z0-9._-]', '_')
    return $u.ToLowerInvariant()
}

function Get-UsersFolder {
    return (Join-Path (Join-Path $PSScriptRoot $script:ConfigFolderName) $script:UsersFolderName)
}

function Get-UserProfilePath {
    param([string]$User)
    $k = $(if ("$User".Trim()) { ("$User".Trim() -replace '[^A-Za-z0-9._-]', '_').ToLowerInvariant() } else { Get-UserProfileKey })
    return (Join-Path (Get-UsersFolder) "$k.json")
}

function Test-UserProfile {
    <# Has this person customised anything yet? #>
    param([string]$User)
    return (Test-Path -LiteralPath (Get-UserProfilePath -User $User))
}

function Get-UserProfile {
    <#
        The whole profile file, cached in a GLOBAL -- and therefore for the
        life of the POWERSHELL WINDOW, not the life of the console.

        THE CACHE IS ONLY AS GOOD AS THE FILE IT CAME FROM, and it took a
        reported bug to make that concrete. Theme Manager runs DETACHED, in its
        own powershell.exe, and writes the chosen theme into this file. The
        console that launched it is still running in the original window with
        this global already filled -- so when the operator closes the console
        and re-runs it in that same window, or when the relaunch trampoline
        reopens it in-process, the profile it reads is the one from BEFORE the
        save. Theme Manager reported success; a newly launched tool showed the
        new theme because it was a new process; the console alone kept the old
        one, which is exactly how it was reported. Same shape as the
        Themes.json staleness fixed on 2026-09-10, and fixed the same way:
        compare the file's write time against the one the cache was filled
        from.

        Stat-ing the file per call is affordable BECAUSE nothing in a paint
        path comes through here -- the console resolves the layout at startup
        and on save, and Get-ConsoleTheme caches its own answer. Do not add a
        caller inside a Paint handler without revisiting this.

        NEVER THROWS and never returns $null: an unreadable or half-written
        profile reads as 'no customisation', because a layout file is not
        worth refusing to start the console over.
    #>
    param([switch]$Refresh)

    $path = Get-UserProfilePath
    $stamp = $null
    try { if (Test-Path -LiteralPath $path) { $stamp = (Get-Item -LiteralPath $path).LastWriteTimeUtc.Ticks } } catch { }
    # A file that has appeared, changed or been deleted since the cache was
    # filled invalidates it. Still-absent compares $null to $null, so a folder
    # with no profile in it does not churn the cache on every call.
    if ($stamp -ne $global:UserProfileStamp) { $Refresh = $true }

    if ($global:UserProfile -and -not $Refresh) { return $global:UserProfile }
    $global:UserProfileStamp = $stamp

    $empty = [PSCustomObject]@{ User = (Get-UserProfileKey); Default = $null; Tenants = [PSCustomObject]@{}; Loaded = $false }
    if (-not (Test-Path -LiteralPath $path)) { $global:UserProfile = $empty; return $global:UserProfile }
    $j = $null
    try { $j = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { $j = $null }
    if (-not $j) { $global:UserProfile = $empty; return $global:UserProfile }
    Add-Member -InputObject $j -NotePropertyName 'Loaded' -NotePropertyValue $true -Force
    if (-not $j.PSObject.Properties['Tenants'] -or -not $j.Tenants) {
        Add-Member -InputObject $j -NotePropertyName 'Tenants' -NotePropertyValue ([PSCustomObject]@{}) -Force
    }
    $global:UserProfile = $j
    return $global:UserProfile
}

function Get-UserToolGroupMap {
    <#
        Which group THIS operator has filed each tool under.

        GROUPING IS PERSONAL, and Tools.json is shared, so it cannot live
        there. Two engineers on one folder group tools the way each of them
        works -- one files everything calendar-shaped together, the other
        keeps rooms apart from calendars -- and neither wants the other's
        arrangement. Tools.json says what a tool IS; a group is an opinion
        about where to keep it.

        Returned as a hashtable keyed the same way Placement is: the script
        path relative to the toolkit root, via Get-ParkedKey. Project tools
        reuse file names, so the bare leaf cannot tell them apart.

        AN EMPTY STRING IS A REAL ANSWER. It means 'I took this one out of its
        group', which has to be distinguishable from 'I never said anything
        about this tool' -- the first overrides whatever the shared file says,
        the second leaves it alone.
    #>
    $out = @{}
    $g = $null
    try { $g = Get-UserLayoutProperty -Name 'Groups' } catch { }
    if (-not $g) { return $out }
    try {
        foreach ($p in @($g.PSObject.Properties)) {
            if (-not "$($p.Name)".Trim()) { continue }
            $out["$($p.Name)"] = "$($p.Value)".Trim()
        }
    } catch { }
    return $out
}

function Get-UserLayoutOwn {
    <#
        The block Save-UserLayout would OVERWRITE, and nothing else.

        Get-UserLayout answers 'what is in force' -- Default with the tenant
        laid over it -- which is the right answer for rendering and the wrong
        one for a read-modify-write. A window that reads the merged view,
        changes one key and saves it writes a full copy of Default into the
        tenant block, and from then on Default is shadowed for that tenant:
        'apply to every tenant' silently stops reaching it. Saving starts
        from this instead, so only keys that were already in the destination
        come back to it.

        Returns $null when the destination block does not exist yet, which is
        the same 'nothing customised' the callers already handle.

        A COPY, NOT THE LIVE BLOCK. Get-UserProfile caches the parsed file in a
        global, so returning $prof.Default handed callers a reference INTO that
        cache -- and every caller of this function exists to Add-Member onto
        what it gets back. Two things went wrong with that:

          * Save Arrangement edited the cache, then read the cache back to
            recover the placements of parked tools, and found the edit it had
            just made. The back-fill compared an object with itself and carried
            nothing, so from the second save onward a parked tool lost its
            section and order.
          * A save that THREW left the cache holding an arrangement that never
            reached disk. The operator was told "Layout NOT saved"; the next
            thing in that process to write the profile -- closing a view, say --
            wrote it for them.

        Shallow is enough: callers replace whole keys, they do not reach inside
        one and edit it.

        -AllTenants MUST MATCH THE SAVE. Save-UserLayout -AllTenants writes the
        Default block; without the switch this read returned the TENANT's, so a
        caller passing the switch to one and not the other read one block and
        overwrote a different one. With a tenant block present that copies it
        over Default and shadows Default for every tenant; with no tenant block
        -- the normal state for somebody who saved 'for every tenant' -- the
        read returns $null and Default is replaced by whatever single key the
        caller was setting, taking their theme, groups, sections and placement
        with it and leaving no .bak. Pass the same switch to both, always.
    #>
    param([switch]$AllTenants)
    $prof = Get-UserProfile
    if (-not $prof -or -not $prof.Loaded) { return $null }
    $key = Get-ActiveTenantKey
    if ($AllTenants) { $key = '' }

    $block = $null
    if ($key -and $prof.PSObject.Properties['Tenants'] -and $prof.Tenants -and $prof.Tenants.PSObject.Properties["$key"]) {
        $block = $prof.Tenants."$key"
    } elseif (-not $key -and $prof.PSObject.Properties['Default']) {
        $block = $prof.Default
    }
    if (-not $block) { return $null }

    $copy = [PSCustomObject]@{}
    foreach ($p in @($block.PSObject.Properties)) {
        Add-Member -InputObject $copy -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force
    }
    return $copy
}

function Remove-UserLayoutKeyFromTenants {
    <#
        Take one key out of every TENANT block, leaving Default alone.

        WHY THIS HAS TO EXIST. The layout is read Default-first with the active
        tenant laid over it, key by key -- so a key present in a tenant block
        WINS, whatever Default says. Somebody who sets a preference "in this
        tenant only" and later sets it "in every tenant" writes Default while
        the tenant copy is still there, and the every-tenant choice appears to
        do nothing in the very tenant they are sitting in. There is no window
        that would let them clear the tenant copy, so the setting would be
        stuck until somebody edited the JSON.

        So: 'for everyone' means everyone. The narrower copies go.

        Returns the number of blocks changed. Never throws -- a profile that
        cannot be read has no tenant copies to remove, which is the same
        outcome.
    #>
    param([Parameter(Mandatory)][string]$Name)
    $path = Get-UserProfilePath
    if (-not (Test-Path -LiteralPath $path)) { return 0 }
    $prof = $null
    try { $prof = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { return 0 }
    if (-not $prof -or -not $prof.PSObject.Properties['Tenants'] -or -not $prof.Tenants) { return 0 }

    $n = 0
    try {
        foreach ($t in @($prof.Tenants.PSObject.Properties)) {
            if (-not $t.Value) { continue }
            if (-not $t.Value.PSObject.Properties[$Name]) { continue }
            $t.Value.PSObject.Properties.Remove($Name)
            $n++
        }
    } catch { return 0 }
    if ($n -eq 0) { return 0 }

    try {
        $null = Write-JsonFileAtomic -Path $path -Json ($prof | ConvertTo-Json -Depth 10)
        $global:UserProfile = $null
        $global:UserProfileStamp = $null
    } catch { return 0 }
    return $n
}

function Get-UserLayout {
    <#
        The layout in force RIGHT NOW: the Default block, with the active
        tenant's block laid over it key by key.

        KEY BY KEY, NOT MERGED WITHIN A KEY. If the tenant block names
        Sections, its Sections are used whole; it does not get merged
        section-by-section with the Default's. Half-merged layouts are
        impossible to predict and worse to explain.

        Returns $null when this person has customised nothing, which every
        caller treats as 'use the shared config'.
    #>
    $prof = Get-UserProfile
    if (-not $prof -or -not $prof.Loaded) { return $null }

    $out = $null
    if ($prof.PSObject.Properties['Default'] -and $prof.Default) {
        $out = [PSCustomObject]@{}
        foreach ($p in @($prof.Default.PSObject.Properties)) {
            Add-Member -InputObject $out -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force
        }
    }
    $key = Get-ActiveTenantKey
    if ($key -and $prof.Tenants -and $prof.Tenants.PSObject.Properties["$key"]) {
        $t = $prof.Tenants."$key"
        if ($t) {
            if (-not $out) { $out = [PSCustomObject]@{} }
            foreach ($p in @($t.PSObject.Properties)) {
                Add-Member -InputObject $out -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force
            }
        }
    }
    return $out
}

function Get-UserLayoutProperty {
    <# One key out of the effective layout, or $null. Saves every caller the same three lines. #>
    param([Parameter(Mandatory)][string]$Name)
    $l = Get-UserLayout
    if (-not $l -or -not $l.PSObject.Properties[$Name]) { return $null }
    return $l.$Name
}

function Save-UserLayout {
    <#
        Write this person's layout for the active tenant -- or, with
        -AllTenants, into the Default block that every tenant falls back to.

        REPLACES THAT BLOCK, leaves the rest of the file alone. The other
        tenants' blocks, and anything a newer toolkit put in the file that
        this one does not understand, survive untouched.
    #>
    param(
        [Parameter(Mandatory)]$Layout,
        [switch]$AllTenants
    )
    $path = Get-UserProfilePath
    $dir  = Split-Path -Parent $path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }

    $prof = $null
    if (Test-Path -LiteralPath $path) {
        try { $prof = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { $prof = $null }
    }
    if (-not $prof) {
        $prof = [PSCustomObject][ordered]@{
            _comment = @(
                'How THIS Windows account wants the console arranged. The toolkit',
                'itself -- which tools exist -- is Tools.json and is shared; this',
                'file only says how they are laid out, grouped, hidden and coloured.',
                '',
                'Default applies everywhere; a block under Tenants applies to that',
                'tenant only and wins where it says anything.',
                '',
                'Delete this file to go back to the shared layout.'
            )
            User    = (Get-UserProfileKey)
            Default = $null
            Tenants = [PSCustomObject]@{}
        }
    }
    if (-not $prof.PSObject.Properties['Tenants'] -or -not $prof.Tenants) {
        Add-Member -InputObject $prof -NotePropertyName 'Tenants' -NotePropertyValue ([PSCustomObject]@{}) -Force
    }

    $key = Get-ActiveTenantKey
    if ($AllTenants -or -not $key) {
        Add-Member -InputObject $prof -NotePropertyName 'Default' -NotePropertyValue $Layout -Force
    } else {
        Add-Member -InputObject $prof.Tenants -NotePropertyName $key -NotePropertyValue $Layout -Force
    }
    Add-Member -InputObject $prof -NotePropertyName 'Updated' -NotePropertyValue (Get-Date -Format 'yyyy-MM-dd HH:mm') -Force

    $null = Write-JsonFileAtomic -Path $path -Json ($prof | ConvertTo-Json -Depth 10)
    $global:UserProfile = $null
    $global:UserProfileStamp = $null
    return $path
}

function Remove-UserLayout {
    <#
        Back to the shared layout: the active tenant's block, or the whole
        file with -All.

        The file is RENAMED rather than deleted when the whole thing goes, so
        an afternoon of arranging is recoverable by somebody who meant to
        reset one tenant.
    #>
    param([switch]$All)
    $path = Get-UserProfilePath
    if (-not (Test-Path -LiteralPath $path)) { return '' }
    if ($All) {
        $bak = "$path.$(Get-Date -Format 'yyyyMMdd-HHmmss').bak"
        Move-Item -LiteralPath $path -Destination $bak -Force
        $global:UserProfile = $null
        $global:UserProfileStamp = $null
        return $bak
    }
    $prof = $null
    try { $prof = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop } catch { return '' }
    # ONE TENANT MEANS ONE TENANT. This used to fall through to clearing
    # Default whenever the active tenant had no block of its own -- which is
    # the NORMAL state for somebody who saved 'for every tenant', so a button
    # reading "reset '<tenant>' only" reset every tenant they had, with no
    # backup. There is nothing tenant-specific to remove in that case, so the
    # honest answer is to remove nothing and say so; -All is the switch that
    # clears everything, and it keeps a .bak.
    $key = Get-ActiveTenantKey
    if (-not $key) { return '' }
    if (-not ($prof.PSObject.Properties['Tenants'] -and $prof.Tenants -and $prof.Tenants.PSObject.Properties["$key"])) {
        return ''
    }
    $prof.Tenants.PSObject.Properties.Remove("$key")
    $null = Write-JsonFileAtomic -Path $path -Json ($prof | ConvertTo-Json -Depth 10)
    $global:UserProfile = $null
    $global:UserProfileStamp = $null
    return $path
}

#region The view catalogue -----------------------------------------------------
#
# WHAT VIEWS THIS INSTALL HAS, read from Views\ and nothing else.
#
# The console's alternative screens -- the tool list, the palette, the status
# board and the rest -- live one to a file in Views\, each declaring itself in
# five comment lines at the top. There is no registry to keep in step: the
# folder IS the list.
#
# THIS LIVES IN Common.ps1 BECAUSE TWO PROCESSES ASK. The console reads it to
# build the View menu, and View Manager -- a separate powershell.exe, like
# every Dashboard tool -- reads it to let somebody choose a default. A second
# copy of this parser in View Manager would be the same drift the toolkit
# already fixed once.
#
# READING IS NOT LOADING. These functions return METADATA. They never
# dot-source a view file, because a view's function body expects the console's
# scope -- $form, the tool model, the view shell -- and none of that exists in
# a tool's process. The console does the dot-sourcing itself, into its own
# scope, and that is the only place it is correct to do it.

function Get-ConsoleViewFolder {
    return (Join-Path $PSScriptRoot 'Views')
}

function Read-ConsoleViewManifest {
    <#
        The five VIEW- lines at the top of a view file, or $null.

        Reads the head of the file only, and STOPS at the first 'function'
        line: the tags are declarations, and a tag quoted inside a comment
        further down -- in the 'how to add a view' note every one of these
        files carries, for instance -- must not be mistaken for one.

        A file missing VIEW-ID or VIEW-FUNCTION is not a view. It is rejected
        rather than guessed at, because a guess here puts a broken entry on
        the operator's menu.
    #>
    param([Parameter(Mandatory)][string]$Path)
    $m = [ordered]@{ Id = ''; Name = ''; Note = ''; Function = ''; Order = 500 }
    try {
        foreach ($line in @(Get-Content -LiteralPath $Path -TotalCount 40 -ErrorAction Stop)) {
            if ($line -match '^\s*function\s') { break }
            if ($line -match '^\s*#\s*VIEW-ID:\s*(\S.*?)\s*$')       { $m.Id       = $Matches[1]; continue }
            if ($line -match '^\s*#\s*VIEW-NAME:\s*(\S.*?)\s*$')     { $m.Name     = $Matches[1]; continue }
            if ($line -match '^\s*#\s*VIEW-NOTE:\s*(\S.*?)\s*$')     { $m.Note     = $Matches[1]; continue }
            if ($line -match '^\s*#\s*VIEW-FUNCTION:\s*(\S.*?)\s*$') { $m.Function = $Matches[1]; continue }
            if ($line -match '^\s*#\s*VIEW-ORDER:\s*(\d+)\s*$')      { $m.Order    = [int]$Matches[1]; continue }
        }
    } catch { return $null }
    if (-not $m.Id -or -not $m.Function) { return $null }
    if (-not $m.Name) { $m.Name = $m.Id }
    return [PSCustomObject]$m
}

function Get-ConsoleViewCatalogue {
    <#
        Every View_*.ps1 in Views\, in menu order, as metadata.

        Each entry carries Id, Name, Note, Function, Order, Path, File and
        Error. A FILE WITH A PROBLEM IS RETURNED WITH THE PROBLEM IN Error
        rather than dropped -- a view that silently is not there is the
        failure mode this whole arrangement has to avoid, and both callers
        want to say what went wrong.

        Duplicate VIEW-IDs: the first file in name order wins and the second
        is returned as an error. Two views claiming one id would mean a
        default view that means two different things.

        NEVER THROWS. A missing folder is an empty catalogue, which the
        callers report in their own words.

        A FOLDER THAT CANNOT BE LISTED IS NOT AN EMPTY FOLDER, and the
        difference is the whole point. An ACL, or an offline share that still
        answers Test-Path from a cached stat, used to come back through
        -ErrorAction SilentlyContinue as zero files -- so the console said
        "Views\ is empty" about a folder with seven files in it, and the
        operator went looking for a packaging mistake that had not happened.
        That case now comes back as ONE entry with an empty File and the
        reason in Error; both callers test File to tell 'the folder' apart
        from 'a file in it'.
    #>
    $out = @()
    $dir = Get-ConsoleViewFolder
    if (-not (Test-Path -LiteralPath $dir)) { return $out }
    $files = @()
    try {
        $files = @(Get-ChildItem -LiteralPath $dir -Filter 'View_*.ps1' -File -ErrorAction Stop | Sort-Object Name)
    } catch {
        return @([PSCustomObject]@{
            Id = ''; Name = '(the Views folder)'; Note = ''; Function = ''; Order = 0
            Path = $dir; File = ''
            Error = "This folder could not be listed: $($_.Exception.Message)"
        })
    }
    # A .ps1 IN Views\ THAT IS NOT NAMED View_* IS REPORTED, NOT IGNORED.
    #
    # The filter above is the whole naming rule, and a rule enforced by silence
    # is a rule nobody can see they have broken. Somebody saving their own
    # screen as Bobsview.ps1 gets a file in the right folder, no view on the
    # menu, and nothing anywhere saying why -- so they conclude views are
    # broken rather than that the name is wrong. Listed with the fix in the
    # message instead. View Manager's Copy button never produces one of these;
    # this is for files put here by hand.
    try {
        foreach ($x in @(Get-ChildItem -LiteralPath $dir -Filter '*.ps1' -File -ErrorAction Stop)) {
            if ($x.Name -like 'View_*') { continue }
            $want = 'View_' + ($x.BaseName -replace '[^A-Za-z0-9]', '') + '.ps1'
            $out += [PSCustomObject]@{
                Id = ''; Name = $x.Name; Note = ''; Function = ''; Order = 9998
                Path = $x.FullName; File = $x.Name
                Error = "Not picked up -- a view file must be named View_<name>.ps1. Rename it to $want."
            }
        }
    } catch { }

    $seen = @{}
    foreach ($f in $files) {
        $man = Read-ConsoleViewManifest -Path $f.FullName
        if (-not $man) {
            $out += [PSCustomObject]@{
                Id = ''; Name = $f.Name; Note = ''; Function = ''; Order = 9999
                Path = $f.FullName; File = $f.Name
                Error = 'No VIEW-ID / VIEW-FUNCTION header -- this file is not a view.'
            }
            continue
        }
        $err = ''
        if ($seen.ContainsKey($man.Id)) {
            $err = "Another file already claims the id '$($man.Id)': $($seen[$man.Id])."
        } else {
            $seen[$man.Id] = $f.Name
        }
        $out += [PSCustomObject]@{
            Id = $man.Id; Name = $man.Name; Note = $man.Note; Function = $man.Function
            Order = [int]$man.Order; Path = $f.FullName; File = $f.Name; Error = $err
        }
    }
    return @($out | Sort-Object @{ Expression = { [int]$_.Order } }, @{ Expression = { "$($_.Name)" } })
}
#endregion

#region User profile -- the view this person opens into ------------------------
#
# WHICH VIEW, AND EACH VIEW'S OWN SETTINGS, ARE PERSONAL. Two engineers share
# one toolkit folder. One wants the button grid; the other finds it small and
# wants the large-text list, in a bigger font, four columns wide. Neither
# answer is right for the folder -- they are facts about a person, so they
# live beside that person's theme and groups in Config\Users\<account>.json
# and not in any shared file.
#
#   "Default": {
#       "Theme":       "Slate",
#       "DefaultView": "List",
#       "Views": { "List": { "Font": 2, "Cols": 4 } }
#   }
#
# A THIRD STORE WAS THE OBVIOUS WRONG ANSWER. ConsoleState.<operator>.json
# already exists and already holds per-operator things (pinned, recent, window
# position), and adding the view there would have meant two personal files
# keyed on two different identities -- Get-LogOperator for one and the Windows
# account for the other -- which drift apart the first time somebody runs the
# console as a different admin. The profile is where 'how this person wants
# the console' lives; this belongs with it.

function Get-UserDefaultView {
    <#
        The VIEW-ID this operator opens into, or '' for the button grid.

        Never $null and never throws: '' means 'the console itself', which is
        what everyone gets until they choose otherwise.

        NOT VALIDATED HERE. Common.ps1 does not know which views exist -- that
        is the console, which scans Views\ at startup. An id naming a view file
        that has since been removed comes back as written; the console falls
        back to the button grid and says why. Validating here would mean this
        file needing a list of views, which is the thing the folder scan
        exists to avoid.
    #>
    $v = ''
    try { $v = "$(Get-UserLayoutProperty -Name 'DefaultView')".Trim() } catch { $v = '' }
    return $v
}

function Set-UserDefaultView {
    <#
        Remember the view this operator opens into. '' or 'Buttons' clears it.

        CARRIES THE REST OF THE BLOCK. Save-UserLayout replaces the block it
        writes whole, so the arrangement, theme, groups and parked list have
        to come along or choosing a view would quietly reset them -- the same
        trap Theme Manager hit.
    #>
    param(
        [string]$Id,
        # Write to the Default block rather than the active tenant's, so the
        # choice follows this person into every tenant.
        [switch]$AllTenants
    )
    $layout = Get-UserLayoutOwn -AllTenants:$AllTenants
    if (-not $layout) { $layout = [PSCustomObject]@{} }
    $val = "$Id".Trim()
    if ($val -eq 'Buttons') { $val = '' }
    Add-Member -InputObject $layout -NotePropertyName 'DefaultView' -NotePropertyValue $val -Force
    $out = Save-UserLayout -Layout $layout -AllTenants:$AllTenants
    # 'EVERY TENANT' MEANS EVERY TENANT. A tenant block wins over Default key
    # by key, so leaving an older per-tenant choice in place would make this
    # one appear to do nothing in the tenant the operator is looking at -- with
    # no window anywhere that could clear it.
    if ($AllTenants) { $null = Remove-UserLayoutKeyFromTenants -Name 'DefaultView' }
    return $out
}

function Get-UserViewSetting {
    <#
        One remembered setting for one view -- font step, column count,
        whatever a view decides to keep.

        KEYED ON THE VIEW-ID, so two views can both remember a 'Cols' without
        agreeing about what it means. An absent view, an absent key, or an
        unreadable profile all return $Default, because 'this person has never
        opened that view' is the normal case and not an error.
    #>
    param(
        [Parameter(Mandatory)][string]$ViewId,
        [Parameter(Mandatory)][string]$Name,
        $Default = $null
    )
    try {
        $all = Get-UserLayoutProperty -Name 'Views'
        if (-not $all) { return $Default }
        if (-not $all.PSObject.Properties["$ViewId"]) { return $Default }
        $v = $all."$ViewId"
        if (-not $v -or -not $v.PSObject.Properties["$Name"]) { return $Default }
        return $v."$Name"
    } catch { return $Default }
}

function Set-UserViewSetting {
    <#
        Remember one setting for one view.

        ONE WRITE PER CALL, so callers must not call this from a click handler
        that fires per keystroke. The views save on close, which is once.

        Other views' settings and every other key in the block survive: this
        reads the block it is about to replace, edits the one leaf, and writes
        it back.
    #>
    param(
        [Parameter(Mandatory)][string]$ViewId,
        [Parameter(Mandatory)][string]$Name,
        $Value,
        [switch]$AllTenants
    )
    $layout = Get-UserLayoutOwn -AllTenants:$AllTenants
    if (-not $layout) { $layout = [PSCustomObject]@{} }

    $all = $null
    if ($layout.PSObject.Properties['Views']) { $all = $layout.Views }
    if (-not $all) { $all = [PSCustomObject]@{} }

    $v = $null
    if ($all.PSObject.Properties["$ViewId"]) { $v = $all."$ViewId" }
    if (-not $v) { $v = [PSCustomObject]@{} }

    Add-Member -InputObject $v   -NotePropertyName "$Name"   -NotePropertyValue $Value -Force
    Add-Member -InputObject $all -NotePropertyName "$ViewId" -NotePropertyValue $v     -Force
    Add-Member -InputObject $layout -NotePropertyName 'Views' -NotePropertyValue $all  -Force
    $out = Save-UserLayout -Layout $layout -AllTenants:$AllTenants
    # Same reason as Set-UserDefaultView: a tenant's Views block would shadow
    # this one WHOLE -- the merge is key by key, not within a key -- so one
    # stale per-tenant copy would silently revert every view's saved size.
    if ($AllTenants) { $null = Remove-UserLayoutKeyFromTenants -Name 'Views' }
    return $out
}
#endregion

function Get-UserParked {
    <#
        What this PERSON has hidden, as { Scripts = @{}; Sections = @() }.

        Layered ON TOP of the tenant's parked list, never instead of it: the
        tenant list says this tenant does not have those tools at all, which
        is a fact about the tenant and not a preference. A person can hide
        more; nobody can un-hide what the tenant parked.
    #>
    $idx = @{}
    $secs = @()
    $l = Get-UserLayout
    if ($l -and $l.PSObject.Properties['Parked'] -and $l.Parked) {
        foreach ($e in @($l.Parked.Scripts)) {
            if ($null -eq $e) { continue }
            $sv = $(if ($e.PSObject.Properties['Script']) { "$($e.Script)" } else { "$e" })
            $k = Get-ParkedKey -Script $sv
            if ($k) { $idx[$k] = $e }
        }
        $secs = @(@($l.Parked.Sections) | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
    }
    return [PSCustomObject]@{ Scripts = $idx; Sections = @($secs) }
}

function Get-EffectiveParked {
    <#
        Tenant parking plus this person's, as one answer for the console to
        filter on. Both lists are keyed the same way (Get-ParkedKey), so the
        union is a plain merge.
    #>
    $idx = @{}
    foreach ($k in (Get-ParkedScriptIndex).Keys) { $idx[$k] = $true }
    $secs = @(Get-ParkedSection)
    $u = Get-UserParked
    foreach ($k in $u.Scripts.Keys) { $idx[$k] = $true }
    foreach ($sname in @($u.Sections)) { $secs += $sname }
    return [PSCustomObject]@{
        Scripts  = $idx
        Sections = @($secs | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    }
}
#endregion

function Get-TenantConfigPath {
    <#
        The tenant file to read.

        Config\Tenants\<active>.json in multi-tenant mode, Config\Tenant.json
        otherwise -- and 'otherwise' includes multi-tenant mode with nothing
        selected, so a half-configured install still opens something rather than
        nothing.
    #>
    $k = Get-ActiveTenantKey
    if ($k) {
        $p = Join-Path (Get-TenantsFolder) "$k.json"
        if (Test-Path -LiteralPath $p) { return $p }
    }
    return (Join-Path (Join-Path $PSScriptRoot $script:ConfigFolderName) $script:TenantConfigFileName)
}

function Get-TenantConfig {
    <#
        The tenant file, cached in a GLOBAL -- so for the life of the
        POWERSHELL WINDOW, not the life of the console.

        AND THEREFORE RE-READ WHEN THE FILE CHANGES, for the same reason
        Get-UserProfile is. Three DETACHED tools write this file -- Tenant
        Manager, Theme Manager saving 'for everyone', and Park In Place
        parking for the tenant -- and each of them clears this cache in ITS
        OWN process, which does nothing for the console that launched it. Park
        a tool for the tenant, close the console, re-run it in the same window,
        and the parked list it reads is the one from before the save; the
        button is still there and the save looked fine. Same trap as the
        per-user theme, one file over.

        The stat is cheap and this is not called from a paint path; the
        capability lookups that ARE hot go through Test-TenantCapability,
        which reads this once and is itself called from state passes, not
        from rendering.

        NEVER THROWS. A toolkit that will not start because a config file is
        missing is worse than one that starts and says every capability is off:
        the first cannot tell you what is wrong, the second can.
    #>
    param([switch]$Refresh)

    $cp = Get-TenantConfigPath
    $cstamp = $null
    try { if ($cp -and (Test-Path -LiteralPath $cp)) { $cstamp = (Get-Item -LiteralPath $cp).LastWriteTimeUtc.Ticks } } catch { }
    # Also keyed on the PATH, so switching tenant re-reads even if two tenant
    # files happen to share a write time.
    $ckey = "$cp|$cstamp"
    if ($ckey -ne "$($global:TenantConfigStamp)") { $Refresh = $true }

    if ($global:TenantConfig -and -not $Refresh) { return $global:TenantConfig }
    $global:TenantConfigStamp = $ckey

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
          [ValidateSet('ExchangeOnline','SharePointOnline','OnPremAD','OnPremExchange','Pim','PimTimer','MicrosoftGraph')]
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

        These used to be destined for a script: 'contoso.com/Accounts/
        Resource Mailbox' branched on an "if <directory>". That is one AD
        layout written into a tool, and the same mistake as the ValidateSet on
        directory names -- Test-Toolkit flags a domain in a .ps1 for exactly
        this reason.

        Config-driven instead, and the behaviour falls out rather than being
        branched. A directory that HAS the OU configured uses it; one that does
        not gets asked. A directory prompts because nobody filled its OU in,
        not because of what it is called -- so a tenant with
        five directories, or one, behaves sensibly without a code change.

        Returns '' when unset. The caller must prompt rather than invent a path.
    #>
    param(
        [string]$DirectoryId,
        [Parameter(Mandatory)][ValidateSet('SharedMailbox', 'DistributionGroup', 'MailContact')]
        [string]$Kind
    )

    $dir = $null
    if ($DirectoryId) { $dir = Get-TenantDirectory -Id $DirectoryId }
    if (-not $dir)    { $dir = Get-DefaultTenantDirectory }
    if (-not $dir)    { return '' }

    # 'MailContact' reads MailContactOu, added 2026-08-27 for MailContacts.ps1.
    # Falls back to nothing, not to the shared-mailbox OU: a contact in the
    # resource-mailbox OU is the kind of tidy-looking mistake that stays for
    # years. Add MailContactOu to the directory in Tenant.json.
    $key = switch ($Kind) {
        'SharedMailbox'     { 'SharedMailboxOu' }
        'DistributionGroup' { 'DistributionGroupOu' }
        'MailContact'       { 'MailContactOu' }
    }
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
        enable in the affected directory failed on it.

        Validation, not selection. A candidate is returned only when
        Tenant.json actually names it as a controller. Anything else -- a
        domain, a typo, a plausible-looking hostname -- comes back as '', and
        the caller goes forest-wide.

        Refusing to guess is the whole point. Bulk-DLMembership records the
        opposite failure on 2026-08-06: pinning a controller from one directory
        made a write in another fail as "out of the current user's write
        scope", which
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
        operation in one directory would leave a pin behind and the next
        operation inherited it. That is the 2026-08-12 failure, where an
        account was verified against another directory's controller.

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
        # which for an operator in the local domain is right by accident and
        # for anyone else is wrong without saying so.
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

function Get-TenantExchangeServerList {
    <#
        The on-premises Exchange servers these tools may touch, from
        Tenant.json -- NEVER from a list typed into a .ps1.

        WHY THIS EXISTS. Check Queues, Clear IIS Logs and Exchange Health
        Check each carried the same four hostnames as literals in their own
        code. That breaks the rule this toolkit is built on -- Tenant.json and
        Settings.json are the only files allowed to name anything belonging to
        an organisation -- and it means a toolkit handed to another tenant
        reaches for servers that do not exist there, with an admin-share
        timeout as the only clue.

        RESOLUTION ORDER, first answer wins:

          1. OnPremExchange.Servers in Tenant.json -- an explicit list, which
             is the right answer when only some servers are in scope.
          2. Get-ExchangeServer, when an on-premises session is open. The
             directory knows its own servers; asking is better than being
             told, and a server added last week is included without an edit.
          3. Nothing, and the caller says where to set it.

        ExcludedServers is subtracted from whichever list is produced, so the
        one key already in Tenant.json keeps working the same way for both
        routes. Comparison is case-insensitive and ignores any domain suffix:
        'EXSRV01' and 'exsrv01.corp.example.com' are one server.
    #>
    param([switch]$Refresh)

    $short = {
        param([string]$n)
        $t = "$n".Trim()
        if (-not $t) { return '' }
        return ($t -split '\.')[0].ToLowerInvariant()
    }

    $ex = @{}
    $cfg = $null
    try { $cfg = (Get-TenantConfig).OnPremExchange } catch { }
    if ($cfg -and $cfg.PSObject.Properties['ExcludedServers']) {
        foreach ($e in @($cfg.ExcludedServers)) {
            $k = & $short $e
            if ($k) { $ex[$k] = $true }
        }
    }

    $list = @()
    if ($cfg -and $cfg.PSObject.Properties['Servers']) {
        # A NAME OR AN OBJECT, and this is the one reader that has to accept
        # both. Since 2026-09-17 an entry may be { Name, Site, Role } so that
        # ServerReboot.ps1 can honour the site rule without guessing from the
        # name. Every other tool only wants the NAME -- and "$_" on an object
        # stringifies to '@{Name=X; Site=Y...}', which is a server that does
        # not exist. Read .Name when there is one, the string when there is not.
        $list = @(@($cfg.Servers) | ForEach-Object {
            if ($_ -is [string]) { "$_".Trim() }
            elseif ($_ -and $_.PSObject.Properties['Name']) { "$($_.Name)".Trim() }
            else { "$_".Trim() }
        } | Where-Object { $_ })
    }

    if ($list.Count -eq 0) {
        # Ask the directory. Guarded twice: the cmdlet may not exist, and it
        # throws rather than returning nothing when there is no session.
        try {
            if ((Test-EmsAvailable).Available) {
                $list = @(Invoke-OnPremCommand -Name 'Get-ExchangeServer' -Parameters @{
                            ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' } |
                          Where-Object { "$($_.ServerRole)" -match '(?i)Mailbox|HubTransport|ClientAccess' } |
                          ForEach-Object { "$($_.Name)".Trim() } |
                          Where-Object { $_ })
            }
        } catch { $list = @() }
    }

    $out = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($s in $list) {
        $k = & $short $s
        if (-not $k -or $ex[$k] -or $seen[$k]) { continue }
        $seen[$k] = $true
        [void]$out.Add($s)
    }
    return @($out | Sort-Object)
}

function Get-TenantHealthServerList {
    <#
        The servers Server Health checks, from Tenant.json ServerHealth.Servers.

        A SEPARATE LIST FROM OnPremExchange.Servers, AND THAT IS THE POINT.
        That list means "servers these tools may run Exchange cmdlets against",
        and three tools already read it with that meaning. Server Health asks a
        different question -- disk, patches, uptime -- of a WIDER set: domain
        controllers, file servers, a jump box, things with no Exchange on them
        at all. Overloading one list would either drag non-Exchange servers into
        Exchange tooling or keep them out of the health check.

        SHAPE, and every field but Name is optional:

            "ServerHealth": {
              "Servers": [
                { "Name": "SRV01", "Exchange": true,  "Notes": "hub" },
                { "Name": "SRV02", "Exchange": false, "Notes": "file" }
              ]
            }

        A BARE STRING IS ALSO ACCEPTED -- "Servers": [ "SRV01", "SRV02" ] -- and
        reads as Exchange=false. Absent means off: no ServerHealth block, no
        servers, and the tool says where to set them rather than reporting a
        clean estate of nothing.

        Exchange is a HINT, not a gate. The collector looks for the Exchange
        install on every server it reaches and reports what it finds; this flag
        only decides whether ABSENCE is worth flagging. A server marked
        Exchange=true with no Exchange on it is a finding. Marked false, it is
        simply a server.

        Returns objects with Name / Exchange / Notes, de-duplicated on the short
        name the same way Get-TenantExchangeServerList does, so 'SRV01' and
        'srv01.corp.example.com' cannot both appear.
    #>
    # NO -Refresh PARAMETER: it always refreshes. A switch the caller has to
    # remember, on the one call where forgetting it silently returns the wrong
    # answer, is not a choice worth offering.

    # READ THE FILE AGAIN, NOT THE CACHE.
    #
    # Get-TenantConfig caches in a GLOBAL for the life of the process, and the
    # console launches tools with '&' INTO THAT PROCESS. So a console that
    # started before the servers were added holds a Tenant.json with no
    # ServerHealth block, and every tool it launches sees that stale copy --
    # the tool reports "no servers configured" while the file on disk plainly
    # has them. Reported 2026-09-13, and it is the first thing anybody hits.
    #
    # One file read at tool start is nothing next to a remote sweep, so this
    # always refreshes rather than making the caller remember to.
    $cfg = $null
    try { $cfg = (Get-TenantConfig -Refresh).ServerHealth } catch { }

    # THREE SHAPES ACCEPTED, because all three are what people actually write.
    #
    #   "ServerHealth": { "Servers": [ ... ] }   documented, what Tenant Manager writes
    #   "ServerHealth": [ ... ]                  the obvious hand-edit
    #   absent                                   fall through to the Exchange list below
    #
    # Being strict here buys nothing: a file that says what it means in a
    # slightly different shape is not a config error, it is a person typing the
    # sensible thing. Rejecting it produces "no servers configured" over a file
    # that visibly lists servers, which is the least helpful message available.
    $raw = @()
    if ($cfg) {
        if ($cfg -is [System.Array]) {
            $raw = @($cfg)
        } elseif ($cfg.PSObject.Properties['Servers']) {
            $raw = @($cfg.Servers)
        }
    }

    # NOTHING CONFIGURED? USE THE EXCHANGE SERVERS. They are servers, they are
    # already in Tenant.json, and disk and patch state matter on them at least
    # as much as anywhere else -- so the tool does something useful on a fresh
    # install instead of nothing. They come back flagged Exchange, which is
    # what they are. Add a ServerHealth block to widen the list beyond them.
    $fromExchange = $false
    if (@($raw).Count -eq 0) {
        try { $raw = @(Get-TenantExchangeServerList) } catch { $raw = @() }
        $fromExchange = (@($raw).Count -gt 0)
    }

    # Plain array -- see the note in Get-TenantRegistry on why
    # New-Object ...List[object] unquoted throws.
    $out  = @()
    $seen = @{}
    foreach ($entry in @($raw)) {
        if ($null -eq $entry) { continue }

        # A string entry and an object entry both arrive here. Test for the
        # property rather than the type: ConvertFrom-Json gives PSCustomObject
        # for the object form and String for the other, and -is [string] on a
        # PSObject-wrapped string has surprised this toolkit before.
        $name = ''; $isEx = $false; $notes = ''
        if ($entry.PSObject.Properties['Name']) {
            $name = "$($entry.Name)".Trim()
            if ($entry.PSObject.Properties['Exchange']) {
                # [bool] on the string 'false' is TRUE -- any non-empty string
                # is. JSON gives a real boolean, but a hand-edited file may not,
                # so the string forms are matched explicitly.
                #
                # $exRaw, not $raw: $raw is the collection this loop is walking,
                # and reusing the name here reassigned it on the first object
                # entry. The foreach held its own enumerator so it did not break
                # the loop -- which is exactly what makes it the kind of bug
                # that survives testing and bites later.
                $exRaw = "$($entry.Exchange)".Trim()
                $isEx = ($exRaw -match '(?i)^(true|yes|1)$')
            }
            if ($entry.PSObject.Properties['Notes']) { $notes = "$($entry.Notes)".Trim() }
        } else {
            $name = "$entry".Trim()
        }

        if (-not $name) { continue }
        $key = ($name -split '\.')[0].ToLowerInvariant()
        if (-not $key -or $seen[$key]) { continue }
        $seen[$key] = $true

        # Servers borrowed from the Exchange list ARE Exchange servers, so they
        # come back flagged as such whatever the (absent) entry said.
        if ($fromExchange) { $isEx = $true; $notes = 'from OnPremExchange.Servers' }

        $out += [pscustomobject]@{
            Name     = $name
            Exchange = $isEx
            Notes    = $notes
        }
    }
    return @($out | Sort-Object Name)
}

function Get-TenantHealthServerSource {
    <#
        WHERE THE HEALTH SERVER LIST CAME FROM, in one sentence, for the tool to
        print when the list is empty or surprising.

        "No servers configured" over a Tenant.json that visibly lists servers is
        the worst message this tool can give -- it is true of the object in
        memory and false of the file on disk, and it sends somebody to edit a
        file that is already correct. So the tool says which key it read, how
        many it found, and which file it looked in.
    #>
    $path = ''
    try { $path = Get-TenantConfigPath } catch { }
    $cfg  = $null
    try { $cfg = (Get-TenantConfig -Refresh).ServerHealth } catch { }

    if ($cfg -is [System.Array]) {
        return "ServerHealth (a bare list) in $path -- $(@($cfg).Count) entr(ies)."
    }
    if ($cfg -and $cfg.PSObject.Properties['Servers']) {
        return "ServerHealth.Servers in $path -- $(@($cfg.Servers).Count) entr(ies)."
    }
    $ex = @()
    try { $ex = @(Get-TenantExchangeServerList) } catch { }
    if (@($ex).Count -gt 0) {
        return "No ServerHealth block in $path, so the $(@($ex).Count) server(s) from OnPremExchange.Servers are being used instead."
    }
    return "No ServerHealth block and no OnPremExchange.Servers in $path."
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

        'E5 License A - Full' sat hardcoded in Disable-User and NotDeadYet
        while Tenant.json carried a LicenseGroups list -- Name, Sku, Directory
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
        which entry was for which. The reference tenant has four: E1 and E5 per directory.
        The E1 groups exist so the Disabled/Licensed report can find people who
        are disabled but still licensed -- they are a reporting input, and a
        termination must not touch them. The E5 groups are the ones a leaver
        actually holds.

        With no way to tell them apart, every reader had to guess.
        Resolve-TenantLicenseGroupName guessed "the first one" and removed
        terminated staff from 'E1 License A' -- a group they were never in --
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
        when a directory has exactly one licence group. The reference tenant has two per
        directory -- E1 and E5 -- so that function returned whichever happened
        to be first in Tenant.json and the disable pipeline removed the user
        from 'E1 License A' while an E5 user kept their licence. It then
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

        'E5 License B - Full' typed in Word comes back with the hyphen turned
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

function Get-LicenseGroupMemberUser {
    <#
        Every USER that is a member of a group, directly or through nesting,
        WITHOUT the 5,000-member ceiling.

        Get-ADGroupMember -Recursive is capped by the AD web service's
        MaxGroupOrMemberEntries (default 5,000) and fails with "The size limit
        for this request was exceeded". A licence group is precisely the kind of
        group that crosses that line, and it did: Returning User broke on it on
        2026-08-25, and this function replaces the same call in the two places
        that were still making it.

        LDAP_MATCHING_RULE_IN_CHAIN (1.2.840.113556.1.4.1941) asks the directory
        for "objects whose memberOf, followed transitively, reaches this DN".
        It is a search, not a member enumeration, so it pages like any other
        search and has no group-size limit. It also filters to users
        server-side, so the objectClass check that used to follow is gone.

        -Properties is passed straight through. A caller that needs Enabled,
        mail and LastLogonDate for every member gets them in the SAME query,
        instead of one Get-ADUser per member afterwards -- on a 6,000-member
        group that is 6,000 round trips removed.
    #>
    param(
        [Parameter(Mandatory)][string]$GroupDn,
        [Parameter(Mandatory)][string]$Server,
        [string[]]$Properties
    )

    # The DN goes into an LDAP filter, so it needs RFC 4515 escaping -- which
    # is NOT the OPath doubling the rest of the toolkit uses. Parentheses and
    # asterisks in an OU name are legal and do turn up.
    $esc = "$GroupDn" -replace '\\', '\5c' -replace '\(', '\28' -replace '\)', '\29' -replace '\*', '\2a'

    $p = @{
        LDAPFilter  = "(&(objectClass=user)(objectCategory=person)(memberOf:1.2.840.113556.1.4.1941:=$esc))"
        Server      = $Server
        ErrorAction = 'Stop'
    }
    if ($Properties -and $Properties.Count) { $p.Properties = $Properties }
    return @(Get-ADUser @p)
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
            $members = @(Get-LicenseGroupMemberUser -GroupDn $found.Group.DistinguishedName `
                            -Server $found.Server)
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

        NOT GRAPH. Licensing here is group-based, so the licence a person
        holds is a fact about AD group membership and is readable with the
        permissions the operator already has. This deliberately does not report
        the SKU as Microsoft sees it -- it reports the group that drives the
        assignment, which is the thing an operator can actually act on.

        NESTING IS CHECKED, but only when the direct answer is no.

        Get-LicensedUserMap reads members -Recursive while the disable pipeline
        tests $user.MemberOf, which is direct-only. Those two disagree for a
        user who is licensed through a nested group, and this function would
        have had to pick a side. It checks direct membership first because that
        is what the directory actually does and it costs nothing -- memberOf is
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

    # THE CACHE IS INVALIDATED BY THE FILE'S TIMESTAMP, NOT ONLY BY -Refresh.
    #
    # $global:ToolkitSettings is deliberately global so tools the console
    # launches with '&' share one answer -- but a global OUTLIVES THE SCRIPT.
    # The console re-runs itself in this same process on a theme change, and a
    # relaunch from the same PowerShell window is the normal way to restart it.
    # Either path returned whatever was cached the first time, so an edit to
    # Settings.json appeared to do nothing: the file was right, the screen was
    # wrong, and nothing said why.
    #
    # Get-ConsoleTheme carries the identical guard for the identical reason --
    # see the LastWriteTimeUtc stamp there. A cache keyed only on "have I read
    # this yet" cannot notice the file changing underneath it.
    $stamp = $null
    try {
        $sp = Get-SettingsFilePath
        if ($sp -and (Test-Path -LiteralPath $sp)) { $stamp = (Get-Item -LiteralPath $sp).LastWriteTimeUtc.Ticks }
    } catch { }
    if ($stamp -and $global:ToolkitSettingsStamp -and $stamp -ne $global:ToolkitSettingsStamp) {
        $global:ToolkitSettings = $null
    }

    if ($global:ToolkitSettings -and -not $Refresh) { return $global:ToolkitSettings }
    $global:ToolkitSettingsStamp = $stamp

    $settings = [PSCustomObject]@{
        # SpoAdminUrl and ExchangeServerUri are NOT here, and their absence is
        # the design. Both describe the TENANT, so both live in Tenant.json and
        # are edited in Tenant Manager. Removed 2026-08-13; see Get-SpoAdminUrl
        # for what having them in two files actually cost.
        ReportsPath       = ''
        LogPath           = ''
        RetryQueuePath    = ''
        OperatorEmail     = ''
        # THE CONSOLE'S NAME AND THE FACE IT IS SET IN.
        #
        # Here rather than in AdminConsole.ps1 because the toolkit's rule is
        # that a .ps1 never names a company or a tenant -- Test-Toolkit flags
        # it, and the same script ships to more than one place. 'ABCD Admin
        # Console' is a fact about one installation, so it lives in config with
        # the rest of them. Blank means 'Admin Console'.
        ConsoleTitle      = ''
        ConsoleTitleFont  = ''
        # 'Raised' draws the tool buttons with a bevel; anything else (blank)
        # leaves them flat, which is what every existing installation has and
        # therefore what an absent key has to mean.
        ConsoleButtonStyle = ''
        # THE TWO LOOK SETTINGS BELOW BOTH DEFAULT TO OFF WHEN ABSENT, and that
        # is the rule this file follows for anything that changes how the
        # toolkit DRAWS: a key nobody has added yet must leave the console
        # looking exactly as it did. Turning either on is one dropdown in Tool
        # Manager, and turning it back off is the same dropdown -- nothing is
        # rewritten in any tool to make them work.
        #
        # 'On' puts a Segoe icon glyph on buttons whose caption says what they
        # do. Blank leaves every button text-only.
        ConsoleIcons      = ''
        # 'Rounded' gives raised buttons 6px corners. Blank is square. Flat
        # buttons are square either way -- see Get-ButtonCornerRadius.
        ConsoleCorners    = ''
        # 'Themed' routes Show-ToolDialog through a form this toolkit draws.
        # Blank uses the Windows MessageBox, which is what every call site got
        # before and is still the fallback whenever the themed one cannot be
        # built.
        ConsoleDialogStyle = ''
        # 'On' shows "signed in as <name>" beside the console title. Blank
        # leaves the header to the title alone; the operator is already named
        # in the status strip at the bottom, so the header line is a choice.
        ConsoleHeaderOperator = ''
        BrowserProfiles   = $null
    }
    $path = Get-SettingsFilePath

    if (Test-Path -LiteralPath $path) {
        try {
            $j = (Get-Content -LiteralPath $path -Raw -ErrorAction Stop) | ConvertFrom-Json -ErrorAction Stop
            foreach ($k in @('ReportsPath', 'LogPath', 'RetryQueuePath', 'OperatorEmail',
                             'ConsoleTitle', 'ConsoleTitleFont', 'ConsoleButtonStyle',
                             'ConsoleIcons', 'ConsoleCorners', 'ConsoleDialogStyle',
                             'ConsoleHeaderOperator')) {
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
    <#
        EVERY KEY HAS TO BE A PARAMETER HERE, and that is not a style rule.
        This function rebuilds Settings.json from scratch rather than editing
        it, so a key it does not know about is DELETED the next time anyone
        presses Save in Tool Manager. A setting added to Get-ToolkitSettings
        and not to this function works perfectly until someone opens the
        settings dialog, and then silently reverts.
    #>
    param(
        [string]$ReportsPath = '',
        [string]$LogPath = '',
        [string]$RetryQueuePath = '',
        [string]$OperatorEmail = '',
        [string]$ConsoleTitle = '',
        [string]$ConsoleTitleFont = '',
        [string]$ConsoleButtonStyle = '',
        [string]$ConsoleIcons = '',
        [string]$ConsoleCorners = '',
        [string]$ConsoleDialogStyle = '',
        [string]$ConsoleHeaderOperator = '',
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
        ConsoleTitle      = "$ConsoleTitle".Trim()
        ConsoleTitleFont  = "$ConsoleTitleFont".Trim()
        ConsoleButtonStyle = "$ConsoleButtonStyle".Trim()
        ConsoleIcons      = "$ConsoleIcons".Trim()
        ConsoleCorners    = "$ConsoleCorners".Trim()
        ConsoleDialogStyle = "$ConsoleDialogStyle".Trim()
        ConsoleHeaderOperator = "$ConsoleHeaderOperator".Trim()
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
    param(
        [string]$Configured,
        [Parameter(Mandatory)][string]$DefaultLeaf,
        # Opt OUT of the per-tenant subfolder. Nothing uses this yet; it exists
        # so a future shared folder (one audit log for the whole estate) has a
        # way to say so rather than needing a second resolver.
        [switch]$Shared
    )

    $path = if ($Configured) {
                if ([System.IO.Path]::IsPathRooted($Configured)) { $Configured }
                else { Join-Path $PSScriptRoot $Configured }
            } else { Join-Path $PSScriptRoot $DefaultLeaf }

    # ONE TENANT PER FOLDER.
    #
    # Reports and logs from two organisations in one directory is a data
    # problem, not an untidiness problem: a CSV named for its report and its
    # timestamp says nothing about whose data is in it, and the two tenants'
    # files interleave in the listing. So in multi-tenant mode every resolved
    # folder gains a segment named for the tenant.
    #
    # A TENANT THAT SETS ITS OWN ABSOLUTE PATH STILL GETS THE SEGMENT, because
    # the risk is exactly the same -- two tenants pointed at one share is the
    # likeliest way to end up with mixed reports, and it is the case somebody is
    # most likely to configure by hand and not think about.
    #
    # Single-tenant installs are untouched: no key, no segment, same paths as
    # yesterday. That is what makes deleting Config\Tenants a complete revert.
    if (-not $Shared) {
        $tk = Get-ActiveTenantKey
        if ($tk) { $path = Join-Path $path $tk }
    }

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

function Confirm-TenantForWrite {
    <#
        Name the tenant before a tool changes anything in it.

        THE ONLY FAILURE HERE THAT MATTERS IS A WRITE AGAINST THE WRONG
        ORGANISATION. Everything else multi-tenant mode can get wrong is a
        report in an odd folder; this one disables the wrong person's account,
        and it is not noticed on the day. So the tools that WRITE say which
        tenant they are about to write to, once, and wait.

        OFF BY A SETTING, NOT BY EDITING SCRIPTS. Settings.json:

            "ConfirmTenantOnWrite": false

        turns it off everywhere. Absent or true means on. It is on by default
        because the cost of the prompt is a second and the cost of the mistake
        is not, but it is meant to be turned off once switching tenants is
        second nature -- and turning it off should be one key, not a sweep
        through five scripts undoing a hardcoded dialog.

        SILENT ON SINGLE-TENANT INSTALLS. With one tenant there is nothing to
        get wrong, and a prompt that always says the same thing is one people
        learn to click through -- which would make it useless on the day it
        finally had something to say.

        Returns $true to proceed.
    #>
    param(
        # What the tool is about to do, in a few words: 'disable glenn@...'.
        [string]$Action = 'make changes'
    )

    if (-not (Test-MultiTenantMode)) { return $true }

    $on = $true
    try {
        $s = Get-ToolkitSettings
        if ($s -and $s.PSObject.Properties['ConfirmTenantOnWrite']) {
            $on = -not ("$($s.ConfirmTenantOnWrite)".Trim() -match '(?i)^(false|no|0)$')
        }
    } catch { }
    if (-not $on) { return $true }

    $label  = Get-ActiveTenantLabel
    $cfg    = $null
    $domain = ''
    try { $cfg = Get-TenantConfig; $domain = "$($cfg.PrimaryDomain)" } catch { }

    $msg = "About to $Action in:" + [Environment]::NewLine + [Environment]::NewLine +
           "    $label" + $(if ($domain) { "   ($domain)" } else { '' }) + [Environment]::NewLine + [Environment]::NewLine +
           'Is that the right tenant?'

    $ans = Show-ToolDialog($msg, 'Confirm tenant', 'YesNo', 'Warning', 'Button2')
    return ($ans -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Test-TenantFolderCollision {
    <#
        Would two tenants write into the same folder?

        THE WARNING IS THE POINT. Resolve-ToolkitFolder appends the tenant key,
        so collisions should be impossible -- but a tenant whose Settings.json
        names an absolute path that ALREADY ENDS in another tenant's key would
        land on top of it, and so would two tenants sharing a key by different
        file names. This is cheap and runs at startup, because the alternative
        is finding out when somebody opens a report and it is the wrong
        organisation's data.

        Returns a list of warning lines; empty means nothing to say.
    #>
    $warn = New-Object System.Collections.Generic.List[string]
    if (-not (Test-MultiTenantMode)) { return @() }

    $me = Get-ActiveTenantKey
    if (-not $me) { return @() }

    foreach ($pair in @(
        @{ Label = 'Reports'; Path = (Get-ReportPath) }
        @{ Label = 'Logs';    Path = (Get-LogPath) }
    )) {
        $p = "$($pair.Path)"
        if (-not $p) { continue }
        $leaf = Split-Path -Leaf $p
        # The last segment should BE this tenant. If it is a different tenant's
        # key, this install is about to write into their folder.
        foreach ($t in Get-TenantRegistry) {
            if ($t.Key -eq $me) { continue }
            if ($leaf -eq $t.Key) {
                [void]$warn.Add("$($pair.Label) resolves to '$p', which is tenant '$($t.Key)'. Check Settings.json.")
            }
        }
    }
    return @($warn)
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

    $null = Show-ToolDialog($msg, 'Error', 'OK', 'Error')
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
    # NEVER 'the first tmpEXO_*': with several connections in the process
    # that is a coin toss between live and expired. Refresh the map (which
    # asks Get-ConnectionInformation for the live module) and try once more
    # before falling back to a name match.
    $map = Get-ExchangeModuleMap -Refresh
    if ($map.Cloud) {
        $cmd = Get-Command -Name "$($map.Cloud.Name)\$Name" -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd }
    }
    $cmd = $all | Where-Object { $_.ModuleName -like 'tmpEXO_*' -or $_.Source -like 'tmpEXO_*' } | Select-Object -Last 1
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
# Environment was carrying a job it cannot do. It holds a directory name --
# a DIRECTORY within one tenant, not the tenant itself. This console ships to
# six of them, every copy writes the same CSV format, and the moment two of
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

    # The tools already write the directory in brackets at the end of the
    # sentence -- "(EXO)", "(On-Prem)", or whatever this organisation calls its
    # directories. That is an existing convention at a known position, so
    # reading it back is recovery rather than guesswork.
    #
    # THE NAMES ARE NOT LISTED HERE. They used to be: the pattern named one
    # organisation's two directories outright, which meant the file could not be
    # shipped anywhere else and the second install carried a scrubbed copy with
    # placeholder names that matched nothing. A bracketed token at the end of
    # the sentence is the convention; what it says is the caller's business.
    # Anything with a space or punctuation beyond a hyphen is not a directory
    # label, so the match stays conservative and still returns blank rather
    # than a guess.
    $e = [regex]::Match($Task, '\(([\w][\w \-]{0,30})\)\s*$')
    if ($e.Success) { $out.Environment = $e.Groups[1].Value.Trim() }

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
        Use Get-LogTarget to say that -- it takes candidates in preference
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
        legacy parser recovers it from the bracketed-directory convention the
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

        Was: [ValidateSet(...)] over two literal names -- one organisation's
        structure written into a type constraint, in a function 35 scripts
        call. Any tenant without those two names could not use it, and a tenant
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
        $null = Show-ToolDialog(
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

function Enable-FormAutoScale {
    <#
        Make a window ZOOM. Drag it bigger and the text, the buttons, the boxes
        and the gaps between them all grow together, in proportion.

        WHY THIS RATHER THAN A FONT SLIDER. A slider changes the text and
        nothing else, so the text grows inside controls that did not, and by two
        steps up it is clipping. Making the window itself the zoom control is
        the thing people already reach for, and it needs no setting, no
        preference file and no explaining.

        WHY THIS RATHER THAN ANCHORS. Anchoring REFLOWS -- controls keep their
        size and move apart, so a bigger window is a sparser window with the
        same tiny 9pt text. That is the opposite of the ask. Anchoring solves
        "use the extra space"; this solves "I cannot read it".

        ---------------------------------------------------------------------
        THE FOUR THINGS THAT MAKE THIS WORK, EACH OF WHICH FAILED FIRST
        ---------------------------------------------------------------------

        1. EVERY VALUE IS COMPUTED FROM THE ORIGINAL, NEVER FROM THE CURRENT.

           The obvious loop -- multiply what is there by the new ratio each time
           -- compounds. Rounding error accumulates on every drag, controls
           creep out of alignment, and shrinking back does not return them to
           where they started because integer rounding is not reversible. So the
           design-time Bounds and Font of every control are captured ONCE and
           every later resize is computed from that baseline.

        2. ANCHORS ARE TURNED OFF, AND THIS IS THE HELP BUTTON BUG.

           An anchored control is moved by WinForms on resize, and then moved
           AGAIN here -- it travels twice and lands outside the window. The '?'
           anchors Top+Right, which is exactly the combination that walks off
           the right edge and disappears.

           Anchors and proportional scaling are two answers to the same
           question and only one can be in charge. Scaling wins, so every
           control is pinned to Top,Left for the duration and this function
           places it. The original anchors are captured anyway, so nothing is
           lost if this is ever reversed.

        3. DOCKED CONTROLS KEEP THEIR DOCK.

           A Dock'd control is sized by its parent, and writing Bounds on one
           either does nothing or fights the layout engine. Docked controls get
           their FONT scaled and their bounds left alone -- which is the correct
           result, because Dock already handles the space.

        4. ONE UNIFORM FACTOR, FROM WHICHEVER AXIS GREW LESS.

           Scaling width and height independently stretches text and turns
           square buttons into rectangles. A single factor keeps every
           proportion exactly as designed -- which is the stated requirement --
           and any excess on the other axis becomes margin. Dragging wider alone
           therefore does nothing until the height follows, and that is
           deliberate: it is what stops the layout distorting.

        ---------------------------------------------------------------------
        WHEN IT RUNS
        ---------------------------------------------------------------------

        The baseline is captured on Shown, NOT when this is called. By Shown
        every control exists, including any a tool adds after its help button,
        and the form has settled at its real size. Capturing earlier misses
        controls and pins the baseline to a size the window is about to leave.

        Scaling is applied on ResizeEnd -- once, when the drag finishes --
        rather than on every Resize event. Rebuilding forty Font objects per
        mouse-move is how a smooth drag turns into a stuttering one, and GDI
        objects are a finite resource. WindowState changes (maximise, restore)
        do not raise ResizeEnd at all, so those are caught separately.

        NEVER SMALLER THAN DESIGNED. The floor is 1.0. Set-FormResizable
        already pins MinimumSize to the design size for the same reason: these
        layouts are absolutely positioned and have no smaller arrangement to
        fall back to.

        OPT OUT by putting 'NoAutoScale' anywhere in the form's Tag.
    #>
    param(
        [Parameter(Mandatory)]$Form,
        # Past roughly 3x the text is enormous and the window no longer fits a
        # normal screen. The cap stops a stray double-click on a 4K display
        # producing something unusable.
        [double]$MaxScale = 3.0
    )

    <#
        THIS FUNCTION NEVER WRITES TO $Form.Tag, AND THAT IS NOT FUSSINESS.

        Add-ReportFileControls stores an ArrayList of control bundles on the
        form's Tag and Add_Activated walks it calling Refresh(). Appending a
        marker string to Tag replaces that list with a string. Nothing throws --
        @('some string') is a perfectly good one-element array and the Refresh
        lookup on it just comes back null -- so every report window quietly
        stops refreshing its file list, with no error and no obvious cause.

        The opt-out is READ from Tag because that is the existing convention
        (see Set-ConsoleTheme's 'NoTheme'), but the wired-already marker lives
        on a NoteProperty of its own where it cannot collide with anything.
    #>
    if ("$($Form.Tag)" -like '*NoAutoScale*') { return }
    if ($null -ne $Form.PSObject.Properties['ToolkitScale']) { return }   # already wired

    # A zoomable window must be resizable. Harmless if the tool already did it.
    if ($Form.FormBorderStyle -ne [System.Windows.Forms.FormBorderStyle]::Sizable) {
        $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
        $Form.MaximizeBox     = $true
    }

    <#
        STATE LIVES ON THE FORM, AND THE HANDLERS REACH IT THROUGH $this.

        This is not a style choice. A scriptblock handed to Add_Shown runs long
        after this function has returned, and by then this function's LOCAL
        VARIABLES ARE GONE -- the scope that held them was popped. A handler
        written against a local $state finds $null and silently does nothing,
        which presents as "resizing just doesn't work" with no error to chase.

        The same constraint is why Add-ToolHelpButton rides its data on Tag
        instead of capturing $ScriptPath, and why neither uses GetNewClosure:
        that would fix the variables and break the FUNCTION lookups instead,
        because Common.ps1 is dot-sourced into each tool's script scope rather
        than the global one.

        So: the baseline hangs off the form as a NoteProperty, every handler
        starts from $this, and the work is done by named functions that the
        handler's session state can still see.
    #>
    $Form | Add-Member -NotePropertyName 'ToolkitScale' -Force -NotePropertyValue @{
        BaseW     = 0
        BaseH     = 0
        Items     = New-Object System.Collections.ArrayList
        Ready     = $false
        Busy      = $false
        LastUsed  = 1.0
        LastState = $Form.WindowState
        MaxScale  = $MaxScale
    }

    $Form.Add_Shown({
        $st = $this.ToolkitScale
        if ($null -eq $st -or $st.Ready) { return }
        $st.BaseW = $this.ClientSize.Width
        $st.BaseH = $this.ClientSize.Height
        Initialize-FormScaleBaseline -Form $this
        $st.Ready = $true

        # THE DESIGN SIZE BECOMES THE FLOOR, for every form and not just the 13
        # that already called Set-FormResizable.
        #
        # The scale factor is clamped at 1.0, so dragging smaller cannot shrink
        # the controls -- they would simply be cut off by the window edge, which
        # looks like the tool is broken. Blocking the drag is honest about the
        # fact that these layouts have no smaller arrangement. Set here rather
        # than at wiring time because Shown is the first moment the size is
        # genuinely final; Add-ToolHelpButton can still widen the form after the
        # tool has finished laying it out.
        if ($this.MinimumSize.Width -le 0 -and $this.MinimumSize.Height -le 0) {
            $this.MinimumSize = $this.Size
        }
        # A form that opens maximised is already off its baseline on the first
        # paint, so scale once here rather than waiting for a resize that may
        # never come.
        Update-FormScale -Form $this
    })

    $Form.Add_ResizeEnd({ Update-FormScale -Form $this })

    # Maximise and restore change ClientSize WITHOUT raising ResizeEnd -- that
    # event is about the drag handle finishing, not about the size changing.
    # Without this, maximising gives a big window with small text, which is the
    # exact complaint this function exists to answer.
    $Form.Add_SizeChanged({
        $st = $this.ToolkitScale
        if ($null -eq $st) { return }
        if ($this.WindowState -ne $st.LastState) {
            $st.LastState = $this.WindowState
            Update-FormScale -Form $this
        }
    })
}

function Initialize-FormScaleBaseline {
    <#
        Record where every control started, so later resizes can be computed
        from the design layout instead of from the last one.

        Walks the tree ITERATIVELY. Recursion would need the function to be
        visible from wherever it is called, which is the thing that is awkward
        here; a stack has no such requirement and no depth limit worth worrying
        about.

        Called once, from the form's Shown handler. See Enable-FormAutoScale.
    #>
    param([Parameter(Mandatory)]$Form)

    $st = $Form.ToolkitScale
    if ($null -eq $st) { return }

    $stack = New-Object System.Collections.Stack
    $stack.Push($Form)
    while ($stack.Count -gt 0) {
        $parent = $stack.Pop()
        foreach ($c in $parent.Controls) {
            $cols = $null
            if ($c -is [System.Windows.Forms.ListView]) {
                # ListView columns are not controls, so walking Controls never
                # reaches them. Without this they keep their design width inside
                # a zoomed window and the text no longer fits the column.
                $cols = @()
                foreach ($col in $c.Columns) { $cols += $col.Width }
            }
            $docked = ($c.Dock -ne [System.Windows.Forms.DockStyle]::None)

            # Distance from the control's right edge to the form's right edge,
            # for DIRECT children of the form only. A control inside a panel
            # measures against the panel, and stretching it to the form's edge
            # would push it out of its parent. $null means "not a candidate".
            $gap = $null
            if ([object]::ReferenceEquals($parent, $Form)) {
                $gap = $st.BaseW - ($c.Bounds.X + $c.Bounds.Width)
            }

            [void]$st.Items.Add([pscustomobject]@{
                Control  = $c
                Bounds   = $c.Bounds
                FontSize = $c.Font.Size
                FontFam  = $c.Font.FontFamily
                FontStyl = $c.Font.Style
                FontUnit = $c.Font.Unit
                Docked   = $docked
                ColWidth = $cols
                RightGap = $gap
                Made     = $null
            })

            # ANCHORS ARE NEUTRALISED HERE, ONCE, AND NOT ON EVERY RESIZE.
            #
            # This is the fix for windows that came out stretched and
            # overlapping. Add-ToolTargetBar gives its textbox
            # Anchor = 'Top,Left,Right' and its Look Up button 'Top,Right', so
            # WinForms widens and moves them itself the moment the form grows.
            # Update-FormScale then writes its own bounds -- and the two are
            # fighting over the same controls.
            #
            # Clearing the anchor inside Update-FormScale was not enough,
            # because that happens between SuspendLayout and ResumeLayout: the
            # resumed layout pass still had an anchored control to reconcile,
            # and the textbox came back out stretched to the window edge with
            # the '?' sitting on top of it.
            #
            # Doing it at capture time means WinForms stops moving these
            # controls before the first resize ever happens, so scaling is the
            # ONLY thing positioning them. Bounds are recorded on the line above
            # first -- reading them after the anchor change would read whatever
            # the reconciliation had just done.
            if (-not $docked) {
                try {
                    $c.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                                [System.Windows.Forms.AnchorStyles]::Left
                } catch { }
            }

            if ($c.Controls.Count -gt 0) { $stack.Push($c) }
        }
    }
}

function Update-FormScale {
    <#
        Resize every control to the current zoom factor.

        The factor is the SMALLER of the two axis ratios, so proportions are
        preserved exactly and any excess on the other axis becomes margin. Never
        below 1.0 -- these layouts are absolutely positioned and have no smaller
        arrangement to fall back to, which is also why Set-FormResizable pins
        MinimumSize to the design size.
    #>
    param([Parameter(Mandatory)]$Form)

    $st = $Form.ToolkitScale
    if ($null -eq $st -or -not $st.Ready -or $st.Busy) { return }
    if ($st.BaseW -le 0 -or $st.BaseH -le 0) { return }

    $f = [Math]::Min(($Form.ClientSize.Width / $st.BaseW),
                     ($Form.ClientSize.Height / $st.BaseH))
    if ($f -lt 1.0)          { $f = 1.0 }
    if ($f -gt $st.MaxScale) { $f = $st.MaxScale }

    # Fonts are rebuilt only when the factor moved. Bounds are ALWAYS re-applied:
    # a drag that only changes the width leaves the factor alone and still has
    # to stretch the full-width controls to the new edge -- an early exit here
    # was the console's help-button bug in a different coat.
    $fontChanged = ([Math]::Abs($f - $st.LastUsed) -ge 0.01)

    $st.Busy = $true
    $Form.SuspendLayout()
    try {
        foreach ($it in $st.Items) {
            $c = $it.Control
            if ($null -eq $c -or $c.IsDisposed) { continue }

            # FONT FIRST, BOUNDS SECOND. An AutoSize control resizes itself the
            # moment its font changes, so setting bounds first just gets
            # overwritten.
            #
            # THE FONT IS NEVER DISPOSED, AND THAT IS DELIBERATE.
            #
            # The first version disposed the font it had made on the previous
            # pass, reasoning that it owned it. It does not. A control whose
            # Font was never explicitly set INHERITS its parent's Font object --
            # $child.Font and $parent.Font are the same instance. Assigning a new
            # font to the parent and disposing the old one leaves every
            # not-yet-processed child pointing at a disposed GDI+ object, and the
            # next layout pass calls MeasureString on it and throws
            # "Parameter is not valid" from deep inside Label.SetBoundsCore.
            # On screen that is a window full of red X boxes.
            #
            # Fonts are small and their finalizers release the handle. Leaning on
            # the GC is the correct trade here: this runs once per completed
            # drag, not per mouse-move.
            if ($fontChanged) {
                try {
                    $pt = [float]($it.FontSize * $f)
                    # GDI+ refuses absurd em sizes and reports it as the same
                    # unhelpful "Parameter is not valid".
                    if ($pt -gt 0 -and $pt -le 96) {
                        $c.Font = New-Object System.Drawing.Font(
                            $it.FontFam, $pt, $it.FontStyl, $it.FontUnit)
                    }
                } catch { }
            }

            if (-not $it.Docked) {
                # ANCHORS OFF. An anchored control is moved by WinForms AND by
                # this loop -- it travels twice and leaves the window. The '?'
                # anchors Top+Right, which is precisely the combination that
                # walks off the right edge and vanishes.
                $c.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                            [System.Windows.Forms.AnchorStyles]::Left

                $auto = $false
                try { $auto = [bool]$c.AutoSize } catch { }
                if ($auto) {
                    # An AutoSize control owns its own width and height. Writing
                    # them forces a GetPreferredSize pass that measures the text
                    # again for no benefit -- and that measurement is exactly
                    # where a bad font surfaces as a crash.
                    $c.Location = New-Object System.Drawing.Point(
                        [int][Math]::Round($it.Bounds.X * $f),
                        [int][Math]::Round($it.Bounds.Y * $f))
                } else {
                    $w = [int][Math]::Round($it.Bounds.Width * $f)

                    # FULL-WIDTH CONTROLS FOLLOW THE RIGHT EDGE.
                    #
                    # Uniform zoom is capped by the shorter axis, so a tall
                    # dialog maximised on a wide screen scales to fit the height
                    # and leaves a strip of nothing down the right -- 560px of
                    # it on Out of Office at 1840 wide. The controls that
                    # SPANNED the window at design time are the ones that
                    # should own that space: the output box, the address bar,
                    # the grid. Anything whose design right edge sat within
                    # 24px of the design client width is stretched to sit the
                    # same scaled distance from the NEW right edge.
                    #
                    # Only spanning controls. A button that happens to be near
                    # the edge is not spanning and stays a button.
                    $x = [int][Math]::Round($it.Bounds.X * $f)
                    if ($null -ne $it.RightGap -and $it.RightGap -ge 0) {
                        if ($it.RightGap -le 24 -and $it.Bounds.Width -ge ($st.BaseW / 2)) {
                            # Spanning: keep the left edge, stretch to the right.
                            $w = $Form.ClientSize.Width - $x - [int][Math]::Round($it.RightGap * $f)
                            if ($w -lt 40) { $w = [int][Math]::Round($it.Bounds.Width * $f) }
                        } elseif ($it.RightGap -lt 60 -and $it.Bounds.Width -lt ($st.BaseW / 2)) {
                            # Small and hugging the right edge -- the '?' and
                            # the logo. Their position is a distance from the
                            # RIGHT; scaling X from the left marches them into
                            # the middle of a wide window. Same rule the console
                            # uses for its own corner furniture.
                            $x = $Form.ClientSize.Width - $w - [int][Math]::Round($it.RightGap * $f)
                        }
                    }

                    $c.SetBounds(
                        $x,
                        [int][Math]::Round($it.Bounds.Y      * $f),
                        $w,
                        [int][Math]::Round($it.Bounds.Height * $f))
                }
            }

            if ($null -ne $it.ColWidth -and $c -is [System.Windows.Forms.ListView]) {
                for ($i = 0; $i -lt $c.Columns.Count -and $i -lt $it.ColWidth.Count; $i++) {
                    $c.Columns[$i].Width = [int][Math]::Round($it.ColWidth[$i] * $f)
                }
            }
        }
        $st.LastUsed = $f
    } finally {
        $Form.ResumeLayout($true)
        $st.Busy = $false
    }
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
    $txt.Name     = 'txtToolTarget'
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

    # STATUS FOR FREE. This handler is registered BEFORE the tool's own, so
    # it runs first on every click: the bar goes busy with the address being
    # looked up. The matching Done is POSTED (BeginInvoke) rather than run,
    # so it executes after the tool's handler -- the whole click -- has
    # returned, and stamps how long the look-up took. A tool that started a
    # background job in its handler keeps the bar (JobRunning) and the job
    # marks Done itself. Tools that never heard of the status bar get all of
    # this without an edit; tools that want better words call Set-ToolStatus
    # inside their handler and those words win.
    $btn.Add_Click({
        $f = $this.FindForm()
        if (-not $f) { return }
        $s = Get-ToolStatusState -Form $f
        if (-not $s) { return }
        $who = ''
        try { $who = "$($this.Parent.Controls[$this.Tag].Text)".Trim() } catch { }
        Set-ToolStatus -Form $f -Busy -Text $(if ($who) { "Looking up $who ..." } else { 'Looking up ...' })
        # Not a closure: a closure runs in its own module and could not see
        # Set-ToolStatus (defined in the tool's script scope). The form is
        # passed as the delegate's argument instead.
        try {
            [void]$f.BeginInvoke([System.Action[object]]{
                param($frm)
                $s2 = Get-ToolStatusState -Form $frm
                if ($s2 -and $s2.Busy -and -not $s2.JobRunning) { Set-ToolStatus -Form $frm -Done }
            }, [object[]]@($f))
        } catch { }
    })
    $btn.Tag = $txt.Name

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

function Resolve-ConsoleThemeName {
    <#
        WHICH THEME THIS OPERATOR GETS, and WHY -- as a name and a reason.

        Returns { Name; Source; Detail }. Source is one of 'profile', 'tenant',
        'binding', 'active' or '' when nothing resolves.

        WHY THIS IS ITS OWN FUNCTION. The order -- your profile, then the
        tenant's file, then a connected-domain binding, then the shared
        ActiveTheme -- was written once inside Get-ConsoleTheme and then
        re-implemented, badly, in Theme Manager's "Console uses:" label. The
        label read ActiveTheme and nothing else, so for anybody with a theme of
        their own it named a theme their console was not using. Worse, Theme
        Manager's "has anything changed?" test compared against the same shared
        pointer -- so selecting the theme that HAPPENED to equal ActiveTheme
        looked like "no change", wrote nothing to the profile, and left the
        operator's own theme still winning. Reported 2026-09-16 as "I can set
        the theme to anything except Navy", Navy being what ActiveTheme said.

        One function, two callers, and the answer cannot drift again.

        -Config takes an already-parsed Themes.json, which is how Theme Manager
        asks about the file it is editing rather than the one on disk.
    #>
    param($Config)

    $cfg = $Config
    if (-not $cfg) {
        $f = Get-ThemeFilePath
        if (-not $f -or -not (Test-Path -LiteralPath $f)) {
            return [PSCustomObject]@{ Name = ''; Source = ''; Detail = 'no Themes.json' }
        }
        try { $cfg = Get-Content -LiteralPath $f -Raw | ConvertFrom-Json } catch {
            return [PSCustomObject]@{ Name = ''; Source = ''; Detail = 'Themes.json could not be read' }
        }
    }

    # THIS PERSON'S PICK OUTRANKS EVERYTHING. Two operators sharing a folder
    # have a profile each, so 'the theme' is 'the theme for whoever is at the
    # keyboard'. With no profile this reads empty and the rest decides exactly
    # as it did before profiles existed.
    $userTheme = ''
    try { $userTheme = "$(Get-UserLayoutProperty -Name 'Theme')".Trim() } catch { $userTheme = '' }
    if ($userTheme) {
        return [PSCustomObject]@{ Name = $userTheme; Source = 'profile'; Detail = 'your profile' }
    }

    # Multi-tenant only: ActiveTheme is one value and cannot say 'this tenant
    # is red and that one is blue'.
    $tenantTheme = ''
    try { if (Test-MultiTenantMode) { $tenantTheme = "$((Get-TenantConfig).Theme)".Trim() } } catch { }
    if ($tenantTheme) {
        $lbl = ''
        try { $lbl = "$(Get-ActiveTenantLabel)" } catch { }
        return [PSCustomObject]@{
            Name = $tenantTheme; Source = 'tenant'
            Detail = $(if ($lbl) { "the tenant file for '$lbl'" } else { 'the tenant file' })
        }
    }

    # Known only after a connection, which is why it is last but one.
    if ($cfg.FollowTenant) {
        try {
            $d = Get-ConnectedTenantDomain
            if ($d -and $cfg.TenantThemes -and ($cfg.TenantThemes.PSObject.Properties.Name -contains $d)) {
                return [PSCustomObject]@{
                    Name = "$($cfg.TenantThemes.$d)"; Source = 'binding'
                    Detail = "the tenant binding for $d"
                }
            }
        } catch { }
    }

    $act = "$($cfg.ActiveTheme)".Trim()
    if ($act) {
        return [PSCustomObject]@{ Name = $act; Source = 'active'; Detail = 'ActiveTheme in Themes.json (the shared fallback)' }
    }
    return [PSCustomObject]@{ Name = ''; Source = ''; Detail = 'nothing names a theme' }
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

    # SAFE MODE SHORT-CIRCUITS EVERY THEME LOOKUP, and it is checked here
    # rather than at each call site on purpose: this is the single door every
    # themed control in the toolkit comes through. $null is already the
    # documented "no theme" answer -- Set-ConsoleTheme still applies the
    # typeface and DPI scaling for it, so a safe-mode window is a plain
    # Windows window rather than an unstyled one.
    if (Test-ToolkitSafeMode) { return $null }

    if ($Refresh) {
        $script:ThemeCache = $null
        $global:AdminConsoleTheme = $null
    }
    # THE CACHE IS ONLY AS GOOD AS THE FILE IT CAME FROM. Both caches are
    # invalidated when Themes.json has been written since they were filled
    # -- by Theme Manager, by hand, by anyone. Without this, a console
    # re-run from the same PowerShell window (the global outlives the
    # script) kept serving the theme that was active the FIRST time, and a
    # theme picked in between looked as if it never loaded (2026-09-10).
    $themeFile = Get-ThemeFilePath
    $stamp = $null
    if ($themeFile -and (Test-Path -LiteralPath $themeFile)) {
        try { $stamp = (Get-Item -LiteralPath $themeFile).LastWriteTimeUtc.Ticks } catch { }
    }
    if ($stamp -and $global:AdminConsoleThemeStamp -and $stamp -ne $global:AdminConsoleThemeStamp) {
        $script:ThemeCache = $null
        $global:AdminConsoleTheme = $null
    }
    if (-not $Name -and $global:AdminConsoleTheme) { return $global:AdminConsoleTheme }
    if (-not $script:ThemeCache) {
        if (-not $themeFile) { return $null }
        if (-not (Test-Path -LiteralPath $themeFile)) { return $null }
        try {
            $script:ThemeCache = Get-Content -LiteralPath $themeFile -Raw |
                                 ConvertFrom-Json
            $global:AdminConsoleThemeStamp = $stamp
        } catch { return $null }
    }
    $cfg = $script:ThemeCache

    if (-not $Name) {
        $Name = $cfg.ActiveTheme

        # THE TENANT'S OWN Theme WINS, IN MULTI-TENANT MODE ONLY.
        #
        # ActiveTheme is a single value in Themes.json, so with several tenants
        # in one folder it cannot say "ABCD is red and WXYZ is blue".
        # Tenant.json has always had a Theme key for exactly this; until
        # 2026-09-14 nothing read it.
        #
        # THIS HALF IS USELESS ALONE, AND THAT IS THE LESSON FROM SHIPPING IT
        # ALONE. Read the tenant's Theme without ALSO having Theme Manager write
        # the operator's pick into the tenant file, and whatever stale value the
        # file already carried wins forever -- every theme change appears to do
        # nothing. That is precisely what happened on first attempt. The two
        # changes are one change; do not port one without the other.
        #
        # READ BEFORE FollowTenant, deliberately: this is known at STARTUP from
        # a file, while FollowTenant keys on Get-ConnectedTenantDomain and knows
        # nothing until the first connection. A console that opens in one
        # tenant's colours and switches thirty seconds later is worse than one
        # that was right from the start.
        #
        # SINGLE-TENANT INSTALLS SKIP THIS ENTIRELY. There, ActiveTheme is the
        # answer and Tenant.json's Theme is ignored by design -- reading it would
        # make Theme Manager appear broken on every install that has ever set it.
        # THIS PERSON'S PICK OUTRANKS THE TENANT'S. Two operators sharing a
        # folder have their own profile each, so 'the tenant theme' is now
        # 'the tenant theme for whoever is at the keyboard'. With no profile this
        # reads as empty and the tenant file decides exactly as before.
        #
        # THE ORDER LIVES IN Resolve-ConsoleThemeName, not here. It used to be
        # written out inline, and Theme Manager then re-implemented a WRONG
        # version of it for its "Console uses:" label and for its has-anything-
        # changed test -- see that function for what that cost. One definition,
        # two callers.
        $Name = "$((Resolve-ConsoleThemeName -Config $cfg).Name)"
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

    # AN ICON EATS INTO THE CAPTION'S BOX, so the width has to know about it.
    # Without this the auto-fit would size every button for text alone, the
    # gutter would push each caption off its own right-hand edge, and
    # Add-ButtonGlyph's fit test would then refuse an icon on exactly the
    # longest names -- leaving the grid ragged in the one place a gap looks
    # like a fault. Maximum moves with Inset for the same reason: capping the
    # width at the old ceiling would just reintroduce the clipping there.
    #
    # AFTER the font default above, not before: the gutter scales with the
    # font, so asking a $null font for its size would silently price the icon
    # at the smallest gutter and clip the captions it was added to prevent.
    try {
        if (Test-ConsoleIconStyle) {
            $gutter = (Get-ButtonGlyphMetrics -FontSize $Font.Size).Gutter
            $Inset   += $gutter
            $Maximum += $gutter
        }
    } catch { }

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
    # The panel is WIDER than the grid it holds, and the surplus becomes
    # padding either side (Set-SectionPanelGrid). The grid itself is always
    # exactly ButtonsPerRow cells -- since 2026-09-12 the column count is a
    # hard limit rather than however many happened to fit -- so this number is
    # the panel, not the column count, and the slack cannot add a column.
    return ($cell * $ButtonsPerRow) + 2 + $sb + $script:ConsolePanelSlack
}

# ---------------------------------------------------------------------------
# TYPOGRAPHY AND BUTTON ROLES (2026-09-09, ARCHITECTURE 42)
#
# Fonts are the loudest single tell of an old application. Every form in this
# toolkit was built with 'Microsoft Sans Serif' and Consolas grids; the theme
# now re-faces every control to the theme's FontFamily (default Segoe UI, the
# Windows system face) and keeps monospace only where it carries meaning --
# text boxes holding logs, HTML source or identifiers. Sizes and styles are
# untouched, so layout shifts by the few percent Segoe is wider, and the
# zoom-on-resize baseline is captured AFTER this pass so the two compose.
#
# Fonts are cached per (family, size, style) and NEVER disposed. A font
# assigned to a control is shared with every control that got the same one;
# disposing on replacement is the crash the resize work already hit once
# ("Parameter is not valid" from Label.SetBoundsCore).
# ---------------------------------------------------------------------------
$script:ThemeFontCache = @{}

#region Safe mode and readability ---------------------------------------------
<#
    THE ESCAPE HATCH FOR A CONSOLE YOU CANNOT READ.

    Everything here exists because of one asymmetry: the configuration that
    breaks the console worst is the configuration that makes the console
    unreadable, and the tool for fixing it is the console. A theme whose text
    and background are the same colour leaves an operator clicking blindly at
    a window they cannot see, with the fix two menus deep inside it.

    THE ANSWER IS SAFE MODE RATHER THAN A RESET BUTTON, and the difference
    matters. Reset is destructive, needs a confirmation nobody reads, and
    charges an engineer their paths and their tool layout to fix a colour. Safe
    mode changes nothing on disk: it ignores the theme and the section geometry
    for ONE run, so the operator gets a legible window and can then repair the
    config with the normal editors. Every browser and Office application solves
    this the same way, and for the same reason.

    IT TRAVELS BY ENVIRONMENT VARIABLE, not by a parameter passed down a call
    chain. Tools are launched both in-process with '&' and detached through
    Start-Process, and a child process inherits environment variables over
    both paths -- so a tool opened from a safe-mode console is also in safe
    mode without a single tool needing to know the concept exists.
#>
function Test-ToolkitSafeMode {
    return ("$env:ToolkitSafeMode".Trim() -eq '1')
}

# THE SWITCH IS SET AND CLEARED IN AdminConsole.ps1, BEFORE THIS FILE LOADS.
#
# It has to happen there, not here: the console's -SafeMode parameter is the
# authority, and it must be applied before anything can ask for a theme --
# which includes this file's own load. There is deliberately no
# Set-ToolkitSafeMode helper, because a helper living in the file that is
# dot-sourced AFTER the decision could only ever be called too late.

function Get-ColorLuminance {
    <#
        WCAG relative luminance, 0 (black) to 1 (white).

        The gamma expansion is not decoration: perceived brightness is not
        linear in the sRGB value, so a naive average of R, G and B calls
        yellow-on-white readable and calls blue-on-black unreadable, and both
        answers are wrong in the direction that matters.
    #>
    param([Parameter(Mandatory)][System.Drawing.Color]$Color)
    $f = {
        param([double]$c)
        $c = $c / 255.0
        if ($c -le 0.03928) { return ($c / 12.92) }
        return [Math]::Pow((($c + 0.055) / 1.055), 2.4)
    }
    return (0.2126 * (& $f $Color.R)) + (0.7152 * (& $f $Color.G)) + (0.0722 * (& $f $Color.B))
}

function Get-ColorContrastRatio {
    <# 1.0 (identical) to 21.0 (black on white). WCAG AA wants 4.5 for body text. #>
    param([Parameter(Mandatory)][System.Drawing.Color]$Fore, [Parameter(Mandatory)][System.Drawing.Color]$Back)
    $a = Get-ColorLuminance -Color $Fore
    $b = Get-ColorLuminance -Color $Back
    $hi = [Math]::Max($a, $b); $lo = [Math]::Min($a, $b)
    return [Math]::Round((($hi + 0.05) / ($lo + 0.05)), 2)
}

function Test-ThemeReadable {
    <#
        Check a theme's text/background pairs and report the ones that fail.

        CHECKED AT SAVE, NOT AT LOAD. Refusing to LOAD an unreadable theme
        would be too late and too rude: the file is already written and the
        operator is already locked out. Checking as it is saved is the only
        moment where the person who made the mistake is present, knows what
        they just did, and can undo it in one keystroke.

        Returns an array of findings; empty means every pair is legible.
        4.5 is WCAG AA for body text; below 3.0 is called out separately
        because that is the range where text stops being merely hard work and
        starts being invisible.
    #>
    param([Parameter(Mandatory)]$Theme)
    $pairs = @(
        @{ Name = 'Form text on form background';     F = 'FormText';   B = 'FormBack' },
        @{ Name = 'Button text on button face';       F = 'ButtonText'; B = 'ButtonBack' },
        @{ Name = 'Accent text on accent';            F = 'AccentText'; B = 'AccentBack' },
        @{ Name = 'Group heading on panel';           F = 'GroupText';  B = 'PanelBack' },
        @{ Name = 'Header text on header band';       F = 'HeaderText'; B = 'HeaderBack' }
    )
    $out = @()
    foreach ($p in $pairs) {
        $fc = $null; $bc = $null
        try { $fc = ConvertTo-ThemeColor "$($Theme.($p.F))" } catch { }
        try { $bc = ConvertTo-ThemeColor "$($Theme.($p.B))" } catch { }
        # A pair the theme does not define is not a failure -- HeaderBack is
        # optional and most themes have none. Only judge what was set.
        if (-not $fc -or -not $bc) { continue }
        $r = Get-ColorContrastRatio -Fore $fc -Back $bc
        if ($r -ge 4.5) { continue }
        $out += [PSCustomObject]@{
            Pair     = $p.Name
            Fore     = $p.F
            Back     = $p.B
            Ratio    = $r
            Severity = $(if ($r -lt 3.0) { 'Unreadable' } else { 'Hard to read' })
        }
    }
    return $out
}
#endregion

function Get-ThemeFont {
    param([string]$Family, [single]$Size, [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular)
    $key = "$Family|$Size|$([int]$Style)"
    if (-not $script:ThemeFontCache.ContainsKey($key)) {
        $script:ThemeFontCache[$key] = New-Object System.Drawing.Font($Family, $Size, $Style)
    }
    return $script:ThemeFontCache[$key]
}

function Set-ThemeFontOn {
    <#
        Re-face one control. Monospace survives only on TextBox/RichTextBox --
        a grid of names in Consolas is a terminal, not a table.
    #>
    param($Control, [string]$UiFamily, [string]$MonoFamily)
    if (-not $UiFamily) { return }
    $f = $null
    try { $f = $Control.Font } catch { return }
    if (-not $f) { return }
    $cur = "$($f.FontFamily.Name)"
    $isMono = $cur -match '^(Consolas|Courier|Lucida Console|Cascadia)'
    $keepMono = $isMono -and (($Control -is [System.Windows.Forms.TextBox]) -or ($Control -is [System.Windows.Forms.RichTextBox]))
    $want = if ($keepMono) { $MonoFamily } else { $UiFamily }
    if (-not $want -or $cur -ieq $want) { return }
    try { $Control.Font = Get-ThemeFont -Family $want -Size $f.Size -Style $f.Style } catch { }
}

function Get-ConsoleTitleText {
    <# The console's name. Settings.ConsoleTitle, or the generic default. #>
    $t = ''
    try { $t = "$((Get-ToolkitSettings).ConsoleTitle)".Trim() } catch { }
    if (-not $t) { $t = 'Admin Console' }
    return $t
}

function Test-FontFamilyInstalled {
    <#
        Is this font family actually present on this machine?

        THIS CHECK EXISTS BECAUSE THE FAILURE IS SILENT. New-Object
        System.Drawing.Font('Some Brand Font', 15) does NOT throw when the
        family is missing -- GDI+ quietly substitutes a default and hands back
        a perfectly valid Font object. Set a brand face that was never
        installed and the console looks untouched, with nothing anywhere
        saying why.

        System.Drawing.FontFamily's constructor DOES throw for a name it
        cannot resolve, so it is the honest test.

        THE FontFamily IS NOT DISPOSED, AND THAT IS THE FIX, NOT AN OVERSIGHT.

        There used to be a finally block here that disposed it. That looks like
        good hygiene and is the opposite of it. A FontFamily built from an
        INSTALLED family does not own a private handle -- it wraps the one GDI+
        keeps for the whole process. Disposing it releases a handle that every
        Font already built from that family, and every Font built from it
        afterwards, is still using.

        THE DAMAGE DOES NOT RAISE AN ERROR. Text simply stops being drawn:
        TextRenderer paints nothing and controls come out as empty rectangles
        in their own background colour, which reads as a layout fault rather
        than a font one. That is why the themed message dialog appeared with a
        blank label and blank buttons, and why nothing was ever written to the
        error log -- there was no exception to catch.

        It was intermittent between windows because it depended on whether a
        form's fonts were built before or after the dispose, which is why some
        dialogs in the same session were fine and others were blank.

        Switching Settings.ConsoleIcons on is what made it fire every time:
        Get-ToolkitIconFontName asks this about 'Segoe Fluent Icons' during the
        first Set-ConsoleTheme. With icons off it was still happening once per
        launch for Settings.ConsoleTitleFont, which is how it went unnoticed.

        NOTHING LEAKS. A FontFamily for an installed face is a thin wrapper the
        GC collects; the process-wide handle behind it belongs to GDI+ and is
        not ours to release.
    #>
    param([string]$Family)
    if (-not "$Family".Trim()) { return $false }
    try {
        [void](New-Object System.Drawing.FontFamily("$Family".Trim()))
        return $true
    } catch { return $false }
}

function Get-ConsoleTitleFont {
    <#
        The face for the console's title.

        Settings.ConsoleTitleFont names a family -- a corporate brand face,
        say, pushed to the estate by policy. It is used only if it is really
        installed; otherwise the theme's UI font is used, which is what the
        console looked like before anyone asked for branding.

        FALLING BACK RATHER THAN FAILING is deliberate: the toolkit ships to
        machines that will not all have the same fonts, and a console that
        refuses to draw its own title because of a typeface would be an
        absurd way to lose a morning.
    #>
    param([single]$Size = 15, [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Bold)
    $fam = ''
    try { $fam = "$((Get-ToolkitSettings).ConsoleTitleFont)".Trim() } catch { }
    if ($fam -and (Test-FontFamilyInstalled -Family $fam)) {
        return (Get-ThemeFont -Family $fam -Size $Size -Style $Style)
    }
    return (Get-ThemeUiFont -Size $Size -Style $Style)
}

function Get-ThemeUiFont {
    <#
        The face the theme will apply, for code that has to MEASURE before the
        theme runs -- the console sizes its tool buttons from caption widths
        before Set-ConsoleTheme re-faces them, and Segoe is a few percent wider
        than the face the form starts with.
    #>
    param([single]$Size = 9, [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular)
    $fam = ''
    try { $t = Get-ConsoleTheme; if ($t) { $fam = "$($t.FontFamily)".Trim() } } catch { }
    if (-not $fam) { $fam = 'Segoe UI' }
    return (Get-ThemeFont -Family $fam -Size $Size -Style $Style)
}

function Get-ButtonRole {
    <#
        Primary, Danger or Plain -- from Tag first, caption second.

        Tag wins so a form can say what it means. Console tool tiles carry
        their tool OBJECT in Tag and are always Plain: they are navigation,
        and their section colouring is the signal there.

        Caption rules are deliberately narrow. Primary is the button that
        COMMITS (Save, Apply, Create...). Look Up, Load, Refresh, Preview
        stay Plain: reads are not the thing on the form that deserves the
        accent. Danger is the button that destroys.
    #>
    param($Button)
    $tag = $Button.Tag
    if ($null -ne $tag -and -not ($tag -is [string])) { return 'Plain' }
    $t = "$tag"
    if ($t -match '(?i)\bPrimary\b') { return 'Primary' }
    if ($t -match '(?i)\bDanger\b')  { return 'Danger' }
    if ($t -match '(?i)\bPlain\b')   { return 'Plain' }
    $cap = "$($Button.Text)".Trim()
    if ($cap -match '(?i)^(Delete|Remove|Purge|Wipe|Revoke|Terminate|Decommission)\b') { return 'Danger' }
    if ($cap -match '(?i)^(Save|Apply|Create|Run|Start|Restore|Send|Grant|Connect|Add|Enable|Submit|Activate|Release|Unblock|Assign|Insert)\b') { return 'Primary' }
    return 'Plain'
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

#region Button glyphs ----------------------------------------------------------
<#
    ICONS ON BUTTONS, WITH NOTHING TO SHIP AND NOTHING TO DEPLOY.

    Windows already carries two icon FONTS -- 'Segoe Fluent Icons' on Windows
    11 and 'Segoe MDL2 Assets' on Windows 10 -- whose glyphs sit in the Unicode
    private-use area. An icon is therefore just a CHARACTER drawn in that face,
    which is why this is done with a font rather than with .png files:

      * nothing travels alongside the toolkit, so nothing can arrive missing
        or get stripped by a mail gateway on the way to a new engineer,
      * the glyph takes the button's own ForeColor, so it themes and greys out
        for free instead of needing a second set of images for dark themes,
      * it scales with DPI like text rather than blurring like a bitmap, and
      * there is no GDI image handle per button to leak.

    CODEPOINTS ARE WRITTEN AS [char]0xE715, NEVER AS THE CHARACTER ITSELF.
    Every file in this toolkit is ASCII on purpose. A private-use character
    pasted into source survives exactly until something re-saves the file in
    the wrong encoding, and the failure is a console full of empty boxes with
    nothing anywhere explaining why. The hex form cannot be damaged that way.

    THE GLYPH IS INFERRED FROM THE CAPTION RATHER THAN CONFIGURED PER TOOL.
    Fifty-odd tools would otherwise need fifty-odd edits to Tools.json before
    anybody saw a single icon, and every tool added afterwards would arrive
    blank. A tool can still name one explicitly -- Tools.json 'Icon' -- and
    that always wins over the guess.

    ACTION WORDS AND SUBJECT WORDS ARE MATCHED IN A DIFFERENT ORDER depending
    on what the button IS, and that is the whole reason there are two tables.
    'New User' on the console is a tool for dealing with users, so the subject
    wins and it gets a person. 'New...' on a dialog is the button that creates
    something, so the action wins and it gets a plus. One table matched in one
    order gets one of those two wrong whichever order you pick.

    ALL OF IT IS ONE SETTING AWAY FROM OFF. Settings.ConsoleIcons is 'On' or
    blank, Test-ConsoleIconStyle is the only thing that reads it, and safe mode
    forces it off the way it forces off every other piece of this toolkit's own
    drawing. Nothing in any tool script had to change to gain icons, so nothing
    in any tool script has to change to lose them.
#>

function Get-ToolkitIconFontName {
    <#
        The icon font this machine actually has, or '' if it has neither.

        Cached in a global because it is asked once per button per theme pass
        and the answer cannot change while Windows is running. '' is a real
        answer and is cached as one -- $null is the "not asked yet" state, so
        the test below is against $null rather than against emptiness.

        Windows 10 has MDL2 and not Fluent; Windows 11 has both. The two share
        codepoints for everything used here, so preferring Fluent gets the
        newer drawing where it exists and changes nothing where it does not.
    #>
    if ($null -ne $global:ToolkitIconFontName) { return $global:ToolkitIconFontName }
    $found = ''
    foreach ($name in @('Segoe Fluent Icons', 'Segoe MDL2 Assets')) {
        if (Test-FontFamilyInstalled -Family $name) { $found = $name; break }
    }
    $global:ToolkitIconFontName = $found
    return $found
}

function Test-ConsoleIconStyle {
    <#
        Settings.ConsoleIcons = 'On', AND a font to draw them with.

        The font check is not belt and braces. Asking for a missing family does
        not throw -- GDI+ substitutes a default and hands back a valid Font --
        so with no icon font present the private-use codepoints would render as
        whatever the substitute face has at U+E715, which is nothing, a box, or
        a random letter. Better to stay text-only than to draw rubbish.

        SAFE MODE IS ALWAYS OFF, for the same reason it forces buttons flat.
    #>
    if (Test-ToolkitSafeMode) { return $false }
    $on = $false
    try { $on = ("$((Get-ToolkitSettings).ConsoleIcons)".Trim() -ieq 'On') } catch { return $false }
    if (-not $on) { return $false }
    return [bool](Get-ToolkitIconFontName)
}

function Get-ToolkitIconFont {
    <#
        A cached Font in the icon face.

        CACHED BECAUSE A Paint HANDLER IS NOT A PLACE TO ALLOCATE. These run on
        every invalidate -- every hover, every resize, every grey-out -- and a
        Font created there is a GDI handle leaked there. One object per size is
        all this ever needs.
    #>
    param([single]$Size = 12)
    $fam = Get-ToolkitIconFontName
    if (-not $fam) { return $null }
    if (-not $global:ToolkitIconFonts) { $global:ToolkitIconFonts = @{} }
    $key = "$fam|$Size"
    if ($global:ToolkitIconFonts.ContainsKey($key)) { return $global:ToolkitIconFonts[$key] }
    $f = $null
    try {
        $f = New-Object System.Drawing.Font($fam, $Size, [System.Drawing.FontStyle]::Regular,
                                            [System.Drawing.GraphicsUnit]::Point)
    } catch { $f = $null }
    $global:ToolkitIconFonts[$key] = $f
    return $f
}

function Get-ButtonGlyphMetrics {
    <#
        Where the glyph sits and how much room the caption loses to it.

        A FUNCTION RATHER THAN $script: CONSTANTS, and that is deliberate. Three
        different places need these numbers -- the flat path, the owner-drawn
        path in Add-RaisedButtonChrome, and the button auto-fit -- and $script:
        inside a Common.ps1 function resolves to whichever script is on the call
        stack at the time, which for a tool the console launched is the TOOL.
        The numbers would be present in one caller and empty in the next.

        DERIVED FROM THE BUTTON'S OWN FONT rather than fixed, because the font
        is not fixed. The list view has a size stepper, the console has a zoom,
        and both exist because somebody needs the text bigger -- an icon that
        stayed 16px while the words around it grew would single out the one
        part of the window that ignored the request.

        Gutter is Left + Size + the gap: what the caption is pushed by, so the
        flat and raised paths line their text up at the same x.
    #>
    param([single]$FontSize = 9)
    $k = $FontSize / 9.0
    if ($k -lt 1.0) { $k = 1.0 }
    return [PSCustomObject]@{
        Left   = [int](9 * $k)
        Size   = [int](16 * $k)
        Gap    = [int](5 * $k)
        Glyph  = [single](12 * $k)
        Gutter = [int](30 * $k)
    }
}

function Get-ToolkitGlyphName {
    <#
        Named icons, for Tools.json 'Icon' and for anything that wants to ask
        for one outright instead of being guessed at.

        Every codepoint here was read off Microsoft's published Segoe MDL2
        Assets and Segoe Fluent Icons lists rather than remembered. A wrong
        one does not fail -- it draws a different picture, confidently.
    #>
    if ($global:ToolkitGlyphNames) { return $global:ToolkitGlyphNames }
    $m = @{
        'Mail'        = 0xE715; 'Contact'     = 0xE77B; 'People'      = 0xE716
        'Calendar'    = 0xE787; 'Search'      = 0xE721; 'Permissions' = 0xE8D7
        'Lock'        = 0xE72E; 'Unlock'      = 0xE785; 'Document'    = 0xE8A5
        'Settings'    = 0xE713; 'Refresh'     = 0xE72C; 'Delete'      = 0xE74D
        'Warning'     = 0xE7BA; 'Info'        = 0xE946; 'Help'        = 0xE897
        'Check'       = 0xE73E; 'Blocked'     = 0xECE4; 'Shield'      = 0xEA18
        'Phone'       = 0xE8EA; 'Cloud'       = 0xE753; 'Forward'     = 0xE89C
        'Filter'      = 0xE71C; 'List'        = 0xE8FD; 'Chart'       = 0xE9D2
        'History'     = 0xE81C; 'Folder'      = 0xE8B7; 'Download'    = 0xE896
        'Upload'      = 0xE898; 'Play'        = 0xE768; 'Connect'     = 0xE703
        'Add'         = 0xE710; 'Edit'        = 0xE70F; 'Send'        = 0xE724
        'Star'        = 0xE735; 'Pin'         = 0xE718; 'Home'        = 0xE80F
        'Repair'      = 0xE90F; 'Globe'       = 0xE774; 'ContactCard' = 0xE779
        'Package'     = 0xE7B8; 'Chat'        = 0xE8F2; 'Bug'         = 0xEBE8
        'Tag'         = 0xE8EC; 'Switch'      = 0xE8AB; 'Grid'        = 0xF0E2
        'Website'     = 0xEB41; 'Clipboard'   = 0xF0E3; 'Power'       = 0xE7E8
        'Save'        = 0xE74E; 'Copy'        = 0xE8C8; 'Cancel'      = 0xE711
        'Accept'      = 0xE8FB; 'Back'        = 0xE72B; 'OpenFile'    = 0xE8E5
        'Print'       = 0xE749; 'Console'     = 0xE756; 'Reply'       = 0xE8CA
        'Undo'        = 0xE7A7; 'Sort'        = 0xE8CB; 'Preview'     = 0xE8FF
        'Share'       = 0xE72D; 'Link'        = 0xE71B; 'Attach'      = 0xE723
        'Certificate' = 0xEB95; 'HardDrive'   = 0xEDA2; 'Stopwatch'   = 0xE916
        'Flag'        = 0xE7C1; 'View'        = 0xE890; 'ZoomIn'      = 0xE8A3
        'AddUser'     = 0xE8FA; 'Manage'      = 0xE912; 'Activity'    = 0xE93E
        'Library'     = 0xE8F1; 'Tiles'       = 0xECA5
    }
    $global:ToolkitGlyphNames = $m
    return $m
}

function Get-ToolkitGlyphRule {
    <#
        The two ordered rule tables, most specific first.

        ORDER IS THE WHOLE DESIGN HERE. 'Calendar Permissions' matches both
        'calendar' and 'permission', and the calendar is the better picture, so
        calendar is listed above permissions. Adding a rule at the bottom is
        always safe; adding one at the top can quietly steal matches from
        everything below it.
    #>
    param([ValidateSet('Action', 'Subject')][string]$Kind)
    $key = "ToolkitGlyphRule$Kind"
    if ($global:ToolkitGlyphRules -and $global:ToolkitGlyphRules.ContainsKey($key)) {
        return $global:ToolkitGlyphRules[$key]
    }
    if (-not $global:ToolkitGlyphRules) { $global:ToolkitGlyphRules = @{} }

    if ($Kind -eq 'Action') {
        # ANCHORED AT THE START, every one of them. An action rule is meant to
        # fire on 'Save layout' and not on 'Password reset', and only the
        # anchor keeps 'reset' out of the first word's business.
        $rules = @(
            @{ P = '^(save|apply|commit)\b';                      C = 0xE74E }
            @{ P = '^(ok|accept|confirm|yes|approve)\b';          C = 0xE8FB }
            @{ P = '^(cancel|close|no|dismiss|clear)\b';          C = 0xE711 }
            @{ P = '^(delete|remove|purge|wipe|revoke)\b';        C = 0xE74D }
            @{ P = '^(add|new|create)\b';                         C = 0xE710 }
            @{ P = '^(edit|modify|change|update)\b';              C = 0xE70F }
            @{ P = '^(search|look ?up|find|query)\b';             C = 0xE721 }
            @{ P = '^(refresh|reload|re-?scan|sync)\b';           C = 0xE72C }
            @{ P = '^(run|start|execute|go|launch|play|next)\b';  C = 0xE768 }
            @{ P = '^(send|email|mail to)\b';                     C = 0xE724 }
            @{ P = '^(export|download)\b';                        C = 0xE896 }
            @{ P = '^(import|upload)\b';                          C = 0xE898 }
            @{ P = '^(open|browse)\b';                            C = 0xE8E5 }
            @{ P = '^(copy|duplicate)\b';                         C = 0xE8C8 }
            @{ P = '^(print)\b';                                  C = 0xE749 }
            @{ P = '^(preview|view|show|details)\b';              C = 0xE8FF }
            @{ P = '^(help|about|\?)';                            C = 0xE897 }
            @{ P = '^(back|previous|prev)\b';                     C = 0xE72B }
            @{ P = '^(test|verify|validate|check)\b';             C = 0xE73E }
            @{ P = '^(connect|sign ?in|log ?on|activate)\b';      C = 0xE703 }
            @{ P = '^(disconnect|sign ?out|log ?off|exit|quit)';  C = 0xE7E8 }
            @{ P = '^(undo|revert|restore|recover)\b';            C = 0xE7A7 }
            @{ P = '^(sort|order)\b';                             C = 0xE8CB }
            @{ P = '^(share)\b';                                  C = 0xE72D }
            @{ P = '^(pin|favou?rite)\b';                         C = 0xE718 }
            @{ P = '^(arrange|layout|tiles?|grid)\b';             C = 0xECA5 }
            @{ P = '^(settings?|options?|configure|preferences)'; C = 0xE713 }
            @{ P = '^(zoom|font|text size)\b';                    C = 0xE8A3 }
        )
    } else {
        $rules = @(
            @{ P = 'service\s*now';                                   C = 0xEB41 }
            @{ P = 'quick\s*link';                                    C = 0xE735 }
            @{ P = '^dashboard';                                      C = 0xE80F }
            @{ P = '^console$|powershell|command\s*prompt';           C = 0xE756 }
            @{ P = '^(ex ?o|spo|on-?prem)$|^connect';                 C = 0xE703 }
            @{ P = '^activate$|\bpim\b';                              C = 0xEA18 }
            @{ P = '^edit\b';                                         C = 0xE70F }
            @{ P = 'cert';                                            C = 0xEB95 }
            @{ P = 'calendar|\broom\b|booking|meeting|out of office|\boof\b'; C = 0xE787 }
            @{ P = 'inbox rule|\brules?\b|transport';                 C = 0xE71C }
            @{ P = 'trace|analy|look ?up|search|\bfind\b';            C = 0xE721 }
            @{ P = 'permission|\bperms?\b|\baccess\b|delegat|\brights\b'; C = 0xE8D7 }
            @{ P = 'forward';                                         C = 0xE89C }
            @{ P = 'attribute|\brename\b';                            C = 0xE8EC }
            @{ P = 'audit|term request|checklist|\btask';             C = 0xF0E3 }
            @{ P = 'health|repair|missing|\bfix\b';                   C = 0xE90F }
            @{ P = 'inventory|report|statistic|\bstats?\b|usage';     C = 0xE9D2 }
            @{ P = 'log ?viewer|\blogs?\b|history|recent';            C = 0xE81C }
            @{ P = 'one ?drive|share ?point|\bspo\b|\bsite\b|\bcloud\b'; C = 0xE753 }
            @{ P = 'external|\bweb\b|\burl\b|globe|sharing';          C = 0xE774 }
            @{ P = '\bdls?\b|distribution|group|\bteams?\b|member|owner'; C = 0xE716 }
            # ABOVE the mailbox rule on purpose. 'MB w No activity' is a report
            # about accounts nobody is using rather than a mailbox tool, and the
            # words that say so are the ones at the end of the name.
            @{ P = 'disabled|disable|block|deactivat|decomm|stale|orphan|no activity'; C = 0xECE4 }
            @{ P = 'mailbox|\bmail\b|message|e-?mail|\bmbx?\b|\bsmb\b';   C = 0xE715 }
            @{ P = 'user|account|contact|person|people|returning';    C = 0xE77B }
            @{ P = 'protocol|setting|config|manager|theme';           C = 0xE713 }
            @{ P = 'device|mobile|phone';                             C = 0xE8EA }
            @{ P = 'archive|retention|\bhold\b';                      C = 0xE7B8 }
            @{ P = 'restore|recover|undo';                            C = 0xE7A7 }
            @{ P = 'duplicat|convert|migrat|restamp|\bmove\b|swap|switch'; C = 0xE8AB }
            @{ P = 'sync|refresh|reload';                             C = 0xE72C }
            @{ P = 'browser|folder|\bfiles?\b|\.csv|export';          C = 0xE8B7 }
            @{ P = 'bug|feature request|issue|defect';                C = 0xEBE8 }
            @{ P = 'licen[cs]e|subscription';                         C = 0xEB95 }
            @{ P = 'header|document|\bdocs?\b';                       C = 0xE8A5 }
        )
    }
    $global:ToolkitGlyphRules[$key] = $rules
    return $rules
}

function Get-ToolGlyph {
    <#
        A glyph for this caption, as a STRING, or '' for none.

        A STRING RATHER THAN A [char] so that emptiness is testable. '-not' on
        a char is not the question anyone means to ask, and the one value it
        treats as false is [char]0 rather than "I found nothing".

        NO MATCH MEANS NO ICON, unless -Fallback. That is what stops a dialog
        from sprouting a picture next to 'Browse...' because some rule was
        loose enough to reach it: the guess has to be confident or it does not
        happen. A grid of tool tiles is the opposite case -- one tile without a
        picture in a row of tiles with them reads as a fault, not as restraint
        -- so the console asks for -Fallback and gets a neutral glyph.
    #>
    param(
        [string]$Name,
        [string]$Icon = '',
        [switch]$Fallback
    )
    $named = Get-ToolkitGlyphName
    $want = "$Icon".Trim()
    if ($want) {
        # A PowerShell hashtable is case-insensitive, so 'mail' finds 'Mail'.
        if ($named.ContainsKey($want)) { return [string][char]($named[$want]) }
        # A named icon that does not exist is a typo in Tools.json, and falling
        # through to the guess would hide it. Say nothing, draw nothing.
        return ''
    }

    $cap = "$Name".Trim()
    if (-not $cap) { return '' }
    # '&Save' is what a mnemonic looks like in .Text, and the ampersand would
    # break every anchored rule below it. So would the pin marker the list view
    # puts in front of a favourite -- 'Dashboard' is matched on '^dashboard',
    # and a leading symbol is enough to turn that into no icon at all.
    $cap = ($cap -replace '&', '') -replace '^[^A-Za-z0-9]+', ''

    $order = @('Action', 'Subject')
    if ($Fallback) { $order = @('Subject', 'Action') }
    foreach ($kind in $order) {
        foreach ($r in (Get-ToolkitGlyphRule -Kind $kind)) {
            if ($cap -imatch $r.P) { return [string][char]($r.C) }
        }
    }
    if ($Fallback) { return [string][char]0xECA5 }
    return ''
}

function Write-ButtonGlyph {
    <#
        Draw a button's glyph. Shared by the flat and the raised paths so the
        two cannot drift apart in position or colour.

        THE COLOUR IS READ OFF THE BUTTON AT PAINT TIME, exactly as the bevel
        is and for the same reason: the theme repaints, tools grey and un-grey,
        and anything cached here would be right once and wrong afterwards.
    #>
    param(
        [Parameter(Mandatory)]$Graphics,
        [Parameter(Mandatory)]$Button,
        [bool]$Pressed = $false
    )
    $ch = ''
    try { $ch = "$($Button.ToolGlyph)" } catch { }
    if (-not $ch) { return }

    $m = Get-ButtonGlyphMetrics -FontSize $Button.Font.Size
    $f = Get-ToolkitIconFont -Size $m.Glyph
    if (-not $f) { return }

    $off = $(if ($Pressed) { 1 } else { 0 })
    $col = $Button.ForeColor
    # The same grey Add-RaisedButtonChrome uses for a disabled caption. A
    # crisp icon beside a greyed word reads as a half-drawn button.
    if (-not $Button.Enabled) { $col = [System.Drawing.Color]::FromArgb(150, 155, 165) }

    $y = [int](($Button.ClientSize.Height - $m.Size) / 2)
    $r = New-Object System.Drawing.Rectangle(($m.Left + $off), ($y + $off), $m.Size, $m.Size)
    [System.Windows.Forms.TextRenderer]::DrawText($Graphics, $ch, $f, $r, $col,
        ([System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor
         [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
         [System.Windows.Forms.TextFormatFlags]::NoPrefix))
}

function Add-ButtonGlyph {
    <#
        Put an icon on a button, if one can be worked out and it will fit.

        THE FIT TEST IS NOT COSMETIC. The caption moves right by the gutter, so
        a button sized for centred text with nothing to spare would lose the
        end of its own word -- and a clipped caption is a worse outcome than no
        icon at all. Failing that test therefore leaves the button exactly as
        it was, which is the direction it is safe to fail in.

        IDEMPOTENT, like the bevel, and for the same reason: some forms are
        themed more than once, and a second Paint handler would draw the glyph
        twice and cost a handler on every invalidate.
    #>
    param(
        [Parameter(Mandatory)]$Button,
        [string]$Glyph = '',
        [switch]$Fallback
    )
    if (-not $Button) { return }
    if (-not $Button.PSObject.Properties['Text']) { return }

    if (-not $Glyph) {
        # Tools.json 'Icon' arrives on the console's tool buttons inside the
        # tool object in .Tag. Dialog buttons put a role string there instead,
        # which is why the type is checked before the property is read.
        $icon = ''
        try {
            $tag = $Button.Tag
            if ($tag -and -not ($tag -is [string]) -and $tag.PSObject.Properties['Icon']) {
                $icon = "$($tag.Icon)".Trim()
            }
        } catch { }
        $Glyph = Get-ToolGlyph -Name "$($Button.Text)" -Icon $icon -Fallback:$Fallback
    }
    if (-not $Glyph) { return }

    $m = Get-ButtonGlyphMetrics -FontSize $Button.Font.Size

    # THE FIT TEST NEEDS A WIDTH, AND A DOCKED BUTTON HAS NOT GOT ONE YET.
    # Dock = Fill means the layout decides the width later, so .Width here is
    # whatever the control was constructed at -- usually 75. Measuring against
    # that would reject an icon on every tile in the list view, which are the
    # widest buttons in the toolkit. Skip the test where it cannot be answered.
    $known = $true
    try { $known = ($Button.Dock -eq [System.Windows.Forms.DockStyle]::None -and $Button.Width -gt 0) } catch { }
    if ($known) {
        try {
            # Measured WRAPPED, in the box the caption will actually have once
            # the gutter is taken off it -- the same test Get-ConsoleButtonWidth
            # makes, and for the same reason. 'Mailbox / Onedrive Perms' is meant
            # to be two lines, so measuring it on one would reject an icon it can
            # easily hold.
            $avail = $Button.Width - $m.Gutter - 6
            if ($avail -lt 24) { return }
            $box = New-Object System.Drawing.Size($avail, 4000)
            $sz = [System.Windows.Forms.TextRenderer]::MeasureText("$($Button.Text)", $Button.Font, $box,
                        [System.Windows.Forms.TextFormatFlags]::WordBreak)
            # Width can come back WIDER than proposed: WordBreak breaks between
            # words, so a single long word overflows rather than wrapping.
            if ($sz.Width -gt $avail) { return }
            if ($sz.Height -gt ($Button.Height - 4)) { return }
        } catch { }
    }

    $had = $false
    try { $had = [bool]$Button.PSObject.Properties['ToolGlyph'] } catch { }
    $Button | Add-Member -NotePropertyName 'ToolGlyph' -NotePropertyValue $Glyph -Force

    # Push the caption clear of the glyph. Padding is what a Button lays its
    # own text out inside, so this is all the flat path needs; the raised path
    # draws the caption itself and applies the same gutter by hand.
    try {
        $Button.Padding = New-Object System.Windows.Forms.Padding($m.Gutter, 0, 4, 0)
        $Button.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    } catch { }

    if ($had) { try { $Button.Invalidate() } catch { }; return }

    $Button.Add_Paint({
        param($s, $e)
        # THE RAISED CHROME DRAWS THE GLYPH ITSELF and this handler must stand
        # down when it is present. Paint handlers fire in the order they were
        # attached, the bevel fills the whole face, and whichever of the two
        # ran first would be painted over -- so the decision is made here, at
        # paint time, where the order the two were attached in cannot matter.
        try { if ($s.PSObject.Properties['Raised3D']) { return } } catch { }
        try { Write-ButtonGlyph -Graphics $e.Graphics -Button $s } catch { }
    })

    # THE PADDING IS SET ONCE BUT THE GLYPH IS MEASURED EVERY PAINT, so a font
    # change after this point would move one and not the other. The console's
    # zoom and the list view's size stepper BOTH rewrite Font on live buttons
    # without rebuilding them, and at 150 % the glyph grows past a gutter that
    # did not -- so the icon draws over the first letters of its own caption.
    # Re-deriving here is what keeps the two halves of the layout agreeing.
    # (The raised path needs none of this: it computes the caption's left edge
    # from $s.Font at paint time and so has never been able to drift.)
    $Button.Add_FontChanged({
        try {
            $g = (Get-ButtonGlyphMetrics -FontSize $this.Font.Size).Gutter
            $this.Padding = New-Object System.Windows.Forms.Padding($g, 0, 4, 0)
        } catch { }
        $this.Invalidate()
    })

    # Only when the raised chrome is NOT going to add its own. Both handlers do
    # the same thing, and on a button that has both the state change would cost
    # two full repaints instead of one.
    if (-not $Button.PSObject.Properties['Raised3D']) {
        $Button.Add_EnabledChanged({ $this.Invalidate() })
    }
}
#endregion

function Get-ButtonCornerRadius {
    <#
        Settings.ConsoleCorners = 'Rounded' gives 6px corners; anything else
        (blank) is square, which is what every existing installation has and
        therefore what an absent key has to mean.

        ROUNDED APPLIES TO RAISED BUTTONS ONLY, and that is a rendering fact
        rather than a preference. A rounded corner needs the area outside the
        curve painted in the colour BEHIND the button, and needs antialiasing
        to avoid a staircase edge -- both of which require owning the paint.
        A flat button is drawn by WinForms; the only way to round one is to
        set a Region, which clips with hard aliased edges and looks worse than
        the square corner it replaced.

        SAFE MODE IS ALWAYS SQUARE, like every other piece of this toolkit's
        own drawing.

        ASKED ONCE AT ATTACH TIME, NEVER FROM THE PAINT HANDLER. It reads
        Get-ToolkitSettings, which stats Settings.json to check its timestamp
        -- file I/O that is nothing once per button and is a disk hit per
        repaint, on every hover, if it were asked where it is used.
    #>
    if (Test-ToolkitSafeMode) { return 0 }
    try {
        if ("$((Get-ToolkitSettings).ConsoleCorners)".Trim() -ieq 'Rounded') { return 6 }
    } catch { }
    return 0
}

function New-RoundedRectPath {
    <#
        A rounded rectangle as a GraphicsPath. THE CALLER DISPOSES IT.

        Width and Height lose a pixel because a Pen of width 1 straddles the
        line it is drawn on: at the full rectangle the right and bottom edges
        fall half outside the control and are clipped away, which reads as a
        border that is missing on two sides.

        Falls back to a plain rectangle when the radius will not fit rather
        than producing the pinched shape AddArc gives for a diameter larger
        than the side it is drawn on.
    #>
    param([Parameter(Mandatory)]$Rect, [int]$Radius = 6)
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $Radius * 2
    if ($Radius -le 0 -or $Rect.Width -le $d -or $Rect.Height -le $d) {
        $p.AddRectangle($Rect)
        return $p
    }
    $x = $Rect.X
    $y = $Rect.Y
    $w = $Rect.Width - 1
    $h = $Rect.Height - 1
    $p.AddArc($x, $y, $d, $d, 180, 90)
    $p.AddArc(($x + $w - $d), $y, $d, $d, 270, 90)
    $p.AddArc(($x + $w - $d), ($y + $h - $d), $d, $d, 0, 90)
    $p.AddArc($x, ($y + $h - $d), $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

function Test-RaisedButtonStyle {
    <#
        Settings.ConsoleButtonStyle = 'Raised'. Anything else is flat.

        SAFE MODE IS ALWAYS FLAT. The point of safe mode is a window drawn the
        way Windows would draw it, with as little of this toolkit's own
        rendering in the way as possible -- and owner-drawn buttons are exactly
        the kind of thing you want out of the picture when you are trying to
        work out why a console looks wrong.
    #>
    if (Test-ToolkitSafeMode) { return $false }
    try {
        $v = "$((Get-ToolkitSettings).ConsoleButtonStyle)".Trim()
        return ($v -ieq 'Raised' -or $v -ieq 'Enhanced')
    } catch { return $false }
}

function Get-RaisedButtonDepth {
    <#
        How hard the raised chrome is drawn. 1.0 is 'Raised', the original.

        WHY A MULTIPLIER AND NOT A SECOND DRAWING ROUTINE. The bevel is five
        numbers -- the gradient's two ends, the two edge colours, and the
        The bevel is four colours -- the gradient's two ends and the two edge
        tones -- plus the width of the edge itself. 'Enhanced' is those pushed
        further, not a different shape, so one factor applied at each keeps
        the two styles honest with each other: anything that gets fixed in the
        raised look is fixed in the enhanced one for free. Two drawing
        routines would have drifted apart by the second bug.

        Reported 2026-09-12 as "the 3D is VERY subtle" -- true on the darker
        palettes, where a 20% lift off the face colour is most of nothing.

        READ ONCE, WHEN THE CHROME IS ATTACHED, and parked on the button's
        Raised3D state alongside the corner radius -- for the reason spelled
        out there. It comes from Settings.json, and asking Get-ToolkitSettings
        from a paint handler stats that file every time the mouse crosses a
        button.
    #>
    if (Test-ToolkitSafeMode) { return 1.0 }
    try {
        if ("$((Get-ToolkitSettings).ConsoleButtonStyle)".Trim() -ieq 'Enhanced') { return 2.1 }
    } catch { }
    return 1.0
}

function Test-ToolGroupButton {
    <#
        Is this button a GROUP -- one console button standing for several
        tools -- rather than a single tool?

        Asked live from the Tag, the same way Get-ButtonRole asks about
        Danger, because a button's Tag is set after it is created and the
        console rebuilds its buttons on every search, arrange and refresh.
        Anything stamped on at chrome time would be stale by the first
        repaint.

        Tag is a tool object on an ordinary tool button and a PSCustomObject
        with IsToolGroup on a group one, so the property test is the whole
        question. Wrapped because Tag is frequently $null and a string on the
        buttons that are neither.
    #>
    param($Button)
    if (-not $Button) { return $false }
    try {
        $t = $Button.Tag
        if (-not $t) { return $false }
        if (-not $t.PSObject.Properties['IsToolGroup']) { return $false }
        return [bool]$t.IsToolGroup
    } catch { return $false }
}

function Get-ToolGroupOutlineColor {
    <#
        The colour of the ring that marks a group button. White.

        WHITE ALONE DOES NOT READ ON A PALE FACE, which is the whole problem
        with 'just draw it white': on the light palettes the button face is
        already near-white and the ring disappears exactly where the operator
        most needs to count on it. So the ring is always white, and on a pale
        face a one-pixel darker keyline is drawn just inside it -- which makes
        the white read as a deliberate band rather than as nothing at all.

        Returns both: the ring colour, and the keyline ($null when the face is
        dark enough that white stands on its own).
    #>
    param([Parameter(Mandatory)]$Face)
    $white = [System.Drawing.Color]::White
    $lum = 0.0
    try { $lum = Get-ColorLuminance -Color $Face } catch { }
    $key = $null
    if ($lum -ge 0.55) { $key = Get-BlendedColor -From $Face -To ([System.Drawing.Color]::Black) -Amount 0.45 }
    return [PSCustomObject]@{ Ring = $white; Key = $key }
}

function Set-ToolGroupOutline {
    <#
        Put the white ring on a group button that is NOT raised.

        Two styles, one signal. With raised chrome on, the ring is painted by
        the Paint handler -- it has to be, because the bevel fills the whole
        face and would erase a FlatAppearance border. With raised chrome off,
        there is no handler and FlatAppearance IS the edge. This applies the
        second case and stands down for the first, so callers do not have to
        know which style is in force.

        CALLED FROM TWO PLACES ON PURPOSE. Set-ConsoleTheme reaches every
        button on every form; the console ALSO calls it when it rebuilds its
        tool buttons, which happens after theming on every search, refresh and
        arrange-save. A button built after the theme pass would otherwise be
        the only group on the console without a ring.
    #>
    param($Button)
    if (-not (Test-ToolGroupButton -Button $Button)) { return }
    # The raised handler owns the edge on its buttons.
    try { if ($Button.PSObject.Properties['Raised3D']) { return } } catch { }
    # Idempotent: this is called from the theme pass AND from the console's
    # button rebuild, and attaching twice would paint twice.
    try { if ($Button.PSObject.Properties['ToolGroupRing']) { return } } catch { }

    # PAINTED, NOT SET AS A BORDER -- and the first version got this wrong.
    # FlatAppearance gives exactly ONE border colour, so setting it to white
    # threw away the keyline that makes white readable on a pale face: on
    # every light palette the ring rendered as nothing, which is the case
    # Get-ToolGroupOutlineColor exists to handle. Drawing it means both lines
    # land, and the colours are read from the button at PAINT time, so a theme
    # change or a grey-out cannot leave a stale ring behind.
    try {
        Set-FlatStyleSafely -Control $Button
        $Button.FlatAppearance.BorderSize = 0
        $Button | Add-Member -NotePropertyName 'ToolGroupRing' -NotePropertyValue $true -Force
    } catch { return }

    $Button.Add_Paint({
        param($s, $e)
        try {
            if (-not (Test-ToolGroupButton -Button $s)) { return }
            $w = $s.ClientSize.Width; $h = $s.ClientSize.Height
            if ($w -le 8 -or $h -le 8) { return }
            $gc = Get-ToolGroupOutlineColor -Face $s.BackColor
            $pG = New-Object System.Drawing.Pen($gc.Ring, 2)
            try { $e.Graphics.DrawRectangle($pG, 1, 1, ($w - 3), ($h - 3)) } finally { $pG.Dispose() }
            if ($gc.Key) {
                $pK = New-Object System.Drawing.Pen($gc.Key, 1)
                try { $e.Graphics.DrawRectangle($pK, 3, 3, ($w - 7), ($h - 7)) } finally { $pK.Dispose() }
            }
        } catch { }
    })
}

function Add-RaisedButtonChrome {
    <#
        Give a button a bevelled, raised face.

        APPLIED FROM Set-ConsoleTheme, WHICH IS WHY IT IS IN THIS FILE. Every
        form in the toolkit calls Set-ConsoleTheme -- the console, all 120-odd
        tools, and every dialog they open. Hooking the style in there is the
        only way "raised" means raised everywhere rather than on whichever
        windows somebody remembered to update.

        THE COLOURS ARE READ FROM THE BUTTON AT PAINT TIME, NOT CACHED.
        Set-ConsoleTheme repaints buttons on a theme change, tools grey and
        un-grey them as state changes, and the console resizes them on every
        window resize. A cached highlight would be correct until the first of
        those happened and wrong afterwards, in a way that reads as a rendering
        bug rather than a stale value. $s.BackColor is whatever the theme most
        recently set, so deriving from it cannot go stale -- and it means an
        accent or danger button gets a bevel in its own colour for free.

        THE CAPTION IS REDRAWN RATHER THAN SUPPRESSED. .Text has to stay set:
        search filters, group lookups and accessibility all read it. A Flat
        button paints its caption before Paint fires, so the bevel fill covers
        it and this draws it again on top.

        DISABLED IS DRAWN FLAT, NOT BEVELLED. A greyed control that still looks
        raised reads as clickable; removing the relief is the signal.

        IDEMPOTENT. Set-ConsoleTheme is called more than once on some forms --
        a tool themes itself, then a dialog it owns is themed again. Attaching
        the handlers twice would paint twice and fire every mouse handler
        twice, so the marker property is checked first.
    #>
    param([Parameter(Mandatory)]$Button)
    if (-not $Button) { return }
    if ($Button.PSObject.Properties['Raised3D']) { return }

    try {
        $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $Button.FlatAppearance.BorderSize = 0
        $Button.UseVisualStyleBackColor = $false
    } catch { return }

    # State rides on the control: a Paint handler runs from the message loop,
    # where the scope it was written in is not the scope it executes in.
    #
    # Radius is RESOLVED HERE rather than asked for at paint time. Unlike the
    # colours -- which are read live so they cannot go stale when the theme
    # repaints -- the corner setting comes from Settings.json, and reading it
    # means stat-ing that file. Once per button is free; once per repaint is
    # a disk hit every time the mouse crosses a button.
    #
    # Depth is resolved here for the SAME reason, and the first version of it
    # was not -- it asked Get-RaisedButtonDepth from inside the paint handler,
    # which is a stat of Settings.json on every hover of every button. The
    # comment directly above says not to do that; writing the setting and then
    # reading it the wrong way on the same afternoon is how a rule that is
    # already written down still gets broken.
    $Button | Add-Member -NotePropertyName 'Raised3D' -NotePropertyValue ([PSCustomObject]@{
        Pressed = $false
        Hot     = $false
        Radius  = (Get-ButtonCornerRadius)
        Depth   = (Get-RaisedButtonDepth)
    }) -Force

    $Button.Add_Paint({
        param($s, $e)
        $st = $null
        try { $st = $s.Raised3D } catch { }
        if (-not $st) { return }
        $g = $e.Graphics
        $w = $s.ClientSize.Width; $h = $s.ClientSize.Height
        if ($w -le 4 -or $h -le 4) { return }
        $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
        $face  = $s.BackColor
        $white = [System.Drawing.Color]::White
        $black = [System.Drawing.Color]::Black

        # ROUNDED CORNERS.
        #
        # THE CORNERS HAVE TO BE PAINTED OUT FIRST. WinForms draws the button's
        # BackColor across the whole client rectangle before this handler runs,
        # so filling a rounded path on top leaves four square pips of button
        # colour poking out at the corners -- which reads as a rendering fault,
        # not as a rounded button. Filling the rectangle in the PARENT's colour
        # first is what makes the curve land against the panel behind it.
        #
        # $path stays $null when the setting is off, and every use below tests
        # it. That is deliberate: with corners square the drawing is the same
        # sequence of calls it has always been, so turning the setting off
        # restores the previous rendering exactly rather than approximately.
        $path = $null
        $radius = 0
        try { $radius = [int]$st.Radius } catch { }
        if ($radius -gt 0) {
            $under = $face
            try { if ($s.Parent) { $under = $s.Parent.BackColor } } catch { }
            $ub = New-Object System.Drawing.SolidBrush($under)
            try { $g.FillRectangle($ub, $rect) } finally { $ub.Dispose() }
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $path = New-RoundedRectPath -Rect $rect -Radius $radius
        }

        # THE GLYPH IS DRAWN FROM HERE, NOT FROM Add-ButtonGlyph's OWN HANDLER.
        # The bevel fills the entire face, so anything painted before it is
        # erased -- and handler order is decided by which of the two ran first
        # on this particular button, which is not something either of them can
        # know. Add-ButtonGlyph stands down when it sees Raised3D; this draws
        # it instead, after the fill, and the two can be attached in any order.
        # $capLeft is where the caption starts and follows the same gutter the
        # flat path applies as Padding, so the two styles line up identically.
        $gl = ''
        try { $gl = "$($s.ToolGlyph)" } catch { }
        $capLeft = 4
        $capFlags = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter
        if ($gl) {
            $capLeft = (Get-ButtonGlyphMetrics -FontSize $s.Font.Size).Gutter
            # Left-aligned, because centring the caption in what is left over
            # would put a different gap after each icon and the column of
            # glyphs would be the only straight line on the panel.
            $capFlags = [System.Windows.Forms.TextFormatFlags]::Left
        }

        if (-not $s.Enabled) {
            $behind = $face
            try { if ($s.Parent) { $behind = $s.Parent.BackColor } } catch { }
            $flat = Get-BlendedColor -From $face -To $behind -Amount 0.55
            $br = New-Object System.Drawing.SolidBrush($flat)
            try {
                if ($path) { $g.FillPath($br, $path) } else { $g.FillRectangle($br, $rect) }
            } finally { $br.Dispose() }
            $pn = New-Object System.Drawing.Pen((Get-BlendedColor -From $flat -To $black -Amount 0.18), 1)
            try {
                if ($path) { $g.DrawPath($pn, $path) } else { $g.DrawRectangle($pn, 0, 0, ($w - 1), ($h - 1)) }
            } finally { $pn.Dispose() }
            # THE GROUP RING SURVIVES GREYING OUT. It says what the button IS,
            # not whether it can be clicked -- a group whose tools are all
            # blocked is still a group, and having the ring come and go with
            # tenant capability would make it useless as a thing to scan for.
            # Muted with the rest of the face so it does not out-shout an
            # enabled button beside it.
            if (Test-ToolGroupButton -Button $s) {
                $gd = Get-BlendedColor -From $flat -To ([System.Drawing.Color]::White) -Amount 0.75
                $pG = New-Object System.Drawing.Pen($gd, 2)
                try {
                    if ($path) { $g.DrawPath($pG, $path) } else { $g.DrawRectangle($pG, 1, 1, ($w - 3), ($h - 3)) }
                } finally { $pG.Dispose() }
            }
            if ($path) { try { $path.Dispose() } catch { } }
            if ($gl) { try { Write-ButtonGlyph -Graphics $g -Button $s } catch { } }
            $tr = New-Object System.Drawing.Rectangle($capLeft, 0, ($w - $capLeft - 4), $h)
            [System.Windows.Forms.TextRenderer]::DrawText($g, "$($s.Text)", $s.Font, $tr,
                [System.Drawing.Color]::FromArgb(150, 155, 165),
                ($capFlags -bor
                 [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
                 [System.Windows.Forms.TextFormatFlags]::WordBreak -bor
                 [System.Windows.Forms.TextFormatFlags]::NoPrefix))
            return
        }

        # DEPTH IS A MULTIPLIER ON EVERY NUMBER THAT MAKES THE BEVEL.
        # 1.0 is 'Raised'; 'Enhanced' pushes the same five numbers further.
        # Each is clamped below 1.0 because a blend amount of 1 is the pure
        # colour -- a white top edge and a black bottom one, which stops
        # looking like a lit surface and starts looking like a fault.
        # From the state object, resolved once at attach time -- NOT asked for
        # here. See the note where Raised3D is created.
        $depth = 1.0
        try { if ($st.PSObject.Properties['Depth']) { $depth = [double]$st.Depth } } catch { }
        if ($depth -lt 1.0) { $depth = 1.0 }
        $amt = { param([double]$a) [Math]::Min(0.92, $a * $depth) }

        $top = Get-BlendedColor -From $face -To $white -Amount (& $amt 0.20)
        $bot = Get-BlendedColor -From $face -To $black -Amount (& $amt 0.18)
        # HOVER: 0.26 toward white, up from 0.14 on 2026-09-12 -- "a slight
        # contrast difference" was the complaint, and on the darker palettes
        # 0.14 was slight enough to miss. The outline brightens with it, below,
        # so the whole tile lifts rather than just its face.
        if ($st.Hot -and -not $st.Pressed) {
            $top = Get-BlendedColor -From $top -To $white -Amount 0.26
            $bot = Get-BlendedColor -From $bot -To $white -Amount 0.26
        }
        $c1 = $(if ($st.Pressed) { $bot } else { $top })
        $c2 = $(if ($st.Pressed) { $top } else { $bot })
        $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $c1, $c2, 90.0)
        try {
            if ($path) { $g.FillPath($br, $path) } else { $g.FillRectangle($br, $rect) }
        } finally { $br.Dispose() }

        $hi = Get-BlendedColor -From $face -To $white -Amount (& $amt 0.55)
        $lo = Get-BlendedColor -From $face -To $black -Amount (& $amt 0.42)
        if ($st.Hot -and -not $st.Pressed) {
            $hi = Get-BlendedColor -From $hi -To $white -Amount 0.30
            $lo = Get-BlendedColor -From $lo -To $white -Amount 0.30
        }
        if ($st.Pressed) { $sw = $hi; $hi = $lo; $lo = $sw }

        # A THICKER EDGE AT DEPTH, BECAUSE COLOUR ALONE RUNS OUT. Past about
        # 1.5x the blends are near their clamp and pushing them further stops
        # adding anything; a second pixel of bevel is what actually reads as
        # more depth at a glance.
        $bevelW = $(if ($depth -ge 1.6) { 2 } else { 1 })

        if ($path) {
            # ONE OUTLINE RATHER THAN A TWO-TONE BEVEL. The four-line bevel
            # below says "raised" by lighting the top and left and shadowing
            # the bottom and right -- which only works on a SQUARE edge, where
            # each side is a straight line with a corner you can stop at. On a
            # curve there is no corner: the light and dark would meet mid-arc
            # and read as a seam. The gradient across the face is already
            # carrying the raised impression, and it inverts when pressed, so
            # a single quiet outline is all the edge has to do.
            $pB = New-Object System.Drawing.Pen($lo, $bevelW)
            try { $g.DrawPath($pB, $path) } finally { $pB.Dispose() }
        } else {
            $pHi = New-Object System.Drawing.Pen($hi, $bevelW)
            $pLo = New-Object System.Drawing.Pen($lo, $bevelW)
            try {
                $off = [int]($bevelW - 1)
                $g.DrawLine($pHi, $off, $off, ($w - 1 - $off), $off)
                $g.DrawLine($pHi, $off, $off, $off, ($h - 1 - $off))
                $g.DrawLine($pLo, $off, ($h - 1 - $off), ($w - 1 - $off), ($h - 1 - $off))
                $g.DrawLine($pLo, ($w - 1 - $off), $off, ($w - 1 - $off), ($h - 1 - $off))
            } finally { $pHi.Dispose(); $pLo.Dispose() }
        }

        # THE INNER HIGHLIGHT -- the line that actually says "lit from above".
        # One pixel inside the top edge, in a tone lighter than the face, and
        # its opposite just inside the bottom. Only at depth: on the plain
        # raised style it would compete with the gradient it sits on. Skipped
        # while pressed, because a pressed button is lit from the other side
        # and the gradient has already swapped to say so.
        if ($depth -ge 1.6 -and -not $st.Pressed) {
            $inHi = Get-BlendedColor -From $top -To $white -Amount 0.38
            $inLo = Get-BlendedColor -From $bot -To $black -Amount 0.22
            $y1 = $bevelW
            $y2 = $h - 1 - $bevelW
            $x1 = $bevelW + $(if ($path) { 5 } else { 1 })
            $x2 = $w - 1 - $bevelW - $(if ($path) { 5 } else { 1 })
            if ($x2 -gt $x1 -and $y2 -gt $y1) {
                $pIn = New-Object System.Drawing.Pen($inHi, 1)
                try { $g.DrawLine($pIn, $x1, $y1, $x2, $y1) } finally { $pIn.Dispose() }
                $pIn = New-Object System.Drawing.Pen($inLo, 1)
                try { $g.DrawLine($pIn, $x1, $y2, $x2, $y2) } finally { $pIn.Dispose() }
            }
        }

        # THE DANGER OUTLINE, REDRAWN BECAUSE THE BEVEL ERASED IT.
        #
        # Set-ConsoleTheme gives a destructive button a red FlatAppearance
        # border so the hand pauses before the confirm dialog ever appears --
        # but FlatAppearance is painted by WinForms underneath this handler,
        # and BorderSize is set to 0 above so the bevel can own the edge. The
        # result was that switching the toolkit to raised buttons silently
        # removed a SAFETY signal from every Delete and Purge in it.
        #
        # Asked live rather than stamped on the control, because Get-ButtonRole
        # reads Tag and Text and both can change after theming -- a button
        # relabelled from 'Save' to 'Delete' must become dangerous with it.
        $role = 'Plain'
        try { $role = Get-ButtonRole -Button $s } catch { }
        if ($role -eq 'Danger') {
            # ForeColor is where the theme put the danger colour, so this
            # cannot disagree with the caption it is drawn around.
            $pD = New-Object System.Drawing.Pen($s.ForeColor, 1)
            try {
                if ($path) { $g.DrawPath($pD, $path) } else { $g.DrawRectangle($pD, 0, 0, ($w - 1), ($h - 1)) }
            } finally { $pD.Dispose() }
        }
        # THE GROUP RING. A button that opens several tools instead of running
        # one is the only button on the console whose click does something
        # structurally different, and nothing about its face said so -- the
        # caption is just a word, and a group called 'Calendar' looked exactly
        # like a tool called 'Calendar'. Two pixels of white around the edge,
        # drawn last so the bevel cannot erase it, in the same place and the
        # same way as the danger outline above.
        if (Test-ToolGroupButton -Button $s) {
            $gc = Get-ToolGroupOutlineColor -Face $face
            $pG = New-Object System.Drawing.Pen($gc.Ring, 2)
            try {
                if ($path) { $g.DrawPath($pG, $path) } else { $g.DrawRectangle($pG, 1, 1, ($w - 3), ($h - 3)) }
            } finally { $pG.Dispose() }
            if ($gc.Key -and -not $path -and $w -gt 8 -and $h -gt 8) {
                $pK = New-Object System.Drawing.Pen($gc.Key, 1)
                try { $g.DrawRectangle($pK, 3, 3, ($w - 7), ($h - 7)) } finally { $pK.Dispose() }
            }
        }

        # Last use of the path. A GraphicsPath is a GDI+ handle and this runs on
        # every invalidate -- every hover, every resize, every grey-out -- so
        # leaving it to the GC leaks one per repaint.
        if ($path) { try { $path.Dispose() } catch { }; $path = $null }

        $off = $(if ($st.Pressed) { 1 } else { 0 })
        if ($gl) { try { Write-ButtonGlyph -Graphics $g -Button $s -Pressed $st.Pressed } catch { } }
        $tr = New-Object System.Drawing.Rectangle(($capLeft + $off), $off, ($w - $capLeft - 4), $h)
        [System.Windows.Forms.TextRenderer]::DrawText($g, "$($s.Text)", $s.Font, $tr, $s.ForeColor,
            ($capFlags -bor
             [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
             [System.Windows.Forms.TextFormatFlags]::WordBreak -bor
             [System.Windows.Forms.TextFormatFlags]::NoPrefix))

        if ($s.Focused) {
            $fr = New-Object System.Drawing.Rectangle(3, 3, ($w - 7), ($h - 7))
            [System.Windows.Forms.ControlPaint]::DrawFocusRectangle($g, $fr)
        }
    })

    # These only set a paint flag and never swallow the event, so drag
    # handlers that also own MouseDown -- the console's Arrange mode -- keep
    # working exactly as they did.
    $Button.Add_MouseDown({  param($s, $e) $s.Raised3D.Pressed = $true;  $s.Invalidate() })
    $Button.Add_MouseUp({    param($s, $e) $s.Raised3D.Pressed = $false; $s.Invalidate() })
    $Button.Add_MouseEnter({ $this.Raised3D.Hot = $true;  $this.Invalidate() })
    $Button.Add_MouseLeave({ $this.Raised3D.Hot = $false; $this.Raised3D.Pressed = $false; $this.Invalidate() })
    $Button.Add_EnabledChanged({ $this.Invalidate() })
    $Button.Add_GotFocus({  $this.Invalidate() })
    $Button.Add_LostFocus({ $this.Invalidate() })
}

function Set-FormWindowChrome {
    <#
        The parts of a window the theme could not reach: its title bar and its
        corners. Both are drawn by the Desktop Window Manager, not by WinForms,
        and both are set through one Win32 call on the window handle.

        DARK TITLE BAR. A dark theme with a white title bar sat on top of it
        looks broken, and there was nothing in .NET to fix it -- WinForms has no
        property for it. DwmSetWindowAttribute with USE_IMMERSIVE_DARK_MODE
        turns the caption dark. Whether to ask for it is decided from the
        theme's own FormBack: below half luminance is dark. The attribute number
        changed once (19 on early Windows 10 builds, 20 from 20H1), so both are
        tried and neither failing is an error.

        ROUNDED CORNERS. Windows 11 already rounds a standard-frame window; the
        WINDOW_CORNER_PREFERENCE call asks for it explicitly so a window that
        something has un-rounded gets it back, and it is a silent no-op on
        Windows 10, which has no such thing.

        RUNS ON HandleCreated, NOT NOW. Set-ConsoleTheme is called before
        ShowDialog, when the form has no window handle yet, and the DWM call
        needs one. Attaching to the event means it fires at exactly the right
        moment for every form that calls Set-ConsoleTheme -- which is all of
        them -- without any of them having to know.

        SAFE MODE LEAVES THE WINDOW ALONE, like every other piece of the
        toolkit's own drawing. Every step is guarded: this is cosmetic, and a
        machine where dwmapi is unavailable must still open the window.
    #>
    param([Parameter(Mandatory)]$Form, $FormBack)
    if (-not ($Form -is [System.Windows.Forms.Form])) { return }
    if (Test-ToolkitSafeMode) { return }
    if ($Form.PSObject.Properties['WindowChromeHooked']) { return }
    $Form | Add-Member -NotePropertyName 'WindowChromeHooked' -NotePropertyValue $true -Force

    $dark = $false
    try { if ($FormBack) { $dark = ((Get-ColorLuminance -Color $FormBack) -lt 0.5) } } catch { }

    # The P/Invoke type is compiled once per process. Add-Type throws if the
    # type already exists, so it is checked for first rather than caught.
    try {
        if (-not ([System.Management.Automation.PSTypeName]'ToolkitNative.Dwm').Type) {
            Add-Type -Namespace ToolkitNative -Name Dwm -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("dwmapi.dll")]
public static extern int DwmSetWindowAttribute(System.IntPtr hwnd, int attr, ref int value, int size);
'@ -ErrorAction Stop
        }
    } catch { return }

    $Form | Add-Member -NotePropertyName 'WindowChromeDark' -NotePropertyValue $dark -Force
    $Form.Add_HandleCreated({
        try {
            $h = $this.Handle
            $on = $(if ($this.WindowChromeDark) { 1 } else { 0 })
            # 20 = DWMWA_USE_IMMERSIVE_DARK_MODE (20H1+), 19 = the pre-20H1 number.
            foreach ($attr in @(20, 19)) {
                try { [void][ToolkitNative.Dwm]::DwmSetWindowAttribute($h, $attr, [ref]$on, 4); break } catch { }
            }
            # 33 = DWMWA_WINDOW_CORNER_PREFERENCE; 2 = DWMWCP_ROUND. Windows 10
            # rejects the attribute, which is caught and ignored.
            $round = 2
            try { [void][ToolkitNative.Dwm]::DwmSetWindowAttribute($h, 33, [ref]$round, 4) } catch { }
        } catch { }
    })
}

function Set-ConsoleTheme {
    param(
        [Parameter(Mandatory)]$Form,
        $Theme,
        [switch]$NoGradient
    )
    if (-not $Theme) { $Theme = Get-ConsoleTheme }
    # No theme (the disabled 'Default') still gets the typeface, DPI scaling
    # and button roles -- those are not colours. Every colour lookup below
    # comes back empty from this and is skipped.
    if (-not $Theme) { $Theme = [PSCustomObject]@{ Name = 'Default' } }

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

    # Typography and accent roles. Theme keys are optional; the defaults are
    # what an unedited Themes.json gets. FontFamily 'Microsoft Sans Serif'
    # is the opt-out for a tenant that wants the old face back.
    $uiFamily   = "$($Theme.FontFamily)".Trim();  if (-not $uiFamily)   { $uiFamily   = 'Segoe UI' }
    $monoFamily = "$($Theme.MonoFamily)".Trim();  if (-not $monoFamily) { $monoFamily = 'Consolas' }
    $accentBack = ConvertTo-ThemeColor $Theme.AccentBack;  if (-not $accentBack)  { $accentBack  = [System.Drawing.Color]::FromArgb(46, 107, 214) }
    $accentText = ConvertTo-ThemeColor $Theme.AccentText;  if (-not $accentText)  { $accentText  = [System.Drawing.Color]::White }
    $accentHover = ConvertTo-ThemeColor $Theme.AccentHover; if (-not $accentHover) { $accentHover = [System.Drawing.Color]::FromArgb(63, 125, 232) }
    $danger     = ConvertTo-ThemeColor $Theme.DangerColor; if (-not $danger)      { $danger      = [System.Drawing.Color]::FromArgb(214, 69, 69) }

    # DPI. With the process DPI-aware (Enable-ToolkitDpiAwareness at the
    # bottom of this file) a form designed at 96 dpi must say so, or WinForms
    # assumes it was designed at the CURRENT dpi and scales nothing -- crisp
    # but tiny at 150 %. Both properties must be set before the handle
    # exists, and every tool calls this before ShowDialog.
    if ($Form -is [System.Windows.Forms.Form] -and -not $Form.IsHandleCreated) {
        try {
            $Form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
            $Form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
        } catch { }
    }
    Set-ThemeFontOn -Control $Form -UiFamily $uiFamily -MonoFamily $monoFamily

    # Title bar and corners. Hooked here so every themed window gets them;
    # it fires on HandleCreated because the DWM call needs a real window.
    Set-FormWindowChrome -Form $Form -FormBack $formBack

    if ($formBack) { $Form.BackColor = $formBack }
    if ($formText) { $Form.ForeColor = $formText }

    if (-not $NoGradient -and $Theme.GradientTo -and $Theme.FormBack) {
        Add-FormGradient -Form $Form -FromHex $Theme.FormBack -ToHex $Theme.GradientTo
    }

    # Asked once for the whole form rather than per control: it reads a cached
    # settings object, but the intent is that one window has one button style.
    $raised = Test-RaisedButtonStyle
    $icons  = Test-ConsoleIconStyle

    $apply = {
        param($parent)
        foreach ($c in $parent.Controls) {
            # Tag-based opt-out. Controls whose colour carries MEANING -- the
            # ExO/SPO connection state, the ServiceNow highlight -- set
            # Tag = 'NoTheme' and are skipped, so a theme can never disguise a
            # red "not connected" as decoration.
            # The face is applied even to NoTheme controls: the opt-out is
            # about COLOUR carrying meaning, and a connection pill in a
            # different typeface from the rest of the window is not meaning.
            Set-ThemeFontOn -Control $c -UiFamily $uiFamily -MonoFamily $monoFamily
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
                    # AFTER the colours, so the bevel derives from the face the
                    # theme just chose. It reads BackColor live at paint time,
                    # so the order is not load-bearing -- but a reader should
                    # not have to work that out.
                    if ($raised) { Add-RaisedButtonChrome -Button $c }
                    # ROLE. One filled accent button per form is the convention
                    # every product the operators use follows; the rest are
                    # quiet, and the destructive one is outlined in red so the
                    # hand pauses before the confirm dialog ever appears.
                    switch (Get-ButtonRole -Button $c) {
                        'Primary' {
                            Set-FlatStyleSafely -Control $c
                            $c.BackColor = $accentBack; $c.ForeColor = $accentText
                            $c.FlatAppearance.BorderColor = $accentBack
                            $c.FlatAppearance.MouseOverBackColor = $accentHover
                            $c.FlatAppearance.MouseDownBackColor = $accentBack
                            $c.Font = Get-ThemeFont -Family $uiFamily -Size $c.Font.Size -Style ([System.Drawing.FontStyle]::Bold)
                        }
                        'Danger' {
                            Set-FlatStyleSafely -Control $c
                            if ($btnBack) { $c.BackColor = $btnBack }
                            $c.ForeColor = $danger
                            $c.FlatAppearance.BorderColor = $danger
                            $bg = if ($btnBack) { $btnBack } else { $c.BackColor }
                            $c.FlatAppearance.MouseOverBackColor = (Get-BlendedColor -From $bg -To $danger -Amount 0.25)
                            $c.FlatAppearance.MouseDownBackColor = (Get-BlendedColor -From $bg -To $danger -Amount 0.45)
                        }
                    }
                    # THE GROUP RING, FLAT-BUTTON EDITION. Same signal as the
                    # one the raised chrome paints: a button that opens several
                    # tools is outlined in white. Done here rather than in the
                    # role switch because a group is not a role -- it says what
                    # the button CONTAINS, not how dangerous it is, and a group
                    # of destructive tools should be able to be both.
                    Set-ToolGroupOutline -Button $c
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
                    # LAST, so the fit test measures the face the theme just
                    # applied rather than whatever the form was built with --
                    # Segoe is wider than the default, and a caption measured
                    # in the narrower one would be accepted and then clipped.
                    #
                    # NOT -Fallback here. A dialog button only gets an icon if
                    # its caption genuinely says what it does; the console asks
                    # for the fallback itself, because a grid of tiles is the
                    # one place where a gap looks broken rather than restrained.
                    if ($icons) { Add-ButtonGlyph -Button $c }
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

                    # WIRED ONCE PER CONTROL, AND THE GUARD IS NOT OPTIONAL.
                    # Some forms are themed more than once -- a tool themes
                    # itself, then a dialog it owns is themed again. Running
                    # this twice attached a second CheckedChanged handler, and
                    # worse, re-captured '_FontRegular' from $c.Font -- which
                    # on an already-checked control is the BOLD font this code
                    # just set. The control could then never go back to
                    # regular, and the symptom was a radio group where every
                    # option you had ever selected stayed bold.
                    if (-not $c.PSObject.Properties['_FontRegular']) {

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
                    }

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

                    # TABLE, NOT SPREADSHEET. Horizontal rules only, a flat bold
                    # header band, 24px rows, no outer border -- the DataGridView
                    # defaults (both gridlines, raised 3D headers, 22px rows) are
                    # the second loudest tell after the font.
                    try {
                        $c.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
                        $c.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
                        $c.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
                        $c.ColumnHeadersHeight = 28
                        $c.RowTemplate.Height = 24
                        $c.AllowUserToResizeRows = $false
                        $c.BorderStyle = [System.Windows.Forms.BorderStyle]::None
                        $c.ColumnHeadersDefaultCellStyle.Font = Get-ThemeFont -Family $uiFamily -Size $c.Font.Size -Style ([System.Drawing.FontStyle]::Bold)
                        $c.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4, 0, 4, 0)
                    } catch { }
                    # Numbers right-align themselves. Decided per cell at paint
                    # time because the theme runs before any row exists.
                    if (-not $c.PSObject.Properties['_NumAlign']) {
                        Add-Member -InputObject $c -NotePropertyName '_NumAlign' -NotePropertyValue $true -Force
                        $c.Add_CellFormatting({
                            param($s, $e)
                            $v = $e.Value
                            if ($null -eq $v) { return }
                            if (($v -is [int]) -or ($v -is [long]) -or ($v -is [double]) -or ($v -is [decimal]) -or ($v -is [single]) -or
                                ("$v" -match '^\s*-?[\d,]+(\.\d+)?\s*(%|GB|MB|KB|B)?\s*$')) {
                                $e.CellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
                            }
                        })
                    }

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

    # LAST, after the walk, so the walk cannot recolour it.
    try { Add-TenantBand -Form $Form } catch { }
    # And never taller than the screen it is about to open on.
    try { Limit-FormToScreen -Form $Form } catch { }
}

function Limit-FormToScreen {
    <#
        A form taller or wider than the working area is clamped to it and
        given a scrollbar, so the buttons along its bottom edge stay
        reachable. Called from Set-ConsoleTheme, which every form passes
        through, so no tool has to remember.

        WHY HERE AND NOT PER TOOL. Set-FormScreenCorner has done this for
        years -- for the tools that call it. Park In Place grew to 854px on
        2026-09-15 and did not, and an FH operator works in an 844px session
        (ARCHITECTURE 57c): Save would have sat below the bottom of his
        screen with nothing to scroll. Every form built by absolute pixel
        coordinates has this failure waiting; one clamp in the shared path
        removes it from all of them at once.

        Idempotent, and a no-op for anything already small enough.
        AutoScroll is only switched on when something had to shrink, so a
        form that fits keeps its designed, scrollbar-free look.
    #>
    param([Parameter(Mandatory)]$Form)
    if (-not ($Form -is [System.Windows.Forms.Form])) { return }
    $wa = $null
    try { $wa = [System.Windows.Forms.Screen]::FromControl($Form).WorkingArea } catch { }
    if (-not $wa -or $wa.Height -le 0) {
        try { $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea } catch { return }
    }
    $shrunk = $false
    if ($Form.Height -gt $wa.Height) { $Form.Height = $wa.Height; $shrunk = $true }
    if ($Form.Width  -gt $wa.Width)  { $Form.Width  = $wa.Width;  $shrunk = $true }
    if ($shrunk) { $Form.AutoScroll = $true }
}

#region Tenant band --------------------------------------------------------------
<#
    EVERY WINDOW SAYS WHICH TENANT IT IS ON. Not just the console: every tool
    form is a separate window, and with two consoles open side by side a
    Disable-User form for one tenant looks exactly like the other's. The band
    is a strip of colour along the top of every form, one colour per tenant,
    plus the tenant name on the title bar. It is applied from Set-ConsoleTheme,
    which every form already calls, so no tool had to change.

    The colour comes from the tenant file's 'Band' key (hex). With no key, a
    colour is picked from a fixed palette by hashing the tenant key, so two
    tenants that never set one still differ. Single-tenant installs get no
    band: there is nothing to tell apart.
#>
$script:TenantBandPalette = @('#D64545', '#2E6BD6', '#3D9970', '#C4772E', '#7B61C4', '#1FA3A3', '#B5407A', '#8A8A2E')

function Get-TenantBandColor {
    <# The band colour for the active tenant, or $null on a single-tenant install. #>
    $k = Get-ActiveTenantKey
    if (-not $k) { return $null }
    $hex = ''
    try {
        $cfg = Get-TenantConfig
        if ($cfg -and $cfg.PSObject.Properties['Band']) { $hex = "$($cfg.Band)".Trim() }
    } catch { }
    if ($hex -match '^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$') {
        $c = ConvertTo-ThemeColor $hex
        if ($c) { return $c }
    }
    $h = 0
    foreach ($ch in $k.ToLowerInvariant().ToCharArray()) { $h = (($h * 31) + [int]$ch) -band 0x7FFFFFFF }
    return (ConvertTo-ThemeColor $script:TenantBandPalette[$h % $script:TenantBandPalette.Count])
}

function Add-TenantBand {
    <#
        Put the tenant band on a form. Idempotent: a second call restyles the
        band it already added. Safe on any control that is not a Form.
    #>
    param([Parameter(Mandatory)]$Form)
    if (-not ($Form -is [System.Windows.Forms.Form])) { return }
    $color = Get-TenantBandColor
    if (-not $color) { return }
    $label = Get-ActiveTenantLabel

    $band = $null
    foreach ($c in $Form.Controls) { if ($c.Name -eq 'pnlTenantBand') { $band = $c; break } }
    if (-not $band) {
        $band = New-Object System.Windows.Forms.Panel
        $band.Name     = 'pnlTenantBand'
        $band.Tag      = 'NoTheme'
        $band.Location = New-Object System.Drawing.Point(0, 0)
        $band.Size     = New-Object System.Drawing.Size($Form.ClientSize.Width, 6)
        $band.Anchor   = 'Top,Left,Right'
        $Form.Controls.Add($band)
        $band.BringToFront()
        try {
            $tt = New-Object System.Windows.Forms.ToolTip
            $tt.SetToolTip($band, "Tenant: $label")
        } catch { }
    }
    $band.BackColor = $color

    # The title bar carries the name too, for the taskbar and Alt+Tab, where a
    # strip of colour is not visible. Only once, however many times this runs.
    if ($label -and ("$($Form.Text)".IndexOf($label, [System.StringComparison]::OrdinalIgnoreCase) -lt 0)) {
        $Form.Text = "$($Form.Text)  [$label]"
    }
}
#endregion

#region Message dialogs ---------------------------------------------------------
<#
    THE ONE PART OF THIS UI THE THEME COULD NEVER REACH.

    Show-ToolDialog is drawn by the operating system.
    Set-ConsoleTheme walks a form's controls; a MessageBox has no controls to
    walk, is not a Form, and does not belong to this process's styling at all.
    So every careful thing this toolkit does -- the typeface, the palette, the
    button roles, the raised chrome -- stopped at the edge of the window and
    the next thing the operator saw was a grey box in the system face. That is
    the most dated surface left in the toolkit and it is in front of somebody
    two hundred times a day.

    Show-ToolDialog TAKES MESSAGEBOX'S ARGUMENTS, IN MESSAGEBOX'S ORDER, and
    that is the entire reason the swap across eighty-odd files was a textual
    one rather than eighty-odd hand edits. A call reads identically before and
    after, every argument shape MessageBox accepts still works, and reversing
    it is the same substitution run the other way.

    IT FALLS BACK TO THE REAL MESSAGEBOX and it does so on four separate
    grounds: the setting is off, safe mode is on, the themed form threw while
    being built, or the themed form threw while being shown. A confirmation
    prompt is the last place to be clever -- if anything at all goes wrong
    here the operator must still get asked the question, because the
    alternative is a destructive action taken without one.
#>

function Test-ConsoleDialogStyle {
    <# Settings.ConsoleDialogStyle = 'Themed'. Safe mode is always the system box. #>
    if (Test-ToolkitSafeMode) { return $false }
    try { return ("$((Get-ToolkitSettings).ConsoleDialogStyle)".Trim() -ieq 'Themed') } catch { return $false }
}

function Write-ToolDialogIcon {
    <#
        The severity icon, drawn rather than loaded.

        THE COLOURS ARE FIXED AND ARE NOT THEMED. They are the same rule the
        connection pills follow: a colour that carries MEANING must not be
        something a palette can redecorate. An amber warning that a dark theme
        turned blue would be worse than no icon.

        Two ways of drawing it, in order. The Segoe icon glyph is the modern
        one and is used whenever Windows has the font. SystemIcons is the
        floor -- it is in .NET itself, so it cannot be missing, and a slightly
        old-looking triangle beats an empty square where one should be.
    #>
    param(
        [Parameter(Mandatory)]$Graphics,
        [string]$Kind = 'None',
        [int]$X = 24,
        [int]$Y = 24,
        [int]$Size = 34
    )
    if ($Kind -eq 'None') { return }

    $code = 0
    $col  = [System.Drawing.Color]::FromArgb(46, 107, 214)
    switch ($Kind) {
        'Warning'     { $code = 0xE7BA; $col = [System.Drawing.Color]::FromArgb(202, 138, 4) }
        'Error'       { $code = 0xF13D; $col = [System.Drawing.Color]::FromArgb(196, 43, 43) }
        'Question'    { $code = 0xE897; $col = [System.Drawing.Color]::FromArgb(46, 107, 214) }
        'Information' { $code = 0xE946; $col = [System.Drawing.Color]::FromArgb(46, 107, 214) }
        default       { return }
    }

    $f = Get-ToolkitIconFont -Size 21
    if ($f) {
        $r = New-Object System.Drawing.Rectangle($X, $Y, $Size, $Size)
        [System.Windows.Forms.TextRenderer]::DrawText($Graphics, ([string][char]$code), $f, $r, $col,
            ([System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor
             [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
             [System.Windows.Forms.TextFormatFlags]::NoPrefix))
        return
    }

    $ico = $null
    switch ($Kind) {
        'Warning'     { $ico = [System.Drawing.SystemIcons]::Warning }
        'Error'       { $ico = [System.Drawing.SystemIcons]::Error }
        'Question'    { $ico = [System.Drawing.SystemIcons]::Question }
        'Information' { $ico = [System.Drawing.SystemIcons]::Information }
    }
    if ($ico) {
        try { $Graphics.DrawIcon($ico, (New-Object System.Drawing.Rectangle($X, $Y, $Size, $Size))) } catch { }
    }
}

function Show-ThemedMessageDialog {
    <#
        A message box built out of a Form, so the theme can reach it.

        SIZED FROM THE MESSAGE RATHER THAN FIXED. These carry everything from
        'Saved.' to a forty-line list of what a batch could not process, and a
        fixed size is either enormous for the first or a scrollbar for the
        second. The text is measured wrapped at a comfortable reading width and
        the window is built around the answer.

        THE FOOTER IS A PANEL, and it is what makes this read as 2026 rather
        than 2005. Set-ConsoleTheme paints a Panel in the theme's panel colour,
        so the buttons sit on a band a shade off the body -- the command bar
        every product the operators use all day has along its bottom edge. It
        costs one control.

        KEYBOARD BEHAVIOUR MATCHES THE SYSTEM BOX, deliberately, because muscle
        memory is the thing most likely to cause an accident here. Enter takes
        the default; Escape takes Cancel, or No, or Ignore; and where there is
        no such answer the close box is REMOVED rather than left to return
        something the caller never offered.

        Returns a DialogResult, exactly as MessageBox does.
    #>
    param(
        [string]$Text = '',
        [string]$Caption = '',
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::None,
        [System.Windows.Forms.MessageBoxDefaultButton]$DefaultButton = [System.Windows.Forms.MessageBoxDefaultButton]::Button1
    )

    # MessageBoxIcon has several names per value -- Hand, Stop and Error are
    # all 16; Asterisk and Information are both 64 -- so the numeric value is
    # the honest thing to switch on.
    $kind = 'None'
    switch ([int]$Icon) {
        16 { $kind = 'Error' }
        32 { $kind = 'Question' }
        48 { $kind = 'Warning' }
        64 { $kind = 'Information' }
    }

    $spec = @('OK')
    switch ("$Buttons") {
        'OKCancel'         { $spec = @('OK', 'Cancel') }
        'YesNo'            { $spec = @('Yes', 'No') }
        'YesNoCancel'      { $spec = @('Yes', 'No', 'Cancel') }
        'RetryCancel'      { $spec = @('Retry', 'Cancel') }
        'AbortRetryIgnore' { $spec = @('Abort', 'Retry', 'Ignore') }
    }

    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = "$Caption"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.ShowInTaskbar   = $false

    $bodyFont = Get-ThemeUiFont -Size 9
    $padL = 24
    $iconW = $(if ($kind -eq 'None') { 0 } else { 52 })
    $textX = $padL + $iconW
    # 460px is about ninety characters at this size, which is a comfortable
    # measure. Narrower makes long messages tall; wider makes them hard to
    # track back to the start of the next line.
    $wrapAt = 460

    # THE LABEL SIZES ITSELF. TWO ATTEMPTS AT MEASURING IT BY HAND FAILED.
    #
    # The first version measured with TextRenderer before the theme ran and
    # set a fixed label height; the theme then changed the typeface and the
    # text outgrew the label. The second re-measured after the theme and again
    # at Shown, and the last line was STILL clipped -- TextRenderer.MeasureText
    # and the Label's own renderer do not agree about text carrying embedded
    # newlines, which every one of these messages has.
    #
    # AutoSize with a MaximumSize width is the supported way to ask a Label
    # how tall it really needs to be: WinForms wraps at that width with the
    # renderer it is actually going to draw with, and reports the answer in
    # .Height. No second opinion, nothing to keep in step. The window is then
    # sized FROM the label instead of the label being sized from a guess --
    # which is the same lesson as Set-ConsoleFindRowLayout and the initial
    # fit: place from what the controls measure, never from what they were
    # created as.
    $footerH = 58
    $lbl             = New-Object System.Windows.Forms.Label
    $lbl.AutoSize    = $true
    $lbl.MaximumSize = New-Object System.Drawing.Size($wrapAt, 0)
    $lbl.MinimumSize = New-Object System.Drawing.Size(220, 34)
    $lbl.Text        = "$Text"
    $lbl.Font        = $bodyFont
    $lbl.Location    = New-Object System.Drawing.Point($textX, 22)
    # UseMnemonic off, or an address containing '&' loses it and underlines the
    # next letter -- and these messages quote real addresses and real paths.
    $lbl.UseMnemonic = $false
    $lbl.Name = 'ToolDialogBody'
    $dlg.Controls.Add($lbl)

    # A first pass so the form has a sane size before the theme runs. It is
    # recomputed after the theme and again at Shown, from the label.
    $clientW = [Math]::Max(($textX + $lbl.Width + $padL), 400)
    $clientH = 22 + [Math]::Max($lbl.Height, $(if ($kind -eq 'None') { 0 } else { 34 })) + 20 + $footerH
    $dlg.ClientSize = New-Object System.Drawing.Size($clientW, $clientH)

    # A MESSAGE BOX IS NOT A ZOOMABLE WINDOW. Set-ConsoleTheme wires
    # Enable-FormAutoScale into every form, which makes the form Sizable and
    # captures a scaling baseline at Shown -- neither belongs on a FixedDialog
    # three lines tall, and the baseline fought the re-measure below. The
    # opt-out marker rides on Tag with the icon kind; the Paint handler takes
    # the kind from the part before the bar.
    $dlg.Tag = 'NoAutoScale'

    $footer          = New-Object System.Windows.Forms.Panel
    $footer.Height   = $footerH
    $footer.Dock     = [System.Windows.Forms.DockStyle]::Bottom
    $dlg.Controls.Add($footer)

    # PLACED right to left so each button's x can be worked out from the one
    # after it, but the ORDER is the Windows one: the affirmative answer on the
    # left of the group, Cancel on the right. $made is rebuilt front-to-back as
    # it goes, so it reads in visual order for the default and escape lookups
    # below -- index 0 is the leftmost button, which is the one that commits.
    $made = @()
    $x = $clientW - $padL
    for ($i = $spec.Count - 1; $i -ge 0; $i--) {
        $b          = New-Object System.Windows.Forms.Button
        $b.Text     = $spec[$i]
        $b.Font     = $bodyFont
        $b.Size     = New-Object System.Drawing.Size(104, 32)
        $x -= $b.Width
        $b.Location = New-Object System.Drawing.Point($x, 13)
        # ANCHORED RIGHT. These x positions come from $clientW as it was when
        # the buttons were made, and the window is re-fitted twice after that
        # -- once when the theme has chosen the typeface and once at Shown.
        # Without the anchor the buttons stay where a narrower window put them
        # and the group drifts away from the right edge.
        $b.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        $b.DialogResult = [System.Windows.Forms.DialogResult]$spec[$i]
        # Tag is how Get-ButtonRole is told what a button MEANS rather than
        # left to guess from its caption. 'Yes' on a delete prompt is the
        # destructive one and 'Yes' on a save prompt is not, and the caller
        # said which by choosing the icon.
        if ($i -eq 0 -and $kind -ne 'Warning' -and $kind -ne 'Error') { $b.Tag = 'Primary' }
        $footer.Controls.Add($b)
        $made = @($b) + $made
        $x -= 8
    }

    $defIdx = 0
    switch ("$DefaultButton") {
        'Button2' { $defIdx = 1 }
        'Button3' { $defIdx = 2 }
    }
    if ($defIdx -ge $made.Count) { $defIdx = $made.Count - 1 }

    # 'Ignore' IS DELIBERATELY NOT AN ESCAPE ANSWER, even though it is the last
    # button on an AbortRetryIgnore box. The system box disables its close
    # button for that set and swallows Escape -- and it is right to: on a batch
    # prompt, turning a stray Escape into "skip this one" is a decision the
    # operator did not make about a record they did not look at.
    $escIdx = -1
    for ($i = $made.Count - 1; $i -ge 0; $i--) {
        if (@('Cancel', 'No') -contains $made[$i].Text) { $escIdx = $i; break }
    }
    if ($escIdx -lt 0 -and $made.Count -eq 1) { $escIdx = 0 }

    $dlg.AcceptButton = $made[$defIdx]
    if ($escIdx -ge 0) {
        $dlg.CancelButton = $made[$escIdx]
        $dlg.ControlBox   = $true
    } else {
        # No escape answer exists, so there must be no way to ask for one. The
        # system box greys its close button in exactly this case.
        $dlg.ControlBox = $false
    }

    $owner = $null
    try { $owner = [System.Windows.Forms.Form]::ActiveForm } catch { }
    $dlg.StartPosition = $(if ($owner) {
        [System.Windows.Forms.FormStartPosition]::CenterParent
    } else {
        [System.Windows.Forms.FormStartPosition]::CenterScreen
    })

    Set-ConsoleTheme -Form $dlg

    # THE WINDOW, FROM THE LABEL, NOW THAT THE THEME HAS PICKED THE TYPEFACE.
    # The label has already re-wrapped itself to the new font; all this does
    # is read its height and give the window room for it plus the footer.
    try {
        $w = [Math]::Max(($textX + $lbl.Width + $padL), 400)
        $h = 22 + [Math]::Max($lbl.Height, $(if ($kind -eq 'None') { 0 } else { 34 })) + 20 + $footerH
        $dlg.ClientSize = New-Object System.Drawing.Size($w, $h)
    } catch { }

    # THE ICON IS HOOKED AFTER THE THEME, AND THAT ORDER IS LOAD-BEARING.
    # Set-ConsoleTheme adds the gradient's own Paint handler, handlers run in
    # the order they were attached, and an icon drawn before the gradient is
    # an icon painted over by it.
    $dlg.Add_Paint({
        param($s, $e)
        try { Write-ToolDialogIcon -Graphics $e.Graphics -Kind (("$($s.Tag)" -split '\|')[0]) -X 24 -Y 22 } catch { }
    })
    $dlg.Tag = "$kind|NoAutoScale"

    # Focus the default rather than the first control, so Enter and Space both
    # do what the operator expects the instant the window appears.
    #
    # AND FIT THE TEXT, HERE, WITH THE WINDOW REAL. The measurement before the
    # theme and the re-measure after it both ran before the window had a
    # handle, and DPI scaling (AutoScaleMode = Dpi, applied at handle
    # creation) changed the numbers again -- the last line stayed clipped
    # behind the footer. At Shown every size is final. Everything is reached
    # through $this: this handler runs after the function's locals are gone.
    $dlg.Add_Shown({
        # ONE LAST FIT, WITH THE WINDOW REAL. DPI scaling is applied at handle
        # creation, which is after everything above; the label has re-wrapped
        # itself by now and this reads the final answer. Grows only -- a
        # window that shrinks as it appears looks like a fault.
        try {
            $body = $null
            foreach ($c in $this.Controls) { if ($c.Name -eq 'ToolDialogBody') { $body = $c; break } }
            $foot = $null
            foreach ($c in $this.Controls) {
                if ($c -is [System.Windows.Forms.Panel] -and $c.Dock -eq [System.Windows.Forms.DockStyle]::Bottom) { $foot = $c; break }
            }
            if ($body) {
                $needH = $body.Bottom + 20 + $(if ($foot) { $foot.Height } else { 58 })
                $needW = $body.Right + 24
                $w = [Math]::Max($this.ClientSize.Width, $needW)
                $h = [Math]::Max($this.ClientSize.Height, $needH)
                if ($w -ne $this.ClientSize.Width -or $h -ne $this.ClientSize.Height) {
                    $this.ClientSize = New-Object System.Drawing.Size($w, $h)
                }
            }
        } catch { }
        try { $this.AcceptButton.Focus() } catch { }
    })

    try {
        if ($owner) { [void]$dlg.ShowDialog($owner) } else { [void]$dlg.ShowDialog() }
        $res = $dlg.DialogResult
        # Alt+F4 closes without a DialogResult. Answer as the escape button
        # would have, or as OK where OK was the only thing on offer.
        if ("$res" -eq 'None') {
            $res = $(if ($escIdx -ge 0) {
                [System.Windows.Forms.DialogResult]$made[$escIdx].Text
            } else {
                [System.Windows.Forms.DialogResult]::Cancel
            })
        }
        return $res
    } finally {
        try { $dlg.Dispose() } catch { }
    }
}

function Show-ToolDialog {
    <#
        A drop-in for Show-ToolDialog.

        THE ARGUMENTS ARE POSITIONAL AND UNNAMED ON PURPOSE. Every call site in
        this toolkit was written against MessageBox, and matching its signature
        exactly is what let two hundred of them be converted by replacing a
        type name -- no re-ordering, no re-wrapping, and no chance of a
        confirmation prompt quietly losing its 'Button2' default along the way.

        $args RATHER THAN A param BLOCK, because a call written as
        Show-ToolDialog($a, $b, 'YesNo') hands PowerShell ONE argument that
        happens to be an array, and a call written as Show-ToolDialog $a $b
        hands it two. Adding to an array flattens the first case and preserves
        the second, so both shapes arrive here the same way.

            [0] Text   [1] Caption   [2] Buttons   [3] Icon   [4] DefaultButton
    #>
    $a = @()
    foreach ($x in $args) { $a += $x }

    $text    = $(if ($a.Count -ge 1) { "$($a[0])" } else { '' })
    $caption = $(if ($a.Count -ge 2) { "$($a[1])" } else { '' })
    $bArg    = $(if ($a.Count -ge 3) { $a[2] } else { 'OK' })
    $iArg    = $(if ($a.Count -ge 4) { $a[3] } else { 'None' })
    $dArg    = $(if ($a.Count -ge 5) { $a[4] } else { 'Button1' })

    # Casts, not comparisons, so a caller may pass either the string 'YesNo' or
    # [MessageBoxButtons]::YesNo -- both are used in this toolkit and both have
    # to keep working. A value neither form can resolve falls to the default
    # rather than throwing: this function exists to ASK A QUESTION, and failing
    # to ask it because of an argument is the one outcome worth avoiding.
    $mbB = [System.Windows.Forms.MessageBoxButtons]::OK
    $mbI = [System.Windows.Forms.MessageBoxIcon]::None
    $mbD = [System.Windows.Forms.MessageBoxDefaultButton]::Button1
    try { $mbB = [System.Windows.Forms.MessageBoxButtons]$bArg } catch { }
    try { $mbI = [System.Windows.Forms.MessageBoxIcon]$iArg } catch { }
    try { $mbD = [System.Windows.Forms.MessageBoxDefaultButton]$dArg } catch { }

    if (Test-ConsoleDialogStyle) {
        try {
            return (Show-ThemedMessageDialog -Text $text -Caption $caption `
                        -Buttons $mbB -Icon $mbI -DefaultButton $mbD)
        } catch {
            # Deliberately silent, and deliberately still asking. Whatever went
            # wrong drawing a nicer box is not a reason to leave the operator
            # without the question.
        }
    }
    # THE ONE REMAINING MessageBox CALL IN THE TOOLKIT, AND IT MUST STAY ONE.
    # Everything else was converted by substituting the type name; doing that
    # to this line would make the fallback call itself and recurse until the
    # stack ran out. If that sweep is ever repeated, exclude this line.
    return [System.Windows.Forms.MessageBox]::Show($text, $caption, $mbB, $mbI, $mbD)
}

# Shared one-line prompt. Came across from the second tenant's build on
# 2026-09-12 because Migration Status and Mail Flow Connectors both use it.
function Show-InputDialog {
    <#
        Ask for one line of text.

        Added 2026-08-22 for the Rooms tools, which between them prompt for a
        room list name, a room address, a booking-policy delegate and a
        timezone. Five private copies is FUN001, so it starts shared.

        RETURNS '' FOR CANCEL, never $null. A caller writing
        `if (-not $answer)` then behaves the same for "cancelled" and "typed
        nothing", which is what every caller actually wants -- and the ones that
        need to tell them apart can pass -Distinguish and read the object.

        The OK button is the AcceptButton and Cancel the CancelButton, so Enter
        and Escape both work. Omitting that is the difference between a dialog
        that feels native and one that feels like a script.
    #>
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default = '',
        [int]$Width = 460,
        [switch]$Distinguish
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = $Title
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.StartPosition   = 'CenterParent'

    $pad = 12
    $inner = $Width - ($pad * 2)

    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = $Prompt
    $lbl.Location = New-Object System.Drawing.Point($pad, $pad)
    $lbl.Size     = New-Object System.Drawing.Size($inner, 56)
    [void]$dlg.Controls.Add($lbl)

    $txt          = New-Object System.Windows.Forms.TextBox
    $txt.Text     = $Default
    $txt.Location = New-Object System.Drawing.Point($pad, ($pad + 62))
    $txt.Size     = New-Object System.Drawing.Size($inner, 22)
    [void]$dlg.Controls.Add($txt)

    $btnW = 96
    $ok              = New-Object System.Windows.Forms.Button
    $ok.Text         = 'OK'
    $ok.FlatStyle    = 'Flat'
    $ok.Location     = New-Object System.Drawing.Point(($Width - $pad - ($btnW * 2) - 8), ($pad + 94))
    $ok.Size         = New-Object System.Drawing.Size($btnW, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    [void]$dlg.Controls.Add($ok)

    $cancel              = New-Object System.Windows.Forms.Button
    $cancel.Text         = 'Cancel'
    $cancel.FlatStyle    = 'Flat'
    $cancel.Location     = New-Object System.Drawing.Point(($Width - $pad - $btnW), ($pad + 94))
    $cancel.Size         = New-Object System.Drawing.Size($btnW, 28)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    [void]$dlg.Controls.Add($cancel)

    $dlg.AcceptButton = $ok
    $dlg.CancelButton = $cancel
    $dlg.ClientSize   = New-Object System.Drawing.Size($Width, ($pad + 134))

    Set-ConsoleTheme -Form $dlg
    $result = $dlg.ShowDialog()
    $value  = "$($txt.Text)"
    $dlg.Dispose()

    $cancelled = ($result -ne [System.Windows.Forms.DialogResult]::OK)
    if ($Distinguish) {
        return [PSCustomObject]@{
            Cancelled = $cancelled
            Value     = $(if ($cancelled) { '' } else { $value })
        }
    }
    if ($cancelled) { return '' }
    return $value
}
#endregion

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

        IT REPAINTS. IT DOES NOT PUMP INPUT. That distinction is the whole
        reason this note is long.

        This used to call Application::DoEvents(), on the reasoning that a
        single-threaded tool never repaints without it and Windows greys the
        window out as "Not Responding". The first half is true. DoEvents is the
        wrong cure, because it does not only repaint -- it DISPATCHES QUEUED
        INPUT. Every click banked while a long run was working got delivered
        into the middle of that run.

        What that cost, in Disable User, on real terminations: a click on Close
        ran $form.Close() between two users, disposing the form while the loop
        still held references to its grid -- so the next row added threw
        ObjectDisposedException from outside any handler, and the operator was
        left restarting PowerShell mid-batch. Clear did the same to the results
        grid while row indices were live.

        Control.Refresh() repaints the one control that changed and dispatches
        nothing. Clicks still queue; they are delivered when the run finishes,
        which is when they can be handled safely.

        -AllowInput restores the old behaviour for a caller that needs a Cancel
        button to be clickable mid-run. Nothing needs it today. If you reach for
        it, disable every OTHER control first -- a live Cancel and a live Close
        are not the same risk.
    #>
    param(
        $Label,
        $Bar,
        [string]$Activity = 'Working',
        [switch]$NoConsole,
        [switch]$AllowInput
    )
    $state = [pscustomobject]@{
        Label      = $Label
        Bar        = $Bar
        Activity   = $Activity
        NoConsole  = [bool]$NoConsole
        AllowInput = [bool]$AllowInput
        Started    = Get-Date
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
        if ($state.AllowInput) {
            [System.Windows.Forms.Application]::DoEvents()
        } else {
            # Repaint only. IsDisposed is checked because a form closed between
            # two progress calls would otherwise throw from inside the reporter,
            # which is the least useful place for an exception to come from.
            if ($state.Label -and -not $state.Label.IsDisposed) { $state.Label.Refresh() }
            if ($state.Bar   -and -not $state.Bar.IsDisposed)   { $state.Bar.Refresh() }
        }
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
    $dlg.Size            = New-Object System.Drawing.Size(620, 786)
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

    # ---- console branding ---------------------------------------------------
    #
    # HERE RATHER THAN IN Tenant.json, AND THE LINE IS THE SAME ONE THIS
    # DIALOG'S HEADER DRAWS. Tenant.json holds facts that are identical on
    # every machine at a tenant; Settings.json holds what this installation
    # looks like. A console's title is the latter -- and it is certainly not
    # something that belongs hardcoded in AdminConsole.ps1, which ships
    # unchanged to every tenant.
    $grpBrand          = New-Object System.Windows.Forms.GroupBox
    $grpBrand.Text     = 'Console branding and look'
    $grpBrand.Location = New-Object System.Drawing.Point(12, 296)
    $grpBrand.Size     = New-Object System.Drawing.Size(584, 256)
    $dlg.Controls.Add($grpBrand)

    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Console title:'; $l.Location = New-Object System.Drawing.Point(12, 27); $l.Size = New-Object System.Drawing.Size(110, 20)
    $grpBrand.Controls.Add($l)
    $txtTitle          = New-Object System.Windows.Forms.TextBox
    $txtTitle.Location = New-Object System.Drawing.Point(126, 24)
    $txtTitle.Size     = New-Object System.Drawing.Size(260, 22)
    $txtTitle.Text     = "$($current.ConsoleTitle)"
    $grpBrand.Controls.Add($txtTitle)

    $lblTitleHint           = New-Object System.Windows.Forms.Label
    $lblTitleHint.Text      = 'Shown top-left on the console. Blank uses "Admin Console".'
    $lblTitleHint.Location  = New-Object System.Drawing.Point(394, 27)
    $lblTitleHint.Size      = New-Object System.Drawing.Size(178, 28)
    $lblTitleHint.ForeColor = [System.Drawing.Color]::DimGray
    $grpBrand.Controls.Add($lblTitleHint)

    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Title font:'; $l.Location = New-Object System.Drawing.Point(12, 59); $l.Size = New-Object System.Drawing.Size(110, 20)
    $grpBrand.Controls.Add($l)

    # A DROPDOWN OF FONTS THAT ARE REALLY INSTALLED, not a free-text box.
    #
    # Setting a face that is not present fails SILENTLY: GDI+ substitutes a
    # default and returns a perfectly valid Font, so a typo in a brand font
    # name looks exactly like the setting having no effect. Offering only what
    # this machine actually has makes that impossible to get wrong here, and
    # Get-ConsoleTitleFont still verifies at run time for the case where
    # Settings.json came from a machine with different fonts.
    $cboFont               = New-Object System.Windows.Forms.ComboBox
    $cboFont.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cboFont.Location      = New-Object System.Drawing.Point(126, 56)
    $cboFont.Size          = New-Object System.Drawing.Size(260, 22)
    [void]$cboFont.Items.Add('(theme default)')
    try {
        foreach ($fam in ([System.Drawing.FontFamily]::Families | ForEach-Object { $_.Name } | Sort-Object -Unique)) {
            [void]$cboFont.Items.Add($fam)
        }
    } catch { }
    $wantFont = "$($current.ConsoleTitleFont)".Trim()
    $cboFont.SelectedIndex = 0
    if ($wantFont -and $cboFont.Items.Contains($wantFont)) { $cboFont.SelectedItem = $wantFont }
    $grpBrand.Controls.Add($cboFont)

    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Tool buttons:'; $l.Location = New-Object System.Drawing.Point(12, 91); $l.Size = New-Object System.Drawing.Size(110, 20)
    $grpBrand.Controls.Add($l)
    $cboBtn               = New-Object System.Windows.Forms.ComboBox
    $cboBtn.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cboBtn.Location      = New-Object System.Drawing.Point(126, 88)
    $cboBtn.Size          = New-Object System.Drawing.Size(260, 22)
    [void]$cboBtn.Items.Add('Flat')
    [void]$cboBtn.Items.Add('Raised (3D)')
    [void]$cboBtn.Items.Add('Enhanced (3D)')
    # Enhanced is Raised with every number that makes the bevel pushed
    # further -- see Get-RaisedButtonDepth. It is a third VALUE, not a third
    # style: everything that treats a button as raised treats an enhanced one
    # the same way, so nothing else in the toolkit had to learn a new word.
    $cboBtn.SelectedItem  = $(switch ("$($current.ConsoleButtonStyle)".Trim()) {
        'Enhanced' { 'Enhanced (3D)' }
        'Raised'   { 'Raised (3D)' }
        default    { 'Flat' }
    })
    $grpBrand.Controls.Add($cboBtn)

    # ---- icons --------------------------------------------------------------
    #
    # OFFERED ONLY WHERE IT CAN ACTUALLY BE DELIVERED. The glyphs come from an
    # icon font that Windows supplies, and asking for a font that is not
    # installed does not fail -- it silently substitutes, and the console fills
    # with empty boxes. So the dropdown reports what this machine has rather
    # than letting somebody switch on something that cannot work here.
    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Tool icons:'; $l.Location = New-Object System.Drawing.Point(12, 123); $l.Size = New-Object System.Drawing.Size(110, 20)
    $grpBrand.Controls.Add($l)
    $cboIcon               = New-Object System.Windows.Forms.ComboBox
    $cboIcon.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cboIcon.Location      = New-Object System.Drawing.Point(126, 120)
    $cboIcon.Size          = New-Object System.Drawing.Size(260, 22)
    [void]$cboIcon.Items.Add('Text only')
    [void]$cboIcon.Items.Add('Icons and text')
    $cboIcon.SelectedItem  = $(if ("$($current.ConsoleIcons)".Trim() -ieq 'On') { 'Icons and text' } else { 'Text only' })
    $grpBrand.Controls.Add($cboIcon)

    $iconFont = ''
    try { $iconFont = Get-ToolkitIconFontName } catch { }
    $lblIconHint           = New-Object System.Windows.Forms.Label
    $lblIconHint.Location  = New-Object System.Drawing.Point(394, 121)
    $lblIconHint.Size      = New-Object System.Drawing.Size(178, 32)
    $lblIconHint.ForeColor = [System.Drawing.Color]::DimGray
    if ($iconFont) {
        $lblIconHint.Text = "Drawn with $iconFont, which is already on this machine."
    } else {
        $lblIconHint.Text = 'No icon font is installed here, so icons will stay off whatever this says.'
        $cboIcon.Enabled  = $false
    }
    $grpBrand.Controls.Add($lblIconHint)

    # ---- dialogs ------------------------------------------------------------
    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Pop-up boxes:'; $l.Location = New-Object System.Drawing.Point(12, 155); $l.Size = New-Object System.Drawing.Size(110, 20)
    $grpBrand.Controls.Add($l)
    $cboDlg               = New-Object System.Windows.Forms.ComboBox
    $cboDlg.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cboDlg.Location      = New-Object System.Drawing.Point(126, 152)
    $cboDlg.Size          = New-Object System.Drawing.Size(260, 22)
    [void]$cboDlg.Items.Add('Windows standard')
    [void]$cboDlg.Items.Add('Match the theme')
    $cboDlg.SelectedItem  = $(if ("$($current.ConsoleDialogStyle)".Trim() -ieq 'Themed') { 'Match the theme' } else { 'Windows standard' })
    $grpBrand.Controls.Add($cboDlg)

    $lblDlgHint           = New-Object System.Windows.Forms.Label
    $lblDlgHint.Text      = 'Confirmations and messages. Windows draws the standard ones, so no theme can reach them.'
    $lblDlgHint.Location  = New-Object System.Drawing.Point(394, 153)
    $lblDlgHint.Size      = New-Object System.Drawing.Size(178, 32)
    $lblDlgHint.ForeColor = [System.Drawing.Color]::DimGray
    $grpBrand.Controls.Add($lblDlgHint)


    # ---- corners ------------------------------------------------------------
    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Button corners:'; $l.Location = New-Object System.Drawing.Point(12, 187); $l.Size = New-Object System.Drawing.Size(110, 20)
    $grpBrand.Controls.Add($l)
    $cboCorner               = New-Object System.Windows.Forms.ComboBox
    $cboCorner.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cboCorner.Location      = New-Object System.Drawing.Point(126, 184)
    $cboCorner.Size          = New-Object System.Drawing.Size(260, 22)
    [void]$cboCorner.Items.Add('Square')
    [void]$cboCorner.Items.Add('Rounded')
    $cboCorner.SelectedItem  = $(if ("$($current.ConsoleCorners)".Trim() -ieq 'Rounded') { 'Rounded' } else { 'Square' })
    $grpBrand.Controls.Add($cboCorner)

    $lblCornerHint           = New-Object System.Windows.Forms.Label
    $lblCornerHint.Location  = New-Object System.Drawing.Point(394, 185)
    $lblCornerHint.Size      = New-Object System.Drawing.Size(178, 32)
    $lblCornerHint.ForeColor = [System.Drawing.Color]::DimGray
    # Said out loud rather than left to be discovered: with Tool buttons on
    # Flat there is nothing owner-drawn to round, so this dropdown does nothing
    # and the operator is entitled to know that before changing it twice.
    $lblCornerHint.Text = $(if (@('Raised', 'Enhanced') -contains "$($current.ConsoleButtonStyle)".Trim()) {
        'Applies to raised and enhanced buttons.'
    } else {
        'Needs Tool buttons set to Raised or Enhanced; flat buttons are always square.'
    })
    $grpBrand.Controls.Add($lblCornerHint)

    # ---- header operator ------------------------------------------------------
    $chkHdrOp          = New-Object System.Windows.Forms.CheckBox
    $chkHdrOp.Text     = 'Show "signed in as <name>" beside the title (the status strip names you regardless)'
    $chkHdrOp.Location = New-Object System.Drawing.Point(12, 222)
    $chkHdrOp.Size     = New-Object System.Drawing.Size(560, 22)
    $chkHdrOp.Checked  = ("$($current.ConsoleHeaderOperator)".Trim() -ieq 'On')
    $grpBrand.Controls.Add($chkHdrOp)
    $lblFontHint           = New-Object System.Windows.Forms.Label
    $lblFontHint.Location  = New-Object System.Drawing.Point(394, 59)
    $lblFontHint.Size      = New-Object System.Drawing.Size(178, 30)
    $lblFontHint.ForeColor = [System.Drawing.Color]::DimGray
    # A configured face that is NOT installed here is worth saying out loud --
    # it is the one state where the file and the screen disagree and the screen
    # gives no clue why.
    $lblFontHint.Text = $(if ($wantFont -and -not $cboFont.Items.Contains($wantFont)) {
        "'$wantFont' is set but is not installed on this machine -- the theme font is being used."
    } else { 'Only fonts installed here are listed.' })
    $grpBrand.Controls.Add($lblFontHint)

    $grpBrowser          = New-Object System.Windows.Forms.GroupBox
    $grpBrowser.Text     = 'Edge profiles  --  the profile NAME as shown in Edge, not "Profile 2"'
    $grpBrowser.Location = New-Object System.Drawing.Point(12, 562)
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
    $btnOK.Location = New-Object System.Drawing.Point(386, 684)
    $btnOK.Size     = New-Object System.Drawing.Size(100, 28)
    $btnOK.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($btnOK)

    $btnCancel          = New-Object System.Windows.Forms.Button
    $btnCancel.Text     = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(494, 684)
    $btnCancel.Size     = New-Object System.Drawing.Size(100, 28)
    $btnCancel.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btnCancel)

    $btnOK.Add_Click({
        $dis = "$($txtDisable.Text)".Trim()
        if ($dis -and -not (Test-Path -LiteralPath $dis)) {
            $ans = Show-ToolDialog(
                "The retry queue folder does not exist:`n`n$dis`n`nA path entered here is never created automatically, so nothing will be able to save a retry queue until it exists.`n`nSave it anyway?",
                'Settings', 'YesNo', 'Warning', 'Button2')
            if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        $oper = "$($txtOper.Text)".Trim()
        if ($oper -and $oper -notmatch '^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$') {
            $null = Show-ToolDialog(
                "That does not look like an email address:`n`n$oper", 'Settings', 'OK', 'Warning')
            return
        }
        $prof = @{}
        foreach ($k in $boxes.Keys) { $prof[$k] = "$($boxes[$k].Text)".Trim() }
        try {
            # '(theme default)' is the dropdown's way of saying "no font
            # configured" -- it must be saved as blank, not as a family name
            # nothing will ever resolve.
            $fontPick = "$($cboFont.SelectedItem)"
            if ($fontPick -eq '(theme default)') { $fontPick = '' }
            # Stored as the bare word, not the label. 'Raised (3D)' is what
            # reads well in a dropdown; 'Raised' is what the console tests for,
            # and putting the display text in the file would make the setting
            # depend on the wording of a label.
            $btnStyle = $(switch -Wildcard ("$($cboBtn.SelectedItem)") {
                'Enhanced*' { 'Enhanced' }
                'Raised*'   { 'Raised' }
                default     { '' }
            })
            $iconPick = $(if ("$($cboIcon.SelectedItem)" -like 'Icons*') { 'On' } else { '' })
            $cornerPick = $(if ("$($cboCorner.SelectedItem)" -like 'Rounded*') { 'Rounded' } else { '' })
            $dlgPick  = $(if ("$($cboDlg.SelectedItem)" -like 'Match*') { 'Themed' } else { '' })
            [void](Save-ToolkitSettings -ReportsPath "$($txtReports.Text)".Trim() `
                        -LogPath "$($txtLogs.Text)".Trim() `
                        -RetryQueuePath $dis `
                        -OperatorEmail "$($txtOper.Text)".Trim() `
                        -ConsoleTitle "$($txtTitle.Text)".Trim() `
                        -ConsoleTitleFont $fontPick `
                        -ConsoleButtonStyle $btnStyle `
                        -ConsoleIcons $iconPick `
                        -ConsoleCorners $cornerPick `
                        -ConsoleDialogStyle $dlgPick `
                        -ConsoleHeaderOperator $(if ($chkHdrOp.Checked) { 'On' } else { '' }) `
                        -BrowserProfiles $prof)
            $dlg.Tag = 'saved'
            $dlg.Close()
        } catch {
            $null = Show-ToolDialog(
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
        [string]$Extension = 'csv',
        # A window with TWO of these rows needs to say which is which, and
        # 'Previous runs:' twice says nothing. Set the caption here rather than
        # writing .Label.Text afterwards, because the combo's left edge is
        # derived from the label's WIDTH -- assigning the text later moves the
        # words but not the box, and a longer caption then runs under the
        # dropdown. The width is MEASURED from the caption in the form's own
        # font, so a theme with a larger typeface widens the gap instead of
        # clipping. Default unchanged, so every existing caller is untouched.
        [string]$LabelText = 'Previous runs:',
        # Force a gutter width. Two rows in one window measure to two different
        # widths, which puts their dropdowns at two different X positions and
        # reads as a wonky form. Pass the same value to both and they line up.
        [int]$LabelWidth = 0
    )

    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = $LabelText
    $lbl.Location = New-Object System.Drawing.Point($X, ($Y + 5))
    $lblW         = 90
    try {
        $m = [System.Windows.Forms.TextRenderer]::MeasureText($LabelText, $Form.Font)
        if ($m.Width + 6 -gt $lblW) { $lblW = $m.Width + 6 }
    } catch { }
    if ($LabelWidth -gt 0) { $lblW = $LabelWidth }
    $lbl.Size     = New-Object System.Drawing.Size($lblW, 18)
    $Form.Controls.Add($lbl)

    $comboX = $X + $lblW + 4
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
        Id       = 'Graph'
        Label    = 'Microsoft Graph'
        Patterns = @(
            # The toolkit's own wrappers first -- every Graph tool goes through
            # them -- then the raw SDK calls for anything hand-rolled.
            '\bInvoke-ToolkitGraph\b'
            '\bAssert-ToolkitGraphScope\b'
            '\bResolve-GraphUser\b'
            '\bInvoke-MgGraphRequest\b'
            '\bGet-Mg[A-Z][A-Za-z]*'
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
    @{ Id = 'Graph'; Pattern = '(?m)^\s*(\$\w+\s*=\s*)?(Connect-ToolkitGraph|Connect-MgGraph)\b' }
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
    <#
        How long an activation lasts, in hours.

        TENANT FIRST. The duration is the tenant's PIM policy, not a
        preference, so Tenant.json (Pim.DurationHours) is the authority and
        travels with the tenant. Settings.json (PimDurationHours) is read next
        for installs that set it there before Tenant.json carried it, and 4
        is the last resort. Until 2026-09-12 only Settings was read, which
        left a Tenant.json saying 9 with a header counting down from 4.
    #>
    $h = 0
    try {
        $t = Get-TenantPimSetting -Name 'DurationHours'
        if ($null -ne $t -and [int]::TryParse("$t", [ref]$h) -and $h -gt 0) { return $h }
    } catch { }
    $s = $null
    try { $s = Get-ToolkitSettings } catch { }
    if ($s -and $s.PSObject.Properties['PimDurationHours']) {
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

function Get-IntegrityCachePath {
    # Beside the capability cache, for the same reasons: per-user, per-machine,
    # and never in the toolkit folder where several operators would write it.
    $dir = Join-Path $env:LOCALAPPDATA 'M365AdminToolkit'
    if (-not (Test-Path -LiteralPath $dir)) {
        try { New-Item -Path $dir -ItemType Directory -Force | Out-Null } catch { return $null }
    }
    return (Join-Path $dir 'integrity-cache.json')
}

function Get-FileHashCached {
    <#
        SHA-256 of a file, remembered between launches.

        Test-ToolkitIntegrity hashed every code file on every console start to
        print "103 files verified" -- a hundred-plus SHA computations whose
        answer was identical to the previous launch in all but the launch after
        a deploy. That was the largest fixed cost between double-clicking the
        console and seeing it.

        THE KEY IS (size, LastWriteTimeUtc ticks). A file whose bytes changed
        has a new mtime or a new size or both; a file that was merely copied
        gets a new mtime and is rehashed once, which is the correct answer for
        "has it changed" and costs one hash. What this cannot catch is a byte
        edit that preserves size and mtime exactly -- which does not happen by
        accident and is not what this check is for. The manifest check is a
        deployment-drift detector, not tamper evidence.

        THE CACHE IS PASSED IN, NOT OPENED HERE. One JSON read and one write per
        integrity check, not one per file.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$Cache
    )
    try {
        $fi  = Get-Item -LiteralPath $Path -ErrorAction Stop
        $key = "$($fi.Length)|$($fi.LastWriteTimeUtc.Ticks)"
        $k   = $Path.ToLowerInvariant()
        $hit = $Cache[$k]
        if ($hit -and "$($hit.Key)" -eq $key -and "$($hit.Hash)") { return "$($hit.Hash)" }
        $h = Get-FileHashHex -Path $Path
        if ($h) { $Cache[$k] = @{ Key = $key; Hash = $h } }
        return $h
    } catch { return (Get-FileHashHex -Path $Path) }
}

function Import-IntegrityCache {
    $p = Get-IntegrityCachePath
    $c = @{}
    if ($p -and (Test-Path -LiteralPath $p)) {
        try {
            $raw = Get-Content -LiteralPath $p -Raw -ErrorAction Stop | ConvertFrom-Json
            foreach ($prop in $raw.PSObject.Properties) {
                $c[$prop.Name] = @{ Key = "$($prop.Value.Key)"; Hash = "$($prop.Value.Hash)" }
            }
        } catch { $c = @{} }   # a corrupt cache costs one full rehash, not a failure
    }
    return $c
}

function Export-IntegrityCache {
    param([Parameter(Mandatory)][hashtable]$Cache)
    $p = Get-IntegrityCachePath
    if (-not $p) { return }
    try {
        $o = [ordered]@{}
        foreach ($k in $Cache.Keys) { $o[$k] = $Cache[$k] }
        ($o | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath $p -Encoding UTF8
    } catch { }
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

    # Hashes come from the cache unless size or mtime moved. See
    # Get-FileHashCached for what that does and does not guarantee.
    $cache = Import-IntegrityCache
    foreach ($rel in $v.Files.Keys) {
        if (-not $onDisk.ContainsKey($rel)) { [void]$missing.Add($rel); continue }
        $h = Get-FileHashCached -Path $onDisk[$rel] -Cache $cache
        if ($h -and $h -eq $v.Files[$rel]) { $matched++ } else { [void]$modified.Add($rel) }
    }
    Export-IntegrityCache -Cache $cache
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

    # ---------------------------------------------------------------------
    # ZOOM-ON-RESIZE, WIRED HERE BECAUSE THIS IS THE ONE CALL EVERY TOOL MAKES
    #
    # 69 of the 70 forms in this toolkit end their layout with this function --
    # it has to be last for z-order -- so it is the only place that can turn a
    # behaviour on everywhere without editing 69 files. Editing 69 files would
    # also mean 69 chances to put the call in the wrong place, and the wrong
    # place is what breaks it.
    #
    # Enable-FormAutoScale captures its baseline on Shown, not now, so a tool
    # that adds more controls after this call is still handled correctly.
    #
    # To opt a form out, put 'NoAutoScale' in its Tag before this runs.
    # ---------------------------------------------------------------------
    Enable-FormAutoScale -Form $Form

    # KEYBOARD DEFAULTS, for the same reason: this is the one call every tool
    # makes. Enter = Look Up (a read; never Save -- Enter must not be able to
    # write), Esc = Close. Multi-line boxes keep their Enter for newlines.
    # Opt out with 'NoAcceptButton' in the form's Tag (the HTML editor does:
    # its Enter belongs to the document).
    Set-FormKeyboardDefaults -Form $Form

    # STATUS BAR, for the same reason again. Every tool gets a strip along
    # the bottom that says what is happening and for how long, and it turns
    # itself on whenever the tool shows the wait cursor -- which every tool
    # already does -- so nothing else has to be edited for the bar to work.
    # Tools that want better words than "Working ..." call Set-ToolStatus.
    # Opt out with 'NoStatusBar' in the form's Tag (the console does: it has
    # its own status row and wires the bar to that instead).
    if ("$($Form.Tag)" -notlike '*NoStatusBar*') {
        try { Add-ToolStatusBar -Form $Form | Out-Null } catch { }
    }

    return $btn
}

#region Status bar and busy feedback -------------------------------------------
<#
    WHY THIS EXISTS. A tool that calls Exchange blocks its own window until
    Exchange answers, and the only sign of life was the wait cursor -- which
    means a fifteen-second Get-MessageTraceV2 looked exactly like a hang.
    Nothing in that window could animate, because the thread that would draw
    the animation is the thread waiting on Exchange.

    WHAT CAN BE DONE WITHOUT A SECOND THREAD, and is done here:
      - a status label that says WHAT is running ("Reading rules ...") and
        repaints synchronously (Refresh, not DoEvents) before the blocking
        call starts, so the words are on screen while the call runs;
      - an elapsed time stamped in when the call ends, so a slow call is
        seen to have been slow rather than suspected of having hung;
      - the wait cursor mirrored into the bar automatically (CursorChanged),
        which covers the 55 tools that set it and never knew about this.

    WHAT NEEDS A SECOND THREAD, and is optional: Start-ToolJob below runs a
    cloud read in a background runspace when the process can support it, and
    then the marquee moves for real. Where it cannot, the same call runs
    inline with the status text, which is exactly today's behaviour plus
    words. Tools never have to know which happened.

    Refresh() rather than [Application]::DoEvents() is deliberate. DoEvents
    pumps INPUT as well as paint, so a second click on "Look Up" during the
    first one's work would run the handler re-entrantly; the tools with a
    Busy guard would survive that, the forty without would not.
    Control.Refresh() paints the control now and pumps nothing.
#>
$script:ToolStatusHeight = 24
$global:ToolkitStatusBars = @{}

function Get-ToolStatusState {
    <# The status-bar state for a form, keyed by the form's identity (the Tag stays free). #>
    param([Parameter(Mandatory)]$Form)
    $key = $Form.GetHashCode()
    if ($global:ToolkitStatusBars.ContainsKey($key)) { return $global:ToolkitStatusBars[$key] }
    return $null
}

function Add-ToolStatusBar {
    <#
        A status strip along the bottom of a form: label, marquee bar,
        elapsed time. Called for every tool from Add-ToolHelpButton.

        Pass -Label (and optionally -Bar / -Elapsed) to wire EXISTING
        controls instead of creating a strip -- the console does that with
        its result label, because it already has a status row.

        WHEN A STRIP IS CREATED the form is made taller by the strip's height
        so nothing is covered, and bottom-anchored controls are moved up so
        they stay clear of it on later resizes. A form that is opened
        maximised or that uses Dock for its own layout is left alone (the
        strip simply docks below).
    #>
    param(
        [Parameter(Mandatory)]$Form,
        $Label = $null,
        $Bar = $null,
        $Elapsed = $null
    )
    $existing = Get-ToolStatusState -Form $Form
    if ($existing) { return $existing }

    $theme = $null
    try { $theme = Get-ConsoleTheme } catch { }
    $created = $false

    if (-not $Label) {
        $created = $true
        $h = $script:ToolStatusHeight
        $panel = New-Object System.Windows.Forms.Panel
        $panel.Name   = 'pnlToolStatus'
        $panel.Height = $h
        $panel.Dock   = [System.Windows.Forms.DockStyle]::Bottom
        $panel.Tag    = 'ToolStatusBar'

        $line = New-Object System.Windows.Forms.Panel
        $line.Height = 1
        $line.Dock   = [System.Windows.Forms.DockStyle]::Top
        $line.Tag    = 'NoTheme'
        $line.BackColor = [System.Drawing.Color]::FromArgb(217, 221, 227)

        $Label = New-Object System.Windows.Forms.Label
        $Label.Name      = 'lblToolStatus'
        $Label.Text      = 'Ready'
        $Label.AutoSize  = $false
        $Label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $Label.Location  = New-Object System.Drawing.Point(10, 1)
        $Label.Height    = $h - 1
        $Label.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

        $Bar = New-Object System.Windows.Forms.ProgressBar
        $Bar.Name    = 'pbToolStatus'
        $Bar.Style   = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $Bar.MarqueeAnimationSpeed = 25
        $Bar.Size    = New-Object System.Drawing.Size(110, 10)
        $Bar.Anchor  = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
        $Bar.Visible = $false

        $Elapsed = New-Object System.Windows.Forms.Label
        $Elapsed.Name      = 'lblToolElapsed'
        $Elapsed.AutoSize  = $false
        $Elapsed.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
        $Elapsed.Size      = New-Object System.Drawing.Size(90, ($h - 1))
        $Elapsed.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

        $panel.Controls.Add($Label)
        $panel.Controls.Add($Bar)
        $panel.Controls.Add($Elapsed)
        $panel.Controls.Add($line)

        # Place the right-hand pieces from the panel's real width, now and on
        # every resize; the label takes what is left.
        $layout = {
            $p = $this
            $w = $p.ClientSize.Width
            $el = $p.Controls['lblToolElapsed']; $pb = $p.Controls['pbToolStatus']; $lb = $p.Controls['lblToolStatus']
            if ($el) { $el.Location = New-Object System.Drawing.Point(($w - $el.Width - 8), 1) }
            if ($pb) { $pb.Location = New-Object System.Drawing.Point(($w - 90 - 8 - $pb.Width - 10), [int](($p.Height - $pb.Height) / 2)) }
            if ($lb) { $lb.Width = [Math]::Max(60, ($w - 90 - 8 - 110 - 10 - 10 - 10)) }
        }
        $panel.Add_Resize($layout)

        if ($theme -and $theme.Enabled) {
            try {
                $pb2 = "$($theme.PanelBack)"; if (-not $pb2) { $pb2 = "$($theme.FormBack)" }
                if ($pb2) { $panel.BackColor = ConvertTo-ThemeColor $pb2 }
                $ft = "$($theme.GroupText)"; if (-not $ft) { $ft = "$($theme.FormText)" }
                if ($ft) { $Label.ForeColor = ConvertTo-ThemeColor $ft; $Elapsed.ForeColor = ConvertTo-ThemeColor $ft }
                $bd = "$($theme.ButtonBorder)"; if ($bd) { $line.BackColor = ConvertTo-ThemeColor $bd }
            } catch { }
        }

        # Grow the form so the strip covers nothing, then keep bottom-anchored
        # controls above it. A control anchored Top+Bottom stretches with the
        # form, so it is shortened instead of moved.
        $Form.Controls.Add($panel)
        try {
            $cs = $Form.ClientSize
            $Form.ClientSize = New-Object System.Drawing.Size($cs.Width, ($cs.Height + $h))
            $bottom = [System.Windows.Forms.AnchorStyles]::Bottom
            $top    = [System.Windows.Forms.AnchorStyles]::Top
            foreach ($c in @($Form.Controls)) {
                if ($c -eq $panel) { continue }
                if ($c.Dock -ne [System.Windows.Forms.DockStyle]::None) { continue }
                if (($c.Anchor -band $bottom) -eq 0) { continue }
                if (($c.Anchor -band $top) -ne 0) { $c.Height = [Math]::Max(10, ($c.Height - $h)) } else { $c.Top = $c.Top - $h }
            }
            if ($Form.MinimumSize.Height -gt 0) {
                $Form.MinimumSize = New-Object System.Drawing.Size($Form.MinimumSize.Width, ($Form.MinimumSize.Height + $h))
            }
        } catch { }
        # Dock sized the panel when it was added and Resize placed the pieces;
        # this covers a host that adds before the form has a size.
        $panel.PerformLayout()
    }

    $st = @{
        Form = $Form; Label = $Label; Bar = $Bar; Elapsed = $Elapsed
        Busy = $false; Guard = $false; Created = $created; JobRunning = $false
        Watch = New-Object System.Diagnostics.Stopwatch
        Timer = $null; IdleText = "$($Label.Text)"; IdleColor = $Label.ForeColor
        Danger = $null
    }
    if ($theme -and $theme.Enabled -and "$($theme.DangerColor)") { try { $st.Danger = ConvertTo-ThemeColor "$($theme.DangerColor)" } catch { } }
    if (-not $st.Danger) { $st.Danger = [System.Drawing.Color]::FromArgb(214, 69, 69) }

    # Ticks only while the message loop runs, i.e. during a background job;
    # during an inline call the elapsed time is stamped when the call ends.
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 250
    $timer.Tag = $st
    $timer.Add_Tick({
        $s = $this.Tag
        if ($s.Busy -and $s.Elapsed) { $s.Elapsed.Text = ('{0:0.0} s' -f $s.Watch.Elapsed.TotalSeconds) }
    })
    $st.Timer = $timer

    $key = $Form.GetHashCode()
    $global:ToolkitStatusBars[$key] = $st

    # The wait cursor IS the busy signal every tool already sends. Mirror it.
    $Form.Add_CursorChanged({
        $s = Get-ToolStatusState -Form $this
        if (-not $s -or $s.Guard) { return }
        if ($this.Cursor -eq [System.Windows.Forms.Cursors]::WaitCursor) {
            if (-not $s.Busy) { Set-ToolStatus -Form $this -Busy }
        } elseif ($s.Busy) {
            Set-ToolStatus -Form $this -Done
        }
    })
    $Form.Add_FormClosed({
        $s = Get-ToolStatusState -Form $this
        if ($s) {
            try { $s.Timer.Stop(); $s.Timer.Dispose() } catch { }
            $global:ToolkitStatusBars.Remove($this.GetHashCode())
        }
    })
    return $st
}

function Set-ToolStatus {
    <#
        Say what the tool is doing.

            Set-ToolStatus -Form $form -Busy -Text 'Reading rules ...'
            ... blocking call ...
            Set-ToolStatus -Form $form -Done -Text '12 rules'

        -Busy shows the marquee and starts the clock (idempotent: a second
        -Busy only changes the words). -Done / -Failed hide it and stamp the
        elapsed time. Text alone, while busy, just updates the words -- the
        way to narrate a multi-step handler. Every path repaints the strip
        synchronously so the words are visible BEFORE the caller blocks.

        No-op on a form without a bar, so tools can call it unconditionally.
    #>
    param(
        [Parameter(Mandatory)]$Form,
        [string]$Text = '',
        [switch]$Busy,
        [switch]$Done,
        [switch]$Failed
    )
    $s = Get-ToolStatusState -Form $Form
    if (-not $s) { return }
    $s.Guard = $true
    try {
        if ($Busy) {
            if (-not $s.Busy) {
                $s.Busy = $true
                $s.Watch.Restart()
                # The theme may have recoloured the label since the bar was
                # made; remember the current idle colour unless it is the
                # failure red from a previous run.
                if ($s.Label.ForeColor.ToArgb() -ne $s.Danger.ToArgb()) { $s.IdleColor = $s.Label.ForeColor }
                $s.Label.ForeColor = $s.IdleColor
                if ($s.Elapsed) { $s.Elapsed.Text = '' }
                if ($s.Bar) { $s.Bar.Visible = $true }
                try { $s.Timer.Start() } catch { }
            }
            $s.Label.Text = $(if ($Text) { $Text } else { 'Working ...' })
        } elseif ($Done -or $Failed) {
            $secs = [Math]::Round($s.Watch.Elapsed.TotalSeconds, 1)
            $s.Busy = $false
            try { $s.Timer.Stop() } catch { }
            $s.Watch.Stop()
            if ($s.Bar) { $s.Bar.Visible = $false }
            if ($s.Elapsed) { $s.Elapsed.Text = $(if ($secs -gt 0) { "$secs s" } else { '' }) }
            if ($Failed) {
                $s.Label.ForeColor = $s.Danger
                $s.Label.Text = $(if ($Text) { $Text } else { 'Failed -- see the message.' })
            } else {
                $s.Label.ForeColor = $s.IdleColor
                $s.Label.Text = $(if ($Text) { $Text } else { 'Done.' })
            }
        } elseif ($Text) {
            $s.Label.Text = $Text
        }
        try { $s.Label.Refresh(); if ($s.Bar) { $s.Bar.Refresh() }; if ($s.Elapsed) { $s.Elapsed.Refresh() } } catch { }
    } finally {
        $s.Guard = $false
    }
}

function Write-ToolStep {
    <#
        Narrate a step from anywhere in a tool without holding the form:

            Write-ToolStep 'Reading mailbox statistics ...'

        Finds the active form (the one whose handler is running) and puts the
        words on its status bar, going busy if it was not already. Safe to
        call from a helper function that has no idea a form exists: with no
        active form, or no bar on it, it does nothing. Done is marked by
        whatever ends the handler (wait cursor, Look Up's posted Done, or an
        explicit Set-ToolStatus -Done).
    #>
    param([Parameter(Mandatory)][string]$Text)
    try {
        $f = [System.Windows.Forms.Form]::ActiveForm
        if (-not $f) { return }
        if (-not (Get-ToolStatusState -Form $f)) { return }
        Set-ToolStatus -Form $f -Busy -Text $Text
    } catch { }
}

function Invoke-ToolStep {
    <#
        Narrate one blocking step and return its output:

            $rules = Invoke-ToolStep -Form $form -Text 'Reading rules ...' { Get-RuleList -Smtp $smtp }

        The bar goes busy (if it was not already), the words go up, the
        block runs, and its output is returned. On an exception the bar is
        marked Failed with the step's words and the exception is rethrown,
        so the tool's own catch still runs. The bar is NOT marked Done here:
        a handler usually has several steps and ends by clearing its own
        wait cursor, which marks Done with the total time.
    #>
    param(
        [Parameter(Mandatory)]$Form,
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][scriptblock]$Script
    )
    Set-ToolStatus -Form $Form -Busy -Text $Text
    try {
        return (& $Script)
    } catch {
        Set-ToolStatus -Form $Form -Failed -Text "$Text failed."
        throw
    }
}

#region Background jobs (cloud reads off the UI thread) ----------------------
<#
    THE CONSTRAINT. Connect-ExchangeOnline binds its session to the runspace
    that made it, and every tool in this toolkit runs its cmdlets on the UI
    thread's runspace. A background runspace therefore starts with no
    Exchange at all.

    THE WAY THROUGH, for Exchange ONLINE only. The V3 module keeps its
    connection context in process-wide static state and exposes the cmdlets
    as a generated module on disk (tmpEXO_<id>). Import that module into a
    worker runspace and its proxies find the same context. That is the
    documented pattern for parallel EXO work in runspace pools, and it is
    what Start-ToolJob relies on. It is NOT assumed: the first call probes
    the worker with a real cmdlet and, if that fails, every job for the rest
    of the session runs inline instead. On-premises sessions (implicit
    remoting) cannot be shared this way and are never attempted.

    The worker is ONE persistent runspace per process, created on first use
    (it dot-sources Common.ps1 once, which is where the first job's extra
    second goes). Jobs run one at a time.
#>
$global:ToolkitWorker = $null          # the runspace, once created
$global:ToolkitWorkerCapable = $null   # $null = untested, $true / $false after the probe

function Get-ExoTempModulePath {
    <# The on-disk path of the generated Exchange Online module, or ''. #>
    foreach ($m in @(Get-Module | Where-Object { $_.Name -like 'tmpEXO_*' })) {
        if ($m.Path -and (Test-Path -LiteralPath $m.Path)) { return "$($m.Path)" }
    }
    return ''
}

function Test-ToolJobCapable {
    <#
        Can this process run cloud reads on a worker? Answered once per
        session, with a real probe, and cached. Anything unexpected answers
        "no": the inline path is always correct, only slower to look at.
    #>
    param([switch]$Refresh)
    if (-not $Refresh -and $null -ne $global:ToolkitWorkerCapable) { return [bool]$global:ToolkitWorkerCapable }
    $global:ToolkitWorkerCapable = $false
    try {
        $tmp = Get-ExoTempModulePath
        if (-not $tmp) { return $false }
        $commonPath = Join-Path $PSScriptRoot 'Common.ps1'
        if (-not (Test-Path -LiteralPath $commonPath)) { return $false }

        if (-not $global:ToolkitWorker -or $global:ToolkitWorker.RunspaceStateInfo.State -ne 'Opened') {
            $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            $rs  = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
            $rs.ApartmentState = [System.Threading.ApartmentState]::STA
            $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
            $rs.Open()
            $rs.SessionStateProxy.SetVariable('ToolkitWorkerCommon', $commonPath)
            $rs.SessionStateProxy.SetVariable('ToolkitWorkerExoModule', $tmp)
            $ps = [System.Management.Automation.PowerShell]::Create()
            $ps.Runspace = $rs
            [void]$ps.AddScript(@'
$global:ToolkitDpiAwareSet = $true
$global:ToolkitAllowDirect = $true
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Import-Module $ToolkitWorkerExoModule -Global -DisableNameChecking -ErrorAction Stop
. $ToolkitWorkerCommon
'@)
            $ps.Invoke() | Out-Null
            $hadError = ($ps.Streams.Error.Count -gt 0)
            $ps.Dispose()
            if ($hadError) { try { $rs.Close() } catch { }; return $false }
            $global:ToolkitWorker = $rs
        }

        # The probe: a cheap read through the same path every job will use.
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $global:ToolkitWorker
        [void]$ps.AddScript('@(Invoke-EXOCommand -Name ''Get-AcceptedDomain'' -Parameters @{ ErrorAction = ''Stop'' }).Count')
        $n = $ps.Invoke()
        $ok = ($ps.Streams.Error.Count -eq 0 -and $n.Count -gt 0 -and [int]$n[0] -gt 0)
        $ps.Dispose()
        $global:ToolkitWorkerCapable = $ok
        return $ok
    } catch {
        $global:ToolkitWorkerCapable = $false
        return $false
    }
}

function Start-ToolJob {
    <#
        Run a CLOUD READ without freezing the form, when possible.

            Start-ToolJob -Form $form -Text 'Tracing ...' -Script {
                param($From, $Start, $End)
                @(Invoke-EXOCommand -Name 'Get-MessageTraceV2' -Parameters @{ ... })
            } -ArgumentList $from, $start, $end -State @{ Start = $start } -OnDone {
                param($r, $st)
                if ($r.Failed) { ...$r.Error... ; return }
                ... paint $r.Output, $st.Start ...
            }

        RULES FOR -Script. It is sent to the worker as TEXT: it sees
        Common.ps1 (Invoke-EXOCommand and friends) and its own parameters,
        and nothing else -- no tool functions, no form, no $script:
        variables. Read only: writes belong on the UI thread where the
        operator confirmed them, and the worker's audit context is not the
        form's. -OnDone runs on the UI thread with Output / Failed / Error.

        When the worker is not available the same block runs inline, the
        status bar narrates it, and -OnDone is called just the same. Callers
        cannot tell the difference and should not try.
    #>
    param(
        [Parameter(Mandatory)]$Form,
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][scriptblock]$Script,
        [object[]]$ArgumentList = @(),
        [Parameter(Mandatory)][scriptblock]$OnDone,
        $State = $null
    )
    # -State is handed to -OnDone as its second argument. Use it for the
    # handler's local values (dates, labels built before the call): when the
    # job completes on a timer tick, the click handler that started it has
    # long returned and its locals are gone. Do NOT reach for GetNewClosure
    # here -- a closure runs in its own module, so "$script:" inside it would
    # write a module variable and the rest of the tool would never see it.
    $capable = $false
    try {
        if ($null -eq $global:ToolkitWorkerCapable) { Set-ToolStatus -Form $Form -Busy -Text 'Preparing background worker (first use) ...' }
        $capable = Test-ToolJobCapable
    } catch { $capable = $false }

    if (-not $capable) {
        Set-ToolStatus -Form $Form -Busy -Text $Text
        $res = [ordered]@{ Output = @(); Failed = $false; Error = ''; Inline = $true }
        try { $res.Output = @(& $Script @ArgumentList) }
        catch { $res.Failed = $true; $res.Error = "$($_.Exception.Message)" }
        if ($res.Failed) { Set-ToolStatus -Form $Form -Failed -Text "$Text failed." } else { Set-ToolStatus -Form $Form -Done }
        & $OnDone ([PSCustomObject]$res) $State
        return
    }

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $global:ToolkitWorker
    [void]$ps.AddScript("$Script")
    foreach ($a in $ArgumentList) { [void]$ps.AddArgument($a) }
    $handle = $ps.BeginInvoke()
    Set-ToolStatus -Form $Form -Busy -Text $Text
    $stJob = Get-ToolStatusState -Form $Form
    if ($stJob) { $stJob.JobRunning = $true }
    $Form.UseWaitCursor = $true

    $poll = New-Object System.Windows.Forms.Timer
    $poll.Interval = 150
    $poll.Tag = @{ PS = $ps; Handle = $handle; OnDone = $OnDone; Form = $Form; Text = $Text; State = $State }
    $poll.Add_Tick({
        $j = $this.Tag
        if (-not $j.Handle.IsCompleted) { return }
        $this.Stop()
        $res = [ordered]@{ Output = @(); Failed = $false; Error = ''; Inline = $false }
        try { $res.Output = @($j.PS.EndInvoke($j.Handle)) }
        catch {
            $res.Failed = $true
            $ex = $_.Exception
            $res.Error = $(if ($ex.InnerException -and $ex.InnerException.Message) { "$($ex.InnerException.Message)" } else { "$($ex.Message)" })
        }
        if (-not $res.Failed -and $j.PS.Streams.Error.Count -gt 0) { $res.Failed = $true; $res.Error = "$($j.PS.Streams.Error[0])" }
        try { $j.PS.Dispose() } catch { }
        $j.Form.UseWaitCursor = $false
        $stDone = Get-ToolStatusState -Form $j.Form
        if ($stDone) { $stDone.JobRunning = $false }
        if ($res.Failed) { Set-ToolStatus -Form $j.Form -Failed -Text "$($j.Text) failed." } else { Set-ToolStatus -Form $j.Form -Done }
        try { & $j.OnDone ([PSCustomObject]$res) $j.State }
        catch { Show-HandlerError -ErrorRecord $_ -Context "After '$($j.Text)'" }
        $this.Dispose()
    })
    $poll.Start()
}
#endregion
#endregion

function Set-FormKeyboardDefaults {
    param([Parameter(Mandatory)]$Form)
    $walk = $null
    $walk = {
        param($parent, [string]$Text)
        foreach ($c in $parent.Controls) {
            if ($c -is [System.Windows.Forms.Button] -and "$($c.Text)".Trim() -ieq $Text) { return $c }
            if ($c.Controls -and $c.Controls.Count -gt 0) { $r = & $walk $c $Text; if ($r) { return $r } }
        }
        return $null
    }
    $fixMultiline = $null
    $fixMultiline = {
        param($parent)
        foreach ($c in $parent.Controls) {
            if ($c -is [System.Windows.Forms.TextBox] -and $c.Multiline -and -not $c.ReadOnly) { $c.AcceptsReturn = $true }
            if ($c.Controls -and $c.Controls.Count -gt 0) { & $fixMultiline $c }
        }
    }
    try {
        if ("$($Form.Tag)" -notlike '*NoAcceptButton*' -and -not $Form.AcceptButton) {
            $lu = & $walk $Form 'Look Up'
            if ($lu) { & $fixMultiline $Form; $Form.AcceptButton = $lu }
        }
        if (-not $Form.CancelButton) {
            $cl = & $walk $Form 'Close'
            if ($cl) { $Form.CancelButton = $cl }
        }
    } catch { }
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

    # THE CACHE IS TRUSTED ONLY WHILE ITS MODULES ARE STILL LOADED. A
    # disconnect-and-reconnect replaces the cloud module with a new one, and
    # a map cached before that names a module that no longer exists -- every
    # cloud call then falls through to a guess. Cheap to check, so checked.
    # CACHE HIT RETURNS BEFORE ANYTHING IS ENUMERATED. This is called through
    # Get-EXOCommand on EVERY cloud cmdlet, so whatever sits above this line
    # runs once per Get-Mailbox in a two-thousand-mailbox loop. For one day
    # (2026-09-15) a Get-Module enumeration sat here as a staleness check;
    # that was a per-call cost for a once-per-connection question.
    #
    # STALENESS IS HANDLED WHERE IT IS CHEAP INSTEAD. Connect-ExO.ps1 calls
    # this with -Refresh after a sign-in, and Get-EXOCommand refreshes on a
    # MISS: a module-qualified Get-Command against a module that has been
    # unloaded returns nothing, and that is the one moment the cache can be
    # wrong. Both paths are free until they are needed.
    if (-not $Refresh -and $script:ExchangeModuleMap) { return $script:ExchangeModuleMap }

    # NEVER Get-Module -Name on these: a tmpEXO module's Name IS a full path
    # (...\Temp\tmpEXO_xxxx.yyy) and Get-Module refuses path-shaped names
    # without -ListAvailable. Enumerate once, here, on a miss or a refresh.
    $loaded = @(Get-Module)

    # THE CLOUD MODULE IS THE ONE BEHIND THE LIVE CONNECTION, not the first one
    # found. ExchangeOnlineManagement keeps a tmpEXO_* module for EVERY
    # connection ever made in the process, expired ones included, and a call
    # through an expired one comes back as 'a server side error has occurred'
    # -- not 'token expired'. Found 2026-09-15: four connections, three dead,
    # Get-Mailbox happened to resolve well and every calendar cmdlet did not.
    # Get-ConnectionInformation names the module each connection owns, so the
    # Active one is asked for first; the export-based scan below is the
    # fallback for module versions that do not report it.
    $liveCloudName = $null
    try {
        $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
        if ($gci) {
            $act = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object {
                "$($_.TokenStatus)" -eq 'Active' -and $_.PSObject.Properties['ModuleName'] -and "$($_.ModuleName)"
            })
            # Leaf only: the connection reports 'tmpEXO_xxxx.yyy', the module's
            # Name is the full path to it. Compared by leaf on both sides.
            if ($act.Count -gt 0) { $liveCloudName = [System.IO.Path]::GetFileName("$($act[-1].ModuleName)") }
        }
    } catch { $liveCloudName = $null }

    $onPrem = $null; $cloud = $null; $firstCloud = $null
    foreach ($m in $loaded) {
        $ec = $m.ExportedCommands
        if (-not $ec -or -not $ec.ContainsKey('Get-Mailbox')) { continue }
        if ($ec.ContainsKey('Get-RemoteMailbox')) {
            if (-not $onPrem) { $onPrem = $m }
        } else {
            if (-not $firstCloud) { $firstCloud = $m }
            if ($liveCloudName -and ([System.IO.Path]::GetFileName("$($m.Name)") -eq $liveCloudName)) { $cloud = $m }
        }
    }
    # A live connection was named but no export-scanned module matched it
    # (a version that reports a different shape of name): fall back to the
    # first cloud module rather than to nothing, and say nothing -- a wrong
    # guess here shows up in Get-EXOSessionReport, which prints both names.
    if (-not $cloud) { $cloud = $firstCloud }

    $script:ExchangeModuleMap = [PSCustomObject]@{
        OnPrem = $onPrem
        Cloud  = $cloud
        Both   = [bool]($onPrem -and $cloud)
    }
    return $script:ExchangeModuleMap
}

function Invoke-EXOCommandNoEA {
    <#
        Invoke-EXOCommand for the cmdlets that CANNOT be given -ErrorAction.

        Proven 2026-09-15 on Invoke-ChangeMeetingOrganizer and
        Get-CalendarDiagnosticObjects: the identical call works bare and fails
        with 'a server side error has occurred' the moment -ErrorAction Stop is
        on it -- typed or splatted. -ErrorVariable is accepted. Get-Mailbox and
        the rest tolerate -ErrorAction, so this is not the default path; it is
        the one to use when a REST cmdlet fails generically for no reason the
        session report can show.

        Errors are still surfaced as a throw, so a caller's try/catch works as
        it would with -ErrorAction Stop: any ErrorAction key in Parameters is
        removed, the call runs with -ErrorVariable, and the first error record
        collected is thrown.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    $p = @{}
    foreach ($k in $Parameters.Keys) { if ($k -ne 'ErrorAction') { $p[$k] = $Parameters[$k] } }
    $cmd = Get-EXOCommand -Name $Name
    $ev  = $null
    $out = & $cmd @p -ErrorVariable ev 2>$null
    if ($ev -and @($ev).Count -gt 0) { throw @($ev)[0] }
    return $out
}

function Get-EXOSessionReport {
    <#
        What this process's Exchange Online plumbing looks like, as text, for
        pasting into an error message. Every connection with its UPN, token
        state and module; which module the map has chosen as the cloud; and,
        when a cmdlet name is given, which module that name resolves to.

        Exists because 'a server side error has occurred' says nothing, and
        the three times it was chased in 2026-09 the answer was in this
        report every time: dead sessions, a stale module map, or the wrong
        account signed in. A tool that prints this with the error saves the
        round trip to a prompt.
    #>
    param([string]$CommandName)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('EXCHANGE ONLINE SESSION IN THIS PROCESS')
    try {
        $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
        if (-not $gci) { [void]$sb.AppendLine('  Get-ConnectionInformation is not available (module v2 or older).') }
        else {
            $conns = @(Get-ConnectionInformation -ErrorAction Stop)
            if ($conns.Count -eq 0) { [void]$sb.AppendLine('  No Exchange Online connection at all.') }
            foreach ($c in $conns) {
                $mn = $(if ($c.PSObject.Properties['ModuleName']) { "$($c.ModuleName)" } else { '?' })
                [void]$sb.AppendLine("  $($c.UserPrincipalName)   token=$($c.TokenStatus)   state=$($c.State)   module=$mn")
            }
        }
    } catch { [void]$sb.AppendLine("  Could not list connections: $($_.Exception.Message)") }
    try {
        $map = Get-ExchangeModuleMap
        [void]$sb.AppendLine("  Module map: cloud=$(if ($map.Cloud) { $map.Cloud.Name } else { '(none)' })  onprem=$(if ($map.OnPrem) { $map.OnPrem.Name } else { '(none)' })")
    } catch { }
    if ($CommandName) {
        try {
            $cmd = Get-EXOCommand -Name $CommandName
            [void]$sb.AppendLine("  $CommandName resolves to: $($cmd.ModuleName)\$($cmd.Name)  ($($cmd.CommandType))")
        } catch { [void]$sb.AppendLine("  $CommandName does not resolve: $($_.Exception.Message)") }
    }
    return $sb.ToString()
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

    if (Test-TenantCapability -Name 'MicrosoftGraph') {
        $defs += [PSCustomObject]@{
            Id      = 'Graph'
            Label   = 'Microsoft Graph'
            # Authentication ONLY. The full Microsoft.Graph SDK is forty-odd
            # modules and several seconds of import; every Graph tool here
            # calls REST through Invoke-MgGraphRequest and needs nothing else.
            Name    = 'Microsoft.Graph.Authentication'
            Probe   = 'Connect-MgGraph'
            Purpose = 'Sign-in Activity, Risky Users, Auth Methods, Licence Assignment, Service Health, and the Graph pill in the console.'
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
        $answer = Show-ToolDialog(
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

    # Graph is a cheap probe -- Get-MgContext is an in-memory read -- so it
    # runs on every pass, including startup.
    $gr = Test-GraphConnectionState
    $map['Graph'] = [PSCustomObject]@{
        Id = 'Graph'; Label = 'Microsoft Graph'
        Available = [bool]$gr.Connected; Unknown = $false; Detail = "$($gr.Detail)"
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
                 EMS = 'OnPremExchange'; AD  = 'OnPremAD'; Graph = 'MicrosoftGraph' }
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
        # Exchange happened to return first -- a secondary office, not the
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


#region Mailbox editing -- shared logic ------------------------------------------
<#
    ONE IMPLEMENTATION OF EVERY MAILBOX WRITE, CALLED FROM MORE THAN ONE PLACE.

    Added 2026-09-18 for MailboxEditor.ps1, which collapses fourteen
    single-purpose tools into one tabbed window. The logic lives HERE rather
    than in that tool for two reasons:

      1. Unified Lookup edits some of the same things. Two implementations of
         'change the primary address' would be two audit trails and one of them
         eventually wrong -- the exact failure the Attributes/permissions split
         was written to avoid.

      2. It gives a safe migration path. A single-purpose tool can be pointed
         at the shared function and proved to still work BEFORE the editor
         calls the same function, so nothing is rewritten from scratch and
         re-tested from zero.
#>

<#
    WHERE DOES THIS FIELD ACTUALLY LIVE?

    THE FAILURE THIS EXISTS TO PREVENT: on a mailbox synced from Active
    Directory, several of these properties are MASTERED ON-PREMISES. Exchange
    Online accepts the write, reports success, and Entra Connect puts the old
    value back on the next cycle -- typically within thirty minutes. Nothing
    errors. Nothing appears in the log. The operator sees the change, walks
    away, and the setting silently reverts.

    Hide-Unhide.ps1 already knew this and refuses the cloud write for a synced
    mailbox; ForwardingEditor.ps1 did not. In a window that can change thirty
    different things, 'the tool that happened to be written by someone who knew'
    is not a strategy. So every writable field is declared here.

    THREE VALUES, AND THE THIRD ONE IS THE HONEST ONE:

      Cloud        Mastered in Exchange Online. Safe to write for any mailbox,
                   synced or not. Mailbox and folder permissions, holds,
                   automatic replies and mailbox type all behave this way --
                   they are properties of the MAILBOX, which only exists in
                   Exchange Online.

      Directory    Mastered in Active Directory when the object is synced.
                   Writing it in the cloud is the silent-revert case above, so
                   for a synced object the editor routes to AD or refuses and
                   says which. For a cloud-only object these are ordinary
                   cloud writes.

      Unverified   Behaviour depends on the hybrid configuration and I have not
                   proved it on this tenant. The write goes to the cloud and
                   the operator is TOLD it may not stick on a synced mailbox.
                   Being visibly unsure is worth more than a confident guess
                   that reverts in half an hour.

    Moving an entry out of Unverified needs one test, not an opinion: set it on
    a synced mailbox, wait for a sync cycle, read it back.
#>
$script:MailboxFieldOrigin = @{
    # --- Mastered in the directory when synced ---------------------------
    # proxyAddresses. The classic one: a cloud-side address change on a synced
    # object is undone wholesale on the next sync.
    'EmailAddresses'                    = 'Directory'
    'PrimarySmtpAddress'                = 'Directory'
    # msExchHideFromAddressLists -- proved by Hide-Unhide.ps1, which refuses
    # the cloud write for a synced mailbox for exactly this reason.
    'HiddenFromAddressListsEnabled'     = 'Directory'
    'DisplayName'                       = 'Directory'
    'Alias'                             = 'Directory'   # mailNickname
    'CustomAttribute'                   = 'Directory'   # extensionAttribute1-15

    # --- Properties of the mailbox itself, so always Exchange Online -----
    # These have no on-premises counterpart to overwrite them: the mailbox
    # lives in the cloud and so do its permissions, folders and policies.
    'FullAccess'                        = 'Cloud'
    'SendAs'                            = 'Cloud'
    'FolderPermission'                  = 'Cloud'
    'LitigationHoldEnabled'             = 'Cloud'
    'RetentionPolicy'                   = 'Cloud'
    'RetentionHoldEnabled'              = 'Cloud'
    'AutomaticReplies'                  = 'Cloud'
    'ArchiveEnabled'                    = 'Cloud'
    'MailboxType'                       = 'Cloud'       # Set-Mailbox -Type
    'MessageCopyForSentAsEnabled'       = 'Cloud'
    'MessageCopyForSendOnBehalfEnabled' = 'Cloud'
    'MailTip'                           = 'Cloud'
    'MaxSendSize'                       = 'Cloud'
    'MaxReceiveSize'                    = 'Cloud'
    'RecipientLimits'                   = 'Cloud'
    'RetainDeletedItemsFor'             = 'Cloud'

    # --- Genuinely uncertain in hybrid -----------------------------------
    # GrantSendOnBehalfTo maps to publicDelegates, and the delivery
    # restrictions map to authOrig / unauthOrig / dLMemSubmitPerms. Whether
    # those flow from AD depends on which attributes the sync is configured to
    # write, which varies per deployment. ForwardingAddress maps to
    # altRecipient; ForwardingSmtpAddress has no AD counterpart and is the
    # safer of the two. Marked Unverified rather than guessed.
    'GrantSendOnBehalfTo'               = 'Unverified'
    'ForwardingAddress'                 = 'Unverified'
    'ForwardingSmtpAddress'             = 'Unverified'
    'DeliverToMailboxAndForward'        = 'Unverified'
    'AcceptMessagesOnlyFrom'            = 'Unverified'
    'RejectMessagesFrom'                = 'Unverified'
    'RequireSenderAuthenticationEnabled'= 'Unverified'
}

function Get-MailboxFieldOrigin {
    <#
        Where a named field is mastered: Cloud, Directory or Unverified.
        An unknown name returns Unverified rather than Cloud -- a field nobody
        has classified is exactly the one most likely to surprise somebody.
    #>
    param([Parameter(Mandatory)][string]$Field)
    $f = "$Field".Trim()
    if ($script:MailboxFieldOrigin.ContainsKey($f)) { return $script:MailboxFieldOrigin[$f] }
    # CustomAttribute1..15 share one entry.
    if ($f -match '^CustomAttribute\d{1,2}$') { return $script:MailboxFieldOrigin['CustomAttribute'] }
    return 'Unverified'
}

function Get-MailboxWriteTarget {
    <#
        Can this field be written for THIS mailbox, and where does the write go?

        Combines the field's origin with whether the object is synced:

            Ok        write it in Exchange Online and expect it to stick
            Directory synced object, directory-mastered field -- the editor
                      must route to AD or refuse; Reason says which and why
            Warn      Unverified field on a synced object: the write is allowed
                      and the caller shows Reason so nobody is surprised later

        A cloud-only mailbox is always Ok, whatever the field: there is no
        directory copy to overwrite it.
    #>
    param(
        [Parameter(Mandatory)][string]$Field,
        [bool]$IsSynced
    )
    $origin = Get-MailboxFieldOrigin -Field $Field
    $r = [PSCustomObject]@{ Field = $Field; Origin = $origin; State = 'Ok'; Reason = '' }
    if (-not $IsSynced) { return $r }

    switch ($origin) {
        'Directory' {
            $r.State  = 'Directory'
            $r.Reason = "$Field is mastered in Active Directory for a synced mailbox. " +
                        'Exchange Online would accept the change and the next directory sync would put the old value back, usually within half an hour. Change it on-premises instead.'
        }
        'Unverified' {
            $r.State  = 'Warn'
            $r.Reason = "$Field may be mastered in Active Directory on a synced mailbox, depending on how this hybrid is configured. " +
                        'The change will be made in Exchange Online -- check it again after a sync cycle before treating it as done.'
        }
    }
    return $r
}

function Resolve-MailboxTarget {
    <#
        The one mailbox lookup every tab shares.

        Returns Ok / Smtp / Display / Type / IsSynced / Raw, or Ok=$false with
        a Reason that names the problem in words rather than handing back the
        cmdlet's own error.

        IsSynced comes from IsDirSynced, which is what Hide-Unhide already
        keys on -- one definition of 'synced' across the toolkit rather than
        two that disagree at the edges.
    #>
    param([Parameter(Mandatory)][string]$Smtp)
    $r = [PSCustomObject]@{
        Ok = $false; Smtp = ''; Display = ''; Type = ''; IsSynced = $false
        Alias = ''; Raw = $null; Reason = ''
    }
    $t = "$Smtp".Trim()
    if (-not $t) { $r.Reason = 'No mailbox given.'; return $r }
    $mb = $null
    try {
        $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $t; ErrorAction = 'Stop' }
    } catch {
        $r.Reason = "not found in Exchange Online: $($_.Exception.Message)"
        return $r
    }
    $r.Ok       = $true
    $r.Raw      = $mb
    $r.Smtp     = "$($mb.PrimarySmtpAddress)"
    $r.Display  = "$($mb.DisplayName)"
    $r.Type     = "$($mb.RecipientTypeDetails)"
    $r.Alias    = "$($mb.Alias)"
    try { $r.IsSynced = [bool]$mb.IsDirSynced } catch { $r.IsSynced = $false }
    return $r
}
<#
    MAILBOX ACCESS -- full access, send as, send on behalf.

    Three different cmdlet families that the operator thinks of as one thing.
    Kept together so a tool showing 'who can get at this mailbox' asks one
    question instead of three, and so the WRITE side cannot drift from the
    READ side the way it does when each tool rolls its own.
#>

function Get-MailboxAccessEntry {
    <#
        Everyone with access to a mailbox, in one list, one row per grant.

        Kind is FullAccess, SendAs or SendOnBehalf. Inherited and self entries
        are dropped: every mailbox has NT AUTHORITY\SELF and a pile of
        inherited admin ACEs, and showing them buries the two rows that
        actually answer the question.

        Best-effort per kind. A tenant where Get-RecipientPermission is not
        available to the operator still gets the full-access list rather than
        an empty grid and no explanation -- the Note says what was missed.
    #>
    param([Parameter(Mandatory)][string]$Smtp)
    $out  = New-Object System.Collections.ArrayList
    $note = ''

    try {
        foreach ($p in @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{
                            Identity = $Smtp; ErrorAction = 'Stop' })) {
            if ($p.IsInherited) { continue }
            $u = "$($p.User)"
            if ($u -match '(?i)NT AUTHORITY\\SELF|^S-1-5-|Discovery Management') { continue }
            if (@($p.AccessRights) -notcontains 'FullAccess') { continue }
            [void]$out.Add([PSCustomObject]@{
                Kind = 'FullAccess'; User = $u
                Detail = $(if ($p.Deny) { 'DENY' } else { '' })
                AutoMap = ''   # not returned by Get-MailboxPermission; see the tab note
                Raw = $p
            })
        }
    } catch { $note += "Full access could not be read ($($_.Exception.Message)). " }

    try {
        foreach ($p in @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{
                            Identity = $Smtp; ErrorAction = 'Stop' })) {
            $u = "$($p.Trustee)"
            if ($u -match '(?i)NT AUTHORITY\\SELF|^S-1-5-') { continue }
            if (@($p.AccessRights) -notcontains 'SendAs') { continue }
            [void]$out.Add([PSCustomObject]@{
                Kind = 'SendAs'; User = $u
                Detail = $(if ("$($p.AccessControlType)" -eq 'Deny') { 'DENY' } else { '' })
                AutoMap = ''; Raw = $p
            })
        }
    } catch { $note += "Send As could not be read ($($_.Exception.Message)). " }

    # Send on Behalf is a property of the mailbox, not an ACL, so it comes off
    # the object rather than from a permission cmdlet.
    try {
        $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        foreach ($d in @($mb.GrantSendOnBehalfTo)) {
            if (-not "$d") { continue }
            [void]$out.Add([PSCustomObject]@{ Kind = 'SendOnBehalf'; User = "$d"; Detail = ''; AutoMap = ''; Raw = $d })
        }
    } catch { $note += "Send on Behalf could not be read ($($_.Exception.Message)). " }

    return [PSCustomObject]@{ Entries = @($out.ToArray()); Note = $note.Trim() }
}

function Add-MailboxAccessEntry {
    <#
        Grant one kind of access to one trustee.

        AUTOMAPPING IS A PARAMETER, NOT A DEFAULT, because the two common
        cases want opposite answers: a person covering a colleague's mailbox
        wants it to appear in Outlook automatically; a shared mailbox that
        thirty people have access to does not want to open in thirty Outlook
        profiles. -AutoMap $false is also how the automap repair works.

        Send on Behalf is an @{Add=} update on the mailbox rather than a new
        ACL, so it goes through Set-Mailbox.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][ValidateSet('FullAccess', 'SendAs', 'SendOnBehalf')][string]$Kind,
        [Parameter(Mandatory)][string]$Trustee,
        [bool]$AutoMap = $true
    )
    switch ($Kind) {
        'FullAccess' {
            Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
                Identity = $Smtp; User = $Trustee; AccessRights = 'FullAccess'
                InheritanceType = 'All'; AutoMapping = $AutoMap; Confirm = $false; ErrorAction = 'Stop'
            } | Out-Null
        }
        'SendAs' {
            Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
                Identity = $Smtp; Trustee = $Trustee; AccessRights = 'SendAs'
                Confirm = $false; ErrorAction = 'Stop'
            } | Out-Null
        }
        'SendOnBehalf' {
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
                Identity = $Smtp; GrantSendOnBehalfTo = @{ Add = $Trustee }; ErrorAction = 'Stop'
            } | Out-Null
        }
    }
}

function Remove-MailboxAccessEntry {
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][ValidateSet('FullAccess', 'SendAs', 'SendOnBehalf')][string]$Kind,
        [Parameter(Mandatory)][string]$Trustee
    )
    switch ($Kind) {
        'FullAccess' {
            Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{
                Identity = $Smtp; User = $Trustee; AccessRights = 'FullAccess'
                InheritanceType = 'All'; Confirm = $false; ErrorAction = 'Stop'
            } | Out-Null
        }
        'SendAs' {
            Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{
                Identity = $Smtp; Trustee = $Trustee; AccessRights = 'SendAs'
                Confirm = $false; ErrorAction = 'Stop'
            } | Out-Null
        }
        'SendOnBehalf' {
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
                Identity = $Smtp; GrantSendOnBehalfTo = @{ Remove = $Trustee }; ErrorAction = 'Stop'
            } | Out-Null
        }
    }
}

function Repair-MailboxAutoMapping {
    <#
        The automap repair, as one function instead of a tool.

        WHAT IS ACTUALLY WRONG WHEN SOMEBODY ASKS FOR THIS: automapping is
        driven by msExchDelegateListLink, which Exchange stamps when Full
        Access is granted. It does not always clear when the permission is
        removed, and it does not always appear when it should -- so a mailbox
        either keeps opening for somebody who no longer has access, or refuses
        to appear for somebody who does. Removing and re-adding the permission
        restamps the link, which is the whole repair.

        -AutoMap chooses which way to restamp it: $true to make it appear,
        $false to stop it appearing while KEEPING the access, which is the
        request for a busy shared mailbox.

        There is a gap between the remove and the add where the trustee has no
        access. That is seconds, and unavoidable through this interface -- but
        it is a gap that MUST close. If the re-add fails, the repair has taken
        access away and not given it back, so the add is retried and a final
        failure says in plain words what state the mailbox is in rather than
        surfacing whatever EXO said about the second attempt.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$Trustee,
        [bool]$AutoMap = $true
    )
    Remove-MailboxAccessEntry -Smtp $Smtp -Kind 'FullAccess' -Trustee $Trustee
    Start-Sleep -Seconds 2

    $last = $null
    foreach ($attempt in 1..3) {
        try {
            Add-MailboxAccessEntry -Smtp $Smtp -Kind 'FullAccess' -Trustee $Trustee -AutoMap $AutoMap
            return
        } catch {
            # Throttling is the usual reason the second call fails, and it
            # clears on its own. Back off and try again rather than leaving
            # the trustee locked out over a transient 429.
            $last = $_
            if ($attempt -lt 3) { Start-Sleep -Seconds (4 * $attempt) }
        }
    }
    throw ("Full Access was REMOVED from $Trustee on $Smtp and could not be granted back " +
           "after three attempts -- that trustee currently has NO access to the mailbox. " +
           "Grant it again on the Access tab. The last error was: $($last.Exception.Message)")
}

<#
    FOLDER PERMISSIONS -- calendar, inbox, contacts, tasks, notes.

    One implementation for all five. CalendarPermissions.ps1 did the calendar
    and nothing did the inbox, which is the gap this closes: the cmdlets are
    identical, only the folder path changes.
#>

# The well-known folders worth offering, and the name each takes in an
# Identity of the form <mailbox>:\<folder>. Microsoft resolve these by
# LOCALISED display name, so a mailbox created in another language can fail on
# the English word -- Get-MailboxFolderPath below falls back to asking the
# mailbox for its own folder names rather than guessing.
$script:MailboxWellKnownFolder = @(
    [PSCustomObject]@{ Key = 'Calendar'; Label = 'Calendar'; Type = 'Calendar' }
    [PSCustomObject]@{ Key = 'Inbox';    Label = 'Inbox';    Type = 'Inbox' }
    [PSCustomObject]@{ Key = 'Contacts'; Label = 'Contacts'; Type = 'Contacts' }
    [PSCustomObject]@{ Key = 'Tasks';    Label = 'Tasks';    Type = 'Tasks' }
    [PSCustomObject]@{ Key = 'Notes';    Label = 'Notes';    Type = 'Notes' }
)

# The roles Exchange accepts on a folder, most useful first. Availability-only
# and LimitedDetails are calendar-specific and harmless elsewhere -- Exchange
# rejects them on a non-calendar folder with a clear message, which is better
# than this list guessing per folder and hiding a role somebody wanted.
$script:MailboxFolderRole = @(
    'Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author',
    'NonEditingAuthor', 'Reviewer', 'Contributor',
    'AvailabilityOnly', 'LimitedDetails', 'None'
)

function Get-MailboxFolderPath {
    <#
        The Identity string for a well-known folder on this mailbox.

        Asks the mailbox for its folder list and matches on FOLDER TYPE, not
        on the English name, then falls back to the English name if the
        statistics call is unavailable. A mailbox created in another language
        has a localised Calendar folder, and '<mailbox>:\Calendar' simply does
        not resolve against it.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$FolderType
    )
    # CACHED PER MAILBOX AND FOLDER, because the Folder Permissions tab asks
    # for five folders in a row and each answer costs a call.
    if (-not $script:MailboxFolderPathCache) { $script:MailboxFolderPathCache = @{} }
    $ck = "$Smtp|$FolderType".ToLowerInvariant()
    if ($script:MailboxFolderPathCache.ContainsKey($ck)) { return $script:MailboxFolderPathCache[$ck] }
    try {
        # -FolderScope, NOT the whole mailbox. Without it this enumerates
        # every folder in the mailbox -- on a large one that is tens of
        # thousands of rows, five times over, on the UI thread, which reads
        # as the window having hung.
        $folders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{
                        Identity = $Smtp; FolderScope = $FolderType; ErrorAction = 'Stop' })
        $hit = @($folders | Where-Object { "$($_.FolderType)" -eq $FolderType }) | Select-Object -First 1
        if ($hit -and "$($hit.Name)") {
            $script:MailboxFolderPathCache[$ck] = "$Smtp`:\$($hit.Name)"
            return $script:MailboxFolderPathCache[$ck]
        }
    } catch { }
    return "$Smtp`:\$FolderType"
}

function Get-MailboxFolderAccess {
    <#
        Who has rights on one folder. Default and Anonymous are KEPT, unlike
        the access list above -- on a calendar they are the two entries most
        likely to be the answer ('why can everyone see my appointments').
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$FolderType
    )
    $path = Get-MailboxFolderPath -Smtp $Smtp -FolderType $FolderType
    $out  = New-Object System.Collections.ArrayList
    try {
        foreach ($p in @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{
                            Identity = $path; ErrorAction = 'Stop' })) {
            [void]$out.Add([PSCustomObject]@{
                User   = "$($p.User)"
                Rights = (@($p.AccessRights) -join ', ')
                Folder = $FolderType
                Path   = $path
                Raw    = $p
            })
        }
    } catch {
        return [PSCustomObject]@{ Ok = $false; Entries = @(); Path = $path; Reason = "$($_.Exception.Message)" }
    }
    return [PSCustomObject]@{ Ok = $true; Entries = @($out.ToArray()); Path = $path; Reason = '' }
}

function Set-MailboxFolderAccess {
    <#
        Grant or change a folder right.

        ADD-OR-SET, deliberately. Exchange has separate Add- and Set- cmdlets
        and Add- fails when an entry already exists; the operator asking for
        'give Jane Reviewer on the calendar' does not care which of those it
        is, and making them find out by reading an error is not useful. Add
        first, and on 'already exists' fall through to Set.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$FolderType,
        [Parameter(Mandatory)][string]$User,
        [Parameter(Mandatory)][string]$Role
    )
    $path = Get-MailboxFolderPath -Smtp $Smtp -FolderType $FolderType
    try {
        Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' -Parameters @{
            Identity = $path; User = $User; AccessRights = $Role; ErrorAction = 'Stop'
        } | Out-Null
        return 'Added'
    } catch {
        if ("$($_.Exception.Message)" -notmatch '(?i)already exist|existing permission') { throw }
        Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{
            Identity = $path; User = $User; AccessRights = $Role; ErrorAction = 'Stop'
        } | Out-Null
        return 'Updated'
    }
}

function Remove-MailboxFolderAccess {
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$FolderType,
        [Parameter(Mandatory)][string]$User
    )
    $path = Get-MailboxFolderPath -Smtp $Smtp -FolderType $FolderType
    Invoke-EXOCommand -Name 'Remove-MailboxFolderPermission' -Parameters @{
        Identity = $path; User = $User; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
<#
    PRIMARY ADDRESS -- MOVED HERE FROM Unified-Lookup.ps1 ON 2026-09-18.

    Unchanged. It moved because MailboxEditor.ps1 edits addresses too, and a
    second implementation would have been a second chance to get the X500
    handling wrong. The docstring below is the original and still describes
    why every entry is carried across; Unified Lookup now calls this copy.
#>
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

#region Microsoft Graph -----------------------------------------------------------
<#
    ONE APP, DELEGATED ONLY, SCOPES DECLARED BY THE TOOL THAT NEEDS THEM.

    The tenant registers a single Graph application and the console signs in
    THROUGH it -- no secrets, no application permissions, and every call runs
    as the operator with the roles the operator holds. The app's client id and
    the tenant id are tenant facts, so they live in Tenant.json (Graph block)
    and nowhere else; a .ps1 that named them would fail the leakage scan and
    would be wrong for the next tenant anyway.

    HOW SCOPES WORK. Each Graph tool declares the delegated permissions it
    needs in its header:

        # TOOLKIT-GRAPH-SCOPES: AuditLog.Read.All, User.Read.All

    Connect-MSGraph.ps1 scans every script in the folder for that marker and
    signs in once with the UNION -- so one prompt covers every tool, and a tool
    still asks for nothing the toolkit does not use. A tool checks its own
    scopes against the live context before it does anything (Assert-
    ToolkitGraphScope) and, when one is missing, says which and points at the
    Connect button rather than failing at the call.

    TOOLS NEVER SIGN IN THEMSELVES. Tools run inside the console's process
    while the console's own window holds the thread in a modal loop, and an
    interactive sign-in raised there draws and never accepts input -- the same
    fault that moved the Exchange Online connect out of the button handler.
    The connect script is the one place that signs in, and the console runs
    it after its own window has closed. A tool run from a bare prompt with no
    context gets the same message and runs Connect-MSGraph.ps1 by hand.

    REST, NOT THE GENERATED CMDLETS. Only Microsoft.Graph.Authentication is
    required; every call goes through Invoke-MgGraphRequest with a URI read
    from the v1.0 reference. The reasoning is written out in PimActivation.ps1
    and has not changed: the SDK is forty modules and the cmdlet names have
    been renamed between versions; the REST contract is the stable thing.
#>

function Get-TenantGraphConfig {
    <#
        The Graph block of Tenant.json, normalised. Configured is true only
        when the capability is on AND a client id is present: one without the
        other is a half-configured tenant and reads as 'absent', which greys
        the Graph tools with a tooltip that says why instead of letting them
        fail at sign-in.
    #>
    $g = $null
    try { $g = (Get-TenantConfig).Graph } catch { }
    $client = ''; $tenant = ''
    if ($g) {
        try { $client = "$($g.ClientId)".Trim() } catch { }
        try { $tenant = "$($g.TenantId)".Trim() } catch { }
    }
    $on = $false
    try { $on = Test-TenantCapability -Name 'MicrosoftGraph' } catch { }
    return [PSCustomObject]@{
        Enabled    = [bool]$on
        ClientId   = $client
        TenantId   = $tenant
        Configured = ([bool]$on -and [string]::IsNullOrWhiteSpace($client) -eq $false)
    }
}

function Test-GraphConnectionState {
    <#
        Is THIS process signed in to Microsoft Graph, and through which app?

        Connected means a context with an account exists. WrongApp is set when
        that context was made through some other application -- the Graph
        PowerShell default app, say, from a hand-typed Connect-MgGraph -- which
        still works for any scope it happens to carry but is not the tenant's
        app, and Connect-MSGraph.ps1 replaces it rather than adding to it.
    #>
    param()
    $r = [ordered]@{ Connected = $false; Detail = ''; Account = ''; ClientId = ''; Scopes = @(); WrongApp = $false }
    if (-not (Get-Command -Name 'Get-MgContext' -ErrorAction SilentlyContinue)) {
        $r.Detail = 'Not connected -- Microsoft.Graph.Authentication is not loaded in this process. Use the MSGraph connect button.'
        return [PSCustomObject]$r
    }
    $ctx = $null
    try { $ctx = Get-MgContext -ErrorAction Stop } catch { $ctx = $null }
    if (-not $ctx -or -not "$($ctx.Account)") {
        $r.Detail = 'Not connected -- no Microsoft Graph sign-in in this window. Use the MSGraph connect button.'
        return [PSCustomObject]$r
    }
    $r.Connected = $true
    $r.Account   = "$($ctx.Account)"
    $r.ClientId  = "$($ctx.ClientId)"
    $r.Scopes    = @($ctx.Scopes | ForEach-Object { "$_" })
    $cfg = Get-TenantGraphConfig
    if ($cfg.ClientId -and $r.ClientId -and ($cfg.ClientId -ne $r.ClientId)) {
        $r.WrongApp = $true
        $r.Detail = "Connected as $($r.Account), but through a different application than Tenant.json names. Use the MSGraph connect button to sign in through the tenant's app."
    } else {
        $r.Detail = "Connected to Microsoft Graph as $($r.Account) ($($r.Scopes.Count) scope(s))."
    }
    return [PSCustomObject]$r
}

function Get-ScriptGraphScope {
    <#
        The scopes a script declares in its first 40 lines:

            # TOOLKIT-GRAPH-SCOPES: A.Read.All, B.ReadWrite.All

        Header only, for the same reason every other marker is read that way:
        this runs over every script in the folder at connect time.
    #>
    param([Parameter(Mandatory)][string]$Path)
    $out = @()
    try {
        $head = Get-Content -LiteralPath $Path -TotalCount 40 -ErrorAction Stop
        foreach ($line in $head) {
            if ($line -match '^\s*#\s*TOOLKIT-GRAPH-SCOPES:\s*(.+)$') {
                foreach ($s in ($Matches[1] -split '[,\s]+')) {
                    $s = "$s".Trim()
                    if ($s -and $out -notcontains $s) { $out += $s }
                }
            }
        }
    } catch { }
    return @($out)
}

function Get-ToolkitGraphScopeCatalogue {
    <#
        Every scope any script in the toolkit folder declares, and which
        scripts declare it. This is what the connect script signs in with.

        The union is the whole point: one sign-in, one consent screen, and the
        set grows only when a tool that needs more is added -- never by an
        operator guessing at a list.
    #>
    param([string]$Root = $PSScriptRoot)
    $byScope = @{}
    foreach ($f in @(Get-ChildItem -LiteralPath $Root -Filter '*.ps1' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        foreach ($s in (Get-ScriptGraphScope -Path $f.FullName)) {
            if (-not $byScope.ContainsKey($s)) { $byScope[$s] = @() }
            $byScope[$s] += $f.Name
        }
    }
    $rows = @()
    foreach ($k in ($byScope.Keys | Sort-Object)) {
        $rows += [PSCustomObject]@{ Scope = $k; Scripts = @($byScope[$k]) }
    }
    return @($rows)
}

function Connect-ToolkitGraph {
    <#
        Sign this process in to Microsoft Graph through the tenant's app.

        Called by Connect-MSGraph.ps1 and by nothing else that runs under the
        console (see the region header). Reconnects only when something is
        actually missing: a context on the right app that already carries
        every requested scope is returned as-is, because every reconnect is a
        sign-in prompt.

        A context on the WRONG app is disconnected first. Connect-MgGraph
        with a different client id on top of a live context does not always
        replace it cleanly, and 'signed in, but not through our app' is the
        state Test-GraphConnectionState exists to report.
    #>
    param(
        [string[]]$Scopes = @(),
        [switch]$Force
    )
    $cfg = Get-TenantGraphConfig
    if (-not $cfg.Configured) {
        throw 'Microsoft Graph is not configured for this tenant: Tenant.json needs Capabilities.MicrosoftGraph = true and a Graph.ClientId.'
    }
    if (-not (Get-Command -Name 'Connect-MgGraph' -ErrorAction SilentlyContinue)) {
        throw 'Microsoft.Graph.Authentication is not loaded. Import-Module Microsoft.Graph.Authentication first.'
    }

    $want = @($Scopes | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    if ($want.Count -eq 0) { $want = @('User.Read') }

    $state = Test-GraphConnectionState
    if ($state.Connected -and -not $state.WrongApp -and -not $Force) {
        $missing = @($want | Where-Object { $state.Scopes -notcontains $_ })
        if ($missing.Count -eq 0) { return (Get-MgContext) }
        # Keep what the context already has -- a reconnect that asks for LESS
        # silently drops scopes another tool was relying on.
        $want = @(@($state.Scopes) + $want | Sort-Object -Unique)
    }
    if ($state.Connected -and $state.WrongApp) {
        try { Disconnect-MgGraph -ErrorAction Stop | Out-Null } catch { }
    }

    $p = @{ ClientId = $cfg.ClientId; Scopes = $want; ErrorAction = 'Stop' }
    if ($cfg.TenantId) { $p.TenantId = $cfg.TenantId }
    $cmd = Get-Command Connect-MgGraph -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Parameters.ContainsKey('NoWelcome')) { $p.NoWelcome = $true }
    Connect-MgGraph @p | Out-Null
    return (Get-MgContext)
}

function Assert-ToolkitGraphScope {
    <#
        The gate every Graph tool runs before its first call.

        Returns an object rather than throwing, because the tool wants to put
        the sentence on its own status line and disable its buttons -- not
        pop an error dialog over a window that has not painted yet.

            Ok       true when the context is on the right app and holds every scope
            Missing  the scopes it does not hold
            Message  one sentence the tool can show verbatim
    #>
    param([string[]]$Scopes = @())
    $state = Test-GraphConnectionState
    if (-not $state.Connected) {
        return [PSCustomObject]@{ Ok = $false; Missing = @($Scopes); Message = $state.Detail; State = $state }
    }
    if ($state.WrongApp) {
        return [PSCustomObject]@{ Ok = $false; Missing = @($Scopes); Message = $state.Detail; State = $state }
    }
    $missing = @($Scopes | Where-Object { $state.Scopes -notcontains $_ })
    if ($missing.Count -gt 0) {
        return [PSCustomObject]@{
            Ok = $false; Missing = $missing; State = $state
            Message = "This sign-in does not hold $($missing -join ', '). Press the MSGraph connect button to sign in again -- it picks the scope up from this tool's header. If the prompt refuses it, the tenant's Graph app does not have that permission yet."
        }
    }
    return [PSCustomObject]@{ Ok = $true; Missing = @(); Message = $state.Detail; State = $state }
}

function Format-GraphError {
    <#
        The sentence from a failed Invoke-MgGraphRequest, with the HTTP status
        and Graph's own message when they are there, and a hint on 403 --
        which is the one that reads as 'Forbidden' and means one of two very
        different things: the app was not granted the scope, or the OPERATOR
        does not hold the Entra role the API also demands. Both look the same
        from here, so both are named.
    #>
    param([Parameter(Mandatory)]$ErrorRecord)
    $msg = "$($ErrorRecord.Exception.Message)"
    $detail = ''
    try {
        $raw = "$($ErrorRecord.ErrorDetails.Message)"
        if ($raw -and $raw.TrimStart().StartsWith('{')) {
            $j = $raw | ConvertFrom-Json
            if ($j.error -and $j.error.message) { $detail = "$($j.error.message)" }
        }
    } catch { }
    if ($detail -and $msg -notlike "*$detail*") { $msg = "$msg -- $detail" }
    if ($msg -match '\b403\b|Forbidden|Authorization_RequestDenied|AccessDenied') {
        $msg += ' (403: either the Graph app was not granted a permission this call needs, or your account does not hold the Entra role the API also requires. The tool help names both.)'
    }
    return $msg
}

function Invoke-ToolkitGraph {
    <#
        One Graph call, relative or absolute URI, with paging on request.

            Invoke-ToolkitGraph -Uri "users/$id/licenseDetails"
            Invoke-ToolkitGraph -Uri "auditLogs/signIns?`$filter=..." -All
            Invoke-ToolkitGraph -Uri "users/$id/assignLicense" -Method POST -Body @{...}

        -All follows @odata.nextLink and returns the concatenated .value; the
        default returns the raw response so callers that want @odata.count or
        a single object get exactly what Graph sent. -MaxRows stops paging at
        a ceiling, because a broad sign-in search can page for minutes and a
        grid past a few thousand rows is not something anyone reads.

        Errors are re-thrown with Format-GraphError's sentence so the tool's
        status line says something an operator can act on.
    #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE', 'PUT')][string]$Method = 'GET',
        $Body,
        [hashtable]$Headers,
        [switch]$All,
        [int]$MaxRows = 0
    )
    if ($Uri -notmatch '^https?://') { $Uri = 'https://graph.microsoft.com/v1.0/' + $Uri.TrimStart('/') }
    $p = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($null -ne $Body) {
        $p.Body        = $(if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 })
        $p.ContentType = 'application/json'
    }
    if ($Headers) { $p.Headers = $Headers }
    # OutputType PSObject keeps property CASE as Graph sent it. The default
    # (hashtable) is case-insensitive on lookup but the keys still come back
    # exactly as Graph wrote them, so either works for reads; PSObject is what
    # every tool here was written against.
    $cmd = Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Parameters.ContainsKey('OutputType')) { $p.OutputType = 'PSObject' }

    try {
        $r = Invoke-MgGraphRequest @p
    } catch {
        throw (Format-GraphError -ErrorRecord $_)
    }

    # ONE SHAPE OUT, WHICHEVER SHAPE CAME BACK.
    #
    # With -OutputType PSObject (SDK v2) the response is already a
    # PSCustomObject. Without that parameter (an older module) it is a
    # Hashtable, and $r.PSObject.Properties['value'] finds nothing on a
    # Hashtable -- so every caller reading .value that way would see an empty
    # result and report 'none found' rather than failing. Normalising here
    # means the six call sites do not each have to know which module version
    # is installed. A JSON round-trip, because the payload IS JSON and the
    # nesting has to survive.
    if ($r -is [System.Collections.IDictionary]) {
        try { $r = $r | ConvertTo-Json -Depth 20 | ConvertFrom-Json } catch { }
    }

    if (-not $All) { return $r }

    # READS THE RESPONSE BOTH WAYS IT CAN ARRIVE.
    #
    # With -OutputType PSObject the response is a PSCustomObject and
    # .PSObject.Properties finds 'value'. WITHOUT it -- an older
    # Microsoft.Graph.Authentication that has no OutputType parameter --
    # Invoke-MgGraphRequest returns a Hashtable, whose PSObject.Properties are
    # Count and Keys rather than the payload's keys. Testing only the PSObject
    # form would page once, find no 'value', and return nothing at all on that
    # build: an empty grid with no error, which is the worst way to be wrong.
    $pick = {
        param($Obj, [string]$Name)
        if ($null -eq $Obj) { return $null }
        if ($Obj -is [System.Collections.IDictionary]) {
            if ($Obj.Contains($Name)) { return $Obj[$Name] }
            return $null
        }
        if ($Obj.PSObject.Properties[$Name]) { return $Obj.$Name }
        return $null
    }

    $rows = New-Object System.Collections.ArrayList
    $page = $r
    while ($true) {
        $val = & $pick $page 'value'
        if ($null -ne $val) {
            foreach ($v in @($val)) { [void]$rows.Add($v) }
        }
        if ($MaxRows -gt 0 -and $rows.Count -ge $MaxRows) { break }
        $next = "$(& $pick $page '@odata.nextLink')"
        if (-not $next) { break }
        $q = @{ Method = 'GET'; Uri = $next; ErrorAction = 'Stop' }
        if ($Headers) { $q.Headers = $Headers }
        if ($cmd -and $cmd.Parameters.ContainsKey('OutputType')) { $q.OutputType = 'PSObject' }
        try { $page = Invoke-MgGraphRequest @q } catch { throw (Format-GraphError -ErrorRecord $_) }
    }
    return @($rows.ToArray())
}

function ConvertTo-GraphFilterLiteral {
    # OData string literal: single quotes are doubled. Nothing else is escaped
    # because nothing else needs to be inside a quoted literal.
    param([string]$Text)
    return ("$Text" -replace "'", "''")
}

function Resolve-GraphUser {
    <#
        A user object from whatever was typed: UPN, primary SMTP, any proxy
        address, or an object id. Null when nothing matches.

        Three properties in one filter so a mailbox typed by its alias
        address still resolves. proxyAddresses/any with eq is a plain filter
        in v1.0 (no ConsistencyLevel header needed -- checked 2026-09-17);
        endsWith and not would need it, and are not used.
    #>
    param([Parameter(Mandatory)][string]$Text)
    $t = "$Text".Trim()
    if (-not $t) { return $null }
    $sel = 'id,displayName,userPrincipalName,mail,accountEnabled,usageLocation,userType'
    if ($t -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
        try { return (Invoke-ToolkitGraph -Uri "users/$t`?`$select=$sel") } catch { return $null }
    }
    $lit = ConvertTo-GraphFilterLiteral -Text $t
    $f = "userPrincipalName eq '$lit' or mail eq '$lit' or proxyAddresses/any(p:p eq 'smtp:$lit')"
    $hits = @()
    try {
        $r = Invoke-ToolkitGraph -Uri ("users?`$filter=" + [uri]::EscapeDataString($f) + "&`$select=$sel&`$top=5")
        if ($r -and $r.PSObject.Properties['value']) { $hits = @($r.value) }
    } catch {
        # THE THREE-WAY FILTER IS THE ONLY PART OF THIS THAT GRAPH MIGHT
        # REFUSE. proxyAddresses/any() combined with an 'or' across two other
        # properties is documented as a default query, but it is the most
        # exotic thing the toolkit asks of /users, and every Graph tool
        # resolves its target through this function -- a rejection here would
        # read as 'no such user' for a user who plainly exists. So fall back
        # to the two plain equality filters, which cover the address people
        # actually type, and let a genuine failure surface from there.
        $f2 = "userPrincipalName eq '$lit' or mail eq '$lit'"
        $r = Invoke-ToolkitGraph -Uri ("users?`$filter=" + [uri]::EscapeDataString($f2) + "&`$select=$sel&`$top=5")
        if ($r -and $r.PSObject.Properties['value']) { $hits = @($r.value) }
    }
    if ($hits.Count -eq 0) { return $null }
    # Prefer the exact UPN match when an alias search returned more than one.
    $exact = @($hits | Where-Object { "$($_.userPrincipalName)" -ieq $t })
    if ($exact.Count -gt 0) { return $exact[0] }
    return $hits[0]
}

function ConvertTo-LocalGraphTime {
    # Graph timestamps are ISO 8601 UTC strings. Local, short, or '' when null.
    param($Value, [string]$Format = 'yyyy-MM-dd HH:mm')
    if (-not "$Value") { return '' }
    try { return ([DateTimeOffset]::Parse("$Value")).ToLocalTime().ToString($Format) } catch { return "$Value" }
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
    # This read '$Environment -in @(<two literal names>)', so a tenant that
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
        $null = Show-ToolDialog(
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

#region DPI awareness ----------------------------------------------------------
function Enable-ToolkitDpiAwareness {
    <#
        Make this process per-monitor DPI aware, once, before any window
        exists.

        WHY. powershell.exe carries no DPI manifest, so on a 125 / 150 %
        display Windows renders every form at 96 dpi and bitmap-stretches the
        result: blurry text, soft edges, the whole console looking like a
        screenshot of itself. That is the single biggest reason an internal
        tool reads as old next to Outlook on the same screen. Awareness makes
        the rendering crisp; Set-ConsoleTheme's AutoScaleMode = Dpi makes the
        LAYOUT follow, so a 900px form designed at 96 dpi becomes 1350px at
        150 % rather than 900px of tiny controls.

        WHEN. Runs at dot-source time, because Common.ps1 is loaded before any
        form in every process this toolkit starts -- the console, an OwnProcess
        tool, a project step. A tool run inside the console's process finds
        awareness already set and the call is a no-op. Hosts that set their
        own awareness (ISE) refuse the call; that is fine and swallowed.

        SetProcessDpiAwareness(2) = PER_MONITOR_DPI_AWARE (Windows 8.1+), with
        the older SetProcessDPIAware() as the fallback.
    #>
    if ($global:ToolkitDpiAwareSet) { return }
    $global:ToolkitDpiAwareSet = $true
    try {
        if (-not ('Toolkit.Dpi' -as [type])) {
            Add-Type -Namespace Toolkit -Name Dpi -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("shcore.dll")]
public static extern int SetProcessDpiAwareness(int value);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool SetProcessDPIAware();
'@ -ErrorAction Stop
        }
        $hr = [Toolkit.Dpi]::SetProcessDpiAwareness(2)
        if ($hr -ne 0 -and $hr -ne -2147024891) { [void][Toolkit.Dpi]::SetProcessDPIAware() }   # -2147024891 = E_ACCESSDENIED: already set
    } catch {
        try { [void][Toolkit.Dpi]::SetProcessDPIAware() } catch { }
    }
}
Enable-ToolkitDpiAwareness
#endregion

#region Additional helpers -------------------------------------------------------
<#
    HELPERS THAT ARRIVED WITH THE SECOND CONSOLE, now part of the shared file.

    These were written for the second tenant's toolkit and lived below a
    'Tenant additions' marker, which meant Common.ps1 was not the same file in
    the two installs: every update had to splice the file in two, and anything
    written on one side was invisible to the other. Two consoles that are meant
    to be identical were therefore identical everywhere except the one file
    that matters most. Folded in 2026-09-16; eleven functions that nothing
    called anywhere were dropped rather than carried.

    WHAT IS IN HERE, and why each group exists:

      API SECRETS. Some tenants need credentials for Proofpoint, SER and
      SendGrid that must not be written to disk. The store is in MEMORY ONLY,
      asked for once per session, and Forget-ApiCredentials clears it. Nothing
      here persists a secret anywhere, and nothing should be changed so that
      it does.

      MAILBOX AUDIT. Searching and decoding unified audit records, including
      the operation-name map that turns a raw operation into something an
      operator recognises.

      EMS COMMAND WRAPPERS. Get-EmsCommand / Invoke-EmsCommand, for on-premises
      Exchange calls made through the management shell. Invoke-OnPremCommand is
      the name the rest of the toolkit uses for the same idea; these are kept
      because tools written against them still call them.

      QUERY AND REPORT HELPERS. Bulk Exchange Online queries, CSV report
      writing, OU path formatting, connection-state probes.
#>

function Get-Prop {
    <#
        A property off a ConvertFrom-Json object, or a default if it is absent.

        NOT taken from the reference Common.ps1, because it is not there --
        08-COMMON-CONTRACT says so explicitly, and says an earlier draft of that
        contract listed it as though it existed. It lives in TenantManager.ps1,
        which this build does not ship.

        Needed because 5.1's ConvertFrom-Json returns a PSCustomObject and has
        no -AsHashtable (constraint A5). Reading a property that the JSON did
        not contain is not reliably an error -- under Set-StrictMode it throws,
        and without it you get $null and cannot tell "absent" from "present and
        empty". Every optional key in Tenant.json and Settings.json goes through
        here so the two cases stay distinguishable.
    #>
    param(
        [Parameter(Mandatory)]$Object,
        [Parameter(Mandatory)][string]$Name,
        $Default = $null
    )
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if (-not $p) { return $Default }
    if ($null -eq $p.Value) { return $Default }
    return $p.Value
}
#endregion
$script:TenantConfigFileName = 'Tenant.json'

function Invoke-BulkExoQuery {
    <#
        One enumerating Exchange query, made survivable.

        Generalised from reference/Protocols.ps1 Invoke-CasQuery. Four
        behaviours, each of which exists because the naive version failed in
        production:

        1. ONLY PASSES WHAT THE INSTALLED CMDLET DECLARES (C13). The module
           surface varies by version and a guessed parameter fails the WHOLE
           call rather than being ignored. ResultSize, PageSize, Properties,
           RecipientTypeDetails, Filter and Mailbox are each added only if the
           cmdlet actually has them.

        2. DROPS A REJECTED PROPERTY AND RETRIES. Exchange reports an
           unsupported property as 'InvalidProperties = X, Y' inside the
           exception text. Parsing that and retrying without those properties
           turns a total failure into a report missing two columns -- a partial
           answer instead of no answer.

        3. BACKS OFF ON A TRANSIENT DROP. A long enumeration outlives its own
           connection often enough that treating a closed socket as a failed
           report is simply wrong. Exponential, five attempts.

        4. RETURNS THE ROWS OR THROWS. It does NOT slice -- that is the
           caller's decision, because only the caller knows what a safe subset
           looks like.

        C15: nothing here retries an OVERSIZED query. A query that failed for
        size fails again the same way; the answer is a smaller query, not the
        same one twice.
    #>
    param(
        [Parameter(Mandatory)][string]$Cmdlet,
        [string[]]$Properties,
        [string]$RecipientTypeDetails,
        [string]$Filter,
        [string]$Mailbox,
        [int]$PageSize = 500,
        [scriptblock]$Progress
    )

    $supported = @(Get-CmdletParameterNames -Name $Cmdlet)
    if (@($supported).Count -eq 0) {
        throw "$Cmdlet is not available in this session."
    }

    $remaining  = @($Properties)
    $maxAttempt = 5

    for ($attempt = 1; $attempt -le $maxAttempt; $attempt++) {
        $params = @{ ErrorAction = 'Stop' }
        if ($supported -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }
        if ($supported -contains 'PageSize')   { $params['PageSize']   = $PageSize }
        if (@($remaining).Count -gt 0 -and $supported -contains 'Properties') { $params['Properties'] = $remaining }
        if ($RecipientTypeDetails -and $supported -contains 'RecipientTypeDetails') { $params['RecipientTypeDetails'] = $RecipientTypeDetails }
        if ($Filter  -and $supported -contains 'Filter')  { $params['Filter']  = $Filter }
        # SCOPE IS NOT A PROPERTY, AND DROPPING IT IS NOT A DEGRADED ANSWER.
        #
        # This line used to read:
        #
        #     if ($Mailbox -and $supported -contains 'Mailbox') { ... }
        #
        # Get-MailboxPermission and Get-RecipientPermission do not declare
        # -Mailbox. They declare -Identity. So the scope was silently dropped,
        # ResultSize=Unlimited stayed, and the call became "every mailbox
        # permission in the tenant" instead of "this one mailbox".
        #
        # FOUND 2026-08-23 IN PRODUCTION. A 1,039-mailbox report appeared to
        # hang on the first row. It had not hung -- it was enumerating the
        # whole tenant, once per mailbox, and would have done so 1,039 times.
        #
        # THE C13 RULE -- only pass what the cmdlet declares -- IS RIGHT FOR
        # PROPERTIES AND WRONG FOR SCOPE. Dropping an unsupported property
        # costs a column. Dropping the thing that says WHICH OBJECT costs the
        # difference between a question and a census, and nothing on screen
        # says so.
        #
        # So: find whichever scope parameter this cmdlet actually declares, and
        # REFUSE if it declares none. A caller that asked for one mailbox must
        # never silently receive the tenant.
        if ($Mailbox) {
            $scopeParam = @('Mailbox', 'Identity') | Where-Object { $supported -contains $_ } | Select-Object -First 1
            if (-not $scopeParam) {
                throw ("$Cmdlet was asked for a single mailbox ('$Mailbox') but declares neither -Mailbox nor -Identity. " +
                       'Refusing to run it unscoped: without one of those it would enumerate the entire tenant, ' +
                       'which is slow enough to look like a hang and wrong enough to be worth stopping.')
            }
            $params[$scopeParam] = $Mailbox

            # ResultSize is meaningless once scoped to one object, and on some
            # builds combining the two is what triggers the full enumeration.
            if ($params.ContainsKey('ResultSize')) { [void]$params.Remove('ResultSize') }
        }

        try {
            return @(Invoke-EXOCommand -Name $Cmdlet -Parameters $params)
        } catch {
            $msg = "$($_.Exception.Message)"

            $m = [regex]::Match($msg, 'InvalidProperties\s*=\s*(.+?)\s*$')
            if ($m.Success) {
                $bad = @($m.Groups[1].Value -split '\s*,\s*' | Where-Object { $_ })
                $remaining = @($remaining | Where-Object { $bad -notcontains $_ })
                if (@($remaining).Count -eq 0) { throw }
                if ($Progress) {
                    $word = 'ies'
                    if (@($bad).Count -eq 1) { $word = 'y' }
                    & $Progress ("Dropped unsupported propert{0}: {1}. Retrying." -f $word, ($bad -join ', '))
                }
                continue
            }

            $transient = 'underlying connection was closed|operation has timed out|' +
                         'unexpected error occurred on a receive|task was canceled|' +
                         'Unable to read data from the transport|forcibly closed'
            if (($msg -match $transient) -and $attempt -lt $maxAttempt) {
                $wait = [int][Math]::Pow(2, $attempt)
                if ($Progress) { & $Progress "Connection dropped. Retrying in ${wait}s (attempt $($attempt + 1) of $maxAttempt)." }
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
    throw "Could not query $Cmdlet after $maxAttempt attempts."
}


function Write-ReportCsv {
    <#
        The Stage D file-writing rule, in one place.

        NEVER WRITES A ZERO-ROW FILE (D6). An empty CSV and a failed run look
        identical in a folder listing a week later, and the empty one gets
        cited as evidence there was nothing to find.

        LABELS A PARTIAL RUN IN THE FILENAME, not only inside the file. The
        filename is what somebody sees in a month; a note in a file nobody
        opens is not labelling. 09: "A report covering 80% of mailboxes that
        does not say so is worse than no report, because it will be used as
        though it were complete."

        Returns the path, or '' when nothing was written -- so the caller can
        say which happened instead of assuming.
    #>
    param(
        [Parameter(Mandatory)][string]$Directory,
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)][AllowEmptyCollection()]$Rows,
        [switch]$Partial,
        [string]$Coverage = ''
    )

    if (@($Rows).Count -eq 0) { return '' }

    if (-not (Test-Path -LiteralPath $Directory)) {
        [void](New-Item -ItemType Directory -Path $Directory -Force -ErrorAction Stop)
    }

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $name  = "$Prefix-$stamp.csv"
    if ($Partial) { $name = "$Prefix-$stamp-PARTIAL.csv" }
    $path = Join-Path $Directory $name

    $note = $Coverage
    if (-not $note) {
        $note = 'complete'
        if ($Partial) { $note = 'PARTIAL -- see the tool window for what was not covered' }
    }

    # The coverage note goes on EVERY row, not in a trailing comment. A CSV
    # opened in Excel and sorted loses anything that is not per-row.
    $out = @($Rows | ForEach-Object {
        $o = $_ | Select-Object -Property *
        Add-Member -InputObject $o -NotePropertyName 'Coverage' -NotePropertyValue $note -Force
        $o
    })
    $out | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    return $path
}

function Get-ToolkitReportDirectory {
    <#
        Where report CSVs go.

        DELEGATES TO Get-ReportPath, which is the canonical definition lifted
        from the reference. This wrapper exists only because ten report tools
        were written against this name before Get-ReportPath was known to be
        available -- and two functions computing one value is precisely the
        "one value, many copies" shape that Get-ReportPath's own docstring says
        it was created to end.

        They agreed on the same folder when this was checked (2026-08-23), which
        is luck rather than design: Get-ReportPath resolves 'reports' and this
        resolved 'Reports', and only Windows being case-insensitive about paths
        made those the same directory. On any case-sensitive share they would
        have been two folders, and Report Browser would have listed one of them
        while the tools wrote to the other.

        -Fallback is accepted and ignored, so the ten existing callers keep
        working unchanged.
    #>
    param([string]$Fallback)
    return (Get-ReportPath)
}

function ConvertTo-OuPath {
    <#
        The OU an object sits in, from its distinguished name.

        MOVED HERE FROM Unified-Lookup.ps1 on 2026-08-22, because DLInventory
        needed the same conversion. Two files wanting one function is the moment
        it stops being a private helper -- 08 records that seven of thirteen
        local copies in the source toolkit had drifted, and every drift was a fix
        silently reverted. FUN001 exists to catch the second copy; this is the
        second copy, so it moves.

        Splits on unescaped commas only: an OU or CN containing a literal comma
        escapes it as '\,' and splitting naively would cut a name in half and
        report an OU that does not exist.

        The leading CN is dropped -- that is the object, not its location.
    #>
    param([string]$DistinguishedName)

    if (-not "$DistinguishedName".Trim()) { return '' }
    $parts = @("$DistinguishedName" -split '(?<!\\),')
    $ou    = @($parts | Where-Object { $_ -match '^\s*(OU|CN)=' } | Select-Object -Skip 1)
    $dc    = @($parts | Where-Object { $_ -match '^\s*DC=' } | ForEach-Object { ($_ -replace '^\s*DC=', '').Trim() })
    if (@($dc).Count -eq 0) { return '' }
    $domain = ($dc -join '.')
    if (@($ou).Count -eq 0) { return $domain }
    # Reversed so it reads top-down like a path, not bottom-up like a DN.
    [array]::Reverse($ou)
    $trail = @($ou | ForEach-Object { ($_ -replace '^\s*(OU|CN)=', '').Trim() })
    return "$domain/$($trail -join '/')"
}





function Get-EmsCommand {
    <#
        Resolve a cmdlet to the ON-PREMISES copy. The mirror of Get-EXOCommand,
        and the reason Stage F can be written safely at all.

        THIS IS C1/C2 MADE MECHANICAL. Both worlds export Get-Mailbox,
        Set-Mailbox, Get-DistributionGroup and dozens more, and THE LAST IMPORT
        WINS THE BARE NAME. A hybrid tool calling the bare name is relying on
        connection ORDER -- which works right up until somebody connects in the
        other order, and then a write lands in the wrong directory WITH NO
        ERROR. Nothing fails; the change simply happens somewhere else.

        Resolved through Get-ExchangeModuleMap, which tells the two apart by
        what they EXPORT rather than by module name: on-premises exports
        Get-RemoteMailbox, the cloud does not. Implicit remoting generates
        names like tmpEXO_a1b2c3 and they are not stable across versions.

        THROWS RATHER THAN GUESSING. If no on-premises module is loaded, there
        is no safe fallback -- falling back to the bare name would resolve to
        the CLOUD copy and write to the wrong place, which is precisely the
        failure this function exists to prevent. An error the operator can read
        beats a change they cannot see.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $map = Get-ExchangeModuleMap
    if (-not $map.OnPrem) { $map = Get-ExchangeModuleMap -Refresh }

    if ($map.OnPrem) {
        $cmd = Get-Command -Name "$($map.OnPrem.Name)\$Name" -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd }
        throw "The on-premises Exchange module '$($map.OnPrem.Name)' is loaded but does not export '$Name'. Check the cmdlet name -- this is not a connection problem."
    }

    # NO BARE-NAME FALLBACK. See the docstring: the bare name would very likely
    # resolve to Exchange Online, and an on-premises tool silently writing to
    # the cloud is the worst outcome available here.
    $all = @(Get-Command -Name $Name -All -ErrorAction SilentlyContinue)
    if (@($all).Count -eq 0) {
        throw "No command named '$Name' exists in this session, and no on-premises Exchange module is loaded. Connect to on-premises Exchange first."
    }
    $where = ($all | ForEach-Object { "$($_.Name) [$($_.ModuleName)$($_.Source)]" }) -join '  |  '
    throw "No ON-PREMISES Exchange module is loaded, so '$Name' cannot be resolved safely. It does exist here: $where -- but that is almost certainly the Exchange Online copy, and running an on-premises operation against it would write to the wrong directory without reporting an error. Refusing rather than guessing."
}

function Invoke-EmsCommand {
    <#
        Run one on-premises Exchange cmdlet, explicitly resolved.

        The mirror of Invoke-EXOCommand. Every Stage F write goes through this
        rather than calling the bare name, so "which Exchange did that land in"
        is answered by the code instead of by whichever module was imported last.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    $cmd = Get-EmsCommand -Name $Name
    & $cmd @Parameters
}

function Test-IsRealTrustee {
    <#
        Is this a person, or is it plumbing?

        Every mailbox carries NT AUTHORITY\SELF and a set of inherited
        administrative entries. Reporting those is not wrong exactly, but a
        report where every mailbox has four rows nobody cares about is a report
        nobody reads -- and the four real findings are somewhere in the middle
        of it.

        MOVED HERE FROM PermissionsAudit.ps1 ON 2026-08-23, because a second
        tool needed it and quietly failed at run time on every row:

            The term 'Test-IsRealTrustee' is not recognized

        A function defined in one TOOL is not available to another. Only
        Common.ps1 is shared, because only Common.ps1 is dot-sourced by
        everything. That is obvious written down and completely invisible while
        reading a call site that looks exactly like every other helper call.
    #>
    param([string]$Trustee)

    $t = "$Trustee".Trim()
    if (-not $t) { return $false }
    if ($t -match '(?i)^NT AUTHORITY\\') { return $false }
    if ($t -match '(?i)^S-1-5-') { return $false }
    if ($t -match '(?i)\\(Domain Admins|Enterprise Admins|Organization Management|Exchange Servers|Exchange Trusted Subsystem|Public Folder Management|Delegated Setup|Administrator)$') { return $false }
    return $true
}

function Get-MailboxAuditOperationMap {
    <#
        The operations worth offering, plain-language name first.

        BOTH NAMES ARE KEPT EVERYWHERE, and that is not redundancy. The raw
        operation is what Search-UnifiedAuditLog filters on and what you paste
        into a Microsoft support case; the friendly one is what a person reading
        a report under pressure can actually act on. 'MailItemsAccessed' and
        'SendAs' mean nothing to somebody who has been handed an incident.

        Ordered, so the checklist reads alphabetically by the name a human sees
        rather than by the internal one.
    #>
    return [ordered]@{
        'Accessed mailbox attachments'                                  = 'AttachmentAccess'
        'Accessed Mailbox folder'                                       = 'FolderBind'
        'Accessed mailbox items'                                        = 'MailItemsAccessed'
        'Accessed message'                                              = 'MessageBind'
        'Added delegate mailbox permissions'                            = 'Add-MailboxPermission'
        'Added or removed user with delegate access to calendar folder' = 'UpdateCalendarDelegation'
        'Added permissions to folder'                                   = 'AddFolderPermissions'
        'Copied messages to another folder'                             = 'Copy'
        'Created mailbox item'                                          = 'Create'
        'Created new inbox rule in Outlook web app'                     = 'New-InboxRule'
        'Delete priority cleanup label item'                            = 'PriorityCleanupDelete'
        'Deleted messages from Deleted Items folder'                    = 'SoftDelete'
        'Labeled item as a record'                                      = 'ApplyRecord'
        'Labeled message as a record'                                   = 'ApplyRecordLabel'
        'Modified folder permission'                                    = 'ModifyFolderPermissions'
        'Modified inbox rule from Outlook web app'                      = 'Set-InboxRule'
        'Moved messages to another folder'                              = 'Move'
        'Moved messages to Deleted Items folder'                        = 'MoveToDeletedItems'
        'Preserve Mailbox item'                                         = 'PreservedMailItemProactively'
        'Priority cleanup label applied'                                = 'ApplyPriorityCleanup'
        'Purged messages from the mailbox'                              = 'HardDelete'
        'Removed delegate mailbox permissions'                          = 'Remove-MailboxPermission'
        'Removed permissions from folder'                               = 'RemoveFolderPermissions'
        'Search items in a mailbox'                                     = 'SearchQueryInitiated'
        'Sent message'                                                  = 'Send'
        'Sent message using Send As permissions'                        = 'SendAs'
        'Sent message using Send On Behalf permissions'                 = 'SendOnBehalf'
        'Update item label'                                             = 'UpdateComplianceTag'
        'Updated inbox rules from Outlook client'                       = 'UpdateInboxRules'
        'Updated message'                                               = 'Update'
        'User signed in to mailbox'                                     = 'MailboxLogin'
    }
}

function Get-MailboxAuditFriendlyName {
    <#
        Raw operation -> plain language, falling back to the raw value.

        AN UNKNOWN OPERATION RETURNS ITSELF rather than 'Unknown'. The map above
        is a convenience, not a whitelist -- Microsoft adds operations, and a
        record labelled 'Unknown' would hide the one new thing worth seeing.
    #>
    param([string]$Operation)

    if (-not $script:MailboxAuditFriendly) {
        $script:MailboxAuditFriendly = @{}
        $m = Get-MailboxAuditOperationMap
        foreach ($k in $m.Keys) { $script:MailboxAuditFriendly[$m[$k]] = $k }
    }
    $op = "$Operation".Trim()
    if (-not $op) { return '' }
    if ($script:MailboxAuditFriendly.ContainsKey($op)) { return $script:MailboxAuditFriendly[$op] }
    return $op
}

function Invoke-MailboxAuditSearch {
    <#
        WHAT ACTUALLY HAPPENED TO A MAILBOX -- events, not configuration.

        PAGING IS A SESSION, NOT A LOOP OVER -ResultSize, and this is the single
        thing most easily got wrong. Search-UnifiedAuditLog is called REPEATEDLY
        with the SAME SessionId and SessionCommand 'ReturnLargeSet', and each
        call returns the NEXT page. One call without a session returns a
        truncated answer that looks complete -- there is no flag on the result
        saying otherwise, which is why a naive version passes review.

        FOUR WAYS TO STOP, and each is a different fact about the answer:

            empty page      the log had nothing more. Complete.
            short page      fewer than a full page came back. Complete.
            ceiling hit     50,000 records. TRUNCATED, and the caller must say so.
            cancelled       the operator stopped it. PARTIAL, and likewise.
            15 iterations   a guard against a session that never shortens.
                            Also truncated, and reported the same way.

        CANCELLATION IS A SCRIPTBLOCK, NOT A $script: FLAG, and that is a real
        change from the tools this came from. They kept
        $script:AuditCancelRequested and read it directly -- which worked while
        the function lived inside the tool. This function lives in Common.ps1
        now, where $script: is COMMON'S scope, and the flag the tool sets would
        never be seen. The caller passes a predicate that closes over its own
        state instead.

        Checked BETWEEN PAGES, never mid-call. A page in flight cannot be
        abandoned, so the honest granularity is one page.
    #>
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][datetime]$StartUtc,
        [Parameter(Mandatory)][datetime]$EndUtc,
        [string[]]$Operations,

        # FreeText for the shared-mailbox case. UserIds matches the ACTOR;
        # FreeText matches anywhere in the record, which is what finds actions
        # performed BY a delegate ON a shared mailbox. The caller chooses,
        # because only the caller knows which question it is asking.
        [ValidateSet('UserIds', 'FreeText')][string]$MatchOn = 'UserIds',

        [scriptblock]$OnProgress,
        [scriptblock]$IsCancelled,
        [int]$PageSize = 5000,
        [int]$Ceiling = 50000,
        [int]$MaxPages = 15
    )

    $sessionId = "MailboxAudit_$([Guid]::NewGuid().ToString())"
    $all       = New-Object System.Collections.Generic.List[object]
    $stopped   = ''
    $page      = 0

    while ($page -lt $MaxPages) {
        $page++

        $p = @{
            StartDate      = $StartUtc
            EndDate        = $EndUtc
            SessionId      = $sessionId
            SessionCommand = 'ReturnLargeSet'
            ResultSize     = $PageSize
            ErrorAction    = 'Stop'
        }
        if ($MatchOn -eq 'FreeText') { $p['FreeText'] = $Mailbox } else { $p['UserIds'] = $Mailbox }
        if (@($Operations).Count -gt 0) { $p['Operations'] = @($Operations) }

        $rows = @(Invoke-EXOCommand -Name 'Search-UnifiedAuditLog' -Parameters $p)

        if (@($rows).Count -eq 0) { $stopped = 'complete'; break }
        foreach ($r in $rows) { $all.Add($r) }

        if ($OnProgress) { & $OnProgress $all.Count $page }

        # AFTER the page is banked, so a cancel never discards work already paid
        # for.
        if ($IsCancelled -and (& $IsCancelled)) { $stopped = 'cancelled'; break }
        if ($all.Count -ge $Ceiling)            { $stopped = 'ceiling';   break }
        if (@($rows).Count -lt $PageSize)       { $stopped = 'complete';  break }
    }
    if (-not $stopped) { $stopped = 'pagelimit' }

    return [PSCustomObject]@{
        Records   = $all.ToArray()
        Count     = $all.Count
        Pages     = $page
        StoppedBy = $stopped
        # PRE-COMPUTED, so no caller has to re-derive "was this the whole
        # answer" from three fields and get it wrong in one of two tools.
        Complete  = ($stopped -eq 'complete')
        Truncated = ($stopped -eq 'ceiling' -or $stopped -eq 'pagelimit')
        Cancelled = ($stopped -eq 'cancelled')
    }
}

function ConvertTo-MailboxAuditRecord {
    <#
        One raw audit row, flattened into columns.

        AuditData IS A JSON STRING, NOT AN OBJECT. Everything interesting --
        the actor, the client IP, the subject, the folder -- is inside it, and
        the flat properties on the record itself carry only a fraction of that.

        IT PARSES DEFENSIVELY AND DEGRADES, NEVER THROWS. A single record whose
        JSON will not parse must not lose the other four thousand. When parsing
        fails the flat fields are used and RawJson holds the original text, so
        the row is still there, still ordered, still exportable -- just thinner.

        THE RAW TEXT IS ALWAYS KEPT. The columns answer the common questions;
        the raw JSON answers the one question the columns do not, and that is
        the question that gets asked when something has actually gone wrong.

        TIMESTAMPS ARE UTC IN THE LOG. InvariantCulture and AssumeUniversal are
        both load-bearing: without the culture a machine set to a non-US format
        misreads the date, and without AssumeUniversal .NET treats the string as
        LOCAL and shifts it by the offset -- silently, producing a timeline that
        looks plausible and is hours wrong.
    #>
    param([Parameter(Mandatory)]$RawRecord)

    $parsed = $null
    $parseError = ''
    try { $parsed = "$($RawRecord.AuditData)" | ConvertFrom-Json -ErrorAction Stop }
    catch { $parsed = $null; $parseError = "$($_.Exception.Message)" }

    $op = if ($parsed -and $parsed.Operation) { "$($parsed.Operation)" } else { "$($RawRecord.Operations)" }

    $whenUtc = $null
    if ($parsed -and $parsed.CreationTime) {
        try {
            $whenUtc = [datetime]::Parse(
                "$($parsed.CreationTime)",
                [System.Globalization.CultureInfo]::InvariantCulture,
                ([System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal))
        } catch { $whenUtc = $null }
    }
    if (-not $whenUtc -and $RawRecord.CreationDate) {
        try { $whenUtc = ([datetime]$RawRecord.CreationDate).ToUniversalTime() } catch { $whenUtc = $null }
    }

    return [PSCustomObject]@{
        TimeUtc           = $whenUtc
        TimeLocal         = $(if ($whenUtc) { $whenUtc.ToLocalTime() } else { $null })
        Operation         = $op
        OperationFriendly = (Get-MailboxAuditFriendlyName -Operation $op)
        Actor             = $(if ($parsed) { "$($parsed.UserId)" } else { "$($RawRecord.UserIds)" })
        MailboxOwnerUPN   = $(if ($parsed) { "$($parsed.MailboxOwnerUPN)" } else { '' })
        ClientIP          = $(if ($parsed) { "$($parsed.ClientIP)" } else { '' })
        ClientInfo        = $(if ($parsed) { "$($parsed.ClientInfoString)" } else { '' })
        Subject           = $(if ($parsed -and $parsed.Item -and $parsed.Item.Subject) { "$($parsed.Item.Subject)" } else { '' })
        Folder            = $(if ($parsed -and $parsed.Item -and $parsed.Item.ParentFolder -and $parsed.Item.ParentFolder.Name) { "$($parsed.Item.ParentFolder.Name)" } else { '' })
        ResultStatus      = $(if ($parsed) { "$($parsed.ResultStatus)" } else { '' })

        # SAID, NOT SWALLOWED. A row that degraded is a row whose blank columns
        # mean "could not read" rather than "was empty", and those are different
        # answers to an investigator.
        Degraded          = (-not $parsed)
        ParseError        = $parseError

        RawJson           = $(if ($parsed) { ($parsed | ConvertTo-Json -Depth 8) } else { "$($RawRecord.AuditData)" })
    }
}





function Get-ConsoleFontScale {
    <#
        The font multiplier from Themes.json, or 1.0.

        READ FROM THE SAME FILE AS THE COLOURS so a theme and its text size
        travel together, and one cache invalidation clears both at once.
    #>
    $cfg = $null
    try {
        $f = Get-ThemeFilePath
        if ($f -and (Test-Path -LiteralPath $f)) { $cfg = Get-Content -LiteralPath $f -Raw | ConvertFrom-Json }
    } catch { return 1.0 }
    if (-not $cfg) { return 1.0 }

    $s = 1.0
    if ($cfg.PSObject.Properties.Name -contains 'FontScale') {
        try { $s = [double]$cfg.FontScale } catch { $s = 1.0 }
    }
    # CLAMPED. Below 0.8 the console is unreadable and above 1.6 the fixed
    # layouts overlap -- a slider that can make the tool unusable is worse than
    # no slider, and neither bound is recoverable from inside the broken window.
    if ($s -lt 0.8) { $s = 0.8 }
    if ($s -gt 1.6) { $s = 1.6 }
    return $s
}


function Test-TeamsConnectionState {
    <#
        Is there a usable Teams session in THIS process?

        TESTED BY CALLING, NOT BY Get-Command. The cmdlets exist the moment the
        module is imported, connected or not -- so a command check reports
        "connected" and the first real call then fails with an authentication
        error the operator has to decode. Same reasoning as Test-SpoConnectionState.
    #>
    param()
    $r = [ordered]@{ Connected = $false; Detail = '' }
    if (-not (Get-Command -Name 'Get-CsTenant' -ErrorAction SilentlyContinue)) {
        $r.Detail = 'Not connected -- Teams cmdlets are not loaded in this process. Import MicrosoftTeams and run Connect-MicrosoftTeams.'
        return [PSCustomObject]$r
    }
    try {
        $null = Get-CsTenant -ErrorAction Stop
        $r.Connected = $true
        $r.Detail = 'Connected to Microsoft Teams.'
    } catch {
        $r.Detail = "Not connected -- Teams cmdlets are loaded but the session is missing or expired: $($_.Exception.Message)"
    }
    return [PSCustomObject]$r
}

function Test-IppsConnectionState {
    <#
        Is there a Security and Compliance (IPPSSession) connection?

        SHARES CMDLET NAMES WITH EXCHANGE ONLINE, which is the trap here.
        Get-Label, Get-DlpCompliancePolicy and friends come from the
        Security and Compliance endpoint, but Connect-IPPSSession and
        Connect-ExchangeOnline both load ExchangeOnlineManagement -- so probing
        with a name that exists in both would report Security as connected
        whenever ExO is.

        Get-ComplianceSearch is used because it exists ONLY in the Security and
        Compliance session, so a successful call proves which endpoint is live
        rather than merely that the module loaded.
    #>
    param()
    $r = [ordered]@{ Connected = $false; Detail = '' }
    if (-not (Get-Command -Name 'Get-ComplianceSearch' -ErrorAction SilentlyContinue)) {
        $r.Detail = 'Not connected -- Security and Compliance cmdlets are not loaded in this process. Run Connect-IPPSSession.'
        return [PSCustomObject]$r
    }
    try {
        $null = Get-ComplianceSearch -ResultSize 1 -ErrorAction Stop -WarningAction SilentlyContinue
        $r.Connected = $true
        $r.Detail = 'Connected to Security and Compliance (IPPSSession).'
    } catch {
        $r.Detail = "Not connected -- Security and Compliance cmdlets are loaded but the session is missing or expired: $($_.Exception.Message)"
    }
    return [PSCustomObject]$r
}

#region API credentials ---------------------------------------------------------
<#
    CREDENTIAL RETRIEVAL FOR THE THIRD-PARTY MAIL PLATFORMS.

    Built before any key exists, deliberately. Every tool that later talks to
    Proofpoint or SendGrid is written against this interface from its first
    line, so nothing has to be retrofitted when the keys arrive.

    THE THREE RULES, and they are absolute:
      1. No secret is ever written to disk. No file, no temp, no "remember me".
      2. No secret is ever displayed -- not in output, not in an error.
      3. No secret is ever audited. The log records tool, operator, target and
         result, as every other tool does.
#>

# ADDING A SERVICE IS ADDING A ROW. The dialog reads this table; it does not
# branch on the service name. Same reasoning as the Unified Lookup hand-off
# table -- the extension mechanism should be visible rather than buried.
$script:ApiSecretDefs = [ordered]@{

    # PROVISIONAL. The pod's authentication mechanism is NOT CONFIRMED -- the
    # reachable vendor documentation does not state it unambiguously and a
    # support ticket is open. A user/password pair accommodates the widest range
    # of outcomes, so it is the placeholder. See UNVERIFIED.md. When the real
    # mechanism is known, this one row changes.
    'ProofpointPod' = @{
        Label       = 'Proofpoint (mail gateway)'
        NeedsUser   = $true
        UserLabel   = 'API user'
        SecretLabel = 'Password'
        Help        = 'Requires the REST API role. API is on port 10000.'
    }

    'ProofpointSer' = @{
        Label       = 'Proofpoint Secure Email Relay'
        NeedsUser   = $true
        UserLabel   = 'Client ID'
        SecretLabel = 'Client Secret'
        Help        = 'OAuth2 client credentials. Reporting scope only.'
        OAuth       = $true
    }

    'SendGrid' = @{
        Label       = 'SendGrid'
        NeedsUser   = $false
        UserLabel   = ''
        SecretLabel = 'API Key'
        Help        = 'Restricted Access: Suppressions = Full Access, everything else No Access.'
    }
}

# THE CYBERARK SEAM. Flipping this is the only change Get-ApiSecret needs when
# retrieval is approved. It cannot be flipped until all five of these are known,
# and all five are recorded in UNVERIFIED.md:
#     an Application ID, a Safe name, the account object names, the CCP
#     endpoint URL, and a client certificate thumbprint.
$script:UseCcp = $false

function Initialize-ApiSecretStore {
    <#
        THE CACHE LIVES IN $global:, NOT $script:, AND THAT IS NOT A STYLE
        CHOICE.

        The console launches tools with the call operator, and in that
        arrangement Common.ps1's script scope IS NOT the calling tool's script
        scope. A $script: read from a launched tool resolves to $null and the
        operator is prompted on every single call.

        Add-ReportFileControls carries a comment recording the same trap, and it
        is why that function returns state on an object instead of stashing it.

        Initialised defensively because ANY tool may be the first caller in a
        session.
    #>
    if (-not $global:ToolkitApiSecrets)     { $global:ToolkitApiSecrets = @{} }
    if (-not $global:ToolkitApiSecretToken) { $global:ToolkitApiSecretToken = @{} }
}

function Get-SafeApiError {
    <#
        Exception -> display text with anything credential-shaped removed.

        THE SPECIFIC HAZARD IS ERROR HANDLING. Invoke-RestMethod failures can
        surface request detail, and a handler that dumps the exception verbatim
        can print an Authorization header. EVERY API catch block goes through
        here -- including the ones that look harmless, because those are the
        ones that get copied into new code.
    #>
    param($ErrorRecord, [string]$Context = '')

    $text = ''
    try {
        if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
            $text = "$($ErrorRecord.Exception.Message)"
            if ($ErrorRecord.Exception.InnerException) { $text += " -- $($ErrorRecord.Exception.InnerException.Message)" }
        } else { $text = "$ErrorRecord" }
    } catch { $text = 'An error occurred and could not be read.' }

    # ORDER MATTERS. The scheme-prefixed forms are replaced first so the generic
    # long-run rule cannot eat half of one and leave the scheme behind looking
    # like a redaction succeeded when it only partly did.
    $text = [regex]::Replace($text, '(?i)(Bearer)\s+[A-Za-z0-9\-\._~\+/=]+', '$1 [redacted]')
    $text = [regex]::Replace($text, '(?i)(Basic)\s+[A-Za-z0-9\+/=]+',        '$1 [redacted]')
    $text = [regex]::Replace($text, '(?i)(api[_\-]?key["'':\s=]+)[A-Za-z0-9\-\._]{8,}', '$1[redacted]')
    # A SendGrid key is a specific, checkable shape.
    $text = [regex]::Replace($text, 'SG\.[A-Za-z0-9\-\._]{8,}', '[redacted]')
    # Any bare token-shaped run of 32 or more.
    $text = [regex]::Replace($text, '(?<![A-Za-z0-9\-\._])[A-Za-z0-9\-\._]{32,}(?![A-Za-z0-9\-\._])', '[redacted]')

    if ($Context) { return "$Context failed: $text" }
    return $text
}

function Show-ApiSecretDialog {
    <#
        Prompt for one service's credential.

        DRIVEN BY THE TABLE, not by branching on the service. The field count,
        the labels and the help line all come from $script:ApiSecretDefs, so a
        new service needs no change here at all.

        Modelled on the toolkit's existing prompt dialogs -- same construction,
        sizing and OK/Cancel arrangement.
    #>
    param([Parameter(Mandatory)][string]$Service, [Parameter(Mandatory)]$Def)

    $needsUser = [bool]$Def.NeedsUser
    $h = $(if ($needsUser) { 268 } else { 236 })

    $dlgSec                 = New-Object System.Windows.Forms.Form
    $dlgSec.Text            = "$($Def.Label) -- credential"
    $dlgSec.ClientSize      = New-Object System.Drawing.Size(560, $h)
    $dlgSec.FormBorderStyle = 'FixedDialog'
    $dlgSec.MaximizeBox     = $false
    $dlgSec.MinimizeBox     = $false
    $dlgSec.StartPosition   = 'CenterParent'

    $y = 14

    $lblSvc          = New-Object System.Windows.Forms.Label
    $lblSvc.Text     = "$($Def.Label)"
    $lblSvc.Location = New-Object System.Drawing.Point(14, $y)
    $lblSvc.Size     = New-Object System.Drawing.Size(532, 20)
    $lblSvc.Font     = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    [void]$dlgSec.Controls.Add($lblSvc)
    $y += 26

    # WHERE TO GET THE VALUE. Without it the operator has to go and find out
    # which console page issues this, which is the slow part of the job.
    $lblHelp           = New-Object System.Windows.Forms.Label
    $lblHelp.Text      = "$($Def.Help)"
    $lblHelp.Location  = New-Object System.Drawing.Point(14, $y)
    $lblHelp.Size      = New-Object System.Drawing.Size(532, 32)
    $lblHelp.ForeColor = [System.Drawing.Color]::Gray
    $lblHelp.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
    [void]$dlgSec.Controls.Add($lblHelp)
    $y += 36

    $txtUser = $null
    if ($needsUser) {
        $lblU          = New-Object System.Windows.Forms.Label
        $lblU.Text     = "$($Def.UserLabel)"
        $lblU.Location = New-Object System.Drawing.Point(14, ($y + 3))
        $lblU.Size     = New-Object System.Drawing.Size(130, 20)
        [void]$dlgSec.Controls.Add($lblU)

        $txtUser          = New-Object System.Windows.Forms.TextBox
        $txtUser.Location = New-Object System.Drawing.Point(150, $y)
        $txtUser.Size     = New-Object System.Drawing.Size(396, 22)
        [void]$dlgSec.Controls.Add($txtUser)
        $y += 30
    }

    $lblS          = New-Object System.Windows.Forms.Label
    $lblS.Text     = "$($Def.SecretLabel)"
    $lblS.Location = New-Object System.Drawing.Point(14, ($y + 3))
    $lblS.Size     = New-Object System.Drawing.Size(130, 20)
    [void]$dlgSec.Controls.Add($lblS)

    $txtSecret          = New-Object System.Windows.Forms.TextBox
    $txtSecret.Location = New-Object System.Drawing.Point(150, $y)
    $txtSecret.Size     = New-Object System.Drawing.Size(396, 22)
    $txtSecret.UseSystemPasswordChar = $true
    [void]$dlgSec.Controls.Add($txtSecret)
    $y += 32

    # ON THE DIALOG, NOT ONLY IN THE HELP. The person typing a credential into a
    # window is entitled to know what happens to it, and this is the moment they
    # are thinking about it.
    $lblPromise           = New-Object System.Windows.Forms.Label
    $lblPromise.Text      = 'Held in memory for this session only. Never written to disk.'
    $lblPromise.Location  = New-Object System.Drawing.Point(14, $y)
    $lblPromise.Size      = New-Object System.Drawing.Size(532, 18)
    $lblPromise.ForeColor = [System.Drawing.Color]::Gray
    $lblPromise.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
    [void]$dlgSec.Controls.Add($lblPromise)
    $y += 22

    $lblErr           = New-Object System.Windows.Forms.Label
    $lblErr.Location  = New-Object System.Drawing.Point(14, $y)
    $lblErr.Size      = New-Object System.Drawing.Size(532, 18)
    $lblErr.ForeColor = [System.Drawing.Color]::Firebrick
    [void]$dlgSec.Controls.Add($lblErr)
    $y += 24

    $ok           = New-Object System.Windows.Forms.Button
    $ok.Text      = 'OK'
    $ok.Location  = New-Object System.Drawing.Point(312, $y)
    $ok.Size      = New-Object System.Drawing.Size(114, 28)
    $ok.FlatStyle = 'Flat'
    $ok.Enabled   = $false
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    [void]$dlgSec.Controls.Add($ok)

    $no           = New-Object System.Windows.Forms.Button
    $no.Text      = 'Cancel'
    $no.Location  = New-Object System.Drawing.Point(432, $y)
    $no.Size      = New-Object System.Drawing.Size(114, 28)
    $no.FlatStyle = 'Flat'
    $no.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    [void]$dlgSec.Controls.Add($no)

    # OK STAYS OFF UNTIL EVERY REQUIRED FIELD HAS SOMETHING IN IT. An empty
    # credential submitted by accident produces a 401 several seconds later in a
    # completely different part of the tool, and that is miserable to diagnose.
    $validate = {
        $good = ("$($txtSecret.Text)".Length -gt 0)
        if ($needsUser -and "$($txtUser.Text)".Trim().Length -eq 0) { $good = $false }
        $ok.Enabled = $good
        $lblErr.Text = $(if ($good) { '' } else { 'Every field is required.' })
    }
    $txtSecret.Add_TextChanged($validate)
    if ($needsUser) { $txtUser.Add_TextChanged($validate) }

    $dlgSec.AcceptButton = $ok
    $dlgSec.CancelButton = $no
    try { Set-ConsoleTheme -Form $dlgSec } catch { }

    $res = $dlgSec.ShowDialog()

    $result = [PSCustomObject]@{ Ok = $false; User = ''; Secret = $null }
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
        $sec = New-Object System.Security.SecureString
        foreach ($ch in "$($txtSecret.Text)".ToCharArray()) { $sec.AppendChar($ch) }
        $sec.MakeReadOnly()
        $result.Secret = $sec
        if ($needsUser) { $result.User = "$($txtUser.Text)".Trim() }
        $result.Ok = $true
    }

    # CLEARING THE TEXTBOX IS NOT SECURE ERASURE, AND THIS COMMENT WILL NOT
    # PRETEND IT IS.
    #
    # A WinForms TextBox.Text is a managed .NET string: immutable, and it stays
    # in memory until the garbage collector happens to reclaim it. Setting it to
    # '' does not scrub the original.
    #
    # It is done anyway because it removes the value from the visible control
    # and from anything that later reads the form, and the alternative is
    # leaving it there indefinitely. That is the whole of what it achieves.
    $txtSecret.Text = ''
    if ($needsUser) { $txtUser.Text = '' }
    $dlgSec.Dispose()
    return $result
}

function Get-ApiSecretFromCcp {
    <#
        CyberArk Central Credential Provider retrieval.

        DELIBERATELY NOT BUILT. The CCP endpoint and its parameter surface are
        not known here, and plausible code written against an API that cannot be
        checked is worth less than nothing -- it looks finished.

        Five things must be known before $script:UseCcp can be set true, and all
        five are listed in UNVERIFIED.md:
            Application ID, Safe name, account object names, the CCP endpoint
            URL, and a client certificate thumbprint.
    #>
    param([string]$Service)
    throw 'Not implemented -- awaiting CyberArk onboarding.'
}

function Get-ApiSecret {
    <#
        The credential for one service, prompting once per session.

        LAZY. Nothing happens at console startup. The prompt fires on the first
        call for that service, and only for that service -- an operator who only
        releases a quarantined message is never asked for SendGrid.

        CANCEL IS A FIRST-CLASS ANSWER. It returns Ok = false with a Reason and
        caches nothing. The caller shows a normal cancelled state: not an error
        dialog, not a stack trace, and never by proceeding with an empty
        credential and letting the API answer 401.

        NO DISK. This function reads and writes no file. There is deliberately
        no persistence parameter, because a parameter that exists will
        eventually be used.
    #>
    param(
        [Parameter(Mandatory)][string]$Service,
        [string]$Purpose = '',
        [switch]$Force
    )

    Initialize-ApiSecretStore

    $def = $null
    if ($script:ApiSecretDefs.Contains($Service)) { $def = $script:ApiSecretDefs[$Service] }
    if (-not $def) {
        return [PSCustomObject]@{ Ok = $false; Service = $Service; User = ''; Secret = $null
            Source = ''; Reason = "Unknown service '$Service'. Add a row to `$script:ApiSecretDefs." }
    }

    # THE WHOLE CYBERARK CHANGE IS THIS BRANCH. Nothing else moves.
    if ($global:ToolkitApiSecrets.ContainsKey($Service) -and -not $Force) {
        $c = $global:ToolkitApiSecrets[$Service]
        return [PSCustomObject]@{ Ok = $true; Service = $Service; User = $c.User; Secret = $c.Secret
            Source = $c.Source; Reason = '' }
    }
    elseif ($script:UseCcp) {
        return Get-ApiSecretFromCcp -Service $Service
    }
    else {
        $p = Show-ApiSecretDialog -Service $Service -Def $def
        if (-not $p.Ok) {
            return [PSCustomObject]@{ Ok = $false; Service = $Service; User = ''; Secret = $null
                Source = 'Prompt'; Reason = 'Cancelled by the operator. Nothing was stored.' }
        }

        # -Force REPLACES, so the old SecureString is disposed rather than
        # orphaned. A rotated key should not leave its predecessor in memory.
        if ($global:ToolkitApiSecrets.ContainsKey($Service)) {
            try { $global:ToolkitApiSecrets[$Service].Secret.Dispose() } catch { }
        }
        # A NEW CREDENTIAL INVALIDATES ANY TOKEN MINTED FROM THE OLD ONE.
        if ($global:ToolkitApiSecretToken.ContainsKey($Service)) {
            try { $global:ToolkitApiSecretToken[$Service].Token.Dispose() } catch { }
            [void]$global:ToolkitApiSecretToken.Remove($Service)
        }

        $global:ToolkitApiSecrets[$Service] = [PSCustomObject]@{
            User = $p.User; Secret = $p.Secret; Source = 'Prompt'; AcquiredAt = (Get-Date)
        }
        return [PSCustomObject]@{ Ok = $true; Service = $Service; User = $p.User; Secret = $p.Secret
            Source = 'Prompt'; Reason = '' }
    }
}

function Clear-ApiSecret {
    <#
        Forget one service, or all of them.

        CLEARS THE TOKEN AS WELL AS THE CREDENTIAL. A bearer token minted from a
        cleared client secret is still a live credential -- forgetting the
        secret and leaving the token would make the promise false.
    #>
    param([string]$Service)

    Initialize-ApiSecretStore
    $names = @($(if ($Service) { $Service } else { @($global:ToolkitApiSecrets.Keys) + @($global:ToolkitApiSecretToken.Keys) }))
    $names = @($names | Sort-Object -Unique)

    $count = 0
    foreach ($n in $names) {
        if ($global:ToolkitApiSecrets.ContainsKey($n)) {
            try { $global:ToolkitApiSecrets[$n].Secret.Dispose() } catch { }
            [void]$global:ToolkitApiSecrets.Remove($n)
            $count++
        }
        if ($global:ToolkitApiSecretToken.ContainsKey($n)) {
            try { $global:ToolkitApiSecretToken[$n].Token.Dispose() } catch { }
            [void]$global:ToolkitApiSecretToken.Remove($n)
        }
    }
    return $count
}

function Get-ApiSecretState {
    <#
        What is loaded, for the console to display.

        NEVER RETURNS A SECRET, and has no parameter that could make it. The
        console can show what is held without any code path that touches a
        value.
    #>
    Initialize-ApiSecretStore
    $rows = New-Object System.Collections.ArrayList
    foreach ($n in @($script:ApiSecretDefs.Keys)) {
        $loaded = $global:ToolkitApiSecrets.ContainsKey($n)
        [void]$rows.Add([PSCustomObject]@{
            Service    = $n
            Label      = "$($script:ApiSecretDefs[$n].Label)"
            Loaded     = $loaded
            AcquiredAt = $(if ($loaded) { $global:ToolkitApiSecrets[$n].AcquiredAt } else { $null })
            HasToken   = $global:ToolkitApiSecretToken.ContainsKey($n)
        })
    }
    return @($rows.ToArray())
}

function Invoke-WithApiSecret {
    <#
        The ONLY place a SecureString is converted to plaintext.

        The plaintext exists for the shortest possible window and the unmanaged
        buffer is always freed. WITHOUT THE finally A FAILED API CALL LEAKS THE
        BUFFER -- and a failing call is exactly the path most exercised while a
        new integration is being debugged.

        Callers must never convert a SecureString themselves. SEC001 fails the
        build if SecureStringToBSTR appears anywhere else.
    #>
    param(
        [Parameter(Mandatory)][System.Security.SecureString]$Secret,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secret)
    try {
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        & $ScriptBlock $plain
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Get-SerAccessToken {
    <#
        The SER bearer token, cached with its expiry.

        SER DOES NOT ACCEPT THE CLIENT SECRET DIRECTLY. The id and secret are
        exchanged for a short-lived bearer token. THAT TOKEN IS ITSELF A
        CREDENTIAL and obeys every rule here: never on disk, never in output,
        never in the audit log.

        A REFRESH MUST NOT PROMPT. It uses the cached client credentials. Being
        asked to paste a client secret again forty minutes into a session
        because a token expired would mean the design had failed.

        TREATED AS EXPIRED EARLY, by a 60-second margin. A token that expires
        between the check and the call produces a 401 that looks like a
        credential problem and sends somebody debugging in the wrong direction.
    #>
    param([switch]$Force)

    Initialize-ApiSecretStore
    $svc = 'ProofpointSer'

    # SAFETY MARGIN. See above -- this is the difference between a clean refresh
    # and a misleading 401.
    $marginSeconds = 60

    if (-not $Force -and $global:ToolkitApiSecretToken.ContainsKey($svc)) {
        $t = $global:ToolkitApiSecretToken[$svc]
        if ($t.ExpiresAt -and (Get-Date) -lt $t.ExpiresAt.AddSeconds(-1 * $marginSeconds)) {
            return [PSCustomObject]@{ Ok = $true; Token = $t.Token; ExpiresAt = $t.ExpiresAt; Reason = '' }
        }
    }

    $cred = Get-ApiSecret -Service $svc -Purpose 'SER reporting API'
    if (-not $cred.Ok) {
        return [PSCustomObject]@{ Ok = $false; Token = $null; ExpiresAt = $null; Reason = $cred.Reason }
    }

    $tokenUrl = 'https://mail.ser.proofpoint.com/v1/token'
    $result = [PSCustomObject]@{ Ok = $false; Token = $null; ExpiresAt = $null; Reason = '' }

    try {
        $resp = Invoke-WithApiSecret -Secret $cred.Secret -ScriptBlock {
            param($plain)
            $body = @{
                grant_type    = 'client_credentials'
                client_id     = $cred.User
                client_secret = $plain
            }
            Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body `
                -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        }

        if (-not $resp -or -not "$($resp.access_token)") {
            $result.Reason = 'The token endpoint answered without an access_token. Nothing was cached.'
            return $result
        }

        $sec = New-Object System.Security.SecureString
        foreach ($ch in "$($resp.access_token)".ToCharArray()) { $sec.AppendChar($ch) }
        $sec.MakeReadOnly()

        $lifetime = 3600
        if ($resp.expires_in) { try { $lifetime = [int]$resp.expires_in } catch { $lifetime = 3600 } }

        if ($global:ToolkitApiSecretToken.ContainsKey($svc)) {
            try { $global:ToolkitApiSecretToken[$svc].Token.Dispose() } catch { }
        }
        $global:ToolkitApiSecretToken[$svc] = [PSCustomObject]@{
            Token = $sec; ExpiresAt = (Get-Date).AddSeconds($lifetime); AcquiredAt = (Get-Date)
        }

        $result.Ok = $true
        $result.Token = $sec
        $result.ExpiresAt = $global:ToolkitApiSecretToken[$svc].ExpiresAt
        return $result
    } catch {
        # A 400 FROM THE TOKEN ENDPOINT MEANS THE CREDENTIALS ARE WRONG OR
        # EXPIRED, not that the API is down. Saying so is the single most useful
        # thing that can be said at this moment -- the alternative sends people
        # checking network paths for a key that was revoked.
        $status = 0
        try { $status = [int]$_.Exception.Response.StatusCode } catch { }
        $safe = Get-SafeApiError -ErrorRecord $_ -Context 'SER token request'

        if ($status -eq 400 -or $status -eq 401 -or $safe -match '(?i)invalid_client') {
            $result.Reason = 'THE CLIENT ID AND SECRET WERE REJECTED. This is not an outage -- the pair is wrong, revoked or expired, and a replacement must be created in the SER console. ' +
                'Re-run with -Force to enter a different one without restarting the console.' + [Environment]::NewLine + $safe
        } else {
            $result.Reason = $safe
        }
        return $result
    }
}

#endregion
