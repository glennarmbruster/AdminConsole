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

# THE CONSOLE'S CONTEXT, ACCEPTED. Since 2026-09-12 the console's Look Up
# button opens this tool with whatever was typed, so the address arrives here
# and the search runs on its own. Declared, not absorbed into $args, because
# Get-ToolParameterSet in the console only passes what a script declares.
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

# THIS FILE'S OWN COPY, DELIBERATELY NAMED DIFFERENTLY FROM Common.ps1's.
#
# It used to be called ConvertTo-OuPath, the same as the function in
# Common.ps1 -- which this file dot-sources. The local definition wins, so
# the calls below reached THIS one and always had; but the shadowing was a
# trap with two edges. A fix made in Common never arrived here, and anyone
# tidying up an apparent duplicate would have deleted this and left the
# calls binding silently to Common's version, which does not take the same
# parameters. That is a run-time failure at the click, on a Saturday.
#
# The two are genuinely different functions that happened to share a name,
# so they now have different names. Renamed 2026-09-12; behaviour unchanged.
function ConvertTo-AdOuPath {
    <#
        The OU a distinguished name sits in, in the form ADUC shows it.

        'CN=TS-Shared,OU=Shared Mailboxes,OU=Accounts,DC=corp,DC=x,DC=com'
            ->  corp.x.com/Accounts/Shared Mailboxes

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
        OuPath          = (ConvertTo-AdOuPath -DistinguishedName $Dn)
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
        in both directories is not an ambiguity to resolve; it is the
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

        $confirm = Show-ToolDialog(
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
# things that make a name look ambiguous, and 'ABCD' the security group sitting
# beside 'ABCD-Shared' the mailbox is the reason you opened a lookup tool at
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
    # that were asked for. 'Show me cloud user mailboxes matching a common word'
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

# Opened from the console with an address: search it as soon as the window is
# up. At Shown, not now -- the search paints into controls that need a handle,
# and the connection banner has to have been drawn first so a "not connected"
# answer has somewhere to go. Once only; the flag on the form makes a second
# Shown (there is none today) harmless.
if ("$SmtpAddress".Trim()) {
    $txtSearch.Text = "$SmtpAddress".Trim()
    $form.Add_Shown({
        if ($this.PSObject.Properties['AutoSearched']) { return }
        $this | Add-Member -NotePropertyName 'AutoSearched' -NotePropertyValue $true -Force
        try { $btnSearch.PerformClick() } catch { }
    })
}

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
