#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- Get-Recipient is called unqualified against
#                                 on-premises Exchange. Exchange Online exports
#                                 the same cmdlet name and in one process the
#                                 last connection wins, so an ExO session here
#                                 would answer for the cloud and the OU column
#                                 would be empty on every row.
# TOOLKIT-VARIANT: OnPremAD
<#
    Object-Lookup.ps1
    "What is that thing actually called, and what kind of thing is it?"

    A ticket names an object the way a person says it out loud -- 'the HQTS
    shared mailbox' -- and the next tool wants an address. In between sits a
    lookup in Active Directory, a squint at three similar names, and a guess at
    whether the one you want is a mailbox, a distribution list, or a security
    group that only looks mail-enabled.

    Reasoning: ScriptNotes/Object-Lookup.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-20
#>
# TOOLKIT-HELP-BEGIN
# Type part of a name. Get back every matching object, with what it IS and
# where it lives.
#
# Built for the moment a ticket says "give Send As on the HQTS-Shared mailbox"
# and you have no idea what the address is, or whether it is a mailbox at all.
# Search HQTS, read the Type column, double-click the right row, and the address
# is on your clipboard ready to paste into the tool that does the work.
#
# THE TYPE COLUMN IS THE POINT. These all look alike in a name list and behave
# nothing alike:
#
#   UserMailbox / SharedMailbox      Full Access and Send As both work
#   MailUniversalDistributionGroup   Send As works, Full Access is meaningless
#   MailUniversalSecurityGroup       a distribution list that also grants rights
#   MailContact                      no mailbox anywhere; nothing to delegate
#   RemoteSharedMailbox              the mailbox is in the cloud; the object here
#                                    is the on-premises stub that syncs it
#
# Getting this wrong is not a typo, it is twenty minutes and a confused reply.
#
# THE OU COLUMN answers the other half of the question -- which office or
# business unit owns this, and which directory it came from.
#
# ATTRIBUTES opens the selected object in an editor. Everything populated is
# shown; 46 attributes are editable -- names, job and organisation, every phone
# field, the full address, and all fifteen extensionAttributes.
#
# WHAT IS READ-ONLY, AND WHY IT IS A SHORT LIST:
#   addresses      mail and proxyAddresses belong to Exchange. Editing them
#                  behind its back is how an object gets two primaries.
#   identity       sAMAccountName, userPrincipalName, cn, name. Changing one of
#                  these is a migration, not an edit.
#   relationships  member, memberOf, managedBy, manager -- the value is a
#                  distinguished name, and a text box invites somebody to type
#                  a person's name into it.
#   state          userAccountControl and the timestamps. Enabling an account by
#                  hand-editing a bitmask is a way to be wrong quietly.
#   anything with more than one value, which is shown as 'multi (n)'. Writing
#   one value back would silently discard the rest.
#
# IT WORKS ON GROUPS, which the old All Attributes tools did not: they used
# Get-ADUser and simply failed on anything that was not a user. Half of what
# this finds is a distribution list.
#
# 'Also search Active Directory' adds objects with NO mailbox -- the plain
# security group sitting next to the shared mailbox with almost the same name.
# It cannot be delegated to, but seeing it explains why the name looked
# ambiguous. Off by default because it is slower.
#
# Reading is safe. The only thing that writes is the attribute editor, and it
# confirms every change and records it in the audit log.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: On-premises Exchange and Active Directory. Reads only, unless you edit an attribute.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

#region Environment ------------------------------------------------------------
# THE GUARD IS ASYMMETRIC -- see DL-Manager-Editor for the long version.
#
# Get-Recipient is called unqualified because there is no Invoke-EXOCommand
# equivalent for the management shell. An Exchange Online session in this
# process would capture the name and answer for the cloud, where there is no
# distinguishedName to derive an OU from -- so the tool would appear to work and
# quietly stop answering the question it exists for.
#
# WHY NOT USE Exchange Online AND DECODE msExchRecipientTypeDetails FROM AD?
# Because that means encoding a bitmask table -- 4 is SharedMailbox,
# 34359738368 is RemoteSharedMailbox, and so on -- and the Microsoft page that
# published it is archived with its table removed; the values now survive mainly
# in the comments underneath it. Asking Exchange means never maintaining that
# table, and never being subtly wrong about a type nobody thought to test.
$script:Blocked   = $false
$script:BlockText = ''

$emsState = Test-EmsAvailable
$exoState = Test-ExoInProcess

if ($exoState.Present) {
    $script:Blocked   = $true
    $script:BlockText = 'BLOCKED: Exchange Online is loaded in this process and would answer instead of on-premises Exchange. This must run in its own window.'
} elseif (-not $emsState.Available) {
    $script:Blocked   = $true
    $script:BlockText = 'NOT connected to on-premises Exchange -- nothing can run. Use Connect ExSE.'
}
#endregion

#region Actions ----------------------------------------------------------------
# A CEILING, ANNOUNCED WHEN IT IS HIT. Two letters against a directory this size
# would otherwise sit there and then draw a grid nobody can read. Capping is the
# easy half; saying so is the half that matters, or the operator concludes the
# thing they wanted does not exist.
$script:MaxResults = 250

# ATTRIBUTES THAT MAY BE EDITED HERE.
#
# RAW LDAP NAMES, NOT THE POWERSHELL FRIENDLY ONES. This is the one thing to get
# right when comparing against the old All Attributes tools: they used
# Get-ADUser, which renames things on the way out -- 'Office', 'OfficePhone',
# 'POBox', 'HomePage'. Get-ADObject does not, so the same fields are
# physicalDeliveryOfficeName, telephoneNumber, postOfficeBox and wWWHomePage.
# A friendly name in this list would silently match nothing and the field would
# quietly be read-only.
#
# Everything here is descriptive and belongs to AD. What is deliberately ABSENT
# matters more:
#
#   mail, proxyAddresses, targetAddress, mailNickname
#       Exchange's, not AD's. Setting them behind Exchange's back is how an
#       object ends up with two primaries, or a mailbox that no longer matches
#       the address book. Use Mail Attributes or DL Manager.
#   sAMAccountName, userPrincipalName, cn, name, distinguishedName
#       identity. Renaming one of these from a lookup window is not an edit, it
#       is a migration.
#   member, memberOf, managedBy, manager
#       relational. The value is a distinguished name, and a text box invites
#       somebody to type a person's name into it and wonder why it failed.
#   userAccountControl, accountExpires, pwdLastSet, and the timestamps
#       state, not description. Enabling an account by hand-editing a bitmask
#       is a way to be wrong quietly.
$script:EditableAttributes = @(
    # Names and description. displayName is here and 'name' is not -- one is a
    # label, the other is identity.
    'displayName', 'description', 'info', 'comment',
    'givenName', 'sn', 'initials',

    # Job and organisation
    'title', 'department', 'company', 'division', 'o', 'ou',
    'employeeID', 'employeeNumber', 'employeeType',
    'physicalDeliveryOfficeName',

    # Telephone. otherTelephone and url are deliberately absent: both are
    # genuinely multi-valued in normal use, and collapsing a list to one value
    # through a single text box is data loss dressed up as an edit.
    'telephoneNumber', 'homePhone', 'mobile', 'pager',
    'facsimileTelephoneNumber', 'ipPhone', 'wWWHomePage',

    # Address
    'streetAddress', 'postOfficeBox', 'l', 'st', 'postalCode', 'c', 'co',

    # The fifteen an organisation may use for anything at all, which is exactly
    # why they need to be editable -- nothing else in the toolkit knows what
    # they mean here.
    'extensionAttribute1',  'extensionAttribute2',  'extensionAttribute3',
    'extensionAttribute4',  'extensionAttribute5',  'extensionAttribute6',
    'extensionAttribute7',  'extensionAttribute8',  'extensionAttribute9',
    'extensionAttribute10', 'extensionAttribute11', 'extensionAttribute12',
    'extensionAttribute13', 'extensionAttribute14', 'extensionAttribute15'
)

# Noise: operational attributes that are always present, never interesting, and
# would bury the twenty that are.
$script:HiddenAttributes = @(
    'nTSecurityDescriptor', 'PropertyNames', 'AddedProperties', 'RemovedProperties',
    'ModifiedProperties', 'PropertyCount', 'ObjectCategory', 'dSCorePropagationData',
    'uSNChanged', 'uSNCreated', 'msExchVersion', 'replicationSignature'
)

function ConvertTo-OuPath {
    <#
        The OU a distinguished name sits in, in the form ADUC shows it.

        'CN=HQTS-Shared,OU=Shared Mailboxes,OU=HQ,DC=corp,DC=expresspersonnel,DC=com'
            ->  corp.expresspersonnel.com/HQ/Shared Mailboxes

        THE LEADING CN IS DROPPED, deliberately. It is the object itself, which
        the Name column already shows; repeating it in the OU column would push
        the part that differs off the right-hand edge.

        The OU components are REVERSED because a DN reads leaf-first and a person
        reads root-first. Nobody thinks of the object as being in 'Shared
        Mailboxes, which is in HQ'; they think 'HQ, then Shared Mailboxes'.
    #>
    param([string]$DistinguishedName)

    $dn = "$DistinguishedName"
    if (-not $dn) { return '' }

    # SPLIT ON UNESCAPED COMMAS ONLY. A display name with a comma in it --
    # 'CN=Armbruster\, Glenn' -- is escaped with a backslash in the DN, and
    # splitting on every comma would tear it in half and produce an OU path with
    # a fragment of somebody's surname in it.
    $parts = [regex]::Split($dn, '(?<!\\),')

    $ous = New-Object System.Collections.ArrayList
    $dcs = New-Object System.Collections.ArrayList
    foreach ($p in $parts) {
        $t = "$p".Trim()
        # THE INNER PARENTHESES ARE LOAD-BEARING. Inside a method call, a comma
        # is an ARGUMENT SEPARATOR before it is anything else -- so
        # .Add($x -replace '\\,', ',') is read as TWO arguments and fails at run
        # time with "Cannot find an overload for Add and the argument count: 2".
        # It parses cleanly, so nothing catches it until somebody searches.
        if ($t -match '^(?i)OU=(.+)$')      { [void]$ous.Add(($Matches[1] -replace '\\,', ',')) }
        elseif ($t -match '^(?i)DC=(.+)$')  { [void]$dcs.Add($Matches[1]) }
    }

    $domain = ($dcs -join '.')
    if ($ous.Count -eq 0) { return $domain }

    [array]::Reverse($ous)
    return ($domain + '/' + ($ous -join '/'))
}

function Get-MatchingRecipient {
    <#
        Every on-premises Exchange recipient whose name, display name, alias or
        address contains the text, from every directory in Tenant.json.

        Get-Recipient RATHER THAN Get-Mailbox OR Get-DistributionGroup, and that
        is the whole design. Those two each answer for one kind of object, so
        using them means knowing the answer before asking the question -- which
        is exactly what the operator does not have. Get-Recipient returns every
        mail-enabled thing there is and, more usefully, names which it is.

        SUBSTRING, NOT -Anr. Ambiguous name resolution is prefix-based: it finds
        HQTS-Shared from 'HQTS' but not 'Humboldt AR' from 'AR'. An operator with
        half a name in their head is as likely to have the middle of it.

        EVERY DIRECTORY IS SEARCHED, not the first that answers -- unlike the
        rename tool, which pins one because it then WRITES. Here the same name
        existing in HQ and Franchise both is not an ambiguity to resolve; it is
        the answer to the ticket.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $t = ConvertTo-OpathValue -Value $Text
    $filter = "(Name -like '*$t*') -or (DisplayName -like '*$t*') -or " +
              "(Alias -like '*$t*') -or (EmailAddresses -like '*$t*')"

    $rows   = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    $seen   = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)

    $dirs = @()
    try { $dirs = @(Get-TenantDirectory) } catch { }
    if ($dirs.Count -eq 0) { $dirs = @([PSCustomObject]@{ Id = '' }) }

    foreach ($d in $dirs) {
        $id = "$($d.Id)"

        # Resolve-DirectoryDomainController, NOT Get-TenantDirectoryServer:
        # ARCHITECTURE 14f, the latter returns a DOMAIN when a directory names
        # several controllers and Exchange rejects that with error 0x52.
        $dc = ''
        try { if ($id) { $dc = Resolve-DirectoryDomainController -DirectoryId $id } } catch { }

        $p = @{
            Filter      = $filter
            ResultSize  = ($script:MaxResults + 1)
            ErrorAction = 'Stop'
        }
        if ("$dc".Trim()) { $p['DomainController'] = $dc }

        try {
            foreach ($r in @(Get-Recipient @p)) {
                $dn = "$($r.DistinguishedName)"
                # De-duplicated on DN. A directory listed twice in Tenant.json,
                # or two entries pointing at one forest, would otherwise show
                # every object twice and make a clean result look ambiguous.
                if (-not $seen.Add($dn)) { continue }
                [void]$rows.Add([PSCustomObject]@{
                    # FREE. Get-Recipient already returned it, so capturing it
                    # here costs nothing -- unlike the account state below,
                    # which AD has to be asked for separately.
                    HiddenFromGal        = $r.HiddenFromAddressListsEnabled
                    Name                 = "$($r.Name)"
                    DisplayName          = "$($r.DisplayName)"
                    Alias                = "$($r.Alias)"
                    PrimarySmtpAddress   = "$($r.PrimarySmtpAddress)"
                    RecipientTypeDetails = "$($r.RecipientTypeDetails)"
                    DistinguishedName    = $dn
                    OuPath               = (ConvertTo-OuPath -DistinguishedName $dn)
                    EmailAddresses       = @($r.EmailAddresses | ForEach-Object { "$_" })
                    Directory            = $(if ($id) { $id } else { 'default' })
                    Server               = $dc
                    Source               = 'Exchange'
                })
            }
        } catch {
            [void]$errors.Add("$(if ($id) { $id } else { 'default directory' }): $($_.Exception.Message)")
        }
    }

    return [PSCustomObject]@{ Rows = @($rows.ToArray()); Errors = @($errors.ToArray()) }
}

function Get-MatchingAdObject {
    <#
        AD objects with no mailbox, which is the only reason to look here.

        Anything mail-enabled has already been found by Get-Recipient, so
        returning it again would double every row. What Exchange cannot see is
        the plain security group -- the 'HQTS' with no address sitting beside
        'HQTS-Shared' -- and that object is the reason the name looked ambiguous
        in the first place.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $out    = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList

    $esc  = ConvertTo-LdapFilterValue -Value $Text
    $ldap = "(&(|(objectClass=group)(objectClass=user))(|(name=*$esc*)(displayName=*$esc*)(sAMAccountName=*$esc*)))"

    $dirs = @()
    try { $dirs = @(Get-TenantDirectory) } catch { }

    foreach ($d in $dirs) {
        $id = "$($d.Id)"
        $server = ''
        try { if ($id) { $server = Resolve-DirectoryDomainController -DirectoryId $id } } catch { }

        $p = @{
            LDAPFilter    = $ldap
            Properties    = @('displayName', 'mail', 'sAMAccountName', 'objectClass', 'description')
            ResultSetSize = ($script:MaxResults + 1)
            ErrorAction   = 'Stop'
        }
        if ("$server".Trim()) { $p['Server'] = $server }

        try {
            foreach ($o in @(Get-ADObject @p)) {
                [void]$out.Add([PSCustomObject]@{
                    Name              = "$($o.Name)"
                    DisplayName       = "$($o.displayName)"
                    Mail              = "$($o.mail)"
                    Alias             = "$($o.sAMAccountName)"
                    Class             = "$($o.objectClass | Select-Object -Last 1)"
                    DistinguishedName = "$($o.DistinguishedName)"
                    Directory         = $(if ($id) { $id } else { 'default' })
                    Server            = $server
                })
            }
        } catch {
            [void]$errors.Add("$(if ($id) { $id } else { 'default directory' }): $($_.Exception.Message)")
        }
    }

    return [PSCustomObject]@{ Objects = @($out.ToArray()); Errors = @($errors.ToArray()) }
}

function ConvertTo-DnDisplayName {
    <#
        The readable name out of a distinguished name.

        'CN=Miller\, Jennifer L,OU=Users,DC=corp,DC=com'  ->  'Miller, Jennifer L'

        manager and managedBy hold a DN, and a DN is not an answer to "who
        manages this" -- it is the raw form of one. Printing it whole pushes
        everything else off the line.

        The escaped comma is unescaped LAST, after the split has used it to find
        the boundary. Doing it first would put the comma back and the split
        would then cut the surname in half.
    #>
    param([string]$Dn)
    $d = "$Dn"
    if (-not $d) { return '' }
    $first = [regex]::Split($d, '(?<!\\),')[0]
    $first = $first -replace '^(?i)CN=', ''
    return ($first -replace '\\,', ',')
}

function Get-CachedAdDetail {
    <#
        The AD-sourced facts for one row, read ONCE and kept on the row.

        ONE QUERY, NOT FIVE. Account state, job title, department, company,
        manager and managedBy all live on the same object -- asking for them
        separately would be five round trips per click for one object's worth of
        information. The detail pane redraws on every selection change, so
        clicking down a list of twenty results would be a hundred queries.

        Everything comes back as a string or empty. A caller printing 'if
        populated' should not also have to know which properties AD omits
        entirely versus returns as null.
    #>
    param([Parameter(Mandatory)]$Row)

    if ($Row.PSObject.Properties['AdDetail']) { return $Row.AdDetail }

    $d = [PSCustomObject]@{
        Uac        = $null
        Title      = ''
        Department = ''
        Company    = ''
        Manager    = ''
        ManagedBy  = ''
        Error      = ''
    }

    if ("$($Row.DistinguishedName)".Trim()) {
        try {
            $ap = @{
                Identity    = "$($Row.DistinguishedName)"
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
            # NOT FATAL. The address and the type are what most tickets needed;
            # losing the whole pane because one directory read failed would be a
            # poor trade.
            $d.Error = "$($_.Exception.Message)"
        }
    }

    Add-Member -InputObject $Row -NotePropertyName 'AdDetail' -NotePropertyValue $d -Force
    return $d
}

function Format-ObjectDetail {
    <#
        Everything about one row, including the full address list.

        THE ADDRESS LIST IS WHY THIS PANE EXISTS. The grid shows the primary,
        which answers most tickets -- but "mail sent to X is not arriving" is
        usually answered by a secondary address nobody remembered, and that is
        invisible until something prints the lot.
    #>
    param([Parameter(Mandatory)]$Row)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Name                 : $($Row.Name)")
    [void]$sb.AppendLine("DisplayName          : $($Row.DisplayName)")
    [void]$sb.AppendLine("Alias / sAMAccount   : $($Row.Alias)")
    [void]$sb.AppendLine("Type                 : $($Row.RecipientTypeDetails)")
    [void]$sb.AppendLine("OU                   : $($Row.OuPath)")
    [void]$sb.AppendLine("Directory            : $($Row.Directory)")
    [void]$sb.AppendLine("DistinguishedName    : $($Row.DistinguishedName)")

    # ORGANISATION AND OWNERSHIP, printed only when populated.
    #
    # ABOVE the Exchange/AD-only branch on purpose, so both kinds of row get
    # them. managedBy in particular is most useful on the object that has NO
    # mailbox -- a plain security group with an owner is the one you can ask
    # somebody about.
    #
    # Silent when empty rather than printing 'Title : ' with nothing after it.
    # A blank line is a fact about the layout; an absent line is a fact about
    # the object.
    $ad = Get-CachedAdDetail -Row $Row
    if ($ad.Title)      { [void]$sb.AppendLine("Title                : $($ad.Title)") }
    if ($ad.Department) { [void]$sb.AppendLine("Department           : $($ad.Department)") }
    if ($ad.Company)    { [void]$sb.AppendLine("Company              : $($ad.Company)") }
    if ($ad.Manager)    { [void]$sb.AppendLine("Manager              : $($ad.Manager)") }
    if ($ad.ManagedBy)  { [void]$sb.AppendLine("ManagedBy            : $($ad.ManagedBy)") }
    if ($ad.Error)      { [void]$sb.AppendLine("Directory read       : failed -- $($ad.Error)") }

    if ("$($Row.Source)" -eq 'AD') {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('FOUND IN ACTIVE DIRECTORY ONLY -- Exchange does not know about this object.')
        [void]$sb.AppendLine('It has no mailbox and no address, so nothing can be delegated to it and mail')
        [void]$sb.AppendLine('cannot be sent to it. It is listed because its name is why the search looked')
        [void]$sb.AppendLine('ambiguous. Attributes still works on it.')
        return $sb.ToString()
    }

    [void]$sb.AppendLine("PrimarySmtpAddress   : $($Row.PrimarySmtpAddress)")

    # HIDDEN AND DISABLED ARE HERE BECAUSE THEY EXPLAIN SYMPTOMS.
    #
    # Both can be reached through Attributes, and that was the objection worth
    # taking seriously -- this pane should not become a second, worse copy of
    # the attribute grid. These two earn the space anyway:
    #
    #   "they are not in the address book"   -> Hidden
    #   "they cannot sign in"                -> Disabled
    #
    # They are the first question of a whole class of ticket, and neither reads
    # clearly in the grid. userAccountControl is a bitmask -- it shows as 514,
    # not as 'disabled' -- and it is deliberately not editable there, so the
    # grid can only ever show you the number.
    if ($null -ne $Row.HiddenFromGal) {
        $h = if ($Row.HiddenFromGal) { 'YES -- not in the address book' } else { 'No' }
        [void]$sb.AppendLine("Hidden from GAL      : $h")
    }

    # ACCOUNT STATE, for things that have an account. A distribution list and a
    # contact have none, so the line is absent rather than answered wrongly.
    #
    # The value came from the same query as the block above -- see
    # Get-CachedAdDetail.
    if ("$($Row.RecipientTypeDetails)" -match '(?i)mailbox|mailuser' -and $null -ne $ad.Uac) {

        # 0x0002 is ACCOUNTDISABLE. THE RAW NUMBER IS PRINTED BESIDE IT on
        # purpose: if this decode is ever wrong, the value that proves it is on
        # the same line rather than one tool away.
        $acct = if ($ad.Uac -band 2) { "DISABLED  (userAccountControl $($ad.Uac))" }
                else                 { "Enabled  (userAccountControl $($ad.Uac))" }
        [void]$sb.AppendLine("Account              : $acct")
    }

    # LICENCE, FOR THINGS THAT CAN HOLD ONE.
    #
    # Get-UserLicenseState reads AD GROUP MEMBERSHIP, not Graph -- licensing
    # here is group-based, so the licence somebody holds is a fact about the
    # directory and readable with the rights the operator already has. It
    # reports the GROUP that drives the assignment rather than the SKU, because
    # the group is the thing that can be acted on.
    #
    # ONLY FOR MAILBOXES. A distribution list cannot hold a licence and a
    # contact has no account, so asking would cost a directory query per row to
    # print 'no' -- and 'no' next to a distribution list reads like a finding
    # rather than a category error.
    #
    # SHARED AND ROOM MAILBOXES ARE INCLUDED DELIBERATELY. They are the ones
    # worth checking: a licensed shared mailbox is usually somebody's mistake,
    # and Mailbox Health raises it as an issue for exactly that reason.
    if ("$($Row.RecipientTypeDetails)" -match '(?i)mailbox' -and "$($Row.PrimarySmtpAddress)".Trim()) {

        # CACHED ON THE ROW after the first look. The detail pane redraws on
        # every selection change, and clicking down a list of twenty would
        # otherwise be twenty directory queries for answers already known.
        if (-not $Row.PSObject.Properties['LicenseText']) {
            $lic = ''
            try {
                $st  = Get-UserLicenseState -Smtp "$($Row.PrimarySmtpAddress)"
                $lic = Format-UserLicenseState -State $st
            } catch {
                # A licence lookup that fails is not a reason to lose the rest
                # of the pane -- the address and the type are what most tickets
                # actually needed.
                $lic = "could not be read: $($_.Exception.Message)"
            }
            Add-Member -InputObject $Row -NotePropertyName 'LicenseText' -NotePropertyValue $lic -Force
        }
        [void]$sb.AppendLine("Licence              : $($Row.LicenseText)")
    }

    $addrs = @($Row.EmailAddresses)
    if ($addrs.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("All addresses ($($addrs.Count)):")
        # Upper-case SMTP: first -- it is the primary, and the case is the only
        # thing that says so.
        $ordered = $addrs | Sort-Object `
            { if ($_ -cmatch '^SMTP:') { 0 } elseif ($_ -match '^(?i)smtp:') { 1 } else { 2 } },
            { "$_" }
        foreach ($a in $ordered) { [void]$sb.AppendLine("    $a") }
    }
    return $sb.ToString()
}
#endregion

#region Attribute editor -------------------------------------------------------
function Show-AttributeEditor {
    <#
        Every populated attribute on one object, with the curated set editable.

        Get-ADObject / Set-ADObject RATHER THAN Get-ADUser / Set-ADUser, and
        that is not a style choice. Half of what this tool finds is a GROUP --
        the shared mailbox sits next to three distribution lists -- and
        Get-ADUser simply fails on those. Get-ADObject is the only pair that
        works for both without branching.

        THE AD MODULE IS SAFE IN THIS PROCESS even though Exchange is loaded:
        Get-ADObject shares no cmdlet name with either Exchange session, so
        unlike Get-Recipient it cannot be captured by the wrong one.

        Read-only until something is edited, and every edit confirms first and
        is written to the audit log.
    #>
    param(
        [Parameter(Mandatory)]$Row,
        [Parameter(Mandatory)]$Owner
    )

    $server = "$($Row.Server)"
    $dn     = "$($Row.DistinguishedName)"

    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Attributes -- $($Row.Name)"
    $dlg.ClientSize      = New-Object System.Drawing.Size(760, 560)
    $dlg.StartPosition   = 'CenterParent'
    $dlg.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblWho           = New-Object System.Windows.Forms.Label
    $lblWho.Location  = New-Object System.Drawing.Point(12, 12)
    $lblWho.Size      = New-Object System.Drawing.Size(700, 34)
    $lblWho.Text      = "$($Row.DistinguishedName)"
    $lblWho.ForeColor = [System.Drawing.Color]::DimGray
    $dlg.Controls.Add($lblWho)

    $grd = New-Object System.Windows.Forms.DataGridView
    $grd.Location = New-Object System.Drawing.Point(12, 52)
    $grd.Size     = New-Object System.Drawing.Size(736, 380)
    $grd.ReadOnly = $true
    $grd.AllowUserToAddRows    = $false
    $grd.AllowUserToDeleteRows = $false
    $grd.AllowUserToResizeRows = $false
    $grd.RowHeadersVisible     = $false
    $grd.SelectionMode         = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grd.MultiSelect           = $false
    $grd.BackgroundColor       = [System.Drawing.Color]::White
    $grd.Font                  = New-Object System.Drawing.Font('Consolas', 9)
    $grd.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
    foreach ($c in @(
        @{ N = 'Attr';  H = 'Attribute'; W = 230 },
        @{ N = 'Value'; H = 'Value';     W = 420 },
        @{ N = 'Edit';  H = 'Editable';  W = 76  })) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.N; $col.HeaderText = $c.H; $col.Width = $c.W
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        [void]$grd.Columns.Add($col)
    }
    $dlg.Controls.Add($grd)

    $lblNote           = New-Object System.Windows.Forms.Label
    $lblNote.Location  = New-Object System.Drawing.Point(12, 440)
    $lblNote.Size      = New-Object System.Drawing.Size(736, 32)
    $lblNote.Text      = 'Addresses, identity and membership are not editable here -- they belong to Exchange or are relational. Use Mail Attributes or DL Manager for those.'
    $lblNote.ForeColor = [System.Drawing.Color]::DimGray
    $dlg.Controls.Add($lblNote)

    $lblSt          = New-Object System.Windows.Forms.Label
    $lblSt.Location = New-Object System.Drawing.Point(12, 478)
    $lblSt.Size     = New-Object System.Drawing.Size(736, 34)
    $dlg.Controls.Add($lblSt)

    $btnEdit          = New-Object System.Windows.Forms.Button
    $btnEdit.Text     = 'Edit attribute'
    $btnEdit.Location = New-Object System.Drawing.Point(12, 518)
    $btnEdit.Size     = New-Object System.Drawing.Size(140, 30)
    $btnEdit.Enabled  = $false
    $dlg.Controls.Add($btnEdit)

    $btnRefresh          = New-Object System.Windows.Forms.Button
    $btnRefresh.Text     = 'Refresh'
    $btnRefresh.Location = New-Object System.Drawing.Point(160, 518)
    $btnRefresh.Size     = New-Object System.Drawing.Size(110, 30)
    $dlg.Controls.Add($btnRefresh)

    $btnDone          = New-Object System.Windows.Forms.Button
    $btnDone.Text     = 'Close'
    $btnDone.Location = New-Object System.Drawing.Point(638, 518)
    $btnDone.Size     = New-Object System.Drawing.Size(110, 30)
    $dlg.Controls.Add($btnDone)

    function Get-AdParam {
        $p = @{ Identity = $dn; Properties = '*'; ErrorAction = 'Stop' }
        if ("$server".Trim()) { $p['Server'] = $server }
        return $p
    }

    function Update-AttributeGrid {
        $grd.Rows.Clear()
        $btnEdit.Enabled = $false
        try {
            # A PLAIN VARIABLE, because that is the only thing splatting takes.
            # 'Get-ADObject @(Get-AdParam)[0]' looks like it splats and does
            # not: @(...) there is an array subexpression, so the hashtable is
            # passed as a single POSITIONAL argument -- binding to -Identity as
            # a literal object and failing with a message about the identity
            # rather than about the syntax. Same family as the
            # '@script:Foo' trap in ARCHITECTURE 33.
            $adParam = Get-AdParam
            $o = Get-ADObject @adParam
        } catch {
            $lblSt.ForeColor = [System.Drawing.Color]::Firebrick
            $lblSt.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADObject'
            return
        }

        foreach ($prop in ($o.PSObject.Properties | Sort-Object Name)) {
            if ($script:HiddenAttributes -contains $prop.Name) { continue }
            $v = $prop.Value
            if ($null -eq $v) { continue }

            # Multi-valued collections are FLATTENED FOR DISPLAY and never
            # offered for edit. Showing 'member' as a count rather than hiding
            # it matters -- "why can't I edit this" is answered by seeing that
            # it holds 34 things.
            $text = ''
            $valueCount = 1
            if ($v -is [string]) { $text = $v }
            elseif ($v -is [System.Collections.IEnumerable]) {
                $items = @($v | ForEach-Object { "$_" })
                if ($items.Count -eq 0) { continue }
                $valueCount = $items.Count
                $text = if ($items.Count -eq 1) { $items[0] } else { "($($items.Count)) " + ($items -join ' | ') }
            }
            else { $text = "$v" }

            if (-not "$text".Trim()) { continue }

            # EDITABLE IF THE LIST ALLOWS IT AND THERE IS ONE VALUE TO EDIT.
            #
            # 'description' is the reason this is not just a type test. It is
            # multi-valued IN THE SCHEMA even though every tool shows one box,
            # so Get-ADObject can hand it back as a collection -- and a plain
            # '-is [string]' test would make the single most commonly edited
            # field on a group permanently read-only.
            #
            # More than one value stays read-only on purpose: writing one value
            # back would silently discard the rest, which is data loss wearing
            # the costume of an edit.
            $allowed  = ($script:EditableAttributes -contains $prop.Name)
            $editable = $allowed -and ($v -is [string] -or ($valueCount -eq 1))
            # The reason a field is read-only is shown, not left to be guessed
            # at. 'multi (3)' and blank mean different things, and the operator
            # who wants to know why is otherwise reading the source.
            $why = if ($editable) { 'Yes' } elseif ($allowed -and $valueCount -gt 1) { "multi ($valueCount)" } else { '' }
            $i = $grd.Rows.Add($prop.Name, $text, $why)
            $grd.Rows[$i].Tag = [PSCustomObject]@{ Name = $prop.Name; Value = $text; Editable = $editable }
            if (-not $editable) { $grd.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray }
        }
        $lblSt.ForeColor = [System.Drawing.Color]::DimGray
        $lblSt.Text = "$($grd.Rows.Count) populated attribute(s)."
    }

    $grd.Add_SelectionChanged({
        $btnEdit.Enabled = $false
        if ($grd.SelectedRows.Count -eq 0) { return }
        $t = $grd.SelectedRows[0].Tag
        if ($t -and $t.Editable) { $btnEdit.Enabled = $true }
    })

    function Invoke-AttributeEdit {
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
        $l2.Text = 'New value  (leave empty to clear the attribute):'
        $l2.Location = New-Object System.Drawing.Point(12, 66); $l2.Size = New-Object System.Drawing.Size(480, 18)
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

        $newValue = $tn.Text
        if ("$newValue" -ceq "$($t.Value)") {
            $lblSt.ForeColor = [System.Drawing.Color]::DimGray
            $lblSt.Text = 'Unchanged -- nothing written.'
            return
        }

        $shown = if ("$newValue".Trim() -eq '') { '(cleared)' } else { $newValue }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            ("Change $($t.Name) on $($Row.Name)?" + [Environment]::NewLine + [Environment]::NewLine +
             "From: $($t.Value)" + [Environment]::NewLine + "To:   $shown"),
            'Confirm attribute change',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        try {
            $p = @{ Identity = $dn; ErrorAction = 'Stop' }
            if ("$server".Trim()) { $p['Server'] = $server }

            # CLEAR AND SET ARE DIFFERENT OPERATIONS. -Replace with an empty
            # string is not "empty", it is an invalid value, and AD rejects it.
            # Removing the attribute is -Clear.
            if ("$newValue".Trim() -eq '') { $p['Clear']   = $t.Name }
            else                           { $p['Replace'] = @{ $t.Name = $newValue } }

            Set-ADObject @p

            Write-CompletedLog -Action 'EditAttribute' -Target "$($Row.Name)" -Result 'Success' `
                -Environment "$($Row.Directory)" `
                -Detail "$($t.Name): '$($t.Value)' -> '$shown' (Object Lookup)"

            $lblSt.ForeColor = [System.Drawing.Color]::ForestGreen
            $lblSt.Text = "$($t.Name) changed. Grid refreshed from the directory."
            Update-AttributeGrid
        } catch {
            Write-FailedLog -Action 'EditAttribute' -Target "$($Row.Name)" -ErrorRecord $_ `
                -Environment "$($Row.Directory)" -Detail "$($t.Name)"
            $lblSt.ForeColor = [System.Drawing.Color]::Firebrick
            $lblSt.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-ADObject'
        }
    }

    $btnEdit.Add_Click({ Invoke-AttributeEdit })
    $grd.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        Invoke-AttributeEdit
    })
    $btnRefresh.Add_Click({ Update-AttributeGrid })
    $btnDone.Add_Click({ $dlg.Close() })

    Update-AttributeGrid
    Set-ConsoleTheme -Form $dlg
    [void]$dlg.ShowDialog($Owner)
}
#endregion

#region WinForms UI ------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Object Lookup'
$form.ClientSize    = New-Object System.Drawing.Size(1044, 950)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Tag       = 'NoTheme'
$lblWarn.Location  = New-Object System.Drawing.Point(12, 12)
$lblWarn.Size      = New-Object System.Drawing.Size(980, 18)
$lblWarn.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
if ($script:Blocked) {
    $lblWarn.Text      = $script:BlockText
    $lblWarn.ForeColor = [System.Drawing.Color]::Red
} else {
    $lblWarn.Text      = 'ON-PREMISES. Reads Exchange and Active Directory. Double-click a row to copy its address.'
    $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
}
$form.Controls.Add($lblWarn)

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

$chkAd          = New-Object System.Windows.Forms.CheckBox
$chkAd.Text     = 'Also search Active Directory (finds objects with no mailbox)'
$chkAd.Location = New-Object System.Drawing.Point(504, 42)
$chkAd.Size     = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($chkAd)

$lblCount           = New-Object System.Windows.Forms.Label
$lblCount.Location  = New-Object System.Drawing.Point(12, 68)
$lblCount.Size      = New-Object System.Drawing.Size(1020, 18)
$lblCount.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblCount)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 90)
$grid.Size     = New-Object System.Drawing.Size(1020, 300)
$grid.Anchor   = 'Top,Left,Right'
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
foreach ($c in @(
    @{ N = 'Name';    H = 'Name';          W = 170 },
    @{ N = 'Display'; H = 'Display Name';  W = 170 },
    @{ N = 'Address'; H = 'Email Address'; W = 230 },
    @{ N = 'Type';    H = 'Type';          W = 210 },
    @{ N = 'Ou';      H = 'OU';            W = 232 })) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.N; $col.HeaderText = $c.H; $col.Width = $c.W
    # SORTABLE. Twenty rows from one office sort into their kinds in one click,
    # which is the fastest way to see that the shared mailbox is the only
    # SharedMailbox among six distribution groups.
    $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
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
$txtDetail.Size       = New-Object System.Drawing.Size(1020, 480)
$txtDetail.Font       = New-Object System.Drawing.Font('Consolas', 9)

# THE DETAIL BOX ABSORBS ANY RESIZE, and the buttons stay put relative to the
# bottom edge. That is not decoration now that the form is ~989 tall outside:
# Set-FormScreenCorner SHRINKS a window bigger than the working area, and on an
# unanchored form that clamp would put the button row off the bottom -- exactly
# the fault LAY001 exists to catch, arriving at run time on somebody else's
# screen instead of at design time on mine.
#
# The grid stays Top,Left,Right. If both it and the box grew, a resize would
# split the difference between them and neither would end up the size anyone
# asked for.
$txtDetail.Anchor     = 'Top,Left,Right,Bottom'
$form.Controls.Add($txtDetail)

$btnCopy          = New-Object System.Windows.Forms.Button
$btnCopy.Text     = 'Copy address'
$btnCopy.Location = New-Object System.Drawing.Point(12, 908)
$btnCopy.Size     = New-Object System.Drawing.Size(130, 30)
$btnCopy.Anchor   = 'Left,Bottom'
$form.Controls.Add($btnCopy)

$btnCopyDetail          = New-Object System.Windows.Forms.Button
$btnCopyDetail.Text     = 'Copy all detail'
$btnCopyDetail.Location = New-Object System.Drawing.Point(150, 908)
$btnCopyDetail.Size     = New-Object System.Drawing.Size(130, 30)
$btnCopyDetail.Anchor   = 'Left,Bottom'
$form.Controls.Add($btnCopyDetail)

$btnAttrs          = New-Object System.Windows.Forms.Button
$btnAttrs.Text     = 'Attributes...'
$btnAttrs.Location = New-Object System.Drawing.Point(288, 908)
$btnAttrs.Size     = New-Object System.Drawing.Size(130, 30)
$btnAttrs.Enabled  = $false
$btnAttrs.Anchor   = 'Left,Bottom'
$form.Controls.Add($btnAttrs)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(428, 916)
$lblStatus.Size     = New-Object System.Drawing.Size(480, 18)
$lblStatus.Anchor   = 'Left,Bottom'
$form.Controls.Add($lblStatus)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(922, 908)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Anchor   = 'Right,Bottom'
$form.Controls.Add($btnClose)

function Get-SelectedRow {
    if ($grid.SelectedRows.Count -eq 0) { return $null }
    # THE OBJECT RIDES ON THE ROW'S Tag, not on an index into a parallel array.
    # The grid is sortable, so the row at position 3 on screen stops being
    # item 3 the moment a column header is clicked -- and sorting moves the Tag
    # with the row.
    return $grid.SelectedRows[0].Tag
}

function Set-CopiedStatus {
    param([string]$What)
    $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen
    $lblStatus.Text      = "$What copied."
}

function Copy-SelectedAddress {
    $r = Get-SelectedRow
    if (-not $r) { return }
    $a = "$($r.PrimarySmtpAddress)"
    if (-not $a) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = 'That object has no address -- nothing to copy.'
        return
    }
    try { [System.Windows.Forms.Clipboard]::SetText($a); Set-CopiedStatus -What $a }
    catch {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = "Could not copy: $($_.Exception.Message)"
    }
}

$btnSearch.Add_Click({
    if ($script:Blocked) {
        $lblCount.ForeColor = [System.Drawing.Color]::Firebrick
        $lblCount.Text = 'Blocked -- see the message at the top.'
        return
    }
    $text = $txtSearch.Text.Trim()
    if ($text.Length -lt 2) {
        $lblCount.ForeColor = [System.Drawing.Color]::Firebrick
        $lblCount.Text = 'Type at least two characters.'
        return
    }

    $btnSearch.Enabled = $false
    $form.Cursor       = [System.Windows.Forms.Cursors]::WaitCursor
    $grid.Rows.Clear()
    $txtDetail.Text  = ''
    $lblStatus.Text  = ''
    $btnAttrs.Enabled = $false
    $lblCount.ForeColor = [System.Drawing.Color]::DimGray
    $lblCount.Text   = 'Searching...'
    $form.Refresh()

    try {
        $ex     = Get-MatchingRecipient -Text $text
        $rows   = @($ex.Rows)
        $notes  = @($ex.Errors)
        $capped = ($rows.Count -gt $script:MaxResults)
        if ($capped) { $rows = @($rows | Select-Object -First $script:MaxResults) }

        $all = New-Object System.Collections.ArrayList
        foreach ($r in $rows) { [void]$all.Add($r) }

        if ($chkAd.Checked) {
            $ad = Get-MatchingAdObject -Text $text
            # De-duplicated on DN, which is exact. Matching on name would drop a
            # genuinely separate object that happens to share a name across two
            # directories -- the very thing worth seeing.
            $known = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($r in $rows) { [void]$known.Add("$($r.DistinguishedName)") }

            foreach ($o in $ad.Objects) {
                if ($known.Contains($o.DistinguishedName)) { continue }
                if ("$($o.Mail)".Trim()) { continue }
                [void]$all.Add([PSCustomObject]@{
                    Name                 = $o.Name
                    DisplayName          = $o.DisplayName
                    Alias                = $o.Alias
                    PrimarySmtpAddress   = ''
                    RecipientTypeDetails = "$($o.Class) (no mailbox)"
                    DistinguishedName    = $o.DistinguishedName
                    OuPath               = (ConvertTo-OuPath -DistinguishedName $o.DistinguishedName)
                    EmailAddresses       = @()
                    Directory            = $o.Directory
                    Server               = $o.Server
                    Source               = 'AD'
                })
            }
            $notes += @($ad.Errors)
        }

        foreach ($r in $all) {
            $i = $grid.Rows.Add("$($r.Name)", "$($r.DisplayName)", "$($r.PrimarySmtpAddress)",
                                "$($r.RecipientTypeDetails)", "$($r.OuPath)")
            $grid.Rows[$i].Tag = $r
            # AD-only rows are greyed: context, not answers. Nothing can be
            # delegated to them and nothing can be mailed to them.
            if ("$($r.Source)" -eq 'AD') {
                $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
            }
        }

        $note = if ($notes.Count) { '   ' + ($notes -join '; ') } else { '' }

        if ($all.Count -eq 0) {
            $lblCount.ForeColor = [System.Drawing.Color]::Firebrick
            $lblCount.Text = "Nothing matched '$text'." +
                             $(if (-not $chkAd.Checked) { ' If you expected a group with no mailbox, tick the Active Directory box.' } else { '' }) + $note
        } elseif ($capped) {
            $lblCount.ForeColor = [System.Drawing.Color]::Firebrick
            $lblCount.Text = "More than $($script:MaxResults) matches -- showing the first $($script:MaxResults). Type more of the name.$note"
        } else {
            $lblCount.ForeColor = [System.Drawing.Color]::DimGray
            $lblCount.Text = "$($all.Count) match(es).$note"
        }

        # SELECTED, THEN UPDATED EXPLICITLY. ClearSelection first so that
        # selecting row 0 is a real change rather than a no-op -- but the
        # explicit call is what actually guarantees it, and the event is left
        # to handle the operator clicking about afterwards.
        if ($grid.Rows.Count -gt 0) {
            $grid.ClearSelection()
            $grid.Rows[0].Selected = $true
        }
        Update-SelectedDetail
    } catch {
        $txtDetail.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Object lookup'
        $lblCount.ForeColor = [System.Drawing.Color]::Firebrick
        $lblCount.Text = 'Search failed -- see below.'
    } finally {
        $btnSearch.Enabled = $true
        $form.Cursor       = [System.Windows.Forms.Cursors]::Default
    }
})

function Update-SelectedDetail {
    <#
        Redraws the detail pane and the Attributes button for whatever is
        selected now.

        A FUNCTION AND NOT JUST AN EVENT BODY, because the event cannot be
        relied on to fire at the moment the data is ready.

        Rows.Add() selects the first row as a side effect, so SelectionChanged
        fires DURING the loop that fills the grid -- before that row's .Tag has
        been set. Get-SelectedRow returns $null, the pane clears and the button
        greys. The 'Rows[0].Selected = $true' afterwards then changes nothing,
        because the row is already selected, so no second event ever arrives.

        With several matches this was invisible: clicking any other row fired a
        real change and everything worked. With EXACTLY ONE match there is no
        other row to click, and the tool looked broken.
    #>
    $r = Get-SelectedRow
    $btnAttrs.Enabled = [bool]$r
    $txtDetail.Text   = if ($r) { Format-ObjectDetail -Row $r } else { '' }
}

$grid.Add_SelectionChanged({ Update-SelectedDetail })

# DOUBLE-CLICK COPIES. The whole tool exists to get an address into the next
# window, and the shortest path to that is the one people find by accident.
$grid.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    Copy-SelectedAddress
})

$btnCopy.Add_Click({ Copy-SelectedAddress })

$btnCopyDetail.Add_Click({
    if (-not $txtDetail.Text) { return }
    try { [System.Windows.Forms.Clipboard]::SetText($txtDetail.Text); Set-CopiedStatus -What 'Detail' }
    catch {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = "Could not copy: $($_.Exception.Message)"
    }
})

$btnAttrs.Add_Click({
    $r = Get-SelectedRow
    if (-not $r) { return }
    Show-AttributeEditor -Row $r -Owner $form
})

$btnClose.Add_Click({ $form.Close() })

if ($script:Blocked) { $btnSearch.Enabled = $false }

# Enter searches. This is a box people type into and immediately press Enter.
$form.AcceptButton = $btnSearch

Set-ConsoleTheme -Form $form
Set-FormResizable -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Object Lookup' | Out-Null

[void]$form.ShowDialog()
#endregion
