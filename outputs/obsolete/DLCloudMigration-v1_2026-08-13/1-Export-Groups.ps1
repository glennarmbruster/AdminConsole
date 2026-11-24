#Requires -Version 5.1
<#
    1-Export-Groups.ps1
    DL cloud migration -- phase 1 data gathering.

    Reasoning: ScriptNotes/DLCloudMigration.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-08
#>
# TOOLKIT-REQUIRES: OwnProcess  -- reads on-premises Exchange for several
# minutes and must not tie up the console while it does.
# TOOLKIT-HELP-BEGIN
# Exports everything needed to recreate on-premises distribution groups as
# cloud-native groups. READ ONLY -- this script changes nothing.
#
# Run it once per environment. Each run writes a set of CSVs stamped with the
# same timestamp, so attributes, members and owners are always one consistent
# snapshot rather than three files taken minutes apart.
#
# In scope: distribution groups, room lists, dynamic distribution groups.
# Out of scope: mail-enabled security groups and plain security groups -- but
# where one is a MEMBER of an in-scope group it is still exported and flagged,
# because that membership cannot be reproduced in a cloud group.
#
# Read the Blockers CSV before planning anything. A group with no blockers can
# be rebuilt from these files alone; everything else needs a decision first.
# TOOLKIT-HELP-END

# Directory names are NOT in a ValidateSet.
#
# This said [ValidateSet('HQ','Franchise')] -- one organisation's internal
# structure written into a type constraint, in a script meant to be portable.
# It is the exact pattern removed from Get-ADServerString for the same reason,
# reintroduced here by the person who removed it.
#
# Two ways it failed. A tenant whose directories are named anything else could
# not run these at all. And an EMPTY value -- which the console used to pass
# before a lookup -- was rejected by the validator before the default could
# apply, so the script died on its own default value.
#
# Validated against Tenant.json at runtime instead, where the answer lives.
param(
    [string]$Environment = '',
    [switch]$NoPrompt
)

# Common.ps1 lives at the TOOLKIT ROOT and there is exactly one of it. Walk up
# to find it rather than counting folders -- these scripts sit three levels
# down and the next project may not.
$toolkitRoot = $PSScriptRoot
while ($toolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $toolkitRoot 'Common.ps1'))) {
    $toolkitRoot = Split-Path -Parent $toolkitRoot
}
if (-not $toolkitRoot) {
    throw "Common.ps1 was not found in any folder above '$PSScriptRoot'."
}
. (Join-Path $toolkitRoot 'Common.ps1')

# Resolve the directory before anything uses it.
$script:DirIds = @(@(Get-TenantDirectory) | ForEach-Object { "$($_.Id)" } | Where-Object { $_ })
if (-not "$Environment".Trim()) {
    $def = @(@(Get-TenantDirectory) | Where-Object { $_.IsDefault })
    if ($def.Count -gt 0) { $Environment = "$($def[0].Id)" }
    elseif ($script:DirIds.Count -gt 0) { $Environment = $script:DirIds[0] }
}
if ($script:DirIds.Count -gt 0 -and ($script:DirIds -notcontains $Environment)) {
    throw ("Directory '$Environment' is not defined in Tenant.json. " +
           "Configured directories: $($script:DirIds -join ', ')")
}
Add-Type -AssemblyName System.Windows.Forms

#region Scope -------------------------------------------------------------------
# RecipientTypeDetails, not RecipientType. RecipientType lumps every mail-enabled
# group together as MailUniversalDistributionGroup regardless of whether it is a
# security group, so filtering on it would drag every mail-enabled security group
# into a migration that explicitly excludes them.
$script:InScopeTypes = @('MailUniversalDistributionGroup', 'RoomList')

# Types a cloud distribution group can actually hold as a member. Anything else
# is exported and flagged rather than dropped -- a member that silently fails to
# migrate is how a group arrives in the cloud quietly missing people.
$script:CloudMemberTypes = @(
    'UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox',
    'RemoteUserMailbox', 'RemoteSharedMailbox', 'RemoteRoomMailbox',
    'RemoteEquipmentMailbox', 'MailUser', 'MailContact',
    'MailUniversalDistributionGroup', 'RoomList'
)
#endregion

#region Helpers -----------------------------------------------------------------
function Join-Multi {
    <#
        Multi-valued Exchange properties into one CSV cell.

        Semicolon, not comma: these values are addresses and display names, and
        display names contain commas ("Doe, Jane") often enough that a
        comma-joined cell cannot be split back apart reliably.
    #>
    param($Value)
    if ($null -eq $Value) { return '' }
    $items = @($Value | ForEach-Object { "$_" } | Where-Object { $_ -ne '' })
    return ($items -join '; ')
}

#region Data quality capture ----------------------------------------------------
# Exchange warns about objects it considers corrupt or inconsistent, and those
# warnings scroll off the screen and are gone. They are not noise -- they are a
# defect list. An address like 'AR,CoppellTX@ExpressPros.com' has a comma in the
# local part, which is not a legal SMTP address; that group cannot be recreated
# in the cloud until somebody fixes it on-premises.
#
# Finding them during the export is cheap. Finding them during the migration
# weekend is not.
$script:DataQuality     = New-Object System.Collections.ArrayList
$script:LastWarnObject  = ''

function Add-DataQualityNote {
    <#
        One warning, categorised, with the object it belongs to.

        Exchange emits these in RUNS: a header naming the object, then one or
        more detail lines that name only the property. The detail lines are the
        useful part and they carry no identity of their own, so the last object
        seen is carried forward -- otherwise the actionable half of every
        warning arrives orphaned.
    #>
    param([string]$Source, [string]$Text)

    $t = ("$Text" -replace '\s+', ' ').Trim()
    if (-not $t) { return }

    # Object paths appear as 'domain.tld/OU/Name', quoted or bare.
    $obj = ''
    $m = [regex]::Match($t, 'The object ([A-Za-z0-9.\-]+\.[A-Za-z]{2,}/\S[^"]*?) has been')
    if (-not $m.Success) { $m = [regex]::Match($t, 'The group "([^"]+)"') }
    if ($m.Success) {
        $obj = $m.Groups[1].Value.Trim()
        $script:LastWarnObject = $obj
    } else {
        $obj = $script:LastWarnObject
    }

    $prop = ''
    $pm = [regex]::Match($t, 'Could not convert property (\w+)')
    if ($pm.Success) { $prop = $pm.Groups[1].Value }

    $category =
        if ($t -match '(?i)could not convert property')            { 'InvalidAttributeValue' }
        elseif ($t -match '(?i)is not managed by any recipient')    { 'NoManagerButRestrictionSet' }
        elseif ($t -match '(?i)corrupted|inconsistent state')       { 'CorruptOrInconsistent' }
        else                                                        { 'Other' }

    [void]$script:DataQuality.Add([PSCustomObject]@{
        ObjectPath = $obj
        ObjectName = if ($obj) { ($obj -split '/')[-1] } else { '' }
        Category   = $category
        Property   = $prop
        Source     = $Source
        Warning    = $t
    })
}

function Invoke-OnPremCaptured {
    <#
        Invoke-OnPremCommand with the warning stream captured instead of printed.

        3>&1 merges warnings into the success stream so they can be separated by
        type. WarningAction is deliberately NOT set to SilentlyContinue -- that
        would suppress them at the source and there would be nothing to capture.
    #>
    param([string]$Name, [hashtable]$Parameters = @{}, [string]$Source)

    $out = New-Object System.Collections.ArrayList
    foreach ($item in (Invoke-OnPremCommand -Name $Name -Parameters $Parameters 3>&1)) {
        if ($item -is [System.Management.Automation.WarningRecord]) {
            Add-DataQualityNote -Source $Source -Text "$($item.Message)"
        } else {
            [void]$out.Add($item)
        }
    }
    return $out.ToArray()
}
#endregion

$script:RecipientCache = @{}
function Resolve-RecipientSmtp {
    <#
        A name, DN or alias to its PrimarySmtpAddress.

        ManagedBy, GrantSendOnBehalfTo and the restriction lists all return
        identities rather than addresses, and the cloud rebuild needs addresses
        -- an on-premises DN means nothing there.

        Cached because the same handful of managers own hundreds of groups, and
        an uncached lookup per reference turns a five minute export into an
        hour. Failures are cached too: a DN that does not resolve will not
        resolve on the next of its four hundred appearances either.
    #>
    param($Identity)

    $key = "$Identity"
    if ($key -eq '') { return $null }
    if ($script:RecipientCache.ContainsKey($key)) { return $script:RecipientCache[$key] }

    $result = $null
    try {
        $r = Invoke-OnPremCommand -Name 'Get-Recipient' -Parameters @{
            Identity = $key; ErrorAction = 'Stop' }
        if ($r) {
            $result = [PSCustomObject]@{
                Name        = "$($r.Name)"
                Smtp        = "$($r.PrimarySmtpAddress)"
                TypeDetails = "$($r.RecipientTypeDetails)"
                Resolved    = $true
            }
        }
    } catch {
        $result = [PSCustomObject]@{
            Name = $key; Smtp = ''; TypeDetails = ''; Resolved = $false
        }
    }
    $script:RecipientCache[$key] = $result
    return $result
}

function Get-ResolvedSmtpList {
    param($Identities)
    $out = @()
    foreach ($i in @($Identities)) {
        $r = Resolve-RecipientSmtp -Identity $i
        if ($r -and $r.Smtp) { $out += $r.Smtp }
        elseif ($r) { $out += "UNRESOLVED:$($r.Name)" }
    }
    return (Join-Multi $out)
}

function Get-GroupEnvironment {
    <#
        Which directory a group lives in, from Tenant.json rather than from a
        hardcoded OU string.

        Get-DistributionGroup returns OrganizationalUnit as
        'domain.example.com/OU/Sub', so the leading segment is the domain and
        can be matched against the configured directories. Nothing here knows
        the name of any particular directory.
    #>
    param([string]$OrganizationalUnit, $DirectoryMap)

    $ouDomain = ("$OrganizationalUnit" -split '/')[0]
    foreach ($d in $DirectoryMap) {
        if ($ouDomain -and $d.Domain -and $ouDomain -ieq $d.Domain) { return $d.Id }
    }
    # Longest-suffix match, for child domains that are not listed separately.
    $best = $null
    foreach ($d in $DirectoryMap) {
        if ($ouDomain -and $d.Domain -and $ouDomain.EndsWith(".$($d.Domain)", [StringComparison]::OrdinalIgnoreCase)) {
            if (-not $best -or $d.Domain.Length -gt $best.Domain.Length) { $best = $d }
        }
    }
    if ($best) { return $best.Id }
    return 'Unknown'
}

function Test-CloudTranslatableFilter {
    <#
        Whether a dynamic group's filter could be rebuilt in Exchange Online.

        The blocking property is RecipientContainer -- an on-premises OU path.
        Exchange Online has no OUs, so any filter scoped to a container has no
        cloud equivalent and the group needs redesigning rather than migrating.
        LDAP filters referencing on-premises-only attributes are the same
        problem in a different shape.

        This flags for a human. It does not attempt translation: a filter that
        silently selects a different set of people is worse than one that
        obviously needs rewriting.
    #>
    param($Ddg)

    $reasons = @()
    if ("$($Ddg.RecipientContainer)".Trim()) {
        $reasons += 'Scoped to an on-premises OU (RecipientContainer)'
    }
    $filter = "$($Ddg.RecipientFilter)"
    if ($filter -match '(?i)\bDistinguishedName\b|\bOrganizationalUnit\b|\bCanonicalName\b') {
        $reasons += 'Filter references an on-premises directory path'
    }
    if ("$($Ddg.LdapRecipientFilter)".Trim() -and -not $filter.Trim()) {
        $reasons += 'LDAP-only filter, no OPATH equivalent recorded'
    }
    return [PSCustomObject]@{
        Translatable = ($reasons.Count -eq 0)
        Reasons      = (Join-Multi $reasons)
    }
}
#endregion

#region Confirmation ------------------------------------------------------------
$exportRoot = Join-Path $PSScriptRoot 'Exports'
$envFolder  = if ($Environment -eq 'HQ') { 'HQ' } else { 'Fran' }
$outDir     = Join-Path $exportRoot $envFolder

if (-not $NoPrompt) {
    $msg = @"
Export on-premises distribution groups for: $Environment

READ ONLY. This script changes nothing.

In scope   : Distribution Groups, Room Lists, Dynamic Distribution Groups
Out of scope: Mail-enabled security groups, security groups
              (still exported where they are MEMBERS of an in-scope group,
               and flagged as blockers)

Source     : on-premises Exchange. The console must be connected to it.
Written to : Exports\$envFolder

Every file from this run shares one timestamp, so attributes, members and
owners are a single consistent snapshot.

Continue?
"@
    $ans = [System.Windows.Forms.MessageBox]::Show($msg, 'DL Cloud Migration -- Export',
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Information)
    if ($ans -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host 'Cancelled.' -ForegroundColor Yellow
        return
    }
}

if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -Path $outDir -ItemType Directory -Force | Out-Null
}

# ONE timestamp for the whole run. Computed here, before any query, so that a
# long export cannot straddle midnight and produce files that look like they
# came from different days.
$stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
function New-ExportPath { param([string]$Kind) Join-Path $outDir "$Kind`_$Environment`_$stamp.csv" }
#endregion

#region Collect -----------------------------------------------------------------
Write-Host ''
Write-Host "=== DL Cloud Migration -- export ($Environment) ===" -ForegroundColor Cyan
Write-Host ''

try {
    # Groups can live in any domain in the forest, and the default view is the
    # local domain only -- without this an export run from one directory
    # silently reports the other as empty.
    Invoke-OnPremCommand -Name 'Set-ADServerSettings' -Parameters @{ ViewEntireForest = $true } | Out-Null
} catch {
    Write-Host "Could not set forest-wide scope: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Connect to on-premises Exchange first (Connect ExSE), then re-run.' -ForegroundColor Yellow
    return
}

$dirMap = @()
foreach ($d in @(Get-TenantDirectory)) {
    $dirMap += [PSCustomObject]@{ Id = "$($d.Id)"; Domain = "$($d.Domain)" }
}
if ($dirMap.Count -eq 0) {
    Write-Host 'Tenant.json declares no directories, so groups cannot be split by environment.' -ForegroundColor Red
    return
}

# The requested environment must actually be a directory this tenant declares.
#
# Without this the script runs perfectly, matches nothing, and writes seven
# empty CSVs -- which reads as "there are no groups" rather than "you asked for
# a directory that does not exist". The ValidateSet above only proves the value
# is one of two spellings, not that Tenant.json agrees.
if (@($dirMap | Where-Object { $_.Id -ieq $Environment }).Count -eq 0) {
    Write-Host ''
    Write-Host "'$Environment' is not a directory declared in Tenant.json." -ForegroundColor Red
    Write-Host "Declared: $(($dirMap | ForEach-Object { $_.Id }) -join ', ')" -ForegroundColor Yellow
    Write-Host 'Nothing was exported.' -ForegroundColor Yellow
    return
}

Write-Host 'Reading distribution groups from on-premises Exchange...' -ForegroundColor DarkGray
$allGroups = @()
try {
    $allGroups = @(Invoke-OnPremCaptured -Name 'Get-DistributionGroup' -Source 'Get-DistributionGroup' -Parameters @{
        ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
} catch {
    Write-Host "Get-DistributionGroup failed: $($_.Exception.Message)" -ForegroundColor Red
    return
}
Write-Host "  $($allGroups.Count) mail-enabled group(s) in the forest." -ForegroundColor DarkGray

$groups = @()
foreach ($g in $allGroups) {
    $env = Get-GroupEnvironment -OrganizationalUnit "$($g.OrganizationalUnit)" -DirectoryMap $dirMap
    if ($env -ne $Environment) { continue }
    if ($script:InScopeTypes -notcontains "$($g.RecipientTypeDetails)") { continue }
    $groups += [PSCustomObject]@{ Group = $g; Env = $env }
}
Write-Host "  $($groups.Count) in scope for $Environment." -ForegroundColor Green
Write-Host ''

$rowsGroup = @(); $rowsMember = @(); $rowsOwner = @(); $rowsPerm = @(); $rowsBlock = @()

$i = 0
foreach ($entry in $groups) {
    $g = $entry.Group
    $i++
    if ($i % 25 -eq 0 -or $i -eq 1) {
        Write-Host ("  [{0}/{1}] {2}" -f $i, $groups.Count, $g.PrimarySmtpAddress) -ForegroundColor DarkGray
    }

    $smtp = "$($g.PrimarySmtpAddress)"
    $blockers = @()

    # ---- members -------------------------------------------------------------
    $members = @()
    try {
        $members = @(Invoke-OnPremCaptured -Name 'Get-DistributionGroupMember' -Source "Members of $smtp" -Parameters @{
            Identity = "$($g.Guid)"; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
    } catch {
        $blockers += "Membership could not be read: $($_.Exception.Message)"
    }

    $noSmtpCount = 0; $outOfScopeCount = 0; $nestedCount = 0
    foreach ($m in $members) {
        $mType = "$($m.RecipientTypeDetails)"
        $mSmtp = "$($m.PrimarySmtpAddress)"
        $reason = ''

        if ($script:CloudMemberTypes -notcontains $mType) {
            $reason = "Type '$mType' cannot be a member of a cloud distribution group"
            $outOfScopeCount++
        } elseif (-not $mSmtp) {
            $reason = 'No primary SMTP address, so it cannot be re-added by address'
            $noSmtpCount++
        }
        if ($mType -eq 'MailUniversalDistributionGroup' -or $mType -eq 'RoomList') { $nestedCount++ }

        $rowsMember += [PSCustomObject]@{
            GroupPrimarySmtp  = $smtp
            GroupName         = "$($g.Name)"
            MemberName        = "$($m.Name)"
            MemberPrimarySmtp = $mSmtp
            MemberType        = $mType
            MemberAlias       = "$($m.Alias)"
            Reproducible      = [string]([bool](-not $reason))
            BlockerReason     = $reason
        }
    }
    if ($outOfScopeCount -gt 0) { $blockers += "$outOfScopeCount member(s) of a type a cloud group cannot hold" }
    if ($noSmtpCount    -gt 0) { $blockers += "$noSmtpCount member(s) with no SMTP address" }
    if ($members.Count  -eq 0) { $blockers += 'No members' }

    # ---- owners --------------------------------------------------------------
    $ownerCount = 0
    foreach ($o in @($g.ManagedBy)) {
        $r = Resolve-RecipientSmtp -Identity $o
        $ownerCount++
        $rowsOwner += [PSCustomObject]@{
            GroupPrimarySmtp = $smtp
            GroupName        = "$($g.Name)"
            OwnerName        = if ($r) { $r.Name } else { "$o" }
            OwnerPrimarySmtp = if ($r) { $r.Smtp } else { '' }
            OwnerType        = if ($r) { $r.TypeDetails } else { '' }
            Resolved         = [string]([bool]($r -and $r.Resolved -and $r.Smtp))
        }
        if (-not ($r -and $r.Smtp)) { $blockers += "Owner '$o' does not resolve to an address" }
    }
    if ($ownerCount -eq 0) { $blockers += 'No owner (ManagedBy is empty)' }

    # ---- permissions ---------------------------------------------------------
    # Send-As is an ACL on the AD object, NOT a property of the group, so it
    # does not appear anywhere in Get-DistributionGroup output. A migration that
    # reads only the group object silently drops every Send-As delegate.
    try {
        $acl = @(Invoke-OnPremCaptured -Name 'Get-ADPermission' -Source "Send-As on $smtp" -Parameters @{
            Identity = "$($g.DistinguishedName)"; ErrorAction = 'Stop' })
        foreach ($a in $acl) {
            if ("$($a.ExtendedRights)" -notmatch '(?i)send-as') { continue }
            if ("$($a.User)" -match '(?i)\\(NT AUTHORITY|S-1-5|Exchange|Organization Management|Domain Admins|Enterprise Admins)') { continue }
            if ($a.Deny) { continue }
            $r = Resolve-RecipientSmtp -Identity "$($a.User)"
            $rowsPerm += [PSCustomObject]@{
                GroupPrimarySmtp = $smtp
                GroupName        = "$($g.Name)"
                PermissionType   = 'SendAs'
                Trustee          = "$($a.User)"
                TrusteeSmtp      = if ($r) { $r.Smtp } else { '' }
                TrusteeType      = if ($r) { $r.TypeDetails } else { '' }
            }
        }
    } catch {
        $blockers += "Send-As permissions could not be read: $($_.Exception.Message)"
    }

    # SendOnBehalf is a different right from Send-As and is stored on the group.
    # Capturing one and not the other is a very quiet way to break delegates.
    foreach ($sob in @($g.GrantSendOnBehalfTo)) {
        $r = Resolve-RecipientSmtp -Identity $sob
        $rowsPerm += [PSCustomObject]@{
            GroupPrimarySmtp = $smtp
            GroupName        = "$($g.Name)"
            PermissionType   = 'SendOnBehalf'
            Trustee          = "$sob"
            TrusteeSmtp      = if ($r) { $r.Smtp } else { '' }
            TrusteeType      = if ($r) { $r.TypeDetails } else { '' }
        }
    }

    # ---- addressing ----------------------------------------------------------
    $addresses = @($g.EmailAddresses | ForEach-Object { "$_" })
    $x500 = @($addresses | Where-Object { $_ -match '^(?i)x500:' })
    if ($g.EmailAddressPolicyEnabled) {
        $blockers += 'Addresses are policy-generated (EmailAddressPolicyEnabled), so they must be set explicitly in the cloud'
    }

    $rowsGroup += [PSCustomObject]@{
        Environment                          = $entry.Env
        Name                                 = "$($g.Name)"
        DisplayName                          = "$($g.DisplayName)"
        SimpleDisplayName                    = "$($g.SimpleDisplayName)"
        Alias                                = "$($g.Alias)"
        PrimarySmtpAddress                   = $smtp
        RecipientTypeDetails                 = "$($g.RecipientTypeDetails)"
        GroupType                            = "$($g.GroupType)"
        Guid                                 = "$($g.Guid)"
        ExchangeGuid                         = "$($g.ExchangeGuid)"
        ExternalDirectoryObjectId            = "$($g.ExternalDirectoryObjectId)"
        DistinguishedName                    = "$($g.DistinguishedName)"
        OrganizationalUnit                   = "$($g.OrganizationalUnit)"
        IsDirSynced                          = "$($g.IsDirSynced)"

        # The single most important field in this file. The OLD LegacyExchangeDN
        # must be added as an X500 proxy on the NEW cloud group or every cached
        # Outlook autocomplete entry and every reply-all on an existing thread
        # bounces -- days later, looking like an unrelated fault.
        LegacyExchangeDN                     = "$($g.LegacyExchangeDN)"
        EmailAddresses                       = (Join-Multi $addresses)
        X500Addresses                        = (Join-Multi $x500)
        EmailAddressPolicyEnabled            = "$($g.EmailAddressPolicyEnabled)"
        HiddenFromAddressListsEnabled        = "$($g.HiddenFromAddressListsEnabled)"
        WindowsEmailAddress                  = "$($g.WindowsEmailAddress)"

        RequireSenderAuthenticationEnabled   = "$($g.RequireSenderAuthenticationEnabled)"
        AcceptMessagesOnlyFrom               = (Get-ResolvedSmtpList $g.AcceptMessagesOnlyFrom)
        AcceptMessagesOnlyFromDLMembers      = (Get-ResolvedSmtpList $g.AcceptMessagesOnlyFromDLMembers)
        RejectMessagesFrom                   = (Get-ResolvedSmtpList $g.RejectMessagesFrom)
        RejectMessagesFromDLMembers          = (Get-ResolvedSmtpList $g.RejectMessagesFromDLMembers)

        ModerationEnabled                    = "$($g.ModerationEnabled)"
        ModeratedBy                          = (Get-ResolvedSmtpList $g.ModeratedBy)
        BypassModerationFromSendersOrMembers = (Get-ResolvedSmtpList $g.BypassModerationFromSendersOrMembers)
        SendModerationNotifications          = "$($g.SendModerationNotifications)"
        BypassNestedModerationEnabled        = "$($g.BypassNestedModerationEnabled)"

        ManagedBy                            = (Get-ResolvedSmtpList $g.ManagedBy)
        GrantSendOnBehalfTo                  = (Get-ResolvedSmtpList $g.GrantSendOnBehalfTo)
        MemberJoinRestriction                = "$($g.MemberJoinRestriction)"
        MemberDepartRestriction              = "$($g.MemberDepartRestriction)"

        ReportToManagerEnabled               = "$($g.ReportToManagerEnabled)"
        ReportToOriginatorEnabled            = "$($g.ReportToOriginatorEnabled)"
        SendOofMessageToOriginatorEnabled    = "$($g.SendOofMessageToOriginatorEnabled)"

        MaxSendSize                          = "$($g.MaxSendSize)"
        MaxReceiveSize                       = "$($g.MaxReceiveSize)"
        RecipientLimits                      = "$($g.RecipientLimits)"

        MailTip                              = "$($g.MailTip)"
        Notes                                = ("$($g.Notes)" -replace '\r?\n', ' ')
        CustomAttribute1                     = "$($g.CustomAttribute1)"
        CustomAttribute2                     = "$($g.CustomAttribute2)"
        CustomAttribute3                     = "$($g.CustomAttribute3)"
        CustomAttribute4                     = "$($g.CustomAttribute4)"
        CustomAttribute5                     = "$($g.CustomAttribute5)"
        CustomAttribute6                     = "$($g.CustomAttribute6)"
        CustomAttribute7                     = "$($g.CustomAttribute7)"
        CustomAttribute8                     = "$($g.CustomAttribute8)"
        CustomAttribute9                     = "$($g.CustomAttribute9)"
        CustomAttribute10                    = "$($g.CustomAttribute10)"
        CustomAttribute11                    = "$($g.CustomAttribute11)"
        CustomAttribute12                    = "$($g.CustomAttribute12)"
        CustomAttribute13                    = "$($g.CustomAttribute13)"
        CustomAttribute14                    = "$($g.CustomAttribute14)"
        CustomAttribute15                    = "$($g.CustomAttribute15)"
        ExtensionCustomAttribute1            = (Join-Multi $g.ExtensionCustomAttribute1)
        ExtensionCustomAttribute2            = (Join-Multi $g.ExtensionCustomAttribute2)
        ExtensionCustomAttribute3            = (Join-Multi $g.ExtensionCustomAttribute3)
        ExtensionCustomAttribute4            = (Join-Multi $g.ExtensionCustomAttribute4)
        ExtensionCustomAttribute5            = (Join-Multi $g.ExtensionCustomAttribute5)

        WhenCreated                          = "$($g.WhenCreated)"
        WhenChanged                          = "$($g.WhenChanged)"
        MemberCount                          = $members.Count
        NestedGroupCount                     = $nestedCount
        OwnerCount                           = $ownerCount
    }

    $rowsBlock += [PSCustomObject]@{
        GroupPrimarySmtp    = $smtp
        GroupName           = "$($g.Name)"
        RecipientTypeDetails= "$($g.RecipientTypeDetails)"
        MemberCount         = $members.Count
        OwnerCount          = $ownerCount
        NestedGroupCount    = $nestedCount
        UnreproducibleMembers = $outOfScopeCount + $noSmtpCount
        BlockerCount        = $blockers.Count
        Readiness           = if ($blockers.Count -eq 0) { 'Ready' } else { 'NeedsDecision' }
        Blockers            = (Join-Multi $blockers)
    }
}
#endregion

#region Dynamic distribution groups ---------------------------------------------
Write-Host ''
Write-Host 'Reading dynamic distribution groups...' -ForegroundColor DarkGray
$rowsDdg = @()
try {
    $ddgs = @(Invoke-OnPremCaptured -Name 'Get-DynamicDistributionGroup' -Source 'Get-DynamicDistributionGroup' -Parameters @{
        ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
    foreach ($d in $ddgs) {
        $env = Get-GroupEnvironment -OrganizationalUnit "$($d.OrganizationalUnit)" -DirectoryMap $dirMap
        if ($env -ne $Environment) { continue }
        $feas = Test-CloudTranslatableFilter -Ddg $d
        $rowsDdg += [PSCustomObject]@{
            Environment               = $env
            Name                      = "$($d.Name)"
            DisplayName               = "$($d.DisplayName)"
            Alias                     = "$($d.Alias)"
            PrimarySmtpAddress        = "$($d.PrimarySmtpAddress)"
            LegacyExchangeDN          = "$($d.LegacyExchangeDN)"
            EmailAddresses            = (Join-Multi $d.EmailAddresses)
            RecipientFilter           = "$($d.RecipientFilter)"
            LdapRecipientFilter       = "$($d.LdapRecipientFilter)"
            RecipientContainer        = "$($d.RecipientContainer)"
            IncludedRecipients        = "$($d.IncludedRecipients)"
            ConditionalDepartment     = (Join-Multi $d.ConditionalDepartment)
            ConditionalCompany        = (Join-Multi $d.ConditionalCompany)
            ConditionalStateOrProvince= (Join-Multi $d.ConditionalStateOrProvince)
            ManagedBy                 = (Get-ResolvedSmtpList $d.ManagedBy)
            RequireSenderAuthenticationEnabled = "$($d.RequireSenderAuthenticationEnabled)"
            HiddenFromAddressListsEnabled      = "$($d.HiddenFromAddressListsEnabled)"
            CloudTranslatable         = [string]$feas.Translatable
            TranslationBlockers       = $feas.Reasons
        }
    }
    Write-Host "  $($rowsDdg.Count) dynamic group(s) in scope for $Environment." -ForegroundColor DarkGray
} catch {
    Write-Host "  Dynamic groups could not be read: $($_.Exception.Message)" -ForegroundColor Yellow
}
#endregion

#region Transport rule references -----------------------------------------------
# A rule scoped to a group keeps matching the OLD object. Delete and recreate the
# group and the rule stops applying, with no error anywhere -- which is how a
# disclaimer or a block rule quietly stops working weeks after a migration.
Write-Host 'Scanning transport rules for references to these groups...' -ForegroundColor DarkGray
$rowsRule = @()
try {
    $rules = @(Invoke-OnPremCommand -Name 'Get-TransportRule' -Parameters @{ ErrorAction = 'Stop' })
    $scopeSmtp = @{}
    foreach ($r in $rowsGroup) { $scopeSmtp["$($r.PrimarySmtpAddress)".ToLower()] = $r.Name }
    foreach ($rule in $rules) {
        $blob = ($rule | Out-String)
        foreach ($addr in $scopeSmtp.Keys) {
            if ($blob -match [regex]::Escape($addr)) {
                $rowsRule += [PSCustomObject]@{
                    GroupPrimarySmtp = $addr
                    GroupName        = $scopeSmtp[$addr]
                    RuleName         = "$($rule.Name)"
                    RuleState        = "$($rule.State)"
                    RuleMode         = "$($rule.Mode)"
                }
            }
        }
    }
    Write-Host "  $($rowsRule.Count) rule reference(s) found." -ForegroundColor DarkGray
} catch {
    Write-Host "  Transport rules could not be read: $($_.Exception.Message)" -ForegroundColor Yellow
}
#endregion

#region Write -------------------------------------------------------------------
function Save-Csv {
    param($Rows, [string]$Kind)
    $path = New-ExportPath -Kind $Kind
    if (@($Rows).Count -eq 0) {
        # An empty file is written deliberately. A MISSING file is ambiguous --
        # it could mean nothing matched or it could mean the step failed -- and
        # this project cannot afford that ambiguity.
        Set-Content -LiteralPath $path -Value '' -Encoding UTF8
    } else {
        $Rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    }
    Write-Host ("  {0,-14} {1,6} row(s)  {2}" -f $Kind, @($Rows).Count, (Split-Path -Leaf $path)) -ForegroundColor Gray
    return $path
}

Write-Host ''
Write-Host 'Writing CSVs...' -ForegroundColor Cyan
$null = Save-Csv -Rows $rowsGroup  -Kind 'Groups'
$null = Save-Csv -Rows $rowsMember -Kind 'Members'
$null = Save-Csv -Rows $rowsOwner  -Kind 'Owners'
$null = Save-Csv -Rows $rowsPerm   -Kind 'Permissions'
$null = Save-Csv -Rows $rowsDdg    -Kind 'DynamicGroups'
$null = Save-Csv -Rows $rowsRule   -Kind 'TransportRuleRefs'
$null = Save-Csv -Rows $rowsBlock  -Kind 'Blockers'

# Data quality gets an Environment column added at write time, because these
# warnings are raised during the FOREST-WIDE read that happens before the
# environment filter. So this file legitimately contains objects outside the
# environment you asked for. That is useful -- it is a free defect list for the
# other directory -- but only if the file says which is which rather than
# leaving you to guess from the path.
$rowsQuality = @()
foreach ($q in $script:DataQuality) {
    $rowsQuality += [PSCustomObject]@{
        Environment = if ($q.ObjectPath) { Get-GroupEnvironment -OrganizationalUnit $q.ObjectPath -DirectoryMap $dirMap } else { 'Unknown' }
        ObjectName  = $q.ObjectName
        Category    = $q.Category
        Property    = $q.Property
        ObjectPath  = $q.ObjectPath
        Source      = $q.Source
        Warning     = $q.Warning
    }
}
$null = Save-Csv -Rows $rowsQuality -Kind 'DataQuality'

$ready  = @($rowsBlock | Where-Object { $_.Readiness -eq 'Ready' }).Count
$needs  = @($rowsBlock | Where-Object { $_.Readiness -ne 'Ready' }).Count
$noOwn  = @($rowsBlock | Where-Object { [int]$_.OwnerCount -eq 0 }).Count
$nested = @($rowsBlock | Where-Object { [int]$_.NestedGroupCount -gt 0 }).Count
$empty  = @($rowsBlock | Where-Object { [int]$_.MemberCount -eq 0 }).Count

Write-Host ''
Write-Host "--- $Environment summary ---" -ForegroundColor Cyan
Write-Host ("  Groups in scope        : {0}" -f $rowsGroup.Count)
Write-Host ("  Dynamic groups         : {0}   ({1} not cloud-translatable)" -f $rowsDdg.Count, @($rowsDdg | Where-Object { $_.CloudTranslatable -eq 'False' }).Count)
Write-Host ("  Ready to rebuild       : {0}" -f $ready) -ForegroundColor Green
Write-Host ("  Need a decision first  : {0}" -f $needs) -ForegroundColor Yellow
Write-Host ("    no owner             : {0}" -f $noOwn)
Write-Host ("    contains nested group: {0}" -f $nested)
Write-Host ("    empty                : {0}" -f $empty)
Write-Host ("  Transport rule refs    : {0}" -f $rowsRule.Count)

if ($rowsQuality.Count -gt 0) {
    Write-Host ''
    Write-Host "  DATA QUALITY -- $($rowsQuality.Count) warning(s) from Exchange:" -ForegroundColor Yellow
    foreach ($grp in ($rowsQuality | Group-Object Category | Sort-Object Count -Descending)) {
        Write-Host ("    {0,-28} {1}" -f $grp.Name, $grp.Count) -ForegroundColor Yellow
    }
    $mine = @($rowsQuality | Where-Object { $_.Environment -eq $Environment })
    Write-Host ("    ...of which in {0}: {1}" -f $Environment, $mine.Count) -ForegroundColor Yellow
    Write-Host '    These are on-premises defects. Fix them BEFORE migrating those groups --' -ForegroundColor Yellow
    Write-Host '    an invalid address or an inconsistent object will not rebuild in the cloud.' -ForegroundColor Yellow
}
Write-Host ''
Write-Host "Saved to: $outDir" -ForegroundColor Green
Write-Host ''
Write-Host 'Nothing was changed. Read Blockers before planning the pilot.' -ForegroundColor DarkGray
Write-Host ''

try {
    Write-CompletedLog -Action 'DLCloudMigration-Export' -Target $Environment -Result 'Success' `
        -Detail "$($rowsGroup.Count) group(s), $($rowsDdg.Count) dynamic, $ready ready, $needs need a decision"
} catch { }
#endregion
