#Requires -Version 5.1
# Runs IN THE CONSOLE'S PROCESS (no OwnProcess), so it uses the console's
# own on-premises AND Exchange Online sessions. Safe because every Exchange
# call is QUALIFIED (Invoke-OnPremCommand / Invoke-EXOCommand) -- the same
# arrangement Disable-User and the Mailbox Editor use.
# TOOLKIT-CAPABILITY-OPTIONAL: ExO
# TOOLKIT-VARIANT: OnPremAD
<#
    Rename-OfficeGroups.ps1
    Finishes an office rename across the office's groups AND mailboxes,
    wherever each one is anchored.

    An office changes name. Somebody renames the AD group objects and stops
    there, because Name is the field you can see in the tool you renamed them
    in. DisplayName, Alias and the address list keep the old town, so the
    address book, the GAL and every reply-to still say Humboldt while the
    office has been Milan for a fortnight.

    2026-10-02: ONE TOOL FOR BOTH WORLDS. Distribution groups, mail-enabled
    security groups and shared (and room/equipment/user) mailboxes. Synced
    objects are written on-premises (Exchange Online refuses to write them);
    cloud-only objects are written in Exchange Online. See ARCHITECTURE 110.

    Reasoning: ScriptNotes/Rename-OfficeGroups.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-10-02
#>
# TOOLKIT-HELP-BEGIN
# Finish an office rename on its groups and mailboxes -- distribution groups,
# mail-enabled security groups and shared mailboxes -- wherever each one is
# anchored.
#
# Type the OLD CityST and the NEW one -- HumboldtTN and MilanTN -- and press
# Find.
#
# CITYST ONLY. Not 'AR.HumboldtTN', just 'HumboldtTN'. The
# prefixes are kept and rebuilt around the new name, so one pair of entries
# covers the whole office:
#
#     AR.HumboldtTN       ->  AR.MilanTN
#     HumboldtTN          ->  MilanTN
#     Jobs.HumboldtTN     ->  Jobs.MilanTN
#     Invoice.HumboldtTN  ->  Invoice.MilanTN
#
# WHERE EACH OBJECT LIVES -- the Anchor column:
#   On-prem   synced from Active Directory. Written with on-premises Exchange
#             (Set-DistributionGroup / Set-RemoteMailbox) and carried to the
#             cloud by the next sync. Exchange Online refuses these writes.
#   Cloud     cloud-only. Written in Exchange Online (Set-DistributionGroup /
#             Set-Mailbox) and live at once.
# Cloud-only objects need the console connected to Exchange Online; without
# it, Find lists on-prem objects only and says so.
#
# Every object carrying either name is listed, whatever state it is in: not
# renamed at all, half renamed, or already done. Tick the ones you want, read
# the preview, then Apply. USER mailboxes are listed but NOT pre-ticked -- an
# office name on a person's mailbox is unusual, so look before ticking.
#
# WHAT IT CHANGES, per object:
#   DisplayName        HumboldtTN            -> MilanTN
#   Alias              AR.HumboldtTN         -> AR.MilanTN
#   Email addresses    every smtp address containing the old CityST is rewritten,
#                      and the rewritten primary becomes the new primary
#   Name               only if you tick the box -- usually it is already done
#   UPN                mailboxes only, only if you tick the box. It is the
#                      sign-in name, so anything signing in as it must change.
#
# WHAT IT NEVER TOUCHES:
#   X500 addresses     these are what let old replies still resolve. Removing
#                      one breaks reply-to-old-mail, silently, months later.
#   onmicrosoft.com    the routing addresses hybrid mail delivery depends on.
#   Any address        without the old CityST in it.
#   SamAccountName     reported in Notes, never changed.
#   Membership, owners, permissions, or anything that is not a name.
#
# THE OLD ADDRESS. By default it is KEPT as a secondary, so mail sent to it
# keeps arriving. Tick 'Remove the old addresses' to drop it instead -- that is
# the irreversible one, and an NDR is the first anyone hears about it.
#
# EMAIL ADDRESS POLICY (on-prem objects). An object with the policy still
# enabled cannot have its addresses set by hand: Exchange refuses. Those are
# listed with a note and are skipped unless you tick the box to turn the
# policy off for them, which is a real change and is why it is not the default.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: On-premises Exchange in its own window, plus Exchange Online once signed in. Writes DisplayName, Alias and email addresses (Name and UPN optional) on the objects you tick, where each is anchored.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

#region Environment ------------------------------------------------------------
# ON-PREMISES EXCHANGE IS REQUIRED; EXCHANGE ONLINE IS OPTIONAL.
#
# Until 2026-10-02 this tool REFUSED to run with Exchange Online in the
# process, because it called Set-DistributionGroup unqualified and a cloud
# session would have captured the name. Every call now goes through
# Invoke-OnPremCommand or Invoke-EXOCommand, which pick the module explicitly
# (Get-ExchangeModuleMap), so both can be loaded at once -- and must be, for
# cloud-only objects. It runs in the console's process and borrows both of
# the console's connections.
$script:Blocked   = $false
$script:BlockText = ''

$emsState = Test-EmsAvailable
if (-not $emsState.Available) {
    $script:Blocked   = $true
    $script:BlockText = 'NOT connected to on-premises Exchange -- nothing can run. Use Connect ExSE.'
}

# Pinned by Find, and every write reuses it. Splatting an empty hashtable adds
# nothing, so the call sites stay identical whether or not a controller was
# resolved.
$script:DCParam    = @{}
$script:DirectoryId = ''
#endregion

#region Actions ----------------------------------------------------------------
function Rename-CityState {
    <#
        Replaces the old CityST with the new one, case-insensitively,
        leaving the rest of the string alone.

        REGEX WITH AN ESCAPED PATTERN, not .Replace(). .Replace() on a string is
        case-SENSITIVE, so 'humboldttn' in an alias somebody typed in lower case
        would sail straight through and the group would come out half renamed --
        which is the exact condition this tool exists to clean up.

        [regex]::Escape on the CityST because it is operator-typed. A CityST with
        a dot in it -- and these all have one, 'AR.HumboldtTN' -- is a regex
        wildcard otherwise.
    #>
    param([string]$Text, [string]$OldCityST, [string]$NewCityST)
    if (-not "$Text") { return "$Text" }
    return [regex]::Replace("$Text", [regex]::Escape($OldCityST), $NewCityST, 'IgnoreCase')
}

function Get-OfficeTypeLabel {
    # Short words for the grid. The raw RecipientTypeDetails is in the preview.
    param([string]$Rtd)
    switch -Regex ($Rtd) {
        'SharedMailbox$'                   { return 'Shared' }
        'RoomMailbox$'                     { return 'Room' }
        'EquipmentMailbox$'                { return 'Equipment' }
        'UserMailbox$'                     { return 'User' }
        '^MailUniversalDistributionGroup$' { return 'Distribution' }
        '^MailUniversalSecurityGroup$'     { return 'Security' }
        '^MailNonUniversalGroup$'          { return 'Group (non-univ)' }
        '^RoomList$'                       { return 'Room list' }
        default                            { return $Rtd }
    }
}

function ConvertTo-OfficeItem {
    <#
        One object, the same shape whichever cmdlet and whichever Exchange
        found it, so the planner and the preview never branch on where it
        came from. Only the WRITE branches, in Invoke-OfficeRename.

        Identity is what the write binds to: the distinguished name on-prem
        (ARCHITECTURE 14g -- a name is a fresh lookup, and two objects
        briefly sharing a name is this tool's whole job), the object GUID in
        Exchange Online.

        The email address policy only means anything on-premises. Exchange
        Online applies policies to Microsoft 365 Groups only.
    #>
    param(
        [Parameter(Mandatory)]$Obj,
        [Parameter(Mandatory)][string]$Kind,
        [Parameter(Mandatory)][string]$Anchor,
        [Parameter(Mandatory)][string]$Identity,
        [string]$Server = '',
        [string]$DirectoryId = ''
    )
    $upn = ''
    if ($Kind -eq 'Mailbox') {
        if ($Obj.PSObject.Properties['UserPrincipalName']) { $upn = "$($Obj.UserPrincipalName)" }
        if (-not $upn -and $Obj.PSObject.Properties['MicrosoftOnlineServicesID']) { $upn = "$($Obj.MicrosoftOnlineServicesID)" }
    }
    $sam = ''
    if ($Obj.PSObject.Properties['SamAccountName']) { $sam = "$($Obj.SamAccountName)" }
    $eap = $false
    if ($Anchor -eq 'On-prem' -and $Obj.PSObject.Properties['EmailAddressPolicyEnabled']) { $eap = [bool]$Obj.EmailAddressPolicyEnabled }
    $rtd = "$($Obj.RecipientTypeDetails)"

    return [PSCustomObject]@{
        Kind                      = $Kind
        Anchor                    = $Anchor
        Rtd                       = $rtd
        Type                      = (Get-OfficeTypeLabel -Rtd $rtd)
        Identity                  = $Identity
        Server                    = $Server
        DirectoryId               = $(if ($Anchor -eq 'Cloud') { 'Exchange Online' } else { $DirectoryId })
        Name                      = "$($Obj.Name)"
        DisplayName               = "$($Obj.DisplayName)"
        Alias                     = "$($Obj.Alias)"
        PrimarySmtpAddress        = "$($Obj.PrimarySmtpAddress)"
        EmailAddresses            = @($Obj.EmailAddresses)
        EmailAddressPolicyEnabled = $eap
        SamAccountName            = $sam
        UserPrincipalName         = $upn
        DistinguishedName         = "$($Obj.DistinguishedName)"
    }
}

function Get-OfficeObject {
    <#
        Every group and mailbox carrying either CityST, labelled with where it
        is anchored.

        BOTH CITYSTATES, NOT JUST THE OLD ONE. The whole premise is that a rename
        was left half done, so there is no single field that reliably still says
        Humboldt. In the ticket that prompted this, Name already said MilanTN
        while DisplayName, Alias and the address said HumboldtTN.

        ON-PREMISES FIRST: Get-DistributionGroup (distribution AND mail-enabled
        universal security groups) and Get-RemoteMailbox (every cloud mailbox
        whose account is in AD -- shared, user, room, equipment). Anything found
        here is SOURCE-ANCHORED on-premises and is written here. EVERY
        directory in Tenant.json is searched; see the comment in the body.

        THEN EXCHANGE ONLINE, if the console is connected: Get-Recipient
        with the same filter. A synced object already found on-prem is
        skipped (same primary address); a synced one NOT found on-prem is
        reported, not listed -- writing it in the cloud would fail, and the
        object to change is in a directory this search did not reach. What is
        left is cloud-only, and is read again with Get-Mailbox or
        Get-DistributionGroup for the full address list.

        Name, DisplayName and Alias only in the filter. EmailAddresses is
        multi-valued and its OPATH behaviour differs between Exchange versions;
        the alias and the address are generated from each other here.
    #>
    param(
        [Parameter(Mandatory)][string]$OldCityST,
        [Parameter(Mandatory)][string]$NewCityST
    )

    $result = [ordered]@{
        Items         = @()
        DirectoryId   = ''
        Server        = ''
        Errors        = @()
        CloudSearched = $false
    }

    $o = ConvertTo-OpathValue -Value $OldCityST
    $n = ConvertTo-OpathValue -Value $NewCityST
    $filter = "(Name -like '*$o*') -or (DisplayName -like '*$o*') -or (Alias -like '*$o*') -or " +
              "(Name -like '*$n*') -or (DisplayName -like '*$n*') -or (Alias -like '*$n*')"

    $items = New-Object System.Collections.ArrayList

    # ---- on-premises ------------------------------------------------------
    #
    # EVERY DIRECTORY, NOT THE FIRST WITH A HIT. Until 2026-10-02 the search
    # stopped at the first directory that returned anything -- so when an
    # office's groups were in one domain and its shared mailboxes in another,
    # the mailboxes were never looked for. Found the first day it was used.
    #
    # Each object is written through a controller IN ITS OWN DOMAIN, matched
    # from the DC= parts of its distinguished name to the directory's Domain.
    # A controller of another domain is at best a read-only global-catalog
    # copy of the object, and the write fails as if it were a permissions
    # problem. Duplicates (a forest-wide scope returns the same object from
    # both directories) are dropped by DN.
    $dirs = @()
    try { $dirs = @(Get-TenantDirectory) } catch { }
    if ($dirs.Count -eq 0) { $dirs = @([PSCustomObject]@{ Id = ''; Domain = '' }) }

    $dcByDomain = @{}
    $idByDomain = @{}
    $plans      = New-Object System.Collections.ArrayList
    foreach ($d in $dirs) {
        $id = "$($d.Id)"
        # Resolve-DirectoryDomainController, NOT Get-TenantDirectoryServer.
        # ARCHITECTURE 14f: the latter returns a DOMAIN when a directory names
        # several controllers, and Exchange rejects a domain with "Active
        # Directory error 0x52 -- suitability of server". Empty is a valid
        # answer and means "pass nothing".
        $dc = ''
        try { if ($id) { $dc = Resolve-DirectoryDomainController -DirectoryId $id } } catch { }
        $dom = "$($d.Domain)".Trim().ToLowerInvariant()
        if ($dom) { $dcByDomain[$dom] = $dc; $idByDomain[$dom] = $id }
        [void]$plans.Add([PSCustomObject]@{ Id = $id; Dc = $dc })
    }

    $seenDn  = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    $hitDirs = New-Object System.Collections.ArrayList

    $addOnPrem = {
        param($Obj, [string]$Kind, [string]$FoundDc, [string]$FoundId)
        $dn = "$($Obj.DistinguishedName)"
        if (-not $seenDn.Add($dn)) { return }
        $objDom = ((([regex]::Matches($dn, '(?i)(?<=(^|,)DC=)[^,]+')) | ForEach-Object { $_.Value }) -join '.').ToLowerInvariant()
        $dc  = $FoundDc
        $did = $FoundId
        if ($objDom -and $dcByDomain.ContainsKey($objDom)) { $dc = $dcByDomain[$objDom]; $did = $idByDomain[$objDom] }
        if ($did -and -not ($hitDirs -contains $did)) { [void]$hitDirs.Add($did) }
        [void]$items.Add((ConvertTo-OfficeItem -Obj $Obj -Kind $Kind -Anchor 'On-prem' -Identity $dn -Server $dc -DirectoryId $did))
    }

    foreach ($pl in $plans) {
        $label = if ($pl.Id) { $pl.Id } else { 'default directory' }
        $p = @{ Filter = $filter; ResultSize = 'Unlimited'; ErrorAction = 'Stop' }
        if ("$($pl.Dc)".Trim()) { $p['DomainController'] = $pl.Dc }

        try { foreach ($g in @(Invoke-OnPremCommand -Name 'Get-DistributionGroup' -Parameters $p)) { & $addOnPrem $g 'Group' $pl.Dc $pl.Id } }
        catch { $result.Errors += "$label (groups): $($_.Exception.Message)" }
        try { foreach ($m in @(Invoke-OnPremCommand -Name 'Get-RemoteMailbox' -Parameters $p)) { & $addOnPrem $m 'Mailbox' $pl.Dc $pl.Id } }
        catch { $result.Errors += "$label (mailboxes): $($_.Exception.Message)" }
    }
    $result.DirectoryId = (@($hitDirs.ToArray()) -join ', ')
    $searchedDirs = (@($plans | ForEach-Object { if ($_.Id) { $_.Id } else { 'default directory' } }) -join ', ')

    # ---- Exchange Online: cloud-only objects ------------------------------
    $exo = $null
    try { $exo = Test-ExoConnectionState } catch { }
    if ($exo -and $exo.Connected) {
        $result.CloudSearched = $true

        $known = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($it in $items) { if ($it.PrimarySmtpAddress) { [void]$known.Add($it.PrimarySmtpAddress) } }

        $rcp = @()
        try {
            $rcp = @(Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Filter = $filter; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
        } catch {
            $result.Errors += "Exchange Online: $($_.Exception.Message)"
        }

        $groupTypes = @('MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', 'MailNonUniversalGroup', 'RoomList')
        foreach ($r in $rcp) {
            $prim = "$($r.PrimarySmtpAddress)"
            if ($prim -and $known.Contains($prim)) { continue }

            $rtd = "$($r.RecipientTypeDetails)"
            if ([bool]$r.IsDirSynced) {
                $result.Errors += "$prim ($rtd) is synced from on-premises but was not found in $searchedDirs -- not listed. Change it on-premises."
                continue
            }

            $gid = "$($r.Guid)"
            try {
                if ($rtd -like '*Mailbox' -and $rtd -ne 'GroupMailbox') {
                    $full = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $gid; ErrorAction = 'Stop' }
                    [void]$items.Add((ConvertTo-OfficeItem -Obj $full -Kind 'Mailbox' -Anchor 'Cloud' -Identity $gid))
                } elseif ($groupTypes -contains $rtd) {
                    $full = Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{ Identity = $gid; ErrorAction = 'Stop' }
                    [void]$items.Add((ConvertTo-OfficeItem -Obj $full -Kind 'Group' -Anchor 'Cloud' -Identity $gid))
                } else {
                    $result.Errors += "$prim ($rtd) -- this type is not handled here; change it by hand."
                }
            } catch {
                $result.Errors += "$prim ($rtd): $($_.Exception.Message)"
            }
        }
    }

    $result.Items = @($items.ToArray())
    return [PSCustomObject]$result
}

function Get-PlannedChange {
    <#
        What one group would become. Computes only -- writes nothing.

        Everything the operator is asked to approve is decided here, so the
        preview and the write cannot disagree: Invoke-OfficeRename applies this
        object and does no thinking of its own.
    #>
    param(
        [Parameter(Mandatory)]$Group,
        [Parameter(Mandatory)][string]$OldCityST,
        [Parameter(Mandatory)][string]$NewCityST,
        [bool]$RemoveOld,
        [bool]$RenameName,
        [bool]$RenameUpn
    )

    $curName    = "$($Group.Name)"
    $curDisplay = "$($Group.DisplayName)"
    $curAlias   = "$($Group.Alias)"
    $curPrimary = "$($Group.PrimarySmtpAddress)"

    $newName    = Rename-CityState -Text $curName    -OldCityST $OldCityST -NewCityST $NewCityST
    $newDisplay = Rename-CityState -Text $curDisplay -OldCityST $OldCityST -NewCityST $NewCityST
    $newAlias   = Rename-CityState -Text $curAlias   -OldCityST $OldCityST -NewCityST $NewCityST
    $newPrimary = Rename-CityState -Text $curPrimary -OldCityST $OldCityST -NewCityST $NewCityST
    $curUpn     = "$($Group.UserPrincipalName)"
    $newUpn     = Rename-CityState -Text $curUpn     -OldCityST $OldCityST -NewCityST $NewCityST

    # ---- the address list -------------------------------------------------
    #
    # REBUILT WHOLE, because Set-DistributionGroup -EmailAddresses REPLACES the
    # collection rather than merging into it -- the same trap as ManagedBy and
    # AcceptMessagesOnlyFromSendersOrMembers. Anything not carried across here
    # is deleted, so the rule is to copy everything and change only what
    # matches.
    $kept    = New-Object System.Collections.ArrayList
    $dropped = New-Object System.Collections.ArrayList
    $added   = New-Object System.Collections.ArrayList
    $seen    = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)

    $addEntry = {
        param([string]$Prefix, [string]$Value, $Bucket)
        $key = "$Prefix`:$Value"
        if ($seen.Add($Value)) { [void]$Bucket.Add($key) }
    }

    foreach ($raw in @($Group.EmailAddresses)) {
        $s = "$raw"
        if (-not $s) { continue }

        $idx = $s.IndexOf(':')
        if ($idx -lt 1) { [void]$kept.Add($s); continue }

        $prefix = $s.Substring(0, $idx)
        $value  = $s.Substring($idx + 1)

        # X500 AND EVERYTHING THAT IS NOT SMTP PASS THROUGH UNTOUCHED.
        #
        # X500 is the one worth naming: it carries the object's previous
        # legacyExchangeDN, and it is what lets a reply to a two-year-old mail
        # still resolve. Dropping one breaks that silently, and the report
        # arrives months later as "replies to old mail bounce".
        if ($prefix -notmatch '^(?i)smtp$') { [void]$kept.Add($s); continue }

        # ONMICROSOFT ADDRESSES PASS THROUGH TOO. The tenant.mail.onmicrosoft.com
        # address on a remote mailbox is its routing address (targetAddress
        # points at it); renaming or dropping it here would leave the
        # RemoteRoutingAddress naming an address the object no longer has.
        if ($value -match '(?i)\.onmicrosoft\.com$') { & $addEntry 'smtp' $value $kept; continue }

        $renamed = Rename-CityState -Text $value -OldCityST $OldCityST -NewCityST $NewCityST

        if ($renamed -ieq $value) {
            # No CityST in it -- an unrelated address the office also has. Keep.
            & $addEntry 'smtp' $value $kept
            continue
        }

        & $addEntry 'smtp' $renamed $added
        if ($RemoveOld) { [void]$dropped.Add("smtp:$value") }
        else            { & $addEntry 'smtp' $value $kept }
    }

    # ONE PRIMARY, AND IT IS THE CASE THAT DECIDES IT. In an EmailAddresses
    # collection the single upper-case SMTP: entry is the primary; every other
    # is a secondary. Passing two upper-case entries is an error, passing none
    # leaves Exchange to choose.
    $final = New-Object System.Collections.ArrayList
    foreach ($e in @($added.ToArray()) + @($kept.ToArray())) {
        $i = $e.IndexOf(':')
        $v = $e.Substring($i + 1)
        $p = $e.Substring(0, $i)
        if ($p -match '^(?i)smtp$') {
            if ($v -ieq $newPrimary) { [void]$final.Add("SMTP:$v") } else { [void]$final.Add("smtp:$v") }
        } else {
            [void]$final.Add($e)
        }
    }

    # ---- notes: things worth knowing that this tool will not do ------------
    $notes = New-Object System.Collections.ArrayList

    if ($Group.EmailAddressPolicyEnabled) {
        [void]$notes.Add('Email address policy is ENABLED -- addresses cannot be set by hand')
    }
    if ("$($Group.SamAccountName)" -and (Rename-CityState -Text "$($Group.SamAccountName)" -OldCityST $OldCityST -NewCityST $NewCityST) -ine "$($Group.SamAccountName)") {
        [void]$notes.Add("SamAccountName still says $($Group.SamAccountName) -- not changed here")
    }
    if ($curUpn -and $newUpn -ine $curUpn -and -not $RenameUpn) {
        [void]$notes.Add("UPN still says $curUpn -- tick 'Also change the UPN' to change it")
    }
    if ($Group.Type -eq 'User') {
        [void]$notes.Add('User mailbox -- not pre-ticked')
    }
    if (-not ($final | Where-Object { $_ -cmatch '^SMTP:' })) {
        [void]$notes.Add('No primary could be determined -- skipped')
    }

    $changes = New-Object System.Collections.ArrayList
    if ($RenameName -and $newName -cne $curName)   { [void]$changes.Add('Name') }
    if ($newDisplay -cne $curDisplay)              { [void]$changes.Add('DisplayName') }
    if ($newAlias   -cne $curAlias)                { [void]$changes.Add('Alias') }
    if ($added.Count -gt 0 -or $dropped.Count -gt 0) { [void]$changes.Add('EmailAddresses') }
    if ($RenameUpn -and $curUpn -and $newUpn -ine $curUpn) { [void]$changes.Add('UPN') }

    return [PSCustomObject]@{
        Dn           = "$($Group.DistinguishedName)"
        Identity     = "$($Group.Identity)"
        Server       = "$($Group.Server)"
        DirId        = "$($Group.DirectoryId)"
        Kind         = "$($Group.Kind)"
        Anchor       = "$($Group.Anchor)"
        Type         = "$($Group.Type)"
        Rtd          = "$($Group.Rtd)"
        Upn          = $curUpn
        NewUpn       = $newUpn
        Name         = $curName
        NewName      = $newName
        Display      = $curDisplay
        NewDisplay   = $newDisplay
        Alias        = $curAlias
        NewAlias     = $newAlias
        Primary      = $curPrimary
        NewPrimary   = $newPrimary
        Addresses    = @($final.ToArray())
        Added        = @($added.ToArray())
        Dropped      = @($dropped.ToArray())
        PolicyOn     = [bool]$Group.EmailAddressPolicyEnabled
        Notes        = @($notes.ToArray())
        Changes      = @($changes.ToArray())
        HasWork      = ($changes.Count -gt 0)
    }
}

function Format-PlannedChange {
    param([Parameter(Mandatory)]$Plan, [bool]$RenameName, [bool]$RenameUpn)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("=== $($Plan.Name)   [$($Plan.Rtd), $($Plan.Anchor)] ===")

    if (-not $Plan.HasWork) {
        [void]$sb.AppendLine('  Nothing to change -- already carries the new CityST.')
        foreach ($n in $Plan.Notes) { [void]$sb.AppendLine("  NOTE: $n") }
        return $sb.ToString()
    }

    if ($RenameName -and $Plan.NewName -cne $Plan.Name) {
        [void]$sb.AppendLine("  Name         : $($Plan.Name)  ->  $($Plan.NewName)")
    }
    if ($Plan.NewDisplay -cne $Plan.Display) {
        [void]$sb.AppendLine("  DisplayName  : $($Plan.Display)  ->  $($Plan.NewDisplay)")
    }
    if ($Plan.NewAlias -cne $Plan.Alias) {
        [void]$sb.AppendLine("  Alias        : $($Plan.Alias)  ->  $($Plan.NewAlias)")
    }
    if ($Plan.NewPrimary -cne $Plan.Primary) {
        [void]$sb.AppendLine("  Primary SMTP : $($Plan.Primary)  ->  $($Plan.NewPrimary)")
    }
    if ($RenameUpn -and $Plan.Upn -and $Plan.NewUpn -ine $Plan.Upn) {
        [void]$sb.AppendLine("  UPN          : $($Plan.Upn)  ->  $($Plan.NewUpn)")
    }
    foreach ($a in $Plan.Added)   { [void]$sb.AppendLine("  ADD          : $a") }
    foreach ($d in $Plan.Dropped) { [void]$sb.AppendLine("  REMOVE       : $d") }

    # The full resulting list is printed, not just the delta. EmailAddresses is
    # replaced wholesale, so what matters is what the collection will BE -- and
    # an X500 that quietly failed to make the list is only visible here.
    [void]$sb.AppendLine('  Resulting address list:')
    foreach ($a in $Plan.Addresses) { [void]$sb.AppendLine("      $a") }

    foreach ($n in $Plan.Notes) { [void]$sb.AppendLine("  NOTE: $n") }
    return $sb.ToString()
}

function Invoke-OfficeRename {
    <#
        Applies one planned change, WHERE THE OBJECT IS ANCHORED. Decides
        nothing else.

          On-prem group    Set-DistributionGroup  (on-premises)
          On-prem mailbox  Set-RemoteMailbox      (on-premises)
          Cloud group      Set-DistributionGroup  (Exchange Online)
          Cloud mailbox    Set-Mailbox            (Exchange Online)

        A synced object written in Exchange Online fails with "out of the
        current user's write scope ... being synchronized from your
        on-premises organization"; a cloud-only one has nothing on-prem to
        write. Hence one route per anchor, qualified, never a bare name.

        ONE SET CALL PER OBJECT. DisplayName, Alias and the address list go
        together, so an object cannot end up with its alias changed and its
        addresses not -- the half-finished state this tool was written to
        repair.

        -EmailAddresses AND NOT -PrimarySmtpAddress. Microsoft Learn: "You
        can't use the EmailAddresses parameter and the PrimarySmtpAddress
        parameter in the same command." The primary is set by CASE instead --
        the single upper-case SMTP: entry. The caller READS THE OBJECT BACK and
        prints WindowsEmailAddress, because Learn only promises that
        PrimarySmtpAddress keeps it in step.

        UPN: -UserPrincipalName on Set-RemoteMailbox (on-prem), and
        -MicrosoftOnlineServicesID on Set-Mailbox (Exchange Online -- the
        cloud cmdlet's name for the same thing).

        -BypassSecurityGroupManagerCheck on groups because Exchange checks
        ManagedBy on a mail-enabled SECURITY group regardless of RBAC.
        -Confirm:$false on every write: a ShouldProcess prompt has no console
        here and would hang the window.
    #>
    param(
        [Parameter(Mandatory)]$Plan,
        [bool]$RenameName,
        [bool]$DisablePolicy,
        [bool]$RenameUpn
    )

    $p = @{
        Identity       = $Plan.Identity
        DisplayName    = $Plan.NewDisplay
        Alias          = $Plan.NewAlias
        EmailAddresses = $Plan.Addresses
        Confirm        = $false
        ErrorAction    = 'Stop'
    }
    if ($RenameName) { $p['Name'] = $Plan.NewName }
    if ($Plan.Kind -eq 'Group') { $p['BypassSecurityGroupManagerCheck'] = $true }

    $upnChange = ($RenameUpn -and $Plan.Kind -eq 'Mailbox' -and $Plan.Upn -and $Plan.NewUpn -ine $Plan.Upn)

    if ($Plan.Anchor -eq 'On-prem') {
        # THE POLICY GOES OFF IN THE SAME CALL, not a separate one before it.
        # Two calls means a window where the policy is off and the addresses
        # are still the old ones, and if the second call fails the object is
        # left in a state nobody asked for.
        if ($DisablePolicy -and $Plan.PolicyOn) { $p['EmailAddressPolicyEnabled'] = $false }
        if ($upnChange) { $p['UserPrincipalName'] = $Plan.NewUpn }
        if ("$($Plan.Server)".Trim()) { $p['DomainController'] = $Plan.Server }

        $setName = if ($Plan.Kind -eq 'Group') { 'Set-DistributionGroup' } else { 'Set-RemoteMailbox' }
        $getName = if ($Plan.Kind -eq 'Group') { 'Get-DistributionGroup' } else { 'Get-RemoteMailbox' }

        Invoke-OnPremCommand -Name $setName -Parameters $p | Out-Null

        $rp = @{ Identity = $Plan.Identity; ErrorAction = 'Stop' }
        if ("$($Plan.Server)".Trim()) { $rp['DomainController'] = $Plan.Server }
        return (Invoke-OnPremCommand -Name $getName -Parameters $rp)
    }

    if ($upnChange) { $p['MicrosoftOnlineServicesID'] = $Plan.NewUpn }
    $setName = if ($Plan.Kind -eq 'Group') { 'Set-DistributionGroup' } else { 'Set-Mailbox' }
    $getName = if ($Plan.Kind -eq 'Group') { 'Get-DistributionGroup' } else { 'Get-Mailbox' }

    Invoke-EXOCommand -Name $setName -Parameters $p | Out-Null
    return (Invoke-EXOCommand -Name $getName -Parameters @{ Identity = $Plan.Identity; ErrorAction = 'Stop' })
}
#endregion

#region WinForms UI ------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Office Rename -- groups and mailboxes'
$form.ClientSize    = New-Object System.Drawing.Size(944, 652)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Tag       = 'NoTheme'
$lblWarn.Location  = New-Object System.Drawing.Point(12, 12)
$lblWarn.Size      = New-Object System.Drawing.Size(880, 34)
$lblWarn.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
if ($script:Blocked) {
    $lblWarn.Text      = $script:BlockText
    $lblWarn.ForeColor = [System.Drawing.Color]::Red
} else {
    $lblWarn.Text      = 'Writes go where each object is anchored: synced objects ON-PREMISES (they reach the cloud on the next sync), cloud-only objects in EXCHANGE ONLINE.'
    $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
}
$form.Controls.Add($lblWarn)

$lblOld           = New-Object System.Windows.Forms.Label
$lblOld.Text      = 'Old CityST:'
$lblOld.Location  = New-Object System.Drawing.Point(12, 56)
$lblOld.Size      = New-Object System.Drawing.Size(80, 18)
$form.Controls.Add($lblOld)

$txtOld           = New-Object System.Windows.Forms.TextBox
$txtOld.Location  = New-Object System.Drawing.Point(96, 53)
$txtOld.Size      = New-Object System.Drawing.Size(180, 24)
$txtOld.Font      = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOld)

$lblNew           = New-Object System.Windows.Forms.Label
$lblNew.Text      = 'New CityST:'
$lblNew.Location  = New-Object System.Drawing.Point(296, 56)
$lblNew.Size      = New-Object System.Drawing.Size(80, 18)
$form.Controls.Add($lblNew)

$txtNew           = New-Object System.Windows.Forms.TextBox
$txtNew.Location  = New-Object System.Drawing.Point(380, 53)
$txtNew.Size      = New-Object System.Drawing.Size(180, 24)
$txtNew.Font      = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtNew)

$btnFind          = New-Object System.Windows.Forms.Button
$btnFind.Text     = 'Find'
$btnFind.Location = New-Object System.Drawing.Point(576, 52)
$btnFind.Size     = New-Object System.Drawing.Size(100, 26)
$form.Controls.Add($btnFind)

$lblDir           = New-Object System.Windows.Forms.Label
$lblDir.Location  = New-Object System.Drawing.Point(688, 56)
$lblDir.Size      = New-Object System.Drawing.Size(244, 18)
$lblDir.AutoEllipsis = $true
$lblDir.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblDir)

# THE EXAMPLE IS ON THE FORM, not in the help behind the '?'.
#
# The first version labelled these boxes 'CityST', which is a word I made
# up while writing the search and which means nothing to the person holding the
# ticket. The ticket says HumboldtTN and MilanTN; the form should ask for
# exactly that, and show what happens to the prefixes so nobody has to guess
# whether to type 'AR.HumboldtTN' or just 'HumboldtTN'.
$lblHint           = New-Object System.Windows.Forms.Label
$lblHint.Text      = 'CityST only -- HumboldtTN, not AR.HumboldtTN. Prefixes are kept and rebuilt: AR.HumboldtTN becomes AR.MilanTN, Jobs.HumboldtTN becomes Jobs.MilanTN.'
$lblHint.Location  = New-Object System.Drawing.Point(12, 84)
$lblHint.Size      = New-Object System.Drawing.Size(920, 18)
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblHint)

# The grid gives up the 22px rather than the window growing -- it still shows
# about ten rows, and an office has four or five groups.
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 108)
$grid.Size     = New-Object System.Drawing.Size(920, 278)
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

$colSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSel.Name = 'Sel'; $colSel.HeaderText = ''; $colSel.Width = 30
[void]$grid.Columns.Add($colSel)
foreach ($c in @(
    @{ N = 'Anchor';  H = 'Anchor';          W = 64 },
    @{ N = 'Type';    H = 'Type';            W = 96 },
    @{ N = 'Name';    H = 'Name';            W = 110 },
    @{ N = 'Display'; H = 'DisplayName now'; W = 130 },
    @{ N = 'Alias';   H = 'Alias now';       W = 130 },
    @{ N = 'Primary'; H = 'Primary SMTP now'; W = 200 },
    @{ N = 'Notes';   H = 'Notes';           W = 140 })) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.N; $col.HeaderText = $c.H; $col.Width = $c.W; $col.ReadOnly = $true
    [void]$grid.Columns.Add($col)
}
$form.Controls.Add($grid)

$chkRemoveOld          = New-Object System.Windows.Forms.CheckBox
$chkRemoveOld.Text     = 'Remove the old addresses'
$chkRemoveOld.Location = New-Object System.Drawing.Point(12, 396)
$chkRemoveOld.Size     = New-Object System.Drawing.Size(200, 20)
$chkRemoveOld.Checked  = $false
$form.Controls.Add($chkRemoveOld)

$chkRenameName          = New-Object System.Windows.Forms.CheckBox
$chkRenameName.Text     = 'Also rename Name'
$chkRenameName.Location = New-Object System.Drawing.Point(220, 396)
$chkRenameName.Size     = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($chkRenameName)

$chkPolicy          = New-Object System.Windows.Forms.CheckBox
$chkPolicy.Text     = 'Turn off the address policy where it blocks'
$chkPolicy.Location = New-Object System.Drawing.Point(378, 396)
$chkPolicy.Size     = New-Object System.Drawing.Size(282, 20)
$form.Controls.Add($chkPolicy)

$chkUpn          = New-Object System.Windows.Forms.CheckBox
$chkUpn.Text     = 'Also change the UPN'
$chkUpn.Location = New-Object System.Drawing.Point(666, 396)
$chkUpn.Size     = New-Object System.Drawing.Size(166, 20)
$form.Controls.Add($chkUpn)

$btnAll           = New-Object System.Windows.Forms.Button
$btnAll.Text      = 'All'
$btnAll.Location  = New-Object System.Drawing.Point(838, 393)
$btnAll.Size      = New-Object System.Drawing.Size(44, 24)
$form.Controls.Add($btnAll)

$btnNone          = New-Object System.Windows.Forms.Button
$btnNone.Text     = 'None'
$btnNone.Location = New-Object System.Drawing.Point(886, 393)
$btnNone.Size     = New-Object System.Drawing.Size(46, 24)
$form.Controls.Add($btnNone)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(12, 422)
$lblStatus.Size     = New-Object System.Drawing.Size(920, 20)
$form.Controls.Add($lblStatus)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 446)
$txtOut.Size       = New-Object System.Drawing.Size(920, 152)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$btnPreview          = New-Object System.Windows.Forms.Button
$btnPreview.Text     = 'Preview'
$btnPreview.Location = New-Object System.Drawing.Point(12, 610)
$btnPreview.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnPreview)

$btnApply          = New-Object System.Windows.Forms.Button
$btnApply.Text     = 'Apply changes'
$btnApply.Location = New-Object System.Drawing.Point(140, 610)
$btnApply.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnApply)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(822, 610)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnClose)

# Found groups, keyed by grid row. The grid holds display text; this holds the
# objects, because re-reading Exchange to answer "what did row 3 mean" is both
# slow and a chance for the answer to have changed underneath.
$script:Found = @()

function Get-CityStatePair {
    return [PSCustomObject]@{
        Old = $txtOld.Text.Trim()
        New = $txtNew.Text.Trim()
    }
}

function Test-CityStateReady {
    $t = Get-CityStatePair
    if (-not $t.Old -or -not $t.New) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'Enter both the old and the new CityST.'
        return $false
    }
    if ($t.Old -ieq $t.New) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'The two entries are the same -- there would be nothing to change.'
        return $false
    }
    return $true
}

function Get-SelectedPlan {
    <#
        The planned change for every ticked row, recomputed from the current
        checkbox settings.

        RECOMPUTED RATHER THAN CACHED, on purpose. 'Remove the old addresses'
        changes what the plan IS, so a plan built when the box was ticked and
        applied after it was unticked would do the opposite of what the screen
        says.
    #>
    $t = Get-CityStatePair
    $plans = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        if ("$($grid.Rows[$i].Cells['Sel'].Value)" -ne 'True') { continue }
        if ($i -ge $script:Found.Count) { continue }
        [void]$plans.Add((Get-PlannedChange -Group $script:Found[$i] -OldCityST $t.Old -NewCityST $t.New `
                            -RemoveOld $chkRemoveOld.Checked -RenameName $chkRenameName.Checked -RenameUpn $chkUpn.Checked))
    }
    return @($plans.ToArray())
}

$btnFind.Add_Click({
    if ($script:Blocked) { $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick; $lblStatus.Text = 'Blocked -- see the message at the top.'; return }
    if (-not (Test-CityStateReady)) { return }

    $t = Get-CityStatePair
    $btnFind.Enabled = $false
    $form.Cursor     = [System.Windows.Forms.Cursors]::WaitCursor
    $grid.Rows.Clear()
    $script:Found = @()
    $txtOut.Text  = ''
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.Text = 'Searching...'
    $form.Refresh()

    try {
        $r = Get-OfficeObject -OldCityST $t.Old -NewCityST $t.New

        $script:DirectoryId = $r.DirectoryId
        $script:DCParam     = @{}
        if ("$($r.Server)".Trim()) { $script:DCParam['DomainController'] = $r.Server }

        $cloudTag = if ($r.CloudSearched) { ' + Exchange Online' } else { '' }
        $lblDir.Text = if ($r.DirectoryId) { "Found in: $($r.DirectoryId)$cloudTag" } else { '' }

        $script:Found = @($r.Items)
        $work = 0
        foreach ($g in $script:Found) {
            $plan = Get-PlannedChange -Group $g -OldCityST $t.Old -NewCityST $t.New `
                        -RemoveOld $chkRemoveOld.Checked -RenameName $chkRenameName.Checked -RenameUpn $chkUpn.Checked
            $note = if ($plan.Notes.Count) { ($plan.Notes -join '; ') }
                    elseif (-not $plan.HasWork) { 'already renamed' }
                    else { '' }
            if ($plan.HasWork) { $work++ }
            # PRE-TICKED ONLY WHERE THERE IS WORK, and never a person's own
            # mailbox -- that one the operator should look at first.
            $tick = ($plan.HasWork -and $plan.Type -ne 'User')
            [void]$grid.Rows.Add($tick, $plan.Anchor, $plan.Type, $plan.Name, $plan.Display, $plan.Alias, $plan.Primary, $note)
        }

        if ($r.Errors.Count -gt 0) { $txtOut.Text = ($r.Errors -join [Environment]::NewLine) }

        $cloudNote = if ($r.CloudSearched) { '' } else {  ' Exchange Online is not connected in the console, so cloud-only objects were NOT searched.' }
        if ($script:Found.Count -eq 0) {
            $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
            $lblStatus.Text = "Nothing carries '$($t.Old)' or '$($t.New)'. Check the spelling.$cloudNote"
        } else {
            $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
            $lblStatus.Text = "$($script:Found.Count) object(s) found, $work with something to change. Press Preview.$cloudNote"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Find'
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'Search failed -- see below.'
    } finally {
        $btnFind.Enabled = $true
        $form.Cursor     = [System.Windows.Forms.Cursors]::Default
    }
})

$btnAll.Add_Click({  for ($i = 0; $i -lt $grid.Rows.Count; $i++) { $grid.Rows[$i].Cells['Sel'].Value = $true  } })
$btnNone.Add_Click({ for ($i = 0; $i -lt $grid.Rows.Count; $i++) { $grid.Rows[$i].Cells['Sel'].Value = $false } })

$btnPreview.Add_Click({
    if (-not (Test-CityStateReady)) { return }
    $plans = Get-SelectedPlan
    if ($plans.Count -eq 0) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'Nothing ticked.'
        return
    }
    $sb = New-Object System.Text.StringBuilder
    foreach ($p in $plans) { [void]$sb.AppendLine((Format-PlannedChange -Plan $p -RenameName $chkRenameName.Checked -RenameUpn $chkUpn.Checked)) }
    $txtOut.Text = $sb.ToString().TrimEnd()
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.Text = "$($plans.Count) object(s) previewed. Nothing has been written."
})

$btnApply.Add_Click({
    if ($script:Blocked) { $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick; $lblStatus.Text = 'Blocked -- see the message at the top.'; return }
    if (-not (Test-CityStateReady)) { return }

    $plans = @(Get-SelectedPlan | Where-Object { $_.HasWork })
    if ($plans.Count -eq 0) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'Nothing ticked that has anything to change.'
        return
    }

    # BLOCKED BY POLICY IS REPORTED BEFORE THE PROMPT, not discovered during the
    # run. A group whose address policy is on cannot have addresses set, and
    # finding that out one group into a five-group change leaves a half-renamed
    # office -- which is what this tool exists to fix.
    $blocked = @($plans | Where-Object { $_.PolicyOn -and -not $chkPolicy.Checked })
    if ($blocked.Count -gt 0) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = "$($blocked.Count) object(s) have the email address policy enabled. Tick the policy box, or untick those rows."
        $txtOut.Text = "These cannot be changed while their email address policy is enabled:" + [Environment]::NewLine +
                       (($blocked | ForEach-Object { "  $($_.Name)" }) -join [Environment]::NewLine)
        return
    }

    $lines = @($plans | ForEach-Object { "  [$($_.Anchor)] $($_.Name)  ->  $($_.NewPrimary)" })
    $extra = @()
    if ($chkRemoveOld.Checked)  { $extra += 'The old addresses will be REMOVED -- mail to them will bounce.' }
    if ($chkRenameName.Checked) { $extra += 'The Name attribute will be changed as well.' }
    if ($chkUpn.Checked)        { $extra += 'Mailbox UPNs (sign-in names) will be changed as well.' }
    if ($chkPolicy.Checked)     { $extra += 'The email address policy will be turned OFF where it blocks the change.' }

    $confirm = Show-ToolDialog(
        ("Rename $($plans.Count) object(s)?" + [Environment]::NewLine + [Environment]::NewLine +
         ($lines -join [Environment]::NewLine) +
         $(if ($extra.Count) { [Environment]::NewLine + [Environment]::NewLine + ($extra -join [Environment]::NewLine) } else { '' })),
        'Confirm office rename',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
        $lblStatus.Text = 'Cancelled. Nothing was written.'
        return
    }

    $btnApply.Enabled = $false
    $form.Cursor      = [System.Windows.Forms.Cursors]::WaitCursor
    $sb = New-Object System.Text.StringBuilder
    $ok = 0; $bad = 0
    $t  = Get-CityStatePair

    try {
        foreach ($p in $plans) {
            # EACH GROUP IS INDEPENDENT. One that will not write is reported and
            # the rest continue -- abandoning the run on the second of five
            # would leave the office in exactly the mixed state being repaired.
            try {
                $after = Invoke-OfficeRename -Plan $p -RenameName $chkRenameName.Checked -DisablePolicy $chkPolicy.Checked -RenameUpn $chkUpn.Checked

                # WHAT IS ON THE OBJECT NOW, not what was asked for.
                [void]$sb.AppendLine("OK      $($p.Name)   [$($p.Type), $($p.Anchor)]")
                [void]$sb.AppendLine("          DisplayName        : $($after.DisplayName)")
                [void]$sb.AppendLine("          Alias              : $($after.Alias)")
                [void]$sb.AppendLine("          PrimarySmtpAddress : $($after.PrimarySmtpAddress)")
                [void]$sb.AppendLine("          WindowsEmailAddress: $($after.WindowsEmailAddress)")
                if ($p.Kind -eq 'Mailbox') {
                    [void]$sb.AppendLine("          UserPrincipalName  : $($after.UserPrincipalName)")
                }

                # THE ONE MISMATCH WORTH SHOUTING ABOUT. If 'mail' did not
                # follow the primary, every downstream consumer that reads the
                # AD attribute rather than Exchange -- which is most of them,
                # including this toolkit's own directory lookups -- still sees
                # the old town.
                if ("$($after.WindowsEmailAddress)" -and "$($after.WindowsEmailAddress)" -ine "$($after.PrimarySmtpAddress)") {
                    [void]$sb.AppendLine("          ^^ WindowsEmailAddress did NOT follow the primary. Fix it before closing the ticket.")
                }
                $envName = $p.DirId
                Write-CompletedLog -Action 'RenameOfficeGroup' -Target $p.NewPrimary -Result 'Success' `
                    -Environment $envName `
                    -Detail "$($t.Old) -> $($t.New); $($p.Rtd) ($($p.Anchor)); changed $($p.Changes -join ', '); was $($p.Primary)"
                $ok++
            } catch {
                [void]$sb.AppendLine("FAILED  $($p.Name)   [$($p.Type), $($p.Anchor)]  -- $($_.Exception.Message)")
                # $p.Primary, NOT $p.Name: the rename FAILED, so the group still
                # answers to its old address. NewPrimary would name something
                # that does not exist.
                Write-FailedLog -Action 'RenameOfficeGroup' `
                    -Target (Get-LogTarget $p.Primary $p.Name) -ErrorRecord $_ `
                    -Environment $p.DirId `
                    -Detail "$($t.Old) -> $($t.New); $($p.Rtd) ($($p.Anchor))"
                $bad++
            }
        }

        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("$ok succeeded, $bad failed.")
        [void]$sb.AppendLine('Press Find again to re-read and confirm. On-prem objects show the change in Exchange Online after the next sync.')
        $txtOut.Text = $sb.ToString()

        $lblStatus.ForeColor = if ($bad -eq 0) { [System.Drawing.Color]::ForestGreen } else { [System.Drawing.Color]::Firebrick }
        $lblStatus.Text = "$ok changed, $bad failed. On-prem objects reach the cloud on the next sync; cloud-only changes are live now."
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Office rename'
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'Failed -- see below.'
    } finally {
        $btnApply.Enabled = $true
        $form.Cursor      = [System.Windows.Forms.Cursors]::Default
    }
})

$btnClose.Add_Click({ $form.Close() })

if ($script:Blocked) {
    $btnFind.Enabled  = $false
    $btnApply.Enabled = $false
}

$form.AcceptButton = $btnFind

Set-ConsoleTheme -Form $form
Set-FormResizable -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Office Rename' | Out-Null

[void]$form.ShowDialog()
#endregion
