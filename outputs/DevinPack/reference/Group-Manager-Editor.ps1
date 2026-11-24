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
