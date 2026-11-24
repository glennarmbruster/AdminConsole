#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- the -OnPrem side uses Add-/Remove-ADPermission
#                                 and on-premises Get-/Set-DistributionGroup, which
#                                 share cmdlet names with Exchange Online. DL-Manager
#                                 overrides this for the CLOUD side, which must stay
#                                 in the console process to keep its EXO session.
# TOOLKIT-VARIANT: Dual
# ---------------------------------------------------------------------------
# ONE TOOL, BOTH ENVIRONMENTS. Cloud by default; -OnPrem switches it.
#
# This replaces the DL_SendToAs / AD-DL_SendToAs pair. They were 688 and 649
# lines that differed in about six cmdlet calls, and they had already drifted:
# the on-premises copy referenced a variable ($dirId) that was never assigned
# anywhere in the file, so its HQ branch returned a null environment and a null
# directory server. Two copies of one behaviour is how that survives unnoticed.
#
# Everything environment-specific is isolated below in Invoke-DLCmd,
# Get-SendAsState, Add-SendAsGrant, Remove-SendAsGrant and Resolve-DLEnvironment.
# Nothing else in this file knows or cares which side it is running against.
# ---------------------------------------------------------------------------
<#
    DL-Manager-SendToAs.ps1
    Lookup / Add / Remove Send-To (allowed senders) and/or Send-As
    permission on a distribution list, in either environment.

    Opened by DL-Manager.ps1, which decides the side from IsDirSynced and
    passes -OnPrem when Active Directory is the authority. Not in Tools.json:
    it is a helper, not a front door.

    Author   : Glenn Armbruster
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# Controls who may send TO a distribution list, and who may send AS it.
#
# Two separate permissions that get confused constantly:
#
#   * Send To -- the list is restricted, and only approved senders get through.
#     Everyone else gets a bounce.
#   * Send As -- mail appears to come FROM the list rather than from the person.
#
# "Nobody outside can email this list" is Send To. "I need my replies to look
# like they came from the team address" is Send As.
#
# On-premises lists only, and it opens in its own window so it cannot be
# confused with the cloud.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,

    # WRITE ON-PREMISES INSTEAD OF TO EXCHANGE ONLINE.
    #
    # Set by DL-Manager for a directory-synced group. The caller decides,
    # because the caller is the one that read IsDirSynced -- this tool must not
    # guess, and must never fall back to the other side when one fails.
    [switch]$OnPrem
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

#region Environment mode -------------------------------------------------------
$script:OnPremMode = [bool]$OnPrem

# Filled by Resolve-DLEnvironment in on-premises mode. Empty in cloud mode, and
# splatting an empty hashtable adds nothing, so every call site stays identical.
$script:DCParam = @{}

function Invoke-DLCmd {
    <#
        Runs an Exchange cmdlet against whichever environment this instance is
        for. THE ONLY PLACE THAT KNOWS THE DIFFERENCE.

        Cloud goes through Invoke-EXOCommand, which resolves explicitly to the
        Exchange Online module -- so a stray on-premises session in this process
        cannot capture Get-/Set-DistributionGroup.

        On-premises calls the cmdlet directly, with the domain controller
        splatted in. It cannot use Invoke-EXOCommand for the obvious reason, and
        it does not need the same protection for the less obvious one: the tool
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
    <#
        'CONTOSO\bxrubble' -> 'bxrubble'.

        Get-ADPermission reports the trustee in DOMAIN\user form, which is not
        what Add-ADPermission or Get-ADUser want back.
    #>
    param([string]$UserString)
    $s = "$UserString"
    if ($s -match '\\') { return ($s -split '\\')[-1] }
    return $s
}
#endregion


#region Parsing ----------------------------------------------------------------
function ConvertTo-NormalizedUserList {
    param([string]$Text)
    if (-not $Text) { return @() }
    $raw = @($Text -split '[\r\n,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $result = @()
    foreach ($u in $raw) {
        if ($seen.Add($u.ToLowerInvariant())) { $result += $u }
    }
    return $result
}
#endregion

#region Send-To actions --------------------------------------------------------
function Get-SendToState {
    <#
        Reads the allowed-senders list.

        Cloud DLs return AcceptMessagesOnlyFromSendersOrMembers as names, not
        distinguished names, so every entry is re-resolved to a primary SMTP
        address. That address is what the list is rewritten with -- keying on
        anything else means the set comparison below can miss a match that is
        really there, and silently add a duplicate.

        An entry that will not re-resolve is reported rather than dropped: it
        is usually a deleted recipient still sitting on the list, and quietly
        removing it on the next write would be a change nobody asked for.
    #>
    param([Parameter(Mandatory)][string]$DLIdentity)
    $grp = Invoke-DLCmd -Name 'Get-DistributionGroup' -Parameters @{
        Identity = $DLIdentity; ErrorAction = 'Stop' }
    $rawSenders = @($grp.AcceptMessagesOnlyFromSendersOrMembers)
    $resolved = @($rawSenders | ForEach-Object {
        $entry = "$_"
        try {
            $r = Invoke-DLCmd -Name 'Get-Recipient' -Parameters @{
                Identity = $entry; ErrorAction = 'Stop' }
            [PSCustomObject]@{
                Display = "$($r.DisplayName) <$($r.PrimarySmtpAddress)>"
                Key     = "$($r.PrimarySmtpAddress)"
            }
        } catch {
            [PSCustomObject]@{
                Display = "$entry (could not re-resolve -- may be stale/deleted)"
                Key     = $null
            }
        }
    })
    return [PSCustomObject]@{
        DisplayName                        = $grp.DisplayName
        Identity                           = "$($grp.Identity)"
        RequireSenderAuthenticationEnabled = [bool]$grp.RequireSenderAuthenticationEnabled
        Senders                            = $resolved
        RawSenders                         = $rawSenders
    }
}

function Format-SendToState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading ---")
    [void]$sb.AppendLine("  DL              : $($State.DisplayName)  ($($State.Identity))")
    if ($State.RequireSenderAuthenticationEnabled) {
        [void]$sb.AppendLine("  Sender auth req : ON  (this allowed-senders list IS enforced)")
    } else {
        [void]$sb.AppendLine("  Sender auth req : OFF -- this allowed-senders list currently has NO EFFECT until this is turned on.")
    }
    [void]$sb.AppendLine("  Allowed senders ($($State.Senders.Count)):")
    if ($State.Senders.Count -eq 0) {
        [void]$sb.AppendLine("    (none -- DL is not restricted to specific senders)")
    } else {
        $State.Senders | ForEach-Object { [void]$sb.AppendLine("    $($_.Display)") }
    }
    return $sb.ToString()
}

function Invoke-SendToChange {
    <#
        Set-DistributionGroup replaces the whole list, so the current list is
        read, edited in memory, and written back. Unresolvable entries are
        carried through untouched -- see Get-SendToState.
    #>
    param(
        [Parameter(Mandatory)][string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb,
        [bool]$AlsoEnableAuth
    )
    $out = New-Object System.Text.StringBuilder
    $before = Get-SendToState -DLIdentity $DLIdentity
    [void]$out.Append((Format-SendToState -Heading "BEFORE" -State $before))
    [void]$out.AppendLine("")

    $resolvedUsers = @()
    foreach ($u in $Users) {
        try {
            $r = Invoke-DLCmd -Name 'Get-Recipient' -Parameters @{
                Identity = $u; ErrorAction = 'Stop' }
            $resolvedUsers += [PSCustomObject]@{
                Display = "$($r.DisplayName) <$($r.PrimarySmtpAddress)>"
                Key     = "$($r.PrimarySmtpAddress)"
            }
        } catch {
            [void]$out.AppendLine("  ERROR resolving '$u' -- skipped, not included in this change: $($_.Exception.Message)")
        }
    }

    $currentKeys = @($before.Senders | Where-Object { $_.Key } | ForEach-Object { $_.Key })
    $currentSet  = New-Object System.Collections.Generic.HashSet[string]([string[]]$currentKeys, [System.StringComparer]::OrdinalIgnoreCase)
    $newList     = New-Object System.Collections.ArrayList
    foreach ($k in $currentKeys) { [void]$newList.Add($k) }
    foreach ($s in $before.Senders) {
        if (-not $s.Key) { [void]$newList.Add(($s.Display -replace ' \(could not re-resolve.*$', '')) }
    }

    $changed = $false
    $actedOn = @()
    foreach ($ru in $resolvedUsers) {
        $already = $currentSet.Contains($ru.Key)
        if ($Verb -eq 'Add') {
            if ($already) {
                [void]$out.AppendLine("  SKIP (already allowed to send to this DL): $($ru.Display)")
            } else {
                [void]$newList.Add($ru.Key)
                [void]$currentSet.Add($ru.Key)
                [void]$out.AppendLine("  ADDED: $($ru.Display)")
                $actedOn += $ru.Display
                $changed = $true
            }
        } else {
            if (-not $already) {
                [void]$out.AppendLine("  SKIP (not currently on the allowed-senders list): $($ru.Display)")
            } else {
                $newList.Remove($ru.Key)
                [void]$currentSet.Remove($ru.Key)
                [void]$out.AppendLine("  REMOVED: $($ru.Display)")
                $actedOn += $ru.Display
                $changed = $true
            }
        }
    }

    if (-not $changed) {
        [void]$out.AppendLine("")
        [void]$out.AppendLine("No change -- every requested user was already in the desired state. Nothing written.")
        return [PSCustomObject]@{ Output = $out.ToString(); Changed = $false; ActedOn = @() }
    }

    Invoke-DLCmd -Name 'Set-DistributionGroup' -Parameters @{
        Identity = $DLIdentity
        AcceptMessagesOnlyFromSendersOrMembers = $newList.ToArray()
        Confirm = $false; ErrorAction = 'Stop' } | Out-Null

    if ($Verb -eq 'Add' -and $AlsoEnableAuth -and -not $before.RequireSenderAuthenticationEnabled) {
        Invoke-DLCmd -Name 'Set-DistributionGroup' -Parameters @{
            Identity = $DLIdentity; RequireSenderAuthenticationEnabled = $true
            Confirm = $false; ErrorAction = 'Stop' } | Out-Null
        [void]$out.AppendLine("")
        [void]$out.AppendLine("Also enabled 'require sender authentication' (was OFF) -- this list is now actually enforced.")
    }

    $after = Get-SendToState -DLIdentity $DLIdentity
    [void]$out.AppendLine("")
    [void]$out.Append((Format-SendToState -Heading "AFTER" -State $after))

    return [PSCustomObject]@{ Output = $out.ToString(); Changed = $true; ActedOn = $actedOn }
}
#endregion

#region Send-As actions --------------------------------------------------------
function Resolve-CloudRecipient {
    <#
        Confirms a user exists and returns the handle to grant against.

        Cloud-only, so there is no sAMAccountName to find. Send As is granted
        to a recipient and the primary SMTP address is the stable identifier,
        which is also what Get-RecipientPermission reports back as Trustee --
        so the same value compares cleanly on both sides.
    #>
    param([Parameter(Mandatory)][string]$Identity)
    $r = Invoke-DLCmd -Name 'Get-Recipient' -Parameters @{
        Identity = $Identity; ErrorAction = 'Stop' }
    return "$($r.PrimarySmtpAddress)"
}

function Get-SendAsState {
    <#
        Who currently holds Send As on this group.

        Get-RecipientPermission is the cloud equivalent of the on-premises
        Get-ADPermission -ExtendedRights Send-As. NT AUTHORITY\SELF is on every
        object and is the platform's own grant, so it is filtered rather than
        reported as somebody's delegation.

        Deny entries are kept and flagged. A Deny outranks an Allow, so hiding
        one would mean reporting access that does not actually work.
    #>
    param([Parameter(Mandatory)][string]$DLIdentity)
    $grants = @()

    # THE TWO ENVIRONMENTS MODEL THIS DIFFERENTLY, and the shapes do not match.
    #
    # Exchange Online has Get-RecipientPermission, whose trustee is an ADDRESS.
    # On-premises has Get-ADPermission, whose User is DOMAIN\sam and whose
    # rights are a collection to be searched rather than a string to be matched.
    #
    # Both are normalised to the same User/Key/Deny shape here, so everything
    # downstream -- the set comparison, the display, the audit line -- is written
    # once. Key is what gets passed back to the add/remove cmdlet, and it
    # deliberately differs by side: an address in the cloud, a sAMAccountName
    # on-premises.
    if ($script:OnPremMode) {
        $perms = @(Invoke-DLCmd -Name 'Get-ADPermission' -Parameters @{
            Identity = $DLIdentity; ErrorAction = 'Stop' })
        foreach ($p in $perms) {
            $who = "$($p.User)"
            if ($who -like '*\SELF') { continue }
            if ($p.IsInherited) { continue }
            if (-not (@($p.ExtendedRights) -match 'Send-As')) { continue }
            $grants += [PSCustomObject]@{
                User = $who
                Key  = (Get-SamFromUserString -UserString $who)
                Deny = [bool]$p.Deny
            }
        }
    } else {
        $perms = @(Invoke-DLCmd -Name 'Get-RecipientPermission' -Parameters @{
            Identity = $DLIdentity; ErrorAction = 'Stop' })
        foreach ($p in $perms) {
            $trustee = "$($p.Trustee)"
            if ($trustee -match 'NT AUTHORITY\\SELF') { continue }
            if ("$($p.AccessRights)" -notmatch 'SendAs') { continue }
            $grants += [PSCustomObject]@{
                User = $trustee
                Key  = $trustee
                Deny = ("$($p.AccessControlType)" -eq 'Deny')
            }
        }
    }
    return [PSCustomObject]@{
        Identity = $DLIdentity
        Grants   = @($grants | Sort-Object User)
    }
}

function Format-SendAsState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading ---")
    [void]$sb.AppendLine("  DL              : $($State.Identity)")
    [void]$sb.AppendLine("  Send-As grants ($($State.Grants.Count)), SELF excluded:")
    if ($State.Grants.Count -eq 0) {
        [void]$sb.AppendLine("    (none)")
    } else {
        $State.Grants | ForEach-Object {
            $tag = if ($_.Deny) { ' [DENY]' } else { '' }
            [void]$sb.AppendLine("    $($_.User)$tag")
        }
    }
    return $sb.ToString()
}

function Add-SendAsGrant {
    # Different cmdlet, different parameter name, same intent. On-premises uses
    # -ExtendedRights 'Send-As' WITH THE HYPHEN; the cloud right is 'SendAs'
    # without one. They are not interchangeable and neither errors usefully if
    # given the other's spelling.
    param([Parameter(Mandatory)][string]$DLIdentity, [Parameter(Mandatory)][string]$UserIdentity)
    if ($script:OnPremMode) {
        Invoke-DLCmd -Name 'Add-ADPermission' -Parameters @{
            Identity = $DLIdentity; User = $UserIdentity; ExtendedRights = 'Send-As'
            Confirm = $false; ErrorAction = 'Stop' } | Out-Null
    } else {
        Invoke-DLCmd -Name 'Add-RecipientPermission' -Parameters @{
            Identity = $DLIdentity; Trustee = $UserIdentity; AccessRights = 'SendAs'
            Confirm = $false; ErrorAction = 'Stop' } | Out-Null
    }
}

function Remove-SendAsGrant {
    param([Parameter(Mandatory)][string]$DLIdentity, [Parameter(Mandatory)][string]$UserIdentity)
    if ($script:OnPremMode) {
        Invoke-DLCmd -Name 'Remove-ADPermission' -Parameters @{
            Identity = $DLIdentity; User = $UserIdentity; ExtendedRights = 'Send-As'
            Confirm = $false; ErrorAction = 'Stop' } | Out-Null
    } else {
        Invoke-DLCmd -Name 'Remove-RecipientPermission' -Parameters @{
            Identity = $DLIdentity; Trustee = $UserIdentity; AccessRights = 'SendAs'
            Confirm = $false; ErrorAction = 'Stop' } | Out-Null
    }
}

function Invoke-SendAsChange {
    param(
        [Parameter(Mandatory)][string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb
    )
    $out = New-Object System.Text.StringBuilder
    $before = Get-SendAsState -DLIdentity $DLIdentity
    [void]$out.Append((Format-SendAsState -Heading "BEFORE" -State $before))
    [void]$out.AppendLine("")

    $beforeKeys = New-Object System.Collections.Generic.HashSet[string](
        [string[]]@($before.Grants | ForEach-Object { $_.Key }),
        [System.StringComparer]::OrdinalIgnoreCase)

    $changed = $false
    $actedOn = @()
    foreach ($u in $Users) {
        try {
            $key = Resolve-CloudRecipient -Identity $u
        } catch {
            [void]$out.AppendLine("  ERROR resolving '$u' via Get-Recipient -- skipped, not included in this change: $($_.Exception.Message)")
            continue
        }
        $already = $beforeKeys.Contains($key)
        if ($Verb -eq 'Add') {
            if ($already) {
                [void]$out.AppendLine("  SKIP (already has Send-As on this DL): $key")
                continue
            }
            try {
                Add-SendAsGrant -DLIdentity $DLIdentity -UserIdentity $key
                [void]$out.AppendLine("  ADDED: $key")
                $actedOn += $key
                $changed = $true
            } catch {
                [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Add-RecipientPermission ($key -> $DLIdentity)"))
            }
        } else {
            if (-not $already) {
                [void]$out.AppendLine("  SKIP (does not currently have Send-As on this DL): $key")
                continue
            }
            try {
                Remove-SendAsGrant -DLIdentity $DLIdentity -UserIdentity $key
                [void]$out.AppendLine("  REMOVED: $key")
                $actedOn += $key
                $changed = $true
            } catch {
                [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Remove-RecipientPermission ($key -> $DLIdentity)"))
            }
        }
    }

    if (-not $changed) {
        [void]$out.AppendLine("")
        [void]$out.AppendLine("No change -- every requested user was already in the desired state (or failed to resolve). Nothing written.")
        return [PSCustomObject]@{ Output = $out.ToString(); Changed = $false; ActedOn = @() }
    }

    $after = Get-SendAsState -DLIdentity $DLIdentity
    [void]$out.AppendLine("")
    [void]$out.Append((Format-SendAsState -Heading "AFTER" -State $after))

    return [PSCustomObject]@{ Output = $out.ToString(); Changed = $true; ActedOn = $actedOn }
}
#endregion

#region Combined orchestration ---------------------------------------------------

function Invoke-DLAccessChange {
    param(
        [Parameter(Mandatory)][string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb,
        [bool]$DoSendTo,
        [bool]$DoSendAs,
        [bool]$AlsoEnableAuth
    )
    $out = New-Object System.Text.StringBuilder
    $sendToChanged = $false; $sendToActedOn = @()
    $sendAsChanged = $false; $sendAsActedOn = @()

    if ($DoSendTo) {
        [void]$out.AppendLine("=== Send-To (allowed senders) ===")
        $r = Invoke-SendToChange -DLIdentity $DLIdentity -Users $Users -Verb $Verb -AlsoEnableAuth ($Verb -eq 'Add' -and $AlsoEnableAuth)
        [void]$out.Append($r.Output)
        [void]$out.AppendLine("")
        $sendToChanged = $r.Changed
        $sendToActedOn = $r.ActedOn
    }
    if ($DoSendAs) {
        [void]$out.AppendLine("=== Send-As (permission delegation) ===")
        $r2 = Invoke-SendAsChange -DLIdentity $DLIdentity -Users $Users -Verb $Verb
        [void]$out.Append($r2.Output)
        [void]$out.AppendLine("")
        $sendAsChanged = $r2.Changed
        $sendAsActedOn = $r2.ActedOn
    }

    return [PSCustomObject]@{
        Output        = $out.ToString()
        SendToChanged = $sendToChanged
        SendToActedOn = $sendToActedOn
        SendAsChanged = $sendAsChanged
        SendAsActedOn = $sendAsActedOn
    }
}
#endregion

#region WinForms UI -------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = if ($script:OnPremMode) { 'DL Send-To / Send-As  [ON-PREMISES]' }
             else                     { 'DL Send-To / Send-As  [CLOUD]' }
$form.ClientSize = New-Object System.Drawing.Size(764, 703)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
# THE GUARD IS SYMMETRICAL, because the danger is. Whichever side this
# instance is for, the other one loaded in the same process would answer
# instead -- the cmdlet names are shared and the last import wins.
$emsState = Test-EmsAvailable
$exoState = Test-ExoInProcess
$script:Blocked = $false

if ($script:OnPremMode) {
    if ($exoState.Present) {
        $lblWarn.Text   = 'BLOCKED: Exchange Online is loaded in this process and would answer instead of Active Directory. This must run in its own window.'
        $script:Blocked = $true
    } elseif (-not $emsState.Available) {
        $lblWarn.Text   = 'NOT connected to on-premises Exchange -- nothing can run. Use Connect ExSE.'
        $script:Blocked = $true
    } else {
        $lblWarn.Text = 'ON-PREMISES mode. Connected to Exchange on-premises only. The directory is resolved per DL from Tenant.json.'
    }
} else {
    if ($emsState.Available) {
        $lblWarn.Text   = 'BLOCKED: on-premises Exchange is loaded in this process and would answer instead of the cloud.'
        $script:Blocked = $true
    } elseif (-not $exoState.Present) {
        $lblWarn.Text   = 'NOT connected to Exchange Online -- nothing can run. Use Connect ExO.'
        $script:Blocked = $true
    } else {
        $lblWarn.Text = 'CLOUD mode. Every call is routed through the Exchange Online session explicitly.'
    }
}
$lblWarn.ForeColor = if ($script:Blocked) { [System.Drawing.Color]::Red }
                     else                 { [System.Drawing.Color]::FromArgb(0, 128, 0) }
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(740, 32)
$form.Controls.Add($lblWarn)

function Resolve-DLEnvironment {
    <#
        Confirms the distribution list exists before anything tries to change
        it, so a typo produces one clear message instead of two half-finished
        Send-To / Send-As passes.
    #>
    param([Parameter(Mandatory)][string]$DLIdentity)

    $errors = @()

    # CLOUD: one directory, nothing to choose.
    if (-not $script:OnPremMode) {
        try {
            $g = Invoke-DLCmd -Name 'Get-DistributionGroup' -Parameters @{
                Identity = $DLIdentity; ErrorAction = 'Stop' }
            return [PSCustomObject]@{ Found = $true; Label = "$($g.PrimarySmtpAddress)"; Errors = $errors }
        } catch {
            $errors += $_.Exception.Message
        }
        return [PSCustomObject]@{ Found = $false; Label = $null; Errors = $errors }
    }

    # ON-PREMISES: EVERY DIRECTORY IN Tenant.json, TRIED IN TURN.
    #
    # THIS IS A REWRITE, NOT A PORT. The retired AD- version read:
    #
    #     Environment = $dirId ... Get-TenantDirectoryServer -Id $dirId
    #
    # and $dirId was never assigned anywhere in that file. Its first branch
    # therefore returned a NULL environment and asked for the controller of a
    # null directory, and its second hardcoded one directory's name -- a
    # tenant's vocabulary in a .ps1, which is the one thing Test-Toolkit
    # refuses. The whole function is replaced rather than repaired.
    #
    # Directories come from Tenant.json, so a tenant with one gets one and a
    # tenant with three gets three, with no names written down here. The FIRST
    # directory that can resolve the group wins, and its controller is then used
    # for every subsequent call through $script:DCParam -- so the read and the
    # write cannot land in different forests.
    foreach ($d in @(Get-TenantDirectory)) {
        $srv = Get-TenantDirectoryServer -Id $d.Id
        if (-not $srv) { continue }
        try {
            $g = Get-DistributionGroup -Identity $DLIdentity -DomainController $srv -ErrorAction Stop
            $script:DCParam = @{ DomainController = $srv }
            return [PSCustomObject]@{
                Found  = $true
                Label  = "$($g.PrimarySmtpAddress)  [$($d.Id) via $srv]"
                Errors = $errors
            }
        } catch {
            $errors += "$($d.Id) ($srv): $($_.Exception.Message)"
        }
    }

    # NOT FOUND ANYWHERE MEANS DCParam STAYS EMPTY, and that is deliberate: a
    # stale controller left over from a previous lookup would send the next
    # write at the wrong forest.
    $script:DCParam = @{}
    return [PSCustomObject]@{ Found = $false; Label = $null; Errors = $errors }
}

$lblDL = New-Object System.Windows.Forms.Label
$lblDL.Text = "Distribution List (name, not SMTP):"
$lblDL.Location = New-Object System.Drawing.Point(12, 58)
$lblDL.AutoSize = $true
$form.Controls.Add($lblDL)
$txtDL = New-Object System.Windows.Forms.TextBox
$txtDL.Location = New-Object System.Drawing.Point(12, 78)
$txtDL.Size = New-Object System.Drawing.Size(740, 22)
$txtDL.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtDL)

$lblUsers = New-Object System.Windows.Forms.Label
$lblUsers.Text = "User(s) -- optional for Lookup, required for Add/Remove (firstname.lastname, one per line):"
$lblUsers.Location = New-Object System.Drawing.Point(12, 108)
$lblUsers.AutoSize = $true
$form.Controls.Add($lblUsers)
$txtUsers = New-Object System.Windows.Forms.TextBox
$txtUsers.Multiline = $true
$txtUsers.ScrollBars = "Vertical"
$txtUsers.Location = New-Object System.Drawing.Point(12, 128)
$txtUsers.Size = New-Object System.Drawing.Size(740, 110)
$txtUsers.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtUsers)

$grpAccessType = New-Object System.Windows.Forms.GroupBox
$grpAccessType.Text = "Access Type -- select one or both"
$grpAccessType.Size = New-Object System.Drawing.Size(740, 45)
$grpAccessType.Location = New-Object System.Drawing.Point(12, 246)
$form.Controls.Add($grpAccessType)

$chkSendTo = New-Object System.Windows.Forms.CheckBox
$chkSendTo.Text = "Send-To (allowed senders)"
$chkSendTo.AutoSize = $true
$chkSendTo.Location = New-Object System.Drawing.Point(15, 18)
$chkSendTo.Checked = $true
$grpAccessType.Controls.Add($chkSendTo)

$chkSendAs = New-Object System.Windows.Forms.CheckBox
$chkSendAs.Text = "Send-As (permission delegation)"
$chkSendAs.AutoSize = $true
$chkSendAs.Location = New-Object System.Drawing.Point(220, 18)
$grpAccessType.Controls.Add($chkSendAs)

$chkAlsoEnable = New-Object System.Windows.Forms.CheckBox
$chkAlsoEnable.Text = "On Add: also enable 'require sender authentication' if currently OFF (Send-To only -- needed for that list to have any effect; see warning)"
$chkAlsoEnable.Location = New-Object System.Drawing.Point(12, 299)
$chkAlsoEnable.Size = New-Object System.Drawing.Size(740, 22)
$form.Controls.Add($chkAlsoEnable)

$chkSendTo.Add_CheckedChanged({
    $chkAlsoEnable.Enabled = $chkSendTo.Checked
    if (-not $chkSendTo.Checked) { $chkAlsoEnable.Checked = $false }
})

$btnLookup = New-Object System.Windows.Forms.Button
$btnLookup.Text = "Lookup"
$btnLookup.Location = New-Object System.Drawing.Point(12, 331)
$btnLookup.Size = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnLookup)
$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add Access"
$btnAdd.Location = New-Object System.Drawing.Point(130, 331)
$btnAdd.Size = New-Object System.Drawing.Size(140, 30)
$btnAdd.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnAdd)
$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remove Access"
$btnRemove.Location = New-Object System.Drawing.Point(278, 331)
$btnRemove.Size = New-Object System.Drawing.Size(150, 30)
$btnRemove.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRemove)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(436, 331)
$btnClose.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnClose)

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Result:"
$lblResult.Location = New-Object System.Drawing.Point(12, 371)
$lblResult.AutoSize = $true
$form.Controls.Add($lblResult)
$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Both"
$txtOut.WordWrap = $false
$txtOut.Location = New-Object System.Drawing.Point(12, 391)
$txtOut.Size = New-Object System.Drawing.Size(740, 300)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOut)

function Set-ButtonsEnabled {
    param([bool]$Enabled)
    $btnLookup.Enabled = $Enabled
    $btnAdd.Enabled = $Enabled
    $btnRemove.Enabled = $Enabled
}

function Get-AccessTypeSummary {
    $parts = @()
    if ($chkSendTo.Checked) { $parts += 'Send-To' }
    if ($chkSendAs.Checked) { $parts += 'Send-As' }
    return ($parts -join ' and ')
}

$btnLookup.Add_Click({
    # THE GUARD IS ENFORCED HERE, not only displayed.
    #
    # $script:Blocked was set from the environment check and then read by
    # nothing but a label colour -- a red banner over three buttons that all
    # still worked. Same wording as the banner, so the refusal and the
    # explanation cannot drift apart.
    if ($script:Blocked) {
        [void][System.Windows.Forms.MessageBox]::Show($lblWarn.Text,
            'Cannot run in this process',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    Set-ButtonsEnabled -Enabled $false
    try {
        if (-not $chkSendTo.Checked -and -not $chkSendAs.Checked) {
            $txtOut.Text = "Check Send-To and/or Send-As above first."
            return
        }
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Found) {
            $txtOut.Text = "'$dl' was not found. Confirm the name, and confirm this window is connected to Exchange Online.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }
        $out = New-Object System.Text.StringBuilder
        [void]$out.AppendLine("DL: $($envResult.Label)")
        [void]$out.AppendLine("")
        if ($chkSendTo.Checked) {
            $stateTo = Get-SendToState -DLIdentity $dl
            [void]$out.AppendLine("=== Send-To (allowed senders) ===")
            [void]$out.Append((Format-SendToState -Heading "CURRENT STATE" -State $stateTo))
            [void]$out.AppendLine("")
        }
        if ($chkSendAs.Checked) {
            $stateAs = Get-SendAsState -DLIdentity $dl
            [void]$out.AppendLine("=== Send-As (permission delegation) ===")
            [void]$out.Append((Format-SendAsState -Heading "CURRENT STATE" -State $stateAs))
        }
        $txtOut.Text = $out.ToString()
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Lookup DL Access'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnAdd.Add_Click({
    # THE GUARD IS ENFORCED HERE, not only displayed.
    #
    # $script:Blocked was set from the environment check and then read by
    # nothing but a label colour -- a red banner over three buttons that all
    # still worked. Same wording as the banner, so the refusal and the
    # explanation cannot drift apart.
    if ($script:Blocked) {
        [void][System.Windows.Forms.MessageBox]::Show($lblWarn.Text,
            'Cannot run in this process',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    Set-ButtonsEnabled -Enabled $false
    try {
        if (-not $chkSendTo.Checked -and -not $chkSendAs.Checked) {
            $txtOut.Text = "Check Send-To and/or Send-As above first."
            return
        }
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $users = ConvertTo-NormalizedUserList -Text $txtUsers.Text
        if ($users.Count -eq 0) { $txtOut.Text = "Enter at least one user in the second box first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Found) {
            $txtOut.Text = "'$dl' was not found. Confirm the name, and confirm this window is connected to Exchange Online.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }

        $authNote = ''
        if ($chkSendTo.Checked -and -not $chkAlsoEnable.Checked) {
            $pre = Get-SendToState -DLIdentity $dl
            if (-not $pre.RequireSenderAuthenticationEnabled) {
                $authNote = "`n`nNOTE: 'require sender authentication' is currently OFF for this DL -- the Send-To list will have NO EFFECT until that's turned on (check the box above, or enable it separately)."
            }
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Add $($users.Count) user(s) to $(Get-AccessTypeSummary) on '$dl' ($($envResult.Label))?$authNote",
            "Confirm Add Access",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }

        $result = Invoke-DLAccessChange -DLIdentity $dl -Users $users -Verb 'Add' `
            -DoSendTo $chkSendTo.Checked -DoSendAs $chkSendAs.Checked -AlsoEnableAuth $chkAlsoEnable.Checked
        $txtOut.Text = $result.Output

        if ($result.SendToChanged -and $result.SendToActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Added Send-To $dl ($($envResult.Found); users: $($result.SendToActedOn -join '+'))"
        }
        if ($result.SendAsChanged -and $result.SendAsActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Added Send-As $dl ($($envResult.Found); users: $($result.SendAsActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add DL Access'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnRemove.Add_Click({
    # THE GUARD IS ENFORCED HERE, not only displayed.
    #
    # $script:Blocked was set from the environment check and then read by
    # nothing but a label colour -- a red banner over three buttons that all
    # still worked. Same wording as the banner, so the refusal and the
    # explanation cannot drift apart.
    if ($script:Blocked) {
        [void][System.Windows.Forms.MessageBox]::Show($lblWarn.Text,
            'Cannot run in this process',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    Set-ButtonsEnabled -Enabled $false
    try {
        if (-not $chkSendTo.Checked -and -not $chkSendAs.Checked) {
            $txtOut.Text = "Check Send-To and/or Send-As above first."
            return
        }
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $users = ConvertTo-NormalizedUserList -Text $txtUsers.Text
        if ($users.Count -eq 0) { $txtOut.Text = "Enter at least one user in the second box first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Found) {
            $txtOut.Text = "'$dl' was not found. Confirm the name, and confirm this window is connected to Exchange Online.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($users.Count) user(s) from $(Get-AccessTypeSummary) on '$dl' ($($envResult.Label))?",
            "Confirm Remove Access",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }

        $result = Invoke-DLAccessChange -DLIdentity $dl -Users $users -Verb 'Remove' `
            -DoSendTo $chkSendTo.Checked -DoSendAs $chkSendAs.Checked -AlsoEnableAuth $false
        $txtOut.Text = $result.Output

        if ($result.SendToChanged -and $result.SendToActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Removed Send-To $dl ($($envResult.Found); users: $($result.SendToActedOn -join '+'))"
        }
        if ($result.SendAsChanged -and $result.SendAsActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Removed Send-As $dl ($($envResult.Found); users: $($result.SendAsActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove DL Access'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form

Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL Permissions' | Out-Null
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
#endregion
