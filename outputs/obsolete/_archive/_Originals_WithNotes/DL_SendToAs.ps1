#Requires -Version 5.1
<#
.SYNOPSIS
    Lookup / Add / Remove Send-To (allowed senders) and/or Send-As permission
    on an on-prem Distribution List -- merged from the operator's two
    separate tools, SendToDL.ps1 and SendAsDL.ps1, into a single console.
.DESCRIPTION
    REVISED 2026-07-24 -- merge of two previously separate tools:
      - SendToDL.ps1 (a DL's delivery restriction --
        Set-DistributionGroup -AcceptMessagesOnlyFromSendersOrMembers -- who
        is allowed to send TO the DL)
      - SendAsDL.ps1 (Send-As extended right --
        Add-/Remove-ADPermission -ExtendedRights "Send-As" -- who can send
        mail that appears to come FROM the DL)
    Both tools already shared the identical "1 DL x N users" shape (one DL
    box, one users box, Lookup/Add/Remove/Close four-button layout instead
    of the standard Parse/Run/Close triad, both look up the DL by Name with
    no domain-append on the users box, both auto-detect HQ vs Franchise the
    same way via an identical Resolve-DLEnvironment probe), so this merge
    keeps everything about that shape unchanged and adds one thing: which of
    the two independent operations to run.

    UNLIKE the Full Access / Send-As / OneDrive merge (MailboxOneDriveAccess.
    ps1), Send-To and Send-As on a DL have no stated exclusion rule between
    them -- there's no "these two together never make sense" constraint the
    way Send-As + OneDrive access was ruled out there. A DL delegation change
    can legitimately need either, or both, in the same sitting. So Access
    Type here is two independent CHECKBOXES (Send-To, Send-As), not a
    five-way mutually exclusive radio group -- both can be checked at once,
    and Add/Remove runs whichever ones are checked, one after the other, in
    a single confirmation and a single combined result. At least one must be
    checked for Lookup/Add/Remove to do anything.

    Send-To defaults checked; Send-As defaults unchecked -- narrowest single
    default, same reasoning as MailboxOneDriveAccess.ps1 defaulting to Full
    Access alone rather than a combined option.

    What stayed exactly as each source tool built it, moved here unchanged:
      - Get-SendToState / Format-SendToState / Invoke-SendToChange (from
        SendToDL.ps1) -- full-REPLACE semantics on
        AcceptMessagesOnlyFromSendersOrMembers, current-list and typed users
        both resolved to DistinguishedName via Get-Recipient before diffing
        (so differently-typed identities for the same recipient still
        match), one Set-DistributionGroup call per run, SKIP (not error) for
        already-satisfied requests, nothing written at all if nothing
        actually changes.
      - The "On Add: also enable require sender authentication" checkbox --
        AcceptMessagesOnlyFromSendersOrMembers has no effect unless
        RequireSenderAuthenticationEnabled is also true, a well-known
        Exchange gotcha SendToDL.ps1 surfaced explicitly rather than silently
        fixing. It only applies to a Send-To Add, so it's disabled (and
        unchecked, so stale state can't linger) whenever the Send-To
        checkbox itself is unchecked.
      - Get-SamAccountNameFromUserString / Resolve-SamAccountName /
        Get-SendAsState / Format-SendAsState / Add-SendAsGrant /
        Remove-SendAsGrant / Invoke-SendAsChange (from SendAsDL.ps1) --
        naturally incremental Add-/Remove-ADPermission per resolved user
        (not a full-list replace like Send-To), inherited ACEs and
        NT AUTHORITY\SELF excluded from every BEFORE/AFTER snapshot, the
        hyphenated "Send-As" ExtendedRights string (not "Send As" with a
        space -- see SendAsDL.ps1's original KNOWN SYNTAX GOTCHA note, still
        the standards doc's field-tested correction over current public
        Microsoft Learn examples), SKIP (not error) for already-satisfied
        requests.
      - Resolve-DLEnvironment -- byte-for-byte identical in both source
        files already (auto-detect HQ vs Franchise by trying
        Get-DistributionGroup against HQ first, no DC pinned, then Franchise
        pinned to $script:DefaultDC only if HQ doesn't have it), so there was
        nothing to reconcile; kept as the superset shape SendAsDL.ps1's copy
        already returned (Environment, DCParam for the Exchange-side calls,
        AND a separate plain ADServer string for Get-ADUser's -Server
        parameter -- see each source file's own BUG FIX note for why those
        two shapes differ for HQ). Called once per Lookup/Add/Remove click,
        same as both originals; if neither forest has the DL,
        Lookup/Add/Remove report "not found in HQ or Franchise" instead of
        guessing.
      - ConvertTo-NormalizedUserList -- byte-for-byte identical in both
        source files already, kept as one copy.
      - The -SmtpAddress / -Environment / -DomainController params at the
        top: both source tools already declared these ONLY for
        calling-convention consistency with AdminConsole.ps1 (which passes
        the same three parameters to every tool it launches) and never
        actually wired any of them to anything -- the DL box is typed fresh
        every time, matching "standalone, no target lookup" in both
        originals' own headers. That is preserved exactly; this merge does
        not add a use for them that didn't exist before.

    Combining the two into one Add/Remove: when both Send-To and Send-As are
    checked, one click runs Send-To's full BEFORE/action/AFTER block, then
    Send-As's, into one combined result box, behind ONE confirmation dialog
    that names both operations up front. Each operation still gets its own
    Write-CompletedLog line, independently, so a run where only one of the
    two actually changed anything (e.g. Send-To was already satisfied but
    Send-As wasn't) logs exactly the one that did.

    Warning label: kept SendToDL.ps1's stronger "conflict risk" wording
    (Get-/Set-DistributionGroup and Get-Recipient exist in both on-prem
    Exchange and Exchange Online, so an also-connected EXO session in this
    same host could silently resolve against the wrong DL), not SendAsDL.
    ps1's plain "connect before running" wording -- because Resolve-
    DLEnvironment's own HQ/Franchise probe calls Get-DistributionGroup on
    every single Lookup/Add/Remove regardless of which Access Type
    checkbox(es) are selected, so that collision risk applies to the merged
    tool unconditionally, not only when Send-To is checked.

    Deliberate deviation from the standard Parse/Run/Close button triad
    (Section 13), preserved from both originals: Lookup / Add Access /
    Remove Access / Close. No separate Parse step; Lookup IS the preview,
    and Add/Remove each re-run their own BEFORE snapshot(s) internally at
    the start of every run regardless of whether Lookup was clicked first.

    CMDLET COLLISION WARNING (Section 3), same as both originals: every
    cmdlet here is called by its bare name (Get-/Set-DistributionGroup,
    Get-Recipient, Get-/Add-/Remove-ADPermission, Get-ADUser), never through
    Invoke-EXOCommand, because this tool wants whatever on-prem Exchange
    Management Shell session is already loaded and has no Exchange Online
    dependency at all -- same as DLManager.ps1, SendToDL.ps1, and
    SendAsDL.ps1 before it.
.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-07-21 (SendToDL.ps1, SendAsDL.ps1, as two separate
                       tools)
    Revised         : 2026-07-24 (merged into DLAccessManager.ps1 -- see
                       REVISED note above)
    Run from        : on-prem Exchange Management Shell (or Exchange snap-in
                       loaded), with the ActiveDirectory (RSAT) module
                       available for Get-ADUser (needed only if Send-As is
                       used)
    Required modules: ActiveDirectory (Send-As only; Send-To alone needs only
                       the on-prem Exchange Management Shell)
#>
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Parsing ----------------------------------------------------------------
# Tolerates newline, comma, and semicolon as separators (Section 21
# convention). No domain-append -- these are AD logons (firstname.lastname),
# not email addresses. Byte-for-byte identical to both source tools' own
# copy of this function.
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

#region Send-To actions -- from SendToDL.ps1, unchanged ------------------------
# $DCParam is a splatted hashtable -- @{ DomainController = <fqdn> } for
# Franchise, @{} for HQ. Every Exchange cmdlet call below uses @DCParam
# instead of a hardcoded -DomainController so HQ never passes a bare domain
# name where a real server is required.
function Get-SendToState {
    param([string]$DLIdentity, [hashtable]$DCParam)
    $grp = Get-DistributionGroup -Identity $DLIdentity @DCParam -ErrorAction Stop   # on-prem
    $rawSenders = @($grp.AcceptMessagesOnlyFromSendersOrMembers)
    $resolved = @($rawSenders | ForEach-Object {
        try {
            $r = Get-Recipient -Identity "$_" @DCParam -ErrorAction Stop   # on-prem
            [PSCustomObject]@{ Display = "$($r.Name) ($($r.SamAccountName))"; DN = $r.DistinguishedName }
        } catch {
            [PSCustomObject]@{ Display = "$_ (could not re-resolve -- may be stale/deleted)"; DN = $null }
        }
    })
    return [PSCustomObject]@{
        DisplayName                        = $grp.DisplayName
        Identity                           = $grp.Identity.ToString()
        RequireSenderAuthenticationEnabled = [bool]$grp.RequireSenderAuthenticationEnabled
        Senders                             = $resolved
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

# Orchestrates one Add or Remove run against a single DL for Send-To. Resolves
# every current entry AND every typed user to a DistinguishedName so the diff
# is never fooled by two different-looking identities for the same recipient.
function Invoke-SendToChange {
    param(
        [string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb,
        [hashtable]$DCParam,
        [bool]$AlsoEnableAuth
    )
    $out = New-Object System.Text.StringBuilder
    $before = Get-SendToState -DLIdentity $DLIdentity -DCParam $DCParam
    [void]$out.Append((Format-SendToState -Heading "BEFORE" -State $before))
    [void]$out.AppendLine("")

    $resolvedUsers = @()
    foreach ($u in $Users) {
        try {
            $r = Get-Recipient -Identity $u @DCParam -ErrorAction Stop   # on-prem
            $resolvedUsers += [PSCustomObject]@{ Display = "$($r.Name) ($($r.SamAccountName))"; DN = $r.DistinguishedName }
        } catch {
            [void]$out.AppendLine("  ERROR resolving '$u' -- skipped, not included in this change: $($_.Exception.Message)")
        }
    }

    $currentDNs = @($before.Senders | Where-Object { $_.DN } | ForEach-Object { $_.DN })
    $currentSet = New-Object System.Collections.Generic.HashSet[string]([string[]]$currentDNs, [System.StringComparer]::OrdinalIgnoreCase)
    $newDNs = New-Object System.Collections.Generic.List[string]
    $currentDNs | ForEach-Object { [void]$newDNs.Add($_) }

    $changed = $false
    $actedOn = @()
    foreach ($ru in $resolvedUsers) {
        $already = $currentSet.Contains($ru.DN)
        if ($Verb -eq 'Add') {
            if ($already) {
                [void]$out.AppendLine("  SKIP (already allowed to send to this DL): $($ru.Display)")
            } else {
                [void]$newDNs.Add($ru.DN)
                [void]$currentSet.Add($ru.DN)
                [void]$out.AppendLine("  ADDED: $($ru.Display)")
                $actedOn += $ru.Display
                $changed = $true
            }
        } else {
            if (-not $already) {
                [void]$out.AppendLine("  SKIP (not currently on the allowed-senders list): $($ru.Display)")
            } else {
                [void]$newDNs.Remove($ru.DN) | Out-Null
                [void]$currentSet.Remove($ru.DN)
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

    Set-DistributionGroup -Identity $DLIdentity -AcceptMessagesOnlyFromSendersOrMembers $newDNs.ToArray() @DCParam -Confirm:$false -ErrorAction Stop   # on-prem

    if ($Verb -eq 'Add' -and $AlsoEnableAuth -and -not $before.RequireSenderAuthenticationEnabled) {
        Set-DistributionGroup -Identity $DLIdentity -RequireSenderAuthenticationEnabled $true @DCParam -Confirm:$false -ErrorAction Stop   # on-prem
        [void]$out.AppendLine("")
        [void]$out.AppendLine("Also enabled 'require sender authentication' (was OFF) -- this list is now actually enforced.")
    }

    $after = Get-SendToState -DLIdentity $DLIdentity -DCParam $DCParam
    [void]$out.AppendLine("")
    [void]$out.Append((Format-SendToState -Heading "AFTER" -State $after))

    return [PSCustomObject]@{ Output = $out.ToString(); Changed = $true; ActedOn = $actedOn }
}
#endregion

#region Send-As actions -- from SendAsDL.ps1, unchanged -------------------------
# Pulls the SamAccountName back out of whatever form Get-ADPermission reports
# an existing grantee in (commonly "DOMAIN\sam", sometimes just "sam").
function Get-SamAccountNameFromUserString {
    param([string]$UserString)
    if ($UserString -match '\\([^\\]+)$') { return $Matches[1] }
    return $UserString
}

function Resolve-SamAccountName {
    param([string]$Identity, [string]$ADServer)
    $u = Get-ADUser -Identity $Identity -Server $ADServer -ErrorAction Stop
    return $u.SamAccountName
}

function Get-SendAsState {
    param([string]$DLIdentity, [hashtable]$DCParam)
    $perms = @(Get-ADPermission -Identity $DLIdentity @DCParam -ErrorAction Stop)
    # Explicit (non-inherited) Send-As grants only -- inherited ACEs and
    # NT AUTHORITY\SELF are noise for a delegation-tracking tool, not
    # permissions anyone granted.
    $grants = @($perms | Where-Object {
        ($_.ExtendedRights -like '*Send-As*') -and
        (-not $_.IsInherited) -and
        ("$($_.User)" -notlike '*\SELF')
    })
    return [PSCustomObject]@{
        Identity = $DLIdentity
        Grants   = @($grants | ForEach-Object {
            [PSCustomObject]@{ User = "$($_.User)"; Deny = [bool]$_.Deny; Sam = (Get-SamAccountNameFromUserString -UserString "$($_.User)") }
        })
    }
}

function Format-SendAsState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading ---")
    [void]$sb.AppendLine("  DL              : $($State.Identity)")
    [void]$sb.AppendLine("  Send-As grants ($($State.Grants.Count)), inherited/SELF excluded:")
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
    param([string]$DLIdentity, [string]$UserIdentity, [hashtable]$DCParam)
    # ExtendedRights "Send-As" (hyphenated) -- see header KNOWN SYNTAX GOTCHA.
    Add-ADPermission -Identity $DLIdentity -User $UserIdentity -ExtendedRights "Send-As" -Confirm:$false @DCParam -ErrorAction Stop
}

function Remove-SendAsGrant {
    param([string]$DLIdentity, [string]$UserIdentity, [hashtable]$DCParam)
    Remove-ADPermission -Identity $DLIdentity -User $UserIdentity -ExtendedRights "Send-As" -Confirm:$false @DCParam -ErrorAction Stop
}

# Orchestrates one Add or Remove run against a single DL for Send-As, one
# Add-/Remove-ADPermission call per resolved user (naturally incremental,
# unlike Send-To's full-list replace). $ADServer feeds Get-ADUser -Server
# (genuine ActiveDirectory module, bare domain name is fine); $DCParam is the
# splat used for the three Exchange *-ADPermission calls (empty for HQ).
function Invoke-SendAsChange {
    param(
        [string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb,
        [string]$ADServer,
        [hashtable]$DCParam
    )
    $out = New-Object System.Text.StringBuilder
    $before = Get-SendAsState -DLIdentity $DLIdentity -DCParam $DCParam
    [void]$out.Append((Format-SendAsState -Heading "BEFORE" -State $before))
    [void]$out.AppendLine("")

    $beforeSams = New-Object System.Collections.Generic.HashSet[string]([string[]]@($before.Grants | ForEach-Object { $_.Sam }), [System.StringComparer]::OrdinalIgnoreCase)

    $changed = $false
    $actedOn = @()
    foreach ($u in $Users) {
        try {
            $sam = Resolve-SamAccountName -Identity $u -ADServer $ADServer
        } catch {
            [void]$out.AppendLine("  ERROR resolving '$u' via Get-ADUser -- skipped, not included in this change: $($_.Exception.Message)")
            continue
        }
        $already = $beforeSams.Contains($sam)
        if ($Verb -eq 'Add') {
            if ($already) {
                [void]$out.AppendLine("  SKIP (already has Send-As on this DL): $sam")
                continue
            }
            try {
                Add-SendAsGrant -DLIdentity $DLIdentity -UserIdentity $sam -DCParam $DCParam
                [void]$out.AppendLine("  ADDED: $sam")
                $actedOn += $sam
                $changed = $true
            } catch {
                [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Add-ADPermission ($sam -> $DLIdentity)"))
            }
        } else {
            if (-not $already) {
                [void]$out.AppendLine("  SKIP (does not currently have Send-As on this DL): $sam")
                continue
            }
            try {
                Remove-SendAsGrant -DLIdentity $DLIdentity -UserIdentity $sam -DCParam $DCParam
                [void]$out.AppendLine("  REMOVED: $sam")
                $actedOn += $sam
                $changed = $true
            } catch {
                [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Remove-ADPermission ($sam -> $DLIdentity)"))
            }
        }
    }

    if (-not $changed) {
        [void]$out.AppendLine("")
        [void]$out.AppendLine("No change -- every requested user was already in the desired state (or failed to resolve). Nothing written.")
        return [PSCustomObject]@{ Output = $out.ToString(); Changed = $false; ActedOn = @() }
    }

    $after = Get-SendAsState -DLIdentity $DLIdentity -DCParam $DCParam
    [void]$out.AppendLine("")
    [void]$out.Append((Format-SendAsState -Heading "AFTER" -State $after))

    return [PSCustomObject]@{ Output = $out.ToString(); Changed = $true; ActedOn = $actedOn }
}
#endregion

#region Combined orchestration ---------------------------------------------------

# ADDED 2026-07-24 -- the only genuinely new logic in this merge: runs
# whichever of Send-To / Send-As are selected, in that order, into one
# combined output, keeping each side's own Changed/ActedOn separate so the
# caller can log each operation type independently (see header note on why
# a run that only changed one of the two should only log that one).
function Invoke-DLAccessChange {
    param(
        [string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb,
        [bool]$DoSendTo,
        [bool]$DoSendAs,
        [bool]$AlsoEnableAuth,
        [hashtable]$DCParam,
        [string]$ADServer
    )
    $out = New-Object System.Text.StringBuilder
    $sendToChanged = $false; $sendToActedOn = @()
    $sendAsChanged = $false; $sendAsActedOn = @()

    if ($DoSendTo) {
        [void]$out.AppendLine("=== Send-To (allowed senders) ===")
        $r = Invoke-SendToChange -DLIdentity $DLIdentity -Users $Users -Verb $Verb -DCParam $DCParam -AlsoEnableAuth ($Verb -eq 'Add' -and $AlsoEnableAuth)
        [void]$out.Append($r.Output)
        [void]$out.AppendLine("")
        $sendToChanged = $r.Changed
        $sendToActedOn = $r.ActedOn
    }
    if ($DoSendAs) {
        [void]$out.AppendLine("=== Send-As (permission delegation) ===")
        $r2 = Invoke-SendAsChange -DLIdentity $DLIdentity -Users $Users -Verb $Verb -ADServer $ADServer -DCParam $DCParam
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
$form.Text = "DL Access Manager -- Send-To / Send-As (on-prem)"
$form.Size = New-Object System.Drawing.Size(780, 745)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

# Kept SendToDL.ps1's stronger "conflict risk" wording -- see header note on
# why this applies to the merged tool unconditionally, not only when Send-To
# is checked.
$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Connect to on-prem Exchange ONLY (do NOT connect to Exchange Online -- Get-/Set-DistributionGroup and Get-Recipient would resolve against the wrong environment). HQ and Franchise are checked automatically per DL name."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(740, 40)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

# Environment radio group and DC selector -- never existed in this merged
# file; both source tools had already removed them (commented out, not
# deleted, in each source file) before this merge, per the same 2026-07-21 /
# 2026-07-23 revisions described in the header. Nothing to carry over here.

# Fixed Franchise DC. LOCAL to this file (not Common.ps1) -- see
# AdminConsole.ps1's header for why a shared copy there caused a real
# breakage in disable-intake.ps1. Used only by Resolve-DLEnvironment below.
$script:DefaultDC = 'fran-vm-dc01.franchises.expresspersonnel.com'

# Byte-for-byte identical in both source files already -- see header note.
# Tries Get-DistributionGroup against HQ first (no DC pinned, forest-wide),
# then Franchise (pinned to $script:DefaultDC) only if HQ doesn't have it.
# Returns BOTH shapes this file needs -- $DCParam (empty-for-HQ splat) for
# every Exchange cmdlet call, and $ADServer (a plain string) for Get-ADUser's
# -Server parameter, needed only when Send-As is used.
function Resolve-DLEnvironment {
    param([string]$DLIdentity)
    $errors = @()
    try {
        $null = Get-DistributionGroup -Identity $DLIdentity -ErrorAction Stop   # on-prem, HQ (no DC pinned)
        return [PSCustomObject]@{ Environment = 'HQ'; DCParam = @{}; ADServer = $script:HQDomain; Errors = $errors }
    } catch {
        $errors += "HQ: $($_.Exception.Message)"
    }
    try {
        $null = Get-DistributionGroup -Identity $DLIdentity -DomainController $script:DefaultDC -ErrorAction Stop   # on-prem, Franchise
        return [PSCustomObject]@{ Environment = 'Franchise'; DCParam = @{ DomainController = $script:DefaultDC }; ADServer = $script:DefaultDC; Errors = $errors }
    } catch {
        $errors += "Franchise: $($_.Exception.Message)"
    }
    return [PSCustomObject]@{ Environment = $null; DCParam = $null; ADServer = $null; Errors = $errors }
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

# Access Type -- two INDEPENDENT checkboxes, not a mutually exclusive radio
# group. See header DESIGN NOTE for why this merge is different from
# MailboxOneDriveAccess.ps1's radio-button approach: there is no rule here
# saying Send-To and Send-As can never both be needed.
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

# Disabled (and unchecked, so no stale state lingers) whenever Send-To itself
# is unchecked -- this checkbox only ever matters inside the Send-To Add
# path.
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
    Set-ButtonsEnabled -Enabled $false
    try {
        if (-not $chkSendTo.Checked -and -not $chkSendAs.Checked) {
            $txtOut.Text = "Check Send-To and/or Send-As above first."
            return
        }
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }
        $out = New-Object System.Text.StringBuilder
        [void]$out.AppendLine("Environment: $($envResult.Environment)")
        [void]$out.AppendLine("")
        if ($chkSendTo.Checked) {
            $stateTo = Get-SendToState -DLIdentity $dl -DCParam $envResult.DCParam
            [void]$out.AppendLine("=== Send-To (allowed senders) ===")
            [void]$out.Append((Format-SendToState -Heading "CURRENT STATE" -State $stateTo))
            [void]$out.AppendLine("")
        }
        if ($chkSendAs.Checked) {
            $stateAs = Get-SendAsState -DLIdentity $dl -DCParam $envResult.DCParam
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
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }

        $authNote = ''
        if ($chkSendTo.Checked -and -not $chkAlsoEnable.Checked) {
            $pre = Get-SendToState -DLIdentity $dl -DCParam $envResult.DCParam
            if (-not $pre.RequireSenderAuthenticationEnabled) {
                $authNote = "`n`nNOTE: 'require sender authentication' is currently OFF for this DL -- the Send-To list will have NO EFFECT until that's turned on (check the box above, or enable it separately)."
            }
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Add $($users.Count) user(s) to $(Get-AccessTypeSummary) on '$dl' ($($envResult.Environment))?$authNote",
            "Confirm Add Access",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }

        $result = Invoke-DLAccessChange -DLIdentity $dl -Users $users -Verb 'Add' `
            -DoSendTo $chkSendTo.Checked -DoSendAs $chkSendAs.Checked -AlsoEnableAuth $chkAlsoEnable.Checked `
            -DCParam $envResult.DCParam -ADServer $envResult.ADServer
        $txtOut.Text = $result.Output

        if ($result.SendToChanged -and $result.SendToActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Added Send-To $dl ($($envResult.Environment); users: $($result.SendToActedOn -join '+'))"
        }
        if ($result.SendAsChanged -and $result.SendAsActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Added Send-As $dl ($($envResult.Environment); users: $($result.SendAsActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add DL Access'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnRemove.Add_Click({
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
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($users.Count) user(s) from $(Get-AccessTypeSummary) on '$dl' ($($envResult.Environment))?",
            "Confirm Remove Access",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }

        $result = Invoke-DLAccessChange -DLIdentity $dl -Users $users -Verb 'Remove' `
            -DoSendTo $chkSendTo.Checked -DoSendAs $chkSendAs.Checked -AlsoEnableAuth $false `
            -DCParam $envResult.DCParam -ADServer $envResult.ADServer
        $txtOut.Text = $result.Output

        if ($result.SendToChanged -and $result.SendToActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Removed Send-To $dl ($($envResult.Environment); users: $($result.SendToActedOn -join '+'))"
        }
        if ($result.SendAsChanged -and $result.SendAsActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Removed Send-As $dl ($($envResult.Environment); users: $($result.SendAsActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove DL Access'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnClose.Add_Click({ $form.Close() })

# ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin

# Console, which opens top-left. Helper is in Common.ps1.

Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
#endregion
