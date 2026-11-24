#Requires -Version 5.1
<#
.SYNOPSIS
    Lookup / Add / Remove who is allowed to send messages TO an on-prem
    Distribution List -- standalone, no target lookup.
.DESCRIPTION
    One DL, optionally many users -- NOT a many-to-many cross-product tool
    like UnifiedGroupMembership.ps1. Top box is a single Distribution List;
    second box is zero or more users, one per line (commas/semicolons also
    tolerated). Lookup only needs the DL box. Add/Remove need both.

    "Send-To access" here means the DL's delivery restriction --
    Set-DistributionGroup -AcceptMessagesOnlyFromSendersOrMembers -- the
    list of senders permitted to email the DL once it's locked down, as
    confirmed with the operator before building (distinct from "Send-As",
    which SendAsDL.ps1 handles separately via Add-/Remove-ADPermission).
    This tool does NOT touch DL membership, ManagedBy/owners,
    ModeratedBy, or the Reject* delivery-restriction properties -- only
    AcceptMessagesOnlyFromSendersOrMembers (and, optionally on Add,
    RequireSenderAuthenticationEnabled -- see below).

    Lookup key: DL is looked up by Name (Section 5 exception -- mail-enabled
    object -- confirmed with the operator, who explicitly wants to type the
    DL's name, not its SMTP address; Set-/Get-DistributionGroup's -Identity
    accepts Name directly, verified against Microsoft Learn). Users are
    looked up by whatever identifies them to Get-Recipient -- typically
    firstname.lastname (SamAccountName), matching the operator's stated
    input format; no domain is appended to either box, unlike the
    PrimarySmtpAddress-keyed boxes in UnifiedGroupMembership.ps1.

    AcceptMessagesOnlyFromSendersOrMembers is a full-REPLACE property, not
    incremental like DL membership -- Add/Remove both read the DL's current
    list, resolve every current entry AND every typed user to a
    DistinguishedName via Get-Recipient (belt-and-suspenders: current-list
    entries and freshly-typed identities are compared by DN, not by string,
    so "jsmith" typed today reliably matches an entry that was originally
    added as "jsmith@expresspros.com" or a DN), compute the new full list,
    and write it back in ONE Set-DistributionGroup call. Already-satisfied
    operations (already on the list on Add; not on the list on Remove) are
    SKIP, not errors (Section 21 spirit, applied to a single target instead
    of a batch). If nothing in the list actually changes, Set-DistributionGroup
    is never called and nothing is logged.

    IMPORTANT CAVEAT surfaced to the operator, not silently handled:
    AcceptMessagesOnlyFromSendersOrMembers has NO EFFECT unless
    RequireSenderAuthenticationEnabled is also true on the same DL -- a
    well-known Exchange gotcha (the list exists but nothing is actually
    enforced). This tool never flips that setting on its own initiative;
    Lookup and every Add/Remove confirmation surface its current state, and
    an explicit, unchecked-by-default checkbox lets the operator opt into
    also enabling it during an Add, rather than the tool silently changing a
    DL's mail-flow security posture as a side effect.

    Deliberate deviation from the standard Parse/Run/Close button triad
    (Section 13): four buttons -- Lookup, Add Send-To, Remove Send-To,
    Close -- per the operator's explicit request. There's no separate Parse
    step; Lookup IS the preview (it's also run automatically, internally, at
    the start of Add/Remove to fetch the BEFORE state and confirm the DL
    resolves before anything is written).

    CMDLET COLLISION WARNING (Section 3): Get-DistributionGroup,
    Set-DistributionGroup, and Get-Recipient all exist in BOTH on-prem
    Exchange and Exchange Online. This tool calls all three by their bare
    names (like DLManager.ps1, and unlike the EXO-only tools in this suite
    that route through Invoke-EXOCommand) because it wants whatever on-prem
    Exchange Management Shell session is already loaded -- it has no EXO
    dependency at all. If an Exchange Online session is ALSO connected in
    this same PowerShell host, these bare names can resolve against the
    cloud copy instead, silently acting on the wrong DL. See the red warning
    label below, which is the Section 4 "Conflict risk" template, not the
    plain "connect before running" one every EXO-only tool in this suite
    uses.

    Environment: Franchise / HQ radio group (Section 22), no DC dropdown --
    REVISED 2026-07-21 (second same-day revision): the operator asked for
    the DC picker to be removed entirely, the same way it was removed from
    AdminConsole.ps1, with Franchise fixed to fran-vm-dc01. That fixed value
    lives in $script:DefaultDC below, LOCAL to this file (not Common.ps1 --
    see AdminConsole.ps1's header for why a shared copy in Common.ps1 caused
    a real breakage in disable-intake.ps1 earlier the same day).

    REVISED 2026-07-23 (operator request -- auto-detect HQ vs Franchise,
    drop the manual Environment radios, matching the same-day change to
    AdminConsole.ps1 and Common.ps1): DL names are unique across the HQ and
    Franchise forests (same assumption AdminConsole.ps1's mailbox/group
    auto-detect already relies on for mail addresses), so there's no reason
    to make the operator pick an environment before every Lookup/Add/Remove.
      - The Environment radio group ($grpEnv / $radFranchise / $radHQ) is
        commented out, not deleted -- same convention as the DC selector
        immediately below it.
      - New function Resolve-DLEnvironment replaces Get-SelectedEnvLabel /
        Get-CurrentADServer / Get-CurrentExchangeDCSplat (all removed, since
        they read the now-gone radio buttons): it tries
        Get-DistributionGroup -Identity <name> against HQ first (no DC
        pinned, forest-wide), then Franchise (pinned to $script:DefaultDC,
        unchanged from above) only if HQ doesn't have it. Whichever forest
        resolves the DL sets Environment + the $DCParam splat used by every
        Exchange cmdlet call in this file, exactly the same $DCParam contract
        as before -- Get-SendToState / Invoke-SendToChange themselves are
        UNCHANGED. If neither forest has the DL, Lookup/Add/Remove report
        "not found in HQ or Franchise" instead of guessing.
      - Lookup, Add, and Remove each call Resolve-DLEnvironment once, at the
        top of their own click handler, using its own typed DL name -- there
        is no shared "last looked up" state carried between buttons here
        (unlike AdminConsole.ps1's single Look Up step), since Add/Remove
        already re-run their own BEFORE snapshot every time regardless.

    BUG FIX, same revision: the first Franchise/HQ draft of this tool passed
    Get-ADServerString's return value straight into every cmdlet's
    -DomainController parameter, for both environments. That's correct for
    Franchise (a real DC FQDN) but wrong for HQ, where Get-ADServerString
    returns the bare domain name ('expresspersonnel.com') -- Exchange
    cmdlets' -DomainController parameter requires an actual reachable DC,
    and a bare domain name there throws "Active Directory error 0x52 ...
    SuitabilityDirectoryException", exactly what the operator hit running
    Lookup against an HQ DL. Fixed by building the -DomainController value
    as a splatted hashtable (originally Get-CurrentExchangeDCSplat; that
    function was folded into Resolve-DLEnvironment's $DCParam return value
    in the 2026-07-23 revision below, same empty-for-HQ shape) that's empty
    for HQ -- so HQ calls omit -DomainController entirely and resolve via
    the forest-wide view, matching Section 22's -PreferredServer guidance
    applied to Exchange cmdlets' own -DomainController parameter instead.
.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-07-21
    Revised         : 2026-07-21 1700 (removed DC dropdown, fixed the HQ
                       -DomainController bug)
    Revised         : 2026-07-23 (auto-detect HQ vs Franchise by DL name,
                       removed the manual Environment radio group -- see
                       REVISED note above)
    Run from        : on-prem Exchange Management Shell (or Exchange snap-in
                       loaded) -- see cmdlet collision warning above
    Required modules: none (EMS host)
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
# convention). No domain-append here -- unlike UnifiedGroupMembership.ps1's
# boxes, these are AD logons (firstname.lastname), not email addresses.
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

#region Actions -- on-prem Exchange, bare cmdlet names (see header collision warning)
# $DCParam is a splatted hashtable -- @{ DomainController = <fqdn> } for
# Franchise, @{} for HQ (see header BUG FIX note). Every Exchange cmdlet call
# below uses @DCParam instead of a hardcoded -DomainController so HQ never
# passes a bare domain name where a real server is required.
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

# Orchestrates one Add or Remove run against a single DL. Resolves every
# current entry AND every typed user to a DistinguishedName so the diff is
# never fooled by two different-looking identities for the same recipient
# (Section 21 SKIP logic, applied per-user against a live BEFORE snapshot).
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

#region WinForms UI -------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Send-To DL -- Allowed Senders (on-prem)"
$form.Size = New-Object System.Drawing.Size(780, 706)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Connect to on-prem Exchange ONLY (do NOT connect to Exchange Online -- Get-/Set-DistributionGroup and Get-Recipient would resolve against the wrong environment). HQ and Franchise are checked automatically per DL name."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(740, 40)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

# Environment radio group -- COMMENTED OUT 2026-07-23 (operator request):
# DL names are unique across HQ and Franchise, so Resolve-DLEnvironment below
# (an on-prem lookup, HQ first then Franchise) determines environment per-DL
# with no operator input needed. Left in place, not deleted, in case manual
# selection is ever wanted again.
# $grpEnv = New-Object System.Windows.Forms.GroupBox
# $grpEnv.Text = "Environment"
# $grpEnv.Location = New-Object System.Drawing.Point(12, 54)
# $grpEnv.Size = New-Object System.Drawing.Size(740, 50)
# $form.Controls.Add($grpEnv)
# $radFranchise = New-Object System.Windows.Forms.RadioButton
# $radFranchise.Text = "Franchise"
# $radFranchise.Location = New-Object System.Drawing.Point(10, 20)
# $radFranchise.AutoSize = $true
# $radFranchise.Checked = $true
# $grpEnv.Controls.Add($radFranchise)
# $radHQ = New-Object System.Windows.Forms.RadioButton
# $radHQ.Text = "HQ"
# $radHQ.Location = New-Object System.Drawing.Point(120, 20)
# $radHQ.AutoSize = $true
# $grpEnv.Controls.Add($radHQ)

# Domain Controller selector -- REMOVED 2026-07-21 (operator request, same
# pattern as AdminConsole.ps1): commented out, not deleted. Franchise is
# fixed to $script:DefaultDC below; HQ never pins a DC at all (see header
# BUG FIX note -- this is also what makes HQ work now).
# $lblDC = New-Object System.Windows.Forms.Label
# $lblDC.Text = "Domain Controller (Franchise only; HQ is forest-wide):"
# $lblDC.Location = New-Object System.Drawing.Point(12, 108)
# $lblDC.AutoSize = $true
# $form.Controls.Add($lblDC)
# $cboDC = New-Object System.Windows.Forms.ComboBox
# $cboDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
# $script:FranchiseDCs | ForEach-Object { [void]$cboDC.Items.Add($_) }
# $cboDC.SelectedIndex = 0
# $cboDC.Location = New-Object System.Drawing.Point(320, 105)
# $cboDC.Size = New-Object System.Drawing.Size(400, 20)
# $cboDC.Font = New-Object System.Drawing.Font("Consolas", 9)
# $form.Controls.Add($cboDC)

# Fixed Franchise DC, replacing the commented-out selector above. LOCAL to
# this file (not Common.ps1) -- see AdminConsole.ps1's header for why a
# shared copy there caused a real breakage in disable-intake.ps1. Now used
# ONLY by Resolve-DLEnvironment below.
$script:DefaultDC = 'fran-vm-dc01.franchises.expresspersonnel.com'

# Added 2026-07-23 -- see header REVISED note. Auto-detects which forest a
# DL's Name lives in: tries Get-DistributionGroup against HQ first (no DC
# pinned, forest-wide), then Franchise (pinned to $script:DefaultDC) only if
# HQ doesn't have it. Returns the same $DCParam shape Get-SendToState /
# Invoke-SendToChange already expected from Get-CurrentExchangeDCSplat, plus
# an Environment label for confirmation prompts/logging. Environment = $null
# means the DL wasn't found in either forest.
function Resolve-DLEnvironment {
    param([string]$DLIdentity)
    $errors = @()
    try {
        $null = Get-DistributionGroup -Identity $DLIdentity -ErrorAction Stop   # on-prem, HQ (no DC pinned)
        return [PSCustomObject]@{ Environment = 'HQ'; DCParam = @{}; Errors = $errors }
    } catch {
        $errors += "HQ: $($_.Exception.Message)"
    }
    try {
        $null = Get-DistributionGroup -Identity $DLIdentity -DomainController $script:DefaultDC -ErrorAction Stop   # on-prem, Franchise
        return [PSCustomObject]@{ Environment = 'Franchise'; DCParam = @{ DomainController = $script:DefaultDC }; Errors = $errors }
    } catch {
        $errors += "Franchise: $($_.Exception.Message)"
    }
    return [PSCustomObject]@{ Environment = $null; DCParam = $null; Errors = $errors }
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

$chkAlsoEnable = New-Object System.Windows.Forms.CheckBox
$chkAlsoEnable.Text = "On Add: also enable 'require sender authentication' if currently OFF (needed for this list to have any effect -- see warning)"
$chkAlsoEnable.Location = New-Object System.Drawing.Point(12, 246)
$chkAlsoEnable.Size = New-Object System.Drawing.Size(740, 22)
$form.Controls.Add($chkAlsoEnable)

$btnLookup = New-Object System.Windows.Forms.Button
$btnLookup.Text = "Lookup"
$btnLookup.Location = New-Object System.Drawing.Point(12, 274)
$btnLookup.Size = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnLookup)
$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add Send-To"
$btnAdd.Location = New-Object System.Drawing.Point(130, 274)
$btnAdd.Size = New-Object System.Drawing.Size(140, 30)
$btnAdd.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnAdd)
$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remove Send-To"
$btnRemove.Location = New-Object System.Drawing.Point(278, 274)
$btnRemove.Size = New-Object System.Drawing.Size(150, 30)
$btnRemove.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRemove)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(436, 274)
$btnClose.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnClose)

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Result:"
$lblResult.Location = New-Object System.Drawing.Point(12, 314)
$lblResult.AutoSize = $true
$form.Controls.Add($lblResult)
$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Both"
$txtOut.WordWrap = $false
$txtOut.Location = New-Object System.Drawing.Point(12, 334)
$txtOut.Size = New-Object System.Drawing.Size(740, 344)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOut)

function Set-ButtonsEnabled {
    param([bool]$Enabled)
    $btnLookup.Enabled = $Enabled
    $btnAdd.Enabled = $Enabled
    $btnRemove.Enabled = $Enabled
}

$btnLookup.Add_Click({
    Set-ButtonsEnabled -Enabled $false
    try {
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }
        $state = Get-SendToState -DLIdentity $dl -DCParam $envResult.DCParam
        $txtOut.Text = "Environment: $($envResult.Environment)`r`n`r`n" + (Format-SendToState -Heading "CURRENT STATE" -State $state)
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Lookup Send-To'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnAdd.Add_Click({
    Set-ButtonsEnabled -Enabled $false
    try {
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $users = ConvertTo-NormalizedUserList -Text $txtUsers.Text
        if ($users.Count -eq 0) { $txtOut.Text = "Enter at least one user in the second box first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }
        $dcParam = $envResult.DCParam
        $pre = Get-SendToState -DLIdentity $dl -DCParam $dcParam
        $authNote = ''
        if (-not $pre.RequireSenderAuthenticationEnabled -and -not $chkAlsoEnable.Checked) {
            $authNote = "`n`nNOTE: 'require sender authentication' is currently OFF for this DL -- this allowed-senders list will have NO EFFECT until that's turned on (check the box above, or enable it separately)."
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Add $($users.Count) user(s) to the Send-To (allowed senders) list for '$dl' ($($envResult.Environment))?$authNote",
            "Confirm Add Send-To",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $result = Invoke-SendToChange -DLIdentity $dl -Users $users -Verb 'Add' -DCParam $dcParam -AlsoEnableAuth $chkAlsoEnable.Checked
        $txtOut.Text = $result.Output
        if ($result.Changed -and $result.ActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Added Send-To $dl ($($envResult.Environment); users: $($result.ActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add Send-To'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnRemove.Add_Click({
    Set-ButtonsEnabled -Enabled $false
    try {
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $users = ConvertTo-NormalizedUserList -Text $txtUsers.Text
        if ($users.Count -eq 0) { $txtOut.Text = "Enter at least one user in the second box first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }
        $dcParam = $envResult.DCParam
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($users.Count) user(s) from the Send-To (allowed senders) list for '$dl' ($($envResult.Environment))?",
            "Confirm Remove Send-To",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $result = Invoke-SendToChange -DLIdentity $dl -Users $users -Verb 'Remove' -DCParam $dcParam -AlsoEnableAuth $false
        $txtOut.Text = $result.Output
        if ($result.Changed -and $result.ActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Removed Send-To $dl ($($envResult.Environment); users: $($result.ActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove Send-To'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnClose.Add_Click({ $form.Close() })

[void]$form.ShowDialog()
#endregion
