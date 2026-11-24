#Requires -Version 5.1
<#
.SYNOPSIS
    Lookup / Add / Remove Send-As permission on an on-prem Distribution List
    -- standalone, no target lookup.
.DESCRIPTION
    One DL, optionally many users -- same "1 DL x N users" shape as
    SendToDL.ps1, NOT a many-to-many cross-product tool like
    UnifiedGroupMembership.ps1. Top box is a single Distribution List;
    second box is zero or more users, one per line (commas/semicolons also
    tolerated). Lookup only needs the DL box. Add/Remove need both.

    "Send-As" here means the ability to send mail that appears to come FROM
    the DL's address -- Add-/Remove-ADPermission with
    -ExtendedRights "Send-As" -- distinct from SendToDL.ps1's delivery
    restriction (who's allowed to send TO the DL). Confirmed with the
    operator before building. This tool touches ONLY the Send-As extended
    right; it never reads or writes DL membership, ManagedBy/owners,
    delivery restrictions, or any other AD permission.

    Lookup key: DL is looked up by Name (Section 5 exception -- mail-enabled
    object -- confirmed with the operator, who explicitly wants to type the
    DL's name, not its SMTP address). Add-/Remove-/Get-ADPermission's
    -Identity parameter explicitly documents accepting the object's Name or
    DistinguishedName (verified against Microsoft Learn's "Manage
    permissions for recipients" article). Users are typed as
    firstname.lastname (SamAccountName) per the operator's stated format,
    and resolved via Get-ADUser (ActiveDirectory module -- same pattern
    DLManager.ps1 already uses for its own disabled-account check) so
    Add/Remove can reliably tell whether a typed user already has the grant,
    regardless of what form (UPN, alias, DOMAIN\sam, etc.) Get-ADPermission
    happens to report an existing grantee in.

    Add-ADPermission is naturally incremental (one grant per call), unlike
    SendToDL.ps1's full-replace AcceptMessagesOnlyFromSendersOrMembers array
    -- so Add/Remove here call Add-/Remove-ADPermission once per resolved
    user rather than computing and writing a whole new list. Already-
    satisfied operations (already granted on Add; not granted on Remove) are
    SKIP, not errors, determined from a live Get-ADPermission BEFORE
    snapshot (Section 21 spirit, applied to a single target). Lookup, and
    the BEFORE/AFTER snapshots Add/Remove print, exclude inherited ACEs and
    the NT AUTHORITY\SELF entry every AD object has on itself -- both are
    noise for a delegation-tracking tool, not permissions anyone granted.
    This filtering matches Microsoft's own documented Get-ADPermission
    pattern for listing Send connector permissions (-notlike/IsInherited),
    applied here to Send-As on a DL instead.

    KNOWN SYNTAX GOTCHA, followed deliberately (Section 2): this codebase's
    own standards doc already flags Add-ADPermission -ExtendedRights "Send
    As" (with a space) as accepted syntactically but silently wrong --
    "Send-As" (hyphenated) is required. Current public Microsoft Learn
    examples for Add-/Remove-ADPermission actually show the SPACE form
    throughout, which conflicts with that entry. This script follows the
    standards doc's field-tested correction, not the generic public
    example, per the explicit instruction to treat the standards doc as
    authoritative. Worth confirming empirically against Get-ADPermission's
    own ExtendedRights output on first real use (Lookup will show exactly
    what string got stored) and updating the gotcha table if this Exchange
    version's behavior turns out to differ.

    Deliberate deviation from the standard Parse/Run/Close button triad
    (Section 13): four buttons -- Lookup, Add Send-As, Remove Send-As,
    Close -- per the operator's explicit request, mirroring SendToDL.ps1.
    No separate Parse step; Lookup IS the preview, and is also run
    automatically/internally at the start of Add/Remove for the BEFORE
    snapshot.

    Cmdlet collision note (Section 3): Add-ADPermission, Remove-ADPermission,
    and Get-ADPermission are documented as available ONLY in on-premises
    Exchange -- unlike SendToDL.ps1's Get-/Set-DistributionGroup /
    Get-Recipient, there is no cloud-copy-shadowing risk for these three
    specific cmdlets, so the warning label below is the plain "connect
    before running" form (Section 4), not the stronger "conflict risk"
    template SendToDL.ps1 uses. Get-ADUser (ActiveDirectory module, used
    only to resolve a typed user to its SamAccountName) has no Exchange
    Online equivalent to collide with either.

    Environment: Franchise / HQ radio group (Section 22), no DC dropdown --
    REVISED 2026-07-21 (third same-day revision, matching SendToDL.ps1's
    identical fix): the operator asked for the DC picker to be removed
    entirely, the same way it was removed from AdminConsole.ps1 and
    SendToDL.ps1, with Franchise fixed to fran-vm-dc01. That fixed value
    lives in $script:DefaultDC below, LOCAL to this file (not Common.ps1 --
    see AdminConsole.ps1's header for why a shared copy in Common.ps1 caused
    a real breakage in disable-intake.ps1 earlier the same day).

    BUG FIX, same revision: the earlier Franchise/HQ-picker draft of this
    tool passed Get-ADServerString's return value straight into
    Get-/Add-/Remove-ADPermission's -DomainController parameter for both
    environments. That's correct for Franchise (a real DC FQDN) but wrong
    for HQ, where Get-ADServerString returns the bare domain name
    ('expresspersonnel.com') -- despite the "AD" in their names,
    Get-ADPermission/Add-ADPermission/Remove-ADPermission are EXCHANGE
    cmdlets (see header cmdlet collision note below), and Exchange cmdlets'
    -DomainController parameter requires an actual reachable DC, not a bare
    domain name -- the same "Active Directory error 0x52 ...
    SuitabilityDirectoryException" the operator hit in SendToDL.ps1's
    equivalent bug. Fixed the same way: -DomainController for the three
    *-ADPermission calls is now built as a splatted hashtable (originally
    Get-CurrentExchangeDCSplat; folded into Resolve-DLEnvironment's $DCParam
    return value in the 2026-07-23 revision below, same empty-for-HQ shape)
    that's empty for HQ, so HQ calls omit -DomainController entirely and
    resolve via the forest-wide view. This does NOT apply to Get-ADUser's own
    -Server parameter (genuinely the ActiveDirectory module, not Exchange) --
    that parameter accepts a bare domain name fine and now uses
    Resolve-DLEnvironment's $ADServer return value (originally the plain
    Get-CurrentADServer string, same shape) instead.

    Prior revision history: an even earlier draft pinned everything to the
    single fixed $script:DefaultDC AdminConsole.ps1 uses for its own
    Mailbox/Group lookup -- that turned out to be specific to AdminConsole,
    not a suite-wide policy, since these DLs can live in either environment.
    Replaced with a Franchise/HQ + DC picker, which then hit the
    -DomainController bug above and was replaced in turn by the current
    no-dropdown design.

    REVISED 2026-07-23 (operator request -- auto-detect HQ vs Franchise,
    drop the manual Environment radios, matching the same-day change to
    AdminConsole.ps1, Common.ps1, and SendToDL.ps1): DL names are unique
    across the HQ and Franchise forests, so there's no reason to make the
    operator pick an environment before every Lookup/Add/Remove.
      - The Environment radio group ($grpEnv / $radFranchise / $radHQ) is
        commented out, not deleted -- same convention as the DC selector
        immediately below it.
      - New function Resolve-DLEnvironment replaces Get-SelectedEnvLabel /
        Get-CurrentADServer / Get-CurrentExchangeDCSplat (all removed, since
        they read the now-gone radio buttons): it tries
        Get-DistributionGroup -Identity <name> against HQ first (no DC
        pinned, forest-wide), then Franchise (pinned to $script:DefaultDC,
        unchanged from above) only if HQ doesn't have it -- same probe
        SendToDL.ps1 uses, reused here rather than invented separately,
        since "does a DL with this name exist in this forest" doesn't need
        to be Send-As-specific. Whichever forest resolves the DL sets
        Environment, the $DCParam splat the three *-ADPermission calls
        already expected, AND the plain $ADServer string Get-ADUser's
        -Server parameter expects (empty-for-HQ DCParam vs. a real HQ
        domain string for -Server are different shapes, per the header BUG
        FIX note -- Resolve-DLEnvironment returns both). If neither forest
        has the DL, Lookup/Add/Remove report "not found in HQ or Franchise"
        instead of guessing.
      - Lookup, Add, and Remove each call Resolve-DLEnvironment once, at the
        top of their own click handler, same as SendToDL.ps1.
.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-07-21
    Revised         : 2026-07-21 1700 (removed DC dropdown, fixed the
                       *-ADPermission -DomainController HQ bug)
    Revised         : 2026-07-23 (auto-detect HQ vs Franchise by DL name,
                       removed the manual Environment radio group -- see
                       REVISED note above)
    Run from        : on-prem Exchange Management Shell (or Exchange snap-in
                       loaded), with the ActiveDirectory (RSAT) module
                       available for Get-ADUser
    Required modules: ActiveDirectory
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
# not email addresses.
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

#region Actions -- on-prem Exchange ADPermission cmdlets + ActiveDirectory -----
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
    # Explicit (non-inherited) Send-As grants only -- see header for why
    # inherited ACEs and NT AUTHORITY\SELF are excluded.
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

# Orchestrates one Add or Remove run against a single DL, one
# Add-/Remove-ADPermission call per resolved user (naturally incremental,
# unlike SendToDL.ps1's full-list replace). SKIP logic per header. $ADServer
# feeds Get-ADUser -Server (genuine ActiveDirectory module, bare domain name
# is fine); $DCParam is the splat used for the three Exchange *-ADPermission
# calls (empty for HQ -- see header BUG FIX note).
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

#region WinForms UI -------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Send-As DL -- Permission Delegation (on-prem)"
$form.Size = New-Object System.Drawing.Size(780, 702)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Connect to on-prem Exchange before running. HQ and Franchise are checked automatically per DL name."
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
# pattern as AdminConsole.ps1/SendToDL.ps1): commented out, not deleted.
# Franchise is fixed to $script:DefaultDC below; HQ never pins a DC at all
# (see header BUG FIX note -- this is also what makes HQ work now).
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

# Added 2026-07-23 -- see header REVISED note. Same probe SendToDL.ps1 uses:
# tries Get-DistributionGroup against HQ first (no DC pinned, forest-wide),
# then Franchise (pinned to $script:DefaultDC) only if HQ doesn't have it.
# Returns BOTH shapes this file needs -- $DCParam (empty-for-HQ splat) for
# the three Exchange *-ADPermission calls, and $ADServer (a plain string,
# $script:HQDomain or $script:DefaultDC) for Get-ADUser's -Server parameter,
# which is a genuine ActiveDirectory module call and accepts a bare domain
# name fine (see header BUG FIX note for why these two shapes differ).
# Environment = $null means the DL wasn't found in either forest.
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

$btnLookup = New-Object System.Windows.Forms.Button
$btnLookup.Text = "Lookup"
$btnLookup.Location = New-Object System.Drawing.Point(12, 248)
$btnLookup.Size = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnLookup)
$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add Send-As"
$btnAdd.Location = New-Object System.Drawing.Point(130, 248)
$btnAdd.Size = New-Object System.Drawing.Size(140, 30)
$btnAdd.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnAdd)
$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remove Send-As"
$btnRemove.Location = New-Object System.Drawing.Point(278, 248)
$btnRemove.Size = New-Object System.Drawing.Size(150, 30)
$btnRemove.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRemove)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(436, 248)
$btnClose.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnClose)

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Result:"
$lblResult.Location = New-Object System.Drawing.Point(12, 288)
$lblResult.AutoSize = $true
$form.Controls.Add($lblResult)
$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Both"
$txtOut.WordWrap = $false
$txtOut.Location = New-Object System.Drawing.Point(12, 308)
$txtOut.Size = New-Object System.Drawing.Size(740, 346)
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
        $state = Get-SendAsState -DLIdentity $dl -DCParam $envResult.DCParam
        $txtOut.Text = "Environment: $($envResult.Environment)`r`n`r`n" + (Format-SendAsState -Heading "CURRENT STATE" -State $state)
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Lookup Send-As'
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
        $adServer = $envResult.ADServer
        $dcParam = $envResult.DCParam
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Add Send-As permission for $($users.Count) user(s) on '$dl' ($($envResult.Environment))?",
            "Confirm Add Send-As",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $result = Invoke-SendAsChange -DLIdentity $dl -Users $users -Verb 'Add' -ADServer $adServer -DCParam $dcParam
        $txtOut.Text = $result.Output
        if ($result.Changed -and $result.ActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Added Send-As $dl ($($envResult.Environment); users: $($result.ActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add Send-As'
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
        $adServer = $envResult.ADServer
        $dcParam = $envResult.DCParam
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove Send-As permission for $($users.Count) user(s) on '$dl' ($($envResult.Environment))?",
            "Confirm Remove Send-As",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $result = Invoke-SendAsChange -DLIdentity $dl -Users $users -Verb 'Remove' -ADServer $adServer -DCParam $dcParam
        $txtOut.Text = $result.Output
        if ($result.Changed -and $result.ActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Removed Send-As $dl ($($envResult.Environment); users: $($result.ActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove Send-As'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnClose.Add_Click({ $form.Close() })

[void]$form.ShowDialog()
#endregion
