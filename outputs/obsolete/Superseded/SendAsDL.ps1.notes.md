# SendAsDL.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
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
```


## region Parsing

- **L168-170** Tolerates newline, comma, and semicolon as separators (Section 21 convention). No domain-append -- these are AD logons (firstname.lastname), not email addresses.

## region Actions -- on-prem Exchange ADPermission cmdlets + ActiveDirectory

- **L185-186** Pulls the SamAccountName back out of whatever form Get-ADPermission reports an existing grantee in (commonly "DOMAIN\sam", sometimes just "sam").

## function Get-SendAsState

- **L202-203** Explicit (non-inherited) Send-As grants only -- see header for why inherited ACEs and NT AUTHORITY\SELF are excluded.

## function Add-SendAsGrant

- **L236** ExtendedRights "Send-As" (hyphenated) -- see header KNOWN SYNTAX GOTCHA.

## function Remove-SendAsGrant

- **L245-250** Orchestrates one Add or Remove run against a single DL, one Add-/Remove-ADPermission call per resolved user (naturally incremental, unlike SendToDL.ps1's full-list replace). SKIP logic per header. $ADServer feeds Get-ADUser -Server (genuine ActiveDirectory module, bare domain name is fine); $DCParam is the splat used for the three Exchange *-ADPermission calls (empty for HQ -- see header BUG FIX note).

## region WinForms UI

- **L334-354** Environment radio group -- COMMENTED OUT 2026-07-23 (operator request): DL names are unique across HQ and Franchise, so Resolve-DLEnvironment below (an on-prem lookup, HQ first then Franchise) determines environment per-DL with no operator input needed. Left in place, not deleted, in case manual selection is ever wanted again. $grpEnv = New-Object System.Windows.Forms.GroupBox $grpEnv.Text = "Environment" $grpEnv.Location = New-Object System.Drawing.Point(12, 54) $grpEnv.Size = New-Object System.Drawing.Size(740, 50) $form.Controls.Add($grpEnv) $radFranchise = New-Object System.Windows.Forms.RadioButton $radFranchise.Text = "Franchise" $radFranchise.Location = New-Object System.Drawing.Point(10, 20) $radFranchise.AutoSize = $true $radFranchise.Checked = $true $grpEnv.Controls.Add($radFranchise) $radHQ = New-Object System.Windows.Forms.RadioButton $radHQ.Text = "HQ" $radHQ.Location = New-Object System.Drawing.Point(120, 20) $radHQ.AutoSize = $true $grpEnv.Controls.Add($radHQ)
- **L356-372** Domain Controller selector -- REMOVED 2026-07-21 (operator request, same pattern as AdminConsole.ps1/SendToDL.ps1): commented out, not deleted. Franchise is fixed to $script:DefaultDC below; HQ never pins a DC at all (see header BUG FIX note -- this is also what makes HQ work now). $lblDC = New-Object System.Windows.Forms.Label $lblDC.Text = "Domain Controller (Franchise only; HQ is forest-wide):" $lblDC.Location = New-Object System.Drawing.Point(12, 108) $lblDC.AutoSize = $true $form.Controls.Add($lblDC) $cboDC = New-Object System.Windows.Forms.ComboBox $cboDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList $script:FranchiseDCs | ForEach-Object { [void]$cboDC.Items.Add($_) } $cboDC.SelectedIndex = 0 $cboDC.Location = New-Object System.Drawing.Point(320, 105) $cboDC.Size = New-Object System.Drawing.Size(400, 20) $cboDC.Font = New-Object System.Drawing.Font("Consolas", 9) $form.Controls.Add($cboDC)
- **L374-377** Fixed Franchise DC, replacing the commented-out selector above. LOCAL to this file (not Common.ps1) -- see AdminConsole.ps1's header for why a shared copy there caused a real breakage in disable-intake.ps1. Now used ONLY by Resolve-DLEnvironment below.
- **L380-388** Added 2026-07-23 -- see header REVISED note. Same probe SendToDL.ps1 uses: tries Get-DistributionGroup against HQ first (no DC pinned, forest-wide), then Franchise (pinned to $script:DefaultDC) only if HQ doesn't have it. Returns BOTH shapes this file needs -- $DCParam (empty-for-HQ splat) for the three Exchange *-ADPermission calls, and $ADServer (a plain string, $script:HQDomain or $script:DefaultDC) for Get-ADUser's -Server parameter, which is a genuine ActiveDirectory module call and accepts a bare domain name fine (see header BUG FIX note for why these two shapes differ). Environment = $null means the DL wasn't found in either forest.

## function Resolve-DLEnvironment

- **L393** on-prem, HQ (no DC pinned)
- **L399** on-prem, Franchise
