# SendToDL.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
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
```


## region Parsing

- **L147-149** Tolerates newline, comma, and semicolon as separators (Section 21 convention). No domain-append here -- unlike UnifiedGroupMembership.ps1's boxes, these are AD logons (firstname.lastname), not email addresses.

## region Actions -- on-prem Exchange, bare cmdlet names (see header collision warning)

- **L164-167** $DCParam is a splatted hashtable -- @{ DomainController = <fqdn> } for Franchise, @{} for HQ (see header BUG FIX note). Every Exchange cmdlet call below uses @DCParam instead of a hardcoded -DomainController so HQ never passes a bare domain name where a real server is required.

## function Get-SendToState

- **L170** on-prem
- **L174** on-prem

## function Format-SendToState

- **L207-210** Orchestrates one Add or Remove run against a single DL. Resolves every current entry AND every typed user to a DistinguishedName so the diff is never fooled by two different-looking identities for the same recipient (Section 21 SKIP logic, applied per-user against a live BEFORE snapshot).

## function Invoke-SendToChange

- **L227** on-prem
- **L272** on-prem
- **L275** on-prem

## region WinForms UI

- **L303-323** Environment radio group -- COMMENTED OUT 2026-07-23 (operator request): DL names are unique across HQ and Franchise, so Resolve-DLEnvironment below (an on-prem lookup, HQ first then Franchise) determines environment per-DL with no operator input needed. Left in place, not deleted, in case manual selection is ever wanted again. $grpEnv = New-Object System.Windows.Forms.GroupBox $grpEnv.Text = "Environment" $grpEnv.Location = New-Object System.Drawing.Point(12, 54) $grpEnv.Size = New-Object System.Drawing.Size(740, 50) $form.Controls.Add($grpEnv) $radFranchise = New-Object System.Windows.Forms.RadioButton $radFranchise.Text = "Franchise" $radFranchise.Location = New-Object System.Drawing.Point(10, 20) $radFranchise.AutoSize = $true $radFranchise.Checked = $true $grpEnv.Controls.Add($radFranchise) $radHQ = New-Object System.Windows.Forms.RadioButton $radHQ.Text = "HQ" $radHQ.Location = New-Object System.Drawing.Point(120, 20) $radHQ.AutoSize = $true $grpEnv.Controls.Add($radHQ)
- **L325-341** Domain Controller selector -- REMOVED 2026-07-21 (operator request, same pattern as AdminConsole.ps1): commented out, not deleted. Franchise is fixed to $script:DefaultDC below; HQ never pins a DC at all (see header BUG FIX note -- this is also what makes HQ work now). $lblDC = New-Object System.Windows.Forms.Label $lblDC.Text = "Domain Controller (Franchise only; HQ is forest-wide):" $lblDC.Location = New-Object System.Drawing.Point(12, 108) $lblDC.AutoSize = $true $form.Controls.Add($lblDC) $cboDC = New-Object System.Windows.Forms.ComboBox $cboDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList $script:FranchiseDCs | ForEach-Object { [void]$cboDC.Items.Add($_) } $cboDC.SelectedIndex = 0 $cboDC.Location = New-Object System.Drawing.Point(320, 105) $cboDC.Size = New-Object System.Drawing.Size(400, 20) $cboDC.Font = New-Object System.Drawing.Font("Consolas", 9) $form.Controls.Add($cboDC)
- **L343-346** Fixed Franchise DC, replacing the commented-out selector above. LOCAL to this file (not Common.ps1) -- see AdminConsole.ps1's header for why a shared copy there caused a real breakage in disable-intake.ps1. Now used ONLY by Resolve-DLEnvironment below.
- **L349-355** Added 2026-07-23 -- see header REVISED note. Auto-detects which forest a DL's Name lives in: tries Get-DistributionGroup against HQ first (no DC pinned, forest-wide), then Franchise (pinned to $script:DefaultDC) only if HQ doesn't have it. Returns the same $DCParam shape Get-SendToState / Invoke-SendToChange already expected from Get-CurrentExchangeDCSplat, plus an Environment label for confirmation prompts/logging. Environment = $null means the DL wasn't found in either forest.

## function Resolve-DLEnvironment

- **L360** on-prem, HQ (no DC pinned)
- **L366** on-prem, Franchise
