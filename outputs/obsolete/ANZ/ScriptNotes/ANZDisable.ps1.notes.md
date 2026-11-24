# ANZDisable.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    ANZ (cloud-only, expresspros.com.au) termination/disable tool -- hides
    and converts one or more mailboxes to shared, and optionally grants a
    manager Full Access (mailbox) and/or Site Collection Admin (OneDrive).
.DESCRIPTION
    Built the same way as SRGDisable.ps1 -- standalone, multi-user WinForms
    tool, NOT routed through ANZCommon.ps1's Resolve-ToolTarget (that
    function is built for a single resolved target and would force a modal
    prompt / block launching with no target at all). Same shape as
    SendToDL.ps1/SendAsDL.ps1 in the original suite: dot-sources
    ANZCommon.ps1 for its shared helpers (Get-TargetSmtp, Format-ErrorDetail,
    Invoke-EXOCommand, Write-CompletedLog) but owns its own multi-line "one
    or more users" textbox instead of a single-target flow. Accepts
    -SmtpAddress (from ANZAdminConsole's last lookup, if any) purely to
    pre-fill the first line of that textbox -- launching with nothing
    pre-filled is fully supported, unlike every Resolve-ToolTarget-based
    tool in this suite.

    REVISED 2026-07-21 -- license removal and sign-in block, three states
    in one day (matching SRGDisable.ps1). (1) Originally included removing
    the terminated user from an Entra ID license group and blocking
    sign-in, via the AzureAD PowerShell module (Get-AzureADUser /
    Get-AzureADGroup / Remove-AzureADGroupMember / Set-AzureADUser). The
    first real run (of SRGDisable.ps1) hit:
        ERROR [Get-AzureADGroup] -- Access blocked to AAD Graph API for
        this application. https://aka.ms/AzureADGraphMigration.
    Microsoft fully retired the underlying Azure AD Graph API on
    2025-08-31 -- the AzureAD module is permanently dead, not a
    permissions/connection problem. (2) Both steps were dropped, then
    briefly restored using the Microsoft.Entra module (Connect-Entra /
    Get-EntraGroup / Get-EntraUser / Remove-EntraGroupMember /
    Set-EntraUser), a working Graph-backed alternative -- but that required
    one-time tenant-admin consent for the Connect-Entra scopes, and per the
    operator ("graph is not gonna work today") that consent isn't available
    today. (3) Both steps are dropped again -- this script no longer calls
    AzureAD, Microsoft.Entra, or Microsoft Graph in any form.

    In place of the automated steps, a completion popup appears after every
    Disable run, listing the exact users just processed and reminding the
    operator to handle both manually:
        Complete:
        Manually remove from license group and block sign-in for these
        users.
        <one user per line>

    FOUR STEPS PER USER, EACH ATTEMPTED INDEPENDENTLY (a failure in one
    step does not skip the others for that same user -- same "partial data/
    partial success is still useful" philosophy as every other tool in this
    suite, e.g. MailboxHealth's per-field try/catch):
      1+2. Hide from address lists AND convert to a shared mailbox, in one
           Set-Mailbox call (Invoke-EXOCommand-wrapped, per the suite's
           collision-guard convention for every EXO cmdlet).
      3. If a manager was entered AND the "Manager Access" selection
         includes Mailbox: Add-MailboxPermission -AccessRights FullAccess
         only. Per the operator's explicit instruction, this script NEVER
         grants Send-As -- there is no Add-RecipientPermission call
         anywhere in this file.
      4. If a manager was entered AND the "Manager Access" selection
         includes OneDrive: Set-SPOUser -IsSiteCollectionAdmin $true on the
         terminated user's OneDrive personal site, so the manager can
         browse/copy files.

    CMDLET VERIFICATION (per the standards doc -- every cmdlet below was
    checked against Microsoft Learn before this script was delivered, same
    verification pass used for SRGDisable.ps1):
      - Set-Mailbox -Type <Regular|Room|Equipment|Shared> and
        -HiddenFromAddressListsEnabled <$true|$false> -- both confirmed
        Set-Mailbox parameters; combining them in one call changes only the
        mailbox type and GAL visibility, nothing else (existing mail,
        folders, and permissions are retained across the Shared conversion
        per Microsoft's own "Convert a mailbox" documentation).
      - Add-MailboxPermission -AccessRights FullAccess -InheritanceType All
        -- reused verbatim from this suite's own ANZMailboxPermissions.ps1
        (Grant-FullAccessPermission), already a proven call in this
        codebase.
      - Set-SPOUser -Site <url> -LoginName <upn> -IsSiteCollectionAdmin
        $true -- confirmed against Microsoft Learn's Set-SPOUser reference.
        Requires SharePoint Online Management Shell (Connect-SPOService),
        NOT PnP.PowerShell -- same module choice as SRGDisable.ps1
        (operator already has SPO Management Shell set up; it has no
        announced retirement).

    UNVERIFIED / FLAGGED ASSUMPTION -- OneDrive URL construction: the
    personal-site URL for each terminated user is ALGORITHMICALLY BUILT
    from their SMTP address (every '.' and '@' replaced with '_', prefixed
    with "https://$script:OneDriveTenantHost/personal/"), not looked up via
    any API. This matches the operator's ANZ example exactly
    (andy.briggs@expresspros.com.au ->
    .../personal/andy_briggs_expresspros_com_au), but is NOT verified
    against a live tenant beyond that. NOTE the SPO tenant prefix
    ("expressprosau") is NOT simply the mail domain with the TLD stripped
    ("expresspros.com.au" would naively give "expresspros", which is
    wrong) -- it's a separate, hardcoded constant below
    ($script:SpoTenantPrefix) precisely because the two don't match here.
    If a user's OneDrive was provisioned under a different alias than their
    current primary SMTP address (e.g. after a mailbox rename), the
    constructed URL could also be stale or wrong -- Set-SPOUser will simply
    fail with a clear "site not found"-style error in that case (caught and
    reported per user, same as every other step), not silently succeed
    against the wrong site. Confirm the constructed URL shown in the Result
    box matches expectations on first real use.

    Manager Access radio buttons ("Mailbox & OneDrive" / "Mailbox Only" /
    "OneDrive Only") are disabled (greyed out) whenever the Manager textbox
    is blank -- there's nothing for them to control with no manager
    entered. "Mailbox & OneDrive" is the default selection. ONE manager
    field applies to the ENTIRE batch of users being termed in a single run
    (operator's explicit choice, 2026-07-21) -- run the tool again for a
    different manager.

    Connections required (two separate PowerShell sessions/modules, not
    one): Exchange Online (Connect-ExchangeOnline) and SharePoint Online
    Management Shell (Connect-SPOService -Url $script:SpoAdminUrl, shown in
    the warning label). This tool does not attempt to connect either one
    itself -- same "connect before running" convention as every other tool
    in this suite.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21 (built the same way as SRGDisable.ps1)
    Revised : 2026-07-21 (same day) -- dropped license-group-removal and
              sign-in-block steps (AzureAD module permanently dead).
    Revised : 2026-07-21 (same day, later) -- restored license-group-removal
              and sign-in-block using the Microsoft.Entra module instead of
              AzureAD.
    Revised : 2026-07-21 (same day, later still) -- dropped both steps again
              (Entra consent not available today); added a completion popup
              reminding the operator to do both manually. See REVISED note
              above.
    Requires: ExchangeOnlineManagement and Microsoft.Online.SharePoint.PowerShell
              (SharePoint Online Management Shell) modules, connected
              before running.
```


## region ANZ-specific constants

- **L145-147** LOCAL to this file (not ANZCommon.ps1) -- only this tool needs them, same "don't centralize a constant only one script uses" convention as AdminConsole.ps1's own $script:DefaultDC.
- **L151-153** $script:LicenseGroupName ('E5 License ANZ - Full') removed 2026-07-21 -- license removal and sign-in block are handled manually now, see header REVISED note. A completion popup reminds the operator each run.

## region Parsing

- **L157-158** Same tolerant newline/comma/semicolon splitting and case-insensitive dedup as SendToDL.ps1/SendAsDL.ps1's ConvertTo-NormalizedUserList.

## function ConvertTo-NormalizedUserList

- **L171-172** See header "UNVERIFIED / FLAGGED ASSUMPTION" note -- algorithmic, not looked up via any API.

## function Disable-AnzUser

- **L182-187** Runs all four steps for one user, each independently try/caught. Returns the per-user report text (for the Result box) and logs a compact status summary via Write-CompletedLog. License-group-removal and sign-in-block are dropped (again) -- see header REVISED note. The Disable button's click handler shows a completion popup reminding the operator to handle both manually for the whole batch.
- **L197** Steps 1+2: hide + convert to shared, in one call.
- **L209** Step 3: manager Full Access (no Send-As -- see header).
- **L227** Step 4: manager OneDrive Site Collection Admin.
- **L244-248** License-group-removal and sign-in-block (steps 5 and 6) are dropped again -- see header REVISED note. The Disable button's click handler shows a completion popup reminding the operator to handle both manually for the whole batch, so no per-user reminder is duplicated here.

## region WinForms UI

- **L320-321** Radio group is only meaningful once a manager is entered -- greyed out otherwise (operator's explicit request). Re-evaluated on every keystroke.

## function Update-ManagerAccessEnabled

- **L409-411** Automated license removal / sign-in block are out (see header REVISED note) -- remind the operator, for this exact batch, right when the run finishes.
