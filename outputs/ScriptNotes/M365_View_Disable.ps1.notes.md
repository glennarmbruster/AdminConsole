# M365_View_Disable.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Disable / termination view for M365_Host_Web.ps1. DESTRUCTIVE.
.DESCRIPTION
    ADDED 2026-07-24 as part of the single-process host rewrite. Replaces
    standalone M365_Disable_Web.ps1.

    Contains no logic -- ConvertTo-NormalizedUserList,
    Get-OneDrivePersonalSiteUrl, and Disable-M365User all live in
    M365Logic.psm1. Per user, Disable-M365User: hides from address lists and
    converts to a shared mailbox, optionally grants a manager Full Access
    (never Send-As), and optionally grants the manager site-collection admin
    on the user's OneDrive.

    CONFIRMATION IS THIS VIEW'S JOB, and it matters more here than anywhere
    else in the toolkit. Disable-M365User does exactly what it's told
    immediately, with no prompt, so it can be driven by a scheduled job --
    which means the confirm() below listing every affected user is the only
    guard against a misclick converting real mailboxes to shared. Do not
    remove it.

    SHAREPOINT TENANT PREFIX is an operator-supplied field, not a constant.
    The original ANZ tool hardcoded 'expressprosau', and per that tool's own
    header it is NOT derivable from the mail domain by any reliable rule --
    naively stripping the TLD from 'expresspros.com.au' gives 'expresspros',
    which was confirmed wrong for that exact tenant. Guessing would point a
    site-collection-admin grant at the wrong SharePoint site, so it's asked
    for instead. Only needed when the OneDrive step is used; the module
    fails that step with a clear message (and still completes the others) if
    it's missing.

    LICENSE REMOVAL AND SIGN-IN BLOCK ARE NOT AUTOMATED, matching the
    original's final state. CORRECTED 2026-07-24 -- earlier notes in this
    codebase (including this file) said the AzureAD module was "permanently
    dead," which overstates it. What's true: the module is DEPRECATED
    (announced 2024-03-30), and cmdlets that call the retired Azure AD Graph
    endpoint fail -- which is precisely what broke Get-AzureADGroup on a real
    run. But Connect-AzureAD itself still authenticates, confirmed by the
    operator's daily use. The reason these two steps stay manual is therefore
    NOT "the module is gone"; it's that the specific group/user-write cmdlets
    needed are the ones hitting the retired endpoint, and the Microsoft.Entra
    alternative needed tenant-admin consent that wasn't available. The run
    reports both as NOT ATTEMPTED, and this view now renders a per-user Entra
    deep link for each so finishing them by hand is two clicks.
```


## function renderFollowUp

- **L240-245** REVISED 2026-07-24 -- returns per-user objects (with the Entra deep link already built server-side) instead of a flat list of addresses, so the page can render the two manual follow-up steps as real links. Building the URL here rather than in JS keeps every portal URL in one place: M365Logic's Get-M365PortalUrl.


## Embedded HTML / JavaScript comments

- Follow-up actions, populated after a run. The two steps this tool deliberately does not automate get a per-user deep link each, so finishing them is two clicks instead of a hunt through the portal.
- Renders the two manual follow-up steps as per-user deep links. Built
- from the Entra object GUID the server captured during the run (from
- Get-Mailbox's ExternalDirectoryObjectId), so each link lands on that
- specific user's page rather than dumping the operator in the users list.
- When a GUID couldn't be read, the server already substituted the
- all-users URL, so the link still works -- it's just less direct, and the
- row says so.
