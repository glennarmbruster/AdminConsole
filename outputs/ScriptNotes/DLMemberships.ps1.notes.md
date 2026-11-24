# DLMemberships.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Standalone DL Memberships tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\DLMemberships.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the DL Memberships window.

    Fully functional from AdminConsole.ps1 too:
        .\DLMemberships.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the DL Memberships window
    for that mailbox.

    Get-MailboxDLMemberships and Show-DLMembershipsForm below are MOVED,
    UNCHANGED, from the original monolithic Mailbox-Group-Inspector.ps1
    (original lines 1458-1555 and 1852-1954, region "Actions - mailbox
    health"), including every revision-history comment. The only edit
    anywhere in this file that touches original tool behavior is mechanical:
    Show-DLMembershipsForm's closing call changed from
        [void]$dlForm.ShowDialog($form)
    to
        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
        # Console, which opens top-left. Helper is in Common.ps1.
        Set-FormScreenCorner -Form $dlForm -Corner 'BottomLeft'
        [void]$dlForm.ShowDialog()
    because there is no longer a parent $form to own it -- this script IS
    the top-level window now. Every other line, including all comments, is
    byte-for-byte the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
```

- **L48** operator cancelled the prompt

## function Get-MailboxDLMemberships

- **L52-99** REWRITTEN 2026-07-09c (field report: the original version iterated every DL in the tenant with Get-DistributionGroupMember and froze the UI -- see the two 2026-07-09 revision notes in the header. The operator then asked whether the EAC "Group membership" pane does the same full-tenant scan; it does not.) EAC's Group Membership pane, and Set-Mailbox et al under the hood, get this from Active Directory's own "memberOf" attribute on the user object -- a backlink that AD itself maintains automatically every time a group's "member" attribute changes, so reading it is a single indexed object lookup, not a search across every group. This function now does the same thing via Get-ADUser -Properties memberOf (both the -Properties parameter and its behavior confirmed against Microsoft Learn), matched to the mailbox via its "mail" attribute the same way Resolve-Target already matches on-prem GROUPS by mail elsewhere in this script. Cost is now one Get-ADUser call plus one Get-ADGroup call per group the user actually belongs to (typically a handful to a few dozen) -- bounded by the user's real membership count, not by how many DLs exist in the tenant. Two deliberate scope limits, both consistent with how this tool already treats on-prem groups elsewhere: - Direct membership only, not nested/recursive. memberOf lists groups the user is a DIRECT member of; if that group is itself nested inside another group, the outer group is not included. This is what EAC's Group Membership pane and the on-prem AD "Member Of" tab both show. Get-ADPrincipalGroupMembership resolves nested membership too but requires a Global Catalog and behaves differently in multi-domain forests (per Microsoft Learn) -- not used here to keep behavior simple and predictable across the Franchise/HQ environment split this tool already has. - On-prem-synced Distribution / Mail-Enabled Security groups only, via the ActiveDirectory module -- Microsoft 365 Group membership is a separate object type handled by this tool's own Group lookup path (Get-UnifiedGroupLinks), not included here. - A mailbox with no matching on-prem AD user (cloud-only mailbox, never synced) has no memberOf to read -- $result.NoADAccount is set rather than silently returning an empty list, so that case isn't mistaken for "this user is in zero groups." -ProgressAction, if supplied, is invoked once per group-DN being resolved as & $ProgressAction $currentIndex $totalCount $groupDN. -ShouldCancel, if supplied, is checked after each one; returning $true stops early and sets $result.Cancelled, keeping whatever partial results were already resolved. Both are vestigial now that this is a bounded, typically-fast operation, but kept since a handful of memberOf entries can still point at cross-domain groups that are slow to resolve individually.

## function Show-DLMembershipsForm

- **L150-157** Split out of the Health panel on 2026-07-09, then REWRITTEN 2026-07-09c to read AD's memberOf backlink instead of scanning every DL -- see the Get-MailboxDLMemberships header comment and the header .NOTES revision entries for the full story. Now a single Get-ADUser call plus one Get-ADGroup call per actual membership (typically fast), not a full-tenant scan. Scan/Cancel/progress-label plumbing is kept since a cross-domain group DN can still occasionally be slow to resolve, but this should complete quickly in the common case.
- **L251-254** CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the monolith's shared parent window. This script has no parent window, so it shows itself as the top-level dialog instead. Nothing else in this function changed.
