# DLManager.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Distribution List Manager -- scan, browse, and edit on-prem Exchange
    distribution groups, with delegation guardrails.
.DESCRIPTION
    ON-PREM EXCHANGE ONLY. Unlike the rest of this tool suite, this script
    never calls Get-EXOCommand / Invoke-EXOCommand -- it calls
    Get-DistributionGroup, Get-DistributionGroupMember,
    Add-DistributionGroupMember, and Remove-DistributionGroupMember directly
    by name. Those Invoke-EXOCommand wrappers exist elsewhere in this suite
    specifically to force resolution to the CLOUD-sourced copy of a same-
    named cmdlet when an on-prem module might also be loaded (see
    Get-EXOCommand's header comment in Common.ps1) -- here the goal is the
    exact opposite: this tool wants whatever on-prem Exchange Management
    Shell session/snap-in is already loaded, and has no Exchange Online
    dependency at all. Like the original monolith's own assumption
    ("Connect to Exchange Online before running"), this script assumes the
    operator already has an on-prem Exchange Management Shell session open
    (or the Exchange snap-in loaded) before running it.

    Dot-sources Common.ps1 for Format-ErrorDetail, Write-CompletedLog, the
    Franchise DC list / Get-ADServerString, and Get-TargetSmtp -- but does
    NOT use Resolve-ToolTarget, because this tool isn't scoped to one
    mailbox or group the way the other ~10 tool scripts in this suite are.
    It's a standalone utility, launched directly by the operator, in the
    same spirit as ToolManager.ps1 -- not wired into Tools.json / launched
    from AdminConsole.ps1's per-object button flow. (It CAN be added as a
    Tools.json entry later if you want a button for it too; nothing here
    prevents that, it just isn't done by default since it doesn't target a
    single resolved object.)

    The Environment / Domain Controller picker on the main window is used
    ONLY for the two AD-side (not Exchange-side) lookups this tool makes:
    checking whether a member's AD account is disabled (Get-ADUser) before
    letting it be added, and checking whether a DL is synced to Azure AD via
    AAD Connect (Get-ADGroup -Properties mS-DS-ConsistencyGuid). Every
    Exchange cmdlet call (Get-DistributionGroup, Add/Remove-
    DistributionGroupMember, etc.) uses whatever on-prem Exchange session is
    already active and does not depend on this picker at all.

    KNOWN-UNVERIFIED / FIRST-PASS ITEMS (flagged per standards Section 2 --
    "if verification is unavailable, clearly say so before delivery"):
      - "Synced (AAD Connect)" detection uses the presence of a non-empty
        mS-DS-ConsistencyGuid attribute on the underlying AD group as its
        indicator. This is a well-corroborated community convention (AAD
        Connect populates this attribute, by default since sync engine
        v1.5.18.0, as the sourceAnchor for groups in scope of sync) but it
        is NOT a single official Microsoft Learn "this property always means
        synced" reference the way Get-DistributionGroup's own parameters
        are. Treat a "Synced" label as a strong indicator to double-check
        against your AAD Connect sync scope config on first real use, not
        as an unconditional guarantee.
      - "Removing the last owner" is the one guardrail from the original
        request that has no corresponding action in this scaffold yet --
        the membership actions built here are Add / Remove / Bulk Add /
        Bulk Remove / Export MEMBER, not an owner-editing UI (on-prem DL
        owners are the ManagedBy property, a separate field from
        membership). Test-RemovingLastOwner is included and ready to wire
        up, but nothing currently calls it. Add an "Edit Owners" action to
        Show-DLEditorForm and call it there when that's wanted.
      - Get-DistributionGroup -ResultSize Unlimited is fast (it does not
        enumerate members), but a per-DL member COUNT is not a property
        Get-DistributionGroup exposes -- getting it requires
        Get-DistributionGroupMember per group, which is exactly the kind of
        "iterate everything to get one number" scan that already caused a
        documented UI freeze elsewhere in this suite (see
        Get-MailboxDLMemberships's header in Common.ps1/DLMemberships.ps1
        for that history). This script applies the same lesson: Get-
        DLInventory returns the fast metadata only, with MemberCount blank;
        Start-MemberCountScan fills MemberCount in afterward, one DL at a
        time, with a visible status label and a Cancel button -- same
        progress/cancel shape as Show-DLMembershipsForm.
    REVISED 2026-07-11 (design discussion): a full inventory scan (Get-
    DistributionGroup -ResultSize Unlimited, plus one Get-ADGroup call per DL
    for the Source label) is the expensive part of opening this tool, and DL
    membership/metadata genuinely doesn't change often enough to justify
    paying that cost every single time. 'Scan All Distribution Groups' now
    also writes its result to a JSON cache file (-CacheFilePath, default
    '.\DLInventoryCache.json' next to this script) with a timestamp. On
    startup, if that cache file exists and parses, it's loaded automatically
    instead of starting from an empty grid -- no live Exchange call happens
    until 'Scan All Distribution Groups' is clicked, which always does a
    live rescan and overwrites the cache. If member counts had already been
    loaded before a scan/rescan, they're cached too; a scan alone (before
    'Load Member Counts' is clicked) caches with MemberCount still blank,
    same as it displays in the grid.
    IMPORTANT -- this cache is read-only convenience for the BROWSE/SEARCH
    grid, never for editing: Show-DLEditorForm always re-pulls membership and
    re-runs every guardrail check live when it opens, regardless of whether
    the DL row it opened from came from a fresh scan or a week-old cache. A
    stale DisplayName or member count in the grid is a minor inconvenience;
    a stale guardrail check would not be.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-11
```


## region Constants

- **L110-112** Standard Exchange RecipientType values (Microsoft Learn / well-established, already relied on elsewhere in this suite -- e.g. Resolve-Target's GroupCategory check in Common.ps1) used to classify each member row.

## function Get-DLInventory

- **L120-122** Fast pass: metadata only, no per-DL member enumeration. MemberCount is deliberately left $null here -- see header notes -- and is filled in afterward by Start-MemberCountScan.
- **L135-136** filled in by Start-MemberCountScan filled in by Update-DLSourceLabels

## function Get-DLSourceLabel

- **L142-144** See header "KNOWN-UNVERIFIED" note -- mS-DS-ConsistencyGuid presence is a strong, community-corroborated indicator of AAD Connect sync scope, not a single-source-of-truth Microsoft Learn property.

## function Get-DLMemberDetail

- **L161-166** Returns display-ready member rows. IsDisabledAccount is NOT computed here (it needs a live AD lookup per user -- see Test-AddingDisabledAccount) -- computing it for every existing member of every DL up front would repeat the same "scan cost" mistake flagged in the header notes. It's shown as "?" here and only resolved on demand (Refresh Disabled Status button in the editor).

## function Read-DLInventoryCache

- **L199-201** Returns $null (not a throw) on anything short of a fully-readable, non-empty cache -- callers treat "no usable cache" as equivalent to "no cache file at all" and just fall back to an empty grid.
- **L209-213** ConvertFrom-Json collapses a single-item JSON array to a bare value -- re-wrap ManagedBy per DL explicitly, same class of gotcha already documented/handled in ToolManager.ps1's Read-ToolsJson. PSCustomObject properties from ConvertFrom-Json are ordinary settable NoteProperties, so assigning back onto $_ here is safe.

## function Test-RemovingLastOwner

- **L230-232** NOT YET WIRED TO AN ACTION -- see header notes. Included so an "Edit Owners" feature can call it directly when that's built, rather than re-deriving this logic later.

## function Test-AddingDisabledAccount

- **L244-248** Live AD lookup -- only called at the moment a specific member is being added (or on-demand refresh in the editor), never during the bulk inventory/member scans. Only meaningful for user-type recipients; returns $false (no warning) for anything else without attempting a lookup that wouldn't apply.
- **L256-257** Can't confirm -- treat as "unknown" rather than silently claiming enabled. Caller decides how to present this (see Show-DLEditorForm).

## function Get-SyncedDLWarning

- **L268-270** Informational only -- editing a synced DL always happens here (this tool IS on-prem, the authoritative source for a synced object), unlike the EXO-era "view-only, edit on-prem instead" framing this replaced.

## function Get-ModeratedDLWarning

- **L286-291** Centralizes the "block on unacknowledged guardrail, otherwise confirm" pattern for every write action below, so Add/Remove/Bulk Add/Bulk Remove don't each reimplement it. $Guardrails are DANGEROUS conditions -- if any are present and Operator Mode is off, the action is blocked outright. $Warnings are informational (synced/moderated) and never block, only appear in the confirmation text.

## function Add-DLMemberBulk

- **L327-330** Collects per-item success/failure rather than stopping at the first error -- same "don't let one bad row hide the other nine" convention used throughout this suite (e.g. Get-MailboxHealthSnapshot's Errors array).

## region DL Editor window

- **L368-376** Defined BEFORE the main window below on purpose (matching every other script in this suite -- e.g. MailboxHealth.ps1 defines its Show-XForm before calling it, GroupEditor.ps1 defines Show-GroupEditorForm before its trailing call). PowerShell does hoist top-level function definitions across an entire script, so calling Show-DLEditorForm from inside the main window's button handlers below would technically still resolve even with the function defined later in the file -- but this codebase's convention is "define before use," and following it here avoids being the one file that relies on hoisting while everything else doesn't.

## function Update-MembersGrid

- **L529-533** NOTE: RecipientType is assumed 'UserMailbox' here since the operator is typing a raw identity, not selecting from a known-typed row -- this only affects whether the disabled-account guardrail check runs at all (Test-AddingDisabledAccount short-circuits for non-user types); it does not affect what actually gets added.
- **L640-641** ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin Console, which opens top-left. Helper is in Common.ps1.

## function Update-EnvironmentUI

- **L774-775** Right-click context menu on the DL grid -- matches the double-click behavior below plus a couple of read-only conveniences.

## function Update-DLGridRow

- **L794-796** Refreshes a single visible row's cells from $script:AllDLs -- used after the member-count scan updates one DL at a time, and after a source-label refresh, without a full grid rebuild.

## function Update-DLGridFromCurrent

- **L826-828** Real-time "contains" filter across Name and Email -- covers prefix/suffix matching as trivial cases of substring matching, so one filter function handles all three modes the original request called out separately.

## function Update-SearchFilter

- **L952-957** Auto-load from the inventory cache file if one exists and is readable, instead of opening to an empty grid every time. This runs once, at startup, after every control/handler above it is already wired -- 'Scan All Distribution Groups' is still the only thing that ever talks to Exchange; this block only ever reads the cache file. See header REVISED 2026-07-11 note above for the full reasoning.
- **L967-968** else: no cache file at all yet -- $lblStatus already shows the default "Not scanned yet" text set when the label was created above.
