# SRGCommon.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Shared helper library for SRGAdminConsole.ps1 and every SRG tool script
    (SRGMailboxHealth.ps1, SRGMessageTrace.ps1, SRGMailboxPermissions.ps1,
    etc.).
.DESCRIPTION
    Dot-source this file from any SRG tool script:
        . "$PSScriptRoot\SRGCommon.ps1"

    This is a DELIBERATELY SEPARATE sibling of Common.ps1, not a shared
    dependency of it -- operator request (2026-07-21): "I don't want to
    change AdminConsole.ps1 at all." SRG is a second, unrelated engagement
    (cloud-only tenant, mail domain srgpros.com) with its own console
    (SRGAdminConsole.ps1), its own Tools.json (SRGTools.json), and its own
    Tool Manager (SRGToolManager.ps1). Nothing in this file is dot-sourced
    by, or dot-sources, the original Common.ps1 -- the two suites share no
    runtime state, so a change on one side can never surprise the other
    (this is exactly the failure mode the original Common.ps1's own header
    warns about re: disable-intake.ps1 picking up an unwanted shared
    variable). Function names below intentionally mirror Common.ps1's own
    (Format-ErrorDetail, Get-EXOCommand, Invoke-EXOCommand, Resolve-Target,
    Write-CompletedLog, Get-InboxRulesForMailbox, Get-TargetSmtp,
    Resolve-ToolTarget, Show-TargetPromptDialog) since they do the same
    conceptual job -- just never dot-source both files into the same
    PowerShell session, or the second one loaded wins and silently shadows
    the first.

    WHAT'S DIFFERENT FROM Common.ps1 (this is the whole point of this being
    a separate file rather than a parameterized version of the original):
      - No on-prem AD at all. SRG is a cloud-only tenant -- there is no
        Franchise/HQ split, no Domain Controller, no ActiveDirectory module
        dependency anywhere in this file or in any SRG tool script. Every
        SRG tool script's param block is just -SmtpAddress -- no
        -Environment, no -DomainController. Get-ADServerString,
        Show-TargetPromptDialog's Environment/DC controls, and
        Resolve-Target's on-prem Get-ADGroup branch (all present in
        Common.ps1) have no equivalent here because there's nothing for
        them to select between.
      - Mail domain is srgpros.com, not expresspros.com -- see
        $script:SrgMailDomain below. Get-TargetSmtp's "type bob.smith or
        bob.smith@srgpros.com, both resolve the same" behavior is otherwise
        identical to the original's domain-append logic.
      - Resolve-Target detects three types, not the original's three:
        Mailbox, UnifiedGroup (Microsoft 365 Group), and DistributionGroup
        (classic/cloud distribution or mail-enabled security group, via
        Get-DistributionGroup) -- all three checked via EXO only, in that
        order, each wrapped in Invoke-EXOCommand. There is no on-prem
        OnpremGroup type here; a cloud-only tenant has no synced on-prem AD
        groups to find. NOTE: Get-DistributionGroup is a KNOWN cmdlet-name
        collision risk (it exists in both on-prem Exchange and Exchange
        Online Management, exactly like Get-Mailbox) -- this is the main
        reason the Invoke-EXOCommand collision guard was kept in this file
        even though SRG itself never runs on-prem Exchange cmdlets: if an
        operator's PowerShell session ever also has an on-prem Exchange
        module/snap-in loaded (e.g. from working the original AdminConsole
        suite earlier in the same session), a bare Get-DistributionGroup
        call here could silently resolve against the wrong tenant. See
        Get-EXOCommand's own comment below for the full mechanism.
      - Write-CompletedLog writes to a DIFFERENT file --
        SRG-CompletedTasks.csv, not CompletedTasks.csv -- in the same
        E:\Scripts\M365\Logs\Completed folder, so SRG task history never
        mixes with the original suite's log. Change $file below if a
        different folder is wanted instead.
      - Show-TargetPromptDialog / Resolve-ToolTarget are simplified to just
        a single "mailbox or group email address" textbox + OK/Cancel --
        there is no Environment radio group or DC ComboBox to show, since
        there's nothing to choose. Resolve-ToolTarget's returned object has
        only a .Smtp property (no .Environment, no .ADServer) -- every SRG
        tool script reads only $target.Smtp.

    ASSUMPTIONS MADE WITHOUT EXPLICIT CONFIRMATION (flagged so they're easy
    to revisit): the Invoke-EXOCommand collision guard was kept (see above);
    SRGAdminConsole.ps1 uses a single "Tools" button section rather than the
    original's Admin Tools / Daily Scripts split, since only 3 tools exist
    today and the operator said most of the original suite's scripts won't
    carry over. Both are easy to change later and don't require touching
    this file's public function signatures.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21
```


## region Parsing

- **L89-90** "bob.smith" and "bob.smith@srgpros.com" resolve the same -- append the domain only if the operator didn't already type one.

## function Format-ErrorDetail

- **L103-104** Same level of detail on every catch block: cmdlet, script line, exception type, message, category, inner exception if present.

## function Get-EXOCommand

- **L131-140** Resolves a cmdlet name to the Exchange Online-sourced command specifically. Kept from the original Common.ps1 even though SRG never runs on-prem Exchange cmdlets itself -- Get-DistributionGroup (used by Resolve-Target below) is a documented name collision with on-prem Exchange, so if an operator's session also has an on-prem Exchange module/snap-in loaded (e.g. left over from the other, on-prem admin console), a bare call could silently resolve against the wrong environment. Every EXO-flavored call in this file and every SRG tool script goes through Invoke-EXOCommand below instead of calling the cmdlet name directly.

## function Invoke-EXOCommand

- **L157** Splat-friendly wrapper: Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }

## region Object-type resolution (EXO only -- no on-prem branch)

- **L168-173** SRGAdminConsole.ps1 uses this to detect Mailbox / UnifiedGroup / DistributionGroup so it knows which tool buttons apply. Individual SRG tool scripts do NOT need to call this themselves -- each one already assumes its own target type (all three shipped so far assume Mailbox); if handed the wrong type, the underlying Get-Mailbox call in that tool simply fails with a normal Format-ErrorDetail message.

## region Inbox rule read

- **L207-208** Shared with SRGMailboxHealth.ps1's rule count / suspicious-rule scan, same reason Common.ps1 keeps this one function shared rather than duplicated.

## region NEW 2026-07-24 -- which tenant is actually connected

- **L221-233** Added so the console can display WHICH tenant it is pointed at, matching M365_Host_Web.ps1's top bar. "Connected" alone doesn't say connected to what, and this suite is one of several that get used in the same day. NOTE this is the LIVE tenant, deliberately distinct from the $script:*MailDomain constant above. That constant is what gets APPENDED to a bare name typed in the box; this reports what Exchange Online is really connected to. They should normally agree -- and when they don't, that is precisely the mistake worth seeing, because the box would be resolving names into a tenant you are not connected to. Cached after the first success so repeated calls cost nothing. Returns $null instead of throwing when nothing is connected; the caller shows "unknown".

## function Get-ConnectedTenantDomain

- **L245** Not connected, or lookup failed -- not fatal. Deliberately swallowed.

## function Clear-ConnectedTenantDomainCache

- **L251** Needed if the operator reconnects to a different tenant mid-session.

## function Write-CompletedLog

- **L258-260** Separate file from the original suite's CompletedTasks.csv -- SRG is a different tenant/engagement, so its task history is tracked separately rather than interleaved with the original console's log.

## region "Enter a target" prompt for standalone SRG tool launches

- **L278-282** Mirrors Common.ps1's Show-TargetPromptDialog / Resolve-ToolTarget, but with the Environment radios and DC ComboBox removed entirely -- there is nothing for them to select in a cloud-only, single-tenant suite. Every SRG tool script's param block is just -SmtpAddress, and the object this returns carries only .Smtp.

## function Show-TargetPromptDialog

- **L345-349** Every SRG tool script's param block gives it -SmtpAddress. When SRGAdminConsole launches the tool, that value arrives pre-populated and this returns immediately with no UI. When a tool is double-clicked directly (or -SmtpAddress is missing/blank), this pops Show-TargetPromptDialog so the operator can fill it in.

## function Resolve-ToolTarget

- **L358** operator cancelled
