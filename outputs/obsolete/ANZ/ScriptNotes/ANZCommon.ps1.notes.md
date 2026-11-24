# ANZCommon.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Shared helper library for ANZAdminConsole.ps1 and every ANZ tool script
    (ANZMailboxHealth.ps1, ANZMessageTrace.ps1, ANZMailboxPermissions.ps1,
    etc.).
.DESCRIPTION
    Dot-source this file from any ANZ tool script:
        . "$PSScriptRoot\ANZCommon.ps1"

    This is a DELIBERATELY SEPARATE sibling of both the original Common.ps1
    and SRGCommon.ps1 -- same pattern as SRG (operator request, 2026-07-21:
    "same files, done the same way"), built for a second, unrelated
    cloud-only tenant (mail domain expresspros.com.au). Nothing in this file
    is dot-sourced by, or dot-sources, Common.ps1 or SRGCommon.ps1 -- all
    three suites share no runtime state, so a change in one can never
    surprise either of the others. Function names below intentionally
    mirror Common.ps1's/SRGCommon.ps1's own (Format-ErrorDetail,
    Get-EXOCommand, Invoke-EXOCommand, Resolve-Target, Write-CompletedLog,
    Get-InboxRulesForMailbox, Get-TargetSmtp, Resolve-ToolTarget,
    Show-TargetPromptDialog) since they do the same conceptual job -- just
    never dot-source more than one of these three files into the same
    PowerShell session, or whichever is loaded last wins and silently
    shadows the others.

    WHAT'S DIFFERENT FROM Common.ps1 (identical to SRGCommon.ps1's own list,
    since this is the same architecture applied to a second tenant):
      - No on-prem AD at all. ANZ is a cloud-only tenant -- no
        Franchise/HQ split, no Domain Controller, no ActiveDirectory module
        dependency anywhere in this file or in any ANZ tool script. Every
        ANZ tool script's param block is just -SmtpAddress -- no
        -Environment, no -DomainController.
      - Mail domain is expresspros.com.au, not expresspros.com or
        srgpros.com -- see $script:AnzMailDomain below. Get-TargetSmtp's
        "type bob.smith or bob.smith@expresspros.com.au, both resolve the
        same" behavior is otherwise identical to the original's and SRG's
        domain-append logic. NOTE: this domain is deliberately NOT the same
        as the original suite's $script:MailDomain ('expresspros.com') --
        confirm with the operator before assuming they're related tenants;
        as built, ANZ and the original/SRG suites are fully independent.
      - Resolve-Target detects three types, same set as SRGCommon.ps1:
        Mailbox, UnifiedGroup (Microsoft 365 Group), and DistributionGroup
        (classic/cloud distribution or mail-enabled security group, via
        Get-DistributionGroup) -- all three checked via EXO only, in that
        order, each wrapped in Invoke-EXOCommand. There is no on-prem
        OnpremGroup type here; a cloud-only tenant has no synced on-prem AD
        groups to find. NOTE: Get-DistributionGroup is a KNOWN cmdlet-name
        collision risk (it exists in both on-prem Exchange and Exchange
        Online Management, exactly like Get-Mailbox) -- this is the main
        reason the Invoke-EXOCommand collision guard was kept here too,
        even though ANZ itself never runs on-prem Exchange cmdlets: if an
        operator's PowerShell session ever also has an on-prem Exchange
        module/snap-in loaded (e.g. from working the original AdminConsole
        suite earlier in the same session), a bare Get-DistributionGroup
        call here could silently resolve against the wrong tenant. See
        Get-EXOCommand's own comment below for the full mechanism.
      - Write-CompletedLog writes to a DIFFERENT file --
        ANZ-CompletedTasks.csv, not CompletedTasks.csv or
        SRG-CompletedTasks.csv -- in the same E:\Scripts\M365\Logs\Completed
        folder, so ANZ task history never mixes with either other suite's
        log. Change $file below if a different folder is wanted instead.
      - Show-TargetPromptDialog / Resolve-ToolTarget are simplified to just
        a single "mailbox or group email address" textbox + OK/Cancel --
        there is no Environment radio group or DC ComboBox to show, since
        there's nothing to choose. Resolve-ToolTarget's returned object has
        only a .Smtp property (no .Environment, no .ADServer) -- every ANZ
        tool script reads only $target.Smtp.

    ASSUMPTIONS CARRIED OVER FROM THE SRG SUITE WITHOUT RE-CONFIRMING (same
    as SRGCommon.ps1's own list): the Invoke-EXOCommand collision guard is
    kept; ANZAdminConsole.ps1 uses a single "Tools" button section rather
    than the original's Admin Tools / Daily Scripts split, since only 3
    tools exist today. Both are easy to change later and don't require
    touching this file's public function signatures.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21 (built the same way as SRGCommon.ps1, for a second
              cloud-only tenant)
```


## region Parsing

- **L86-87** "bob.smith" and "bob.smith@expresspros.com.au" resolve the same -- append the domain only if the operator didn't already type one.

## function Format-ErrorDetail

- **L100-101** Same level of detail on every catch block: cmdlet, script line, exception type, message, category, inner exception if present.

## function Get-EXOCommand

- **L128-137** Resolves a cmdlet name to the Exchange Online-sourced command specifically. Kept from Common.ps1/SRGCommon.ps1 even though ANZ never runs on-prem Exchange cmdlets itself -- Get-DistributionGroup (used by Resolve-Target below) is a documented name collision with on-prem Exchange, so if an operator's session also has an on-prem Exchange module/snap-in loaded (e.g. left over from working the original, on-prem admin console), a bare call could silently resolve against the wrong environment. Every EXO-flavored call in this file and every ANZ tool script goes through Invoke-EXOCommand below instead of calling the cmdlet name directly.

## function Invoke-EXOCommand

- **L154** Splat-friendly wrapper: Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }

## region Object-type resolution (EXO only -- no on-prem branch)

- **L165-170** ANZAdminConsole.ps1 uses this to detect Mailbox / UnifiedGroup / DistributionGroup so it knows which tool buttons apply. Individual ANZ tool scripts do NOT need to call this themselves -- each one already assumes its own target type (all three shipped so far assume Mailbox); if handed the wrong type, the underlying Get-Mailbox call in that tool simply fails with a normal Format-ErrorDetail message.

## region Inbox rule read

- **L204-206** Shared with ANZMailboxHealth.ps1's rule count / suspicious-rule scan, same reason Common.ps1/SRGCommon.ps1 keep this one function shared rather than duplicated.

## region NEW 2026-07-24 -- which tenant is actually connected

- **L219-231** Added so the console can display WHICH tenant it is pointed at, matching M365_Host_Web.ps1's top bar. "Connected" alone doesn't say connected to what, and this suite is one of several that get used in the same day. NOTE this is the LIVE tenant, deliberately distinct from the $script:*MailDomain constant above. That constant is what gets APPENDED to a bare name typed in the box; this reports what Exchange Online is really connected to. They should normally agree -- and when they don't, that is precisely the mistake worth seeing, because the box would be resolving names into a tenant you are not connected to. Cached after the first success so repeated calls cost nothing. Returns $null instead of throwing when nothing is connected; the caller shows "unknown".

## function Get-ConnectedTenantDomain

- **L243** Not connected, or lookup failed -- not fatal. Deliberately swallowed.

## function Clear-ConnectedTenantDomainCache

- **L249** Needed if the operator reconnects to a different tenant mid-session.

## function Write-CompletedLog

- **L256-259** Separate file from both the original suite's CompletedTasks.csv and SRG's SRG-CompletedTasks.csv -- ANZ is a different tenant/engagement, so its task history is tracked separately rather than interleaved with either other console's log.

## region "Enter a target" prompt for standalone ANZ tool launches

- **L277-281** Mirrors Common.ps1's/SRGCommon.ps1's Show-TargetPromptDialog / Resolve-ToolTarget, but with the Environment radios and DC ComboBox removed entirely -- there is nothing for them to select in a cloud-only, single-tenant suite. Every ANZ tool script's param block is just -SmtpAddress, and the object this returns carries only .Smtp.

## function Show-TargetPromptDialog

- **L344-348** Every ANZ tool script's param block gives it -SmtpAddress. When ANZAdminConsole launches the tool, that value arrives pre-populated and this returns immediately with no UI. When a tool is double-clicked directly (or -SmtpAddress is missing/blank), this pops Show-TargetPromptDialog so the operator can fill it in.

## function Resolve-ToolTarget

- **L357** operator cancelled
