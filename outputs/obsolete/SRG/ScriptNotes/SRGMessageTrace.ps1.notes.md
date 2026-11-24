# SRGMessageTrace.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    SRG (cloud-only, srgpros.com) Message Trace / Quarantine Release tool.
.DESCRIPTION
    REVISED 2026-07-22, part 2 -- no longer tied to a specific mailbox. Same
    change as the original suite's MessageTrace.ps1 (see that script's header
    for the full rationale): a trace is often about a sender, a recipient,
    both, or neither, not necessarily "everything touching one particular
    mailbox". So:
      - This script no longer calls Resolve-ToolTarget at all. It opens
        straight to the Message Trace window every time, standalone or from
        SRGAdminConsole, with no prompt and no gate.
      - The old "Direction" radio group (Sent from / Received to THIS
        mailbox) plus its single "narrow" textbox are replaced by two fully
        independent, optional fields: From (sender) and To (recipient). Fill
        in either, both, or neither.
      - -SmtpAddress is now purely an OPTIONAL convenience prefill (into the
        From field), not a requirement.

    Fully functional on its own:
        .\SRGMessageTrace.ps1
    opens directly to the Message Trace window, no prompt.

    Fully functional from SRGAdminConsole.ps1 too:
        .\SRGMessageTrace.ps1 -SmtpAddress user@srgpros.com
    opens the same window with the From field pre-filled from -SmtpAddress,
    if provided.

    Ported from the original suite's MessageTrace.ps1 -- every function is
    unchanged (Get-MessageTraceV2 / Get-MessageTraceDetailV2 /
    Get-QuarantineMessage / Release-QuarantineMessage are all
    "available only in the cloud-based service" per Microsoft Learn, so
    there was never any on-prem dependency here to remove). Edits vs. the
    original are the same mechanical ones every SRG tool script needs:
      - Dot-sources SRGCommon.ps1, not Common.ps1.
      - No -Environment / -DomainController parameters. Show-MessageTraceForm
        no longer takes -Environment either -- the one place that used it
        (the Write-CompletedLog line in the quarantine-release handler) now
        just omits the environment label entirely, since there's only one
        environment in this suite to log against.

    ADDED 2026-07-22 -- results search/filter box: a new "Filter loaded
    results" textbox + Search/Clear Filter buttons sit just above the grid.
    This does NOT re-query Exchange Online -- it's a client-side filter over
    whatever $script:CurrentTraceResults the last Trace run already pulled
    back, matching the typed text (case-insensitive substring, via -like
    "*text*") against SenderAddress, RecipientAddress, Subject, and Status.
    A new $script:CurrentTraceDisplayResults holds the currently-filtered
    (or unfiltered) subset that the grid actually renders; Update-
    TraceGridDisplay was changed to read from that variable instead of
    $script:CurrentTraceResults directly. Export to CSV was changed to
    export $script:CurrentTraceDisplayResults too, so exporting after a
    filter is applied exports what's on screen, not the full unfiltered
    result set -- this is a judgment call and can be reverted if that's not
    the desired behavior. Filter box is cleared automatically on every new
    Trace run.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21 (ported from MessageTrace.ps1)
    Revised : 2026-07-22 -- added results search/filter box (client-side,
              no new query), ported from MessageTrace.ps1. See ADDED note
              above.
    Revised : 2026-07-22, later -- no longer tied to a specific mailbox;
              removed the Resolve-ToolTarget gate and replaced the
              Sent/Received direction radio with independent optional
              From/To fields. See REVISED note above.
```

- **L78-79** No Resolve-ToolTarget call -- this tool is no longer tied to a specific mailbox. See header REVISED note (2026-07-22, later).

## function Show-MessageTraceForm

- **L305-306** Results search/filter box -- added 2026-07-22, see header ADDED note. Client-side filter over the already-fetched results, no new EXO query.
- **L389-392** Filters $script:CurrentTraceResults (the full, last-fetched result set) down to rows matching $SearchText -- case-insensitive substring match (-like is case-insensitive by default) against sender, recipient, subject, and status. Empty/whitespace search text returns everything.

## function Update-TraceGridDisplay

- **L408-410** Renders $script:CurrentTraceDisplayResults (the current filter's output), not $script:CurrentTraceResults directly -- see the results search/filter box header ADDED note.
- **L442-444** From/To are both independently optional -- see header REVISED note (2026-07-22, later). Leaving both blank traces every message in the date range.
- **L469-471** Results search/filter handlers -- added 2026-07-22, see header ADDED note. Filters the grid over the already-fetched $script:CurrentTraceResults; no new Exchange Online call happens here.

## function Find-TraceResultByKey

- **L497-501** Searches the FULL $script:CurrentTraceResults (not the filtered display subset) -- a filtered-out row is never on screen to double-click/select in the first place, but every row that IS on screen is still a member of the full set, so this lookup doesn't need to change when a filter is active.
- **L557-559** CHANGED FROM ORIGINAL: the original logged "(Get-SelectedEnvLabel)". SRG has no Environment concept, so that label is simply omitted here -- nothing else about this log line changed.
- **L572-575** Exports $script:CurrentTraceDisplayResults -- the currently filtered/displayed set, matching what's on screen in the grid -- not the full unfiltered $script:CurrentTraceResults. See header ADDED note (results search/filter box, 2026-07-22).
