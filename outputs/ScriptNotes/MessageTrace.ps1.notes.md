# MessageTrace.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Standalone Message Trace / Quarantine Release tool -- Layer 2 tool script.
.DESCRIPTION
    REVISED 2026-07-22, part 2 -- no longer tied to a specific mailbox. The
    operator found that requiring a target mailbox before the window would
    even open (via Common.ps1's Resolve-ToolTarget / Show-TargetPromptDialog)
    was more restrictive than useful -- a trace is often about a sender, a
    recipient, both, or neither (e.g. "everything quarantined today"), not
    necessarily "everything touching one particular mailbox". So:
      - This script no longer calls Resolve-ToolTarget at all. It opens
        straight to the Message Trace window every time, standalone or from
        AdminConsole, with no prompt and no gate.
      - The old "Direction" radio group (Sent from / Received to THIS
        mailbox) plus its single "narrow" textbox are replaced by two fully
        independent, optional fields: From (sender) and To (recipient). Fill
        in either, both, or neither -- Get-MessageTraceV2's SenderAddress and
        RecipientAddress parameters are independently optional (see the
        Invoke-MessageTrace comment below), so leaving both blank simply
        traces every message in the date range (still bounded by the
        existing 10-day-window / 90-day-lookback rules).
      - -SmtpAddress is now purely an OPTIONAL convenience prefill (into the
        From field) for whatever AdminConsole's last lookup was, not a
        requirement -- launching with it omitted/blank is fully supported
        and opens the exact same window.
      - -DomainController is still accepted as a parameter (so existing
        AdminConsole launch lines that pass it don't break) but was already
        unused inside this script before this revision (every cmdlet here
        is EXO-only, no on-prem AD calls) and remains unused now.
    Fully functional on its own:
        .\MessageTrace.ps1
    opens directly to the Message Trace window, no prompt.

    Fully functional from AdminConsole.ps1 too:
        .\MessageTrace.ps1 -SmtpAddress user@domain.com -Environment Franchise
    opens the same window with the From field pre-filled from -SmtpAddress,
    if provided.

    Every function below is MOVED, UNCHANGED, from the original monolithic
    Mailbox-Group-Inspector.ps1 (original lines 3100-3651, region "Actions -
    message trace"), including every verification/design comment. The only
    edits anywhere in this file that touch original tool behavior are
    mechanical, both isolated to the very end of Show-MessageTraceForm:
      1. The closing call changed from
             [void]$traceForm.ShowDialog($form)
         to
             # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
             # Console, which opens top-left. Helper is in Common.ps1.
             Set-FormScreenCorner -Form $traceForm -Corner 'BottomLeft'
             [void]$traceForm.ShowDialog()
         because there is no longer a parent $form to own it -- this script
         IS the top-level window now.
      2. The one Write-CompletedLog line inside the Release-from-Quarantine
         handler called Get-SelectedEnvLabel(), which read $radHQ.Checked off
         the monolith's shared form -- a control that no longer exists here.
         Show-MessageTraceForm now takes an -Environment parameter (passed
         from $target.Environment, the same value Resolve-ToolTarget already
         resolved) and that string is logged in its place. Nothing about
         WHEN or WHY the log line fires changed, only where the environment
         label text comes from.
    Every other line of every function, including all comments, is
    byte-for-byte the original.

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
    result set -- this is a judgment call (the alternative, always
    exporting everything regardless of the filter box, seemed less
    intuitive) and can be reverted if that's not the desired behavior.
    Filter box is cleared automatically on every new Trace run.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
    Revised : 2026-07-22 -- added results search/filter box (client-side,
              no new query). See ADDED note above.
    Revised : 2026-07-22, later -- no longer tied to a specific mailbox;
              removed the Resolve-ToolTarget gate and replaced the
              Sent/Received direction radio with independent optional
              From/To fields. See REVISED note above.
```

- **L103-105** No Resolve-ToolTarget call -- this tool is no longer tied to a specific mailbox, see header REVISED note (2026-07-22, later). -SmtpAddress (if passed) is used purely as an optional From-field prefill further down.

## region Actions - message trace (moved unchanged from Mailbox-Group-Inspector.ps1 lines 3100-3651)

- **L108-123** Get-MessageTraceV2 / Get-MessageTraceDetailV2 verified against Microsoft Learn (2026-07-08): both are "available only in the cloud-based service" (no on-prem equivalent by this name -- Get-MessageTrackingLog is the on-prem cmdlet, a completely different name), so there is no collision risk the way Get-Mailbox etc. have. SenderAddress and RecipientAddress on Get-MessageTraceV2 are BOTH independently optional (Mandatory: False on both), which is what makes the Sent/Received one-sided trace below possible with a single cmdlet call -- Sent omits RecipientAddress entirely (any recipient), Received omits SenderAddress entirely (any sender). Still routed through Invoke-EXOCommand for consistency with every other EXO-flavored call in this script, per the standing rule in Get-EXOCommand's header comment. Get-MessageTraceDetailV2 is called exactly the way the operator's own working "EXO Message Trace" reference script already calls it (-MessageTraceId, -RecipientAddress, -StartDate, -EndDate) -- that usage is treated as verified per standards Section 1(a) and is not re-derived here.

## function Invoke-MessageTrace

- **L125-129** -Status accepts any of the values Get-MessageTraceV2 documents for its own -Status parameter (verified against Microsoft Learn, 2026-07-09): Delivered, Expanded, Failed, FilteredAsSpam, GettingStatus, Pending, Quarantined. Omit it entirely (as the "All Messages" dropdown option does) to get the unfiltered default.

## function Invoke-MessageTraceDetail

- **L168-191** Quarantine release -- verified against Microsoft Learn (2026-07-08): - Get-QuarantineMessage / Release-QuarantineMessage are both "available only in the cloud-based service" (Exchange PowerShell module) -- same module family as Get-MessageTraceV2, routed through Invoke-EXOCommand. - Get-QuarantineMessage -MessageId expects the full RFC822 Message-ID header value, angle brackets included (Microsoft Learn Example 3: Get-QuarantineMessage -MessageID "<...@contoso.com>"). This is a DIFFERENT value than Get-MessageTraceV2's MessageTraceId (a GUID) -- the correlation below uses the trace result's MessageId property (the RFC822 header), NOT MessageTraceId. - ASSUMPTION, not verified against a live tenant: Get-MessageTraceV2 result objects expose a MessageId property carrying that RFC822 value (this has been true of the classic Get-MessageTrace schema and is what admins have always used to cross-reference quarantine, but the operator should confirm on first use -- Find-QuarantineMatch throws a clear, specific error if it's missing rather than silently matching the wrong thing). - Release-QuarantineMessage verified via Microsoft Learn Example 2: Release-QuarantineMessage -Identity <id> -ReleaseToAll releases a single quarantined message (by its own Identity, from Get-QuarantineMessage) to ALL of its original recipients. -Confirm $false is added the same way every other write cmdlet in this script suppresses the built-in prompt, since this tool has its own MessageBox confirmation instead.

## function Invoke-QuarantineRelease

- **L219-231** Time zone display for Message Trace results -- Microsoft Learn confirms "The time stamps on the output are in UTC time format" for Get-MessageTraceV2 (and by extension Get-MessageTraceDetailV2), so everything shown to the operator needs an explicit UTC-to-local conversion. Windows time zone IDs (not IANA names) are used with [System.TimeZoneInfo], which is correct for the "Standard PowerShell" Windows environment this script's header already requires -- these IDs carry their DST rules despite the "Standard Time" naming (.NET's own documented behavior, not Exchange-specific, so not re-verified against Microsoft Learn the way Exchange cmdlets are). This ONLY changes how already-fetched results are DISPLAYED (grid, Detail box, CSV export) -- it does not change how the Search From/Through date pickers are interpreted; those are passed to Get-MessageTraceV2 exactly as before.

## function Get-SelectedTraceTimeZone

- **L246-248** Falls back to UTC (unconverted) rather than throwing, so a missing time zone on this particular machine degrades gracefully instead of breaking the whole Trace window.

## function ConvertTo-TraceDisplayTime

- **L253-258** DateTime values coming back over Exchange Online's remoting connection do not reliably keep DateTimeKind.Utc through deserialization, so this forces Kind=Utc explicitly before converting rather than trusting whatever Kind the object already reports -- converting an already-local value as if it were UTC would silently shift every timestamp by several hours, which is worse than not converting at all.

## function Show-MessageTraceForm

- **L284-297** DESIGN NOTE: the operator's pasted reference script binds results to a System.Data.DataTable (auto-generated grid columns, hidden-column pairs for the trace-id/recipient needed by Get Details). This script instead follows the SAME manual-grid convention already used for grdRules and grdCal elsewhere in this file (Columns.Add + Rows.Add, FillWeight sizing). UNLIKE grdRules/grdCal, this grid IS sortable on every column (operator request) -- sorting an unbound DataGridView physically reorders its Rows collection, so row position can no longer be used to find the matching $script:CurrentTraceResults entry after a sort. Fix: a hidden Key column (MessageTraceId|RecipientAddress) on every row, looked up by value via Find-TraceResultByKey -- the same "hidden identity column" approach grdRules already uses for exactly this reason, just adopted here too once sorting made row-index matching unsafe.
- **L301-305** Width increased 30% (780 -> 1014) on request so the DataGridView's Fill-mode columns get more room to show full sender/recipient/subject text without truncation. Right margin (28px, matching the original 780-12-740 layout) preserved -- content-area controls below widen from 740 to 974 (740 + the 234px gained) to keep that same margin.
- **L309-313** From/To fields -- replaced the old Sent/Received "Direction" radio group (tied to one anchor mailbox) 2026-07-22, see header REVISED note. Both independently optional; occupies the same 12,12-252,62 footprint the old grpDir GroupBox used, so nothing else on the form needed to move.
- **L340-344** Time zone display selector -- placed in the space opened up to the right of lblTraceNote by the 30%-wider form. Converts already-fetched results (grid, Detail box, CSV export) on selection change without re-querying Exchange Online -- see the ConvertTo-TraceDisplayTime / Get-SelectedTraceTimeZone header comment above Format-TraceDetail.
- **L357-363** Status filter -- placed in the empty space to the right of the From/To (below lblTraceNote, which only spans y=12-58) so nothing else on the form needs to move. Values verified against Microsoft Learn (Get-MessageTraceV2 -Status): Delivered, Expanded, Failed, FilteredAsSpam, GettingStatus, Pending, Quarantined. "All Messages" is not a real API value -- it just means "don't pass -Status at all", handled in the Trace click handler below.
- **L410-411** 10-day-window business rule ported from the reference script: Through can never be more than 9 days after Search From, and never in the future.
- **L436-437** Results search/filter box -- added 2026-07-22, see header ADDED note. Client-side filter over the already-fetched results, no new EXO query.
- **L480-490** sorts chronologically, not alphabetically MessageTraceId|RecipientAddress -- see below Sortable on every visible column (unlike grdRules/grdCal elsewhere in this script, which are intentionally NotSortable) -- sorting reorders the underlying Rows collection in an unbound grid like this one, so row position can no longer be used to find the matching $script:CurrentTraceResults entry. The hidden Key column below is the fix: each row carries a MessageTraceId+RecipientAddress key that double-click and Release-from-Quarantine look up by value instead of by row index, the same "hidden identity column" approach grdRules already uses for exactly this reason.
- **L529-532** Filters $script:CurrentTraceResults (the full, last-fetched result set) down to rows matching $SearchText -- case-insensitive substring match (-like is case-insensitive by default) against sender, recipient, subject, and status. Empty/whitespace search text returns everything.

## function Update-TraceGridDisplay

- **L545-549** Rebuilds the grid from the already-cached $script:CurrentTraceResults in whatever time zone cboTimeZone currently has selected. Called both right after a Trace run and whenever the time zone dropdown changes -- the latter needs no new Exchange query since the raw (UTC) data is already in hand.
- **L553-555** Renders $script:CurrentTraceDisplayResults (the current filter's output), not $script:CurrentTraceResults directly -- see the results search/filter box header ADDED note.
- **L587-589** From/To are both independently optional -- see header REVISED note (2026-07-22, later). Leaving both blank traces every message in the date range.
- **L614-616** Results search/filter handlers -- added 2026-07-22, see header ADDED note. Filters the grid over the already-fetched $script:CurrentTraceResults; no new Exchange Online call happens here.

## function Find-TraceResultByKey

- **L642-650** Looks up the raw Get-MessageTraceV2 object behind a grid row by its hidden Key column value (MessageTraceId|RecipientAddress) rather than by row position -- required now that the grid is sortable, since sorting an unbound DataGridView physically reorders Rows. Searches the FULL $script:CurrentTraceResults (not the filtered display subset) -- a filtered-out row is never on screen to double-click/select in the first place, but every row that IS on screen is still a member of the full set, so this lookup doesn't need to change when a filter is active.
- **L657-659** Double-click a row to load its detail -- same interaction pattern as grdRules elsewhere in this script (double-click for detail, no separate "open" button).
- **L709-713** CHANGED FROM ORIGINAL: was "(Get-SelectedEnvLabel)" -- that read $radHQ.Checked off the monolith's shared form. This script has no shared form, so it logs the script's own -Environment parameter value instead. Nothing else about this log line changed.
- **L726-729** Exports $script:CurrentTraceDisplayResults -- the currently filtered/displayed set, matching what's on screen in the grid -- not the full unfiltered $script:CurrentTraceResults. See header ADDED note (results search/filter box, 2026-07-22).
- **L737-739** Received is exported in whatever zone the dropdown currently shows (matching the grid on screen), via a calculated property rather than the raw UTC value -- same conversion as the grid.
- **L753-756** CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the monolith's shared parent window. This script has no parent window, so it shows itself as the top-level dialog instead. Nothing else in this function changed.
