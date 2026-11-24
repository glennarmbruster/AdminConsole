# AuditUser.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Mailbox Audit -- traces Exchange Online mailbox activity (Search-UnifiedAuditLog)
    for a single mailbox, across owner/delegate/admin actions, over a date range.
.DESCRIPTION
    Layer 2 tool script, same conventions as the other ~10 tools in this suite:
    dot-sources Common.ps1, uses Resolve-ToolTarget (prompts for a target itself
    when launched with no -SmtpAddress, same as every other tool), Tier2 entry in
    Tools.json, ObjectTypes = ['Mailbox'].
 
    -Environment / -DomainController are accepted ONLY for calling-convention
    consistency with AdminConsole.ps1, which passes the same three parameters to
    every tool it launches. This tool never uses either value for anything -- it
    is pure Exchange Online / Microsoft Purview audit, no on-prem AD dependency
    at all (unlike DLManager.ps1, GroupEditor.ps1's on-prem branch, etc.).
 
    VERIFIED against Microsoft Learn, July 2026 (Search-UnifiedAuditLog reference,
    "Search the audit log for mailbox activities in specific mailboxes," "Audit
    log activities," and "Learn about properties in Exchange audit records" --
    all fetched and read directly while building this script, not recalled from
    training data):
 
      - SCOPING TO "EVERYTHING THAT HAPPENED IN THIS MAILBOX": Search-UnifiedAuditLog's
        -UserIds parameter is generically documented as "the account of the user
        who performed the action" -- but Microsoft's own mailbox-activity search
        guidance explicitly recommends putting the AFFECTED MAILBOX (not a
        suspected actor) into that field for a regular/user mailbox search. That
        is exactly what this script does: -UserIds is set to the resolved
        mailbox's UPN, which (per that same guidance) returns activity performed
        by the owner, a delegate, or an admin against that mailbox -- matching
        the locked design goal of "owner + anyone who touched it," achieved
        server-side, no client-side actor filtering needed.
      - THIS IS DIFFERENT FROM SHARED MAILBOX AUDIT (not built yet, a separate
        planned script): Microsoft's guidance explicitly says shared-mailbox
        searches should put the mailbox's SMTP address in the FreeText/Keywords
        field instead, NOT -UserIds. Do not reuse this script's -UserIds approach
        for that one without re-reading that guidance first -- it's a different
        parameter for a reason.
      - RETENTION: Microsoft's current documented default audit retention is 180
        days (not the 90-day placeholder used earlier in design discussion),
        unless Audit (Premium) is licensed or a custom retention policy extends
        it. This script does NOT default the date range to 180 days -- see
        "DEFAULT DATE RANGE" below -- 180 days is only the outer boundary of
        what's available at all, surfaced in the warning label and in the
        zero-results status message, not used as the picker's starting value.
      - DEFAULT DATE RANGE is last 7 days, not the full 180-day retention window
        -- defaulting to the full window would likely be slow and return far
        more than anyone opening this tool actually wants on first load. 7 days
        is a "what just happened" default; the date pickers allow going back
        further (soft default, not a hard limit).
      - OPERATIONS LIST: the 31 selectable operations below are every Operation
        value with a documented, unambiguous mapping in Microsoft's "Exchange
        mailbox activities" table as of this writing. Three rows from that table
        are deliberately excluded, not overlooked:
          * "Label message as a record" has NO documented Operation value in
            Microsoft's own table -- nothing to map it to.
          * "Priority cleanup label applied on Exchange item" maps to the same
            Operation value (ApplyPriorityCleanup) as "Priority cleanup label
            applied" -- listing both would just be a duplicate checkbox.
          * "Perform search query" maps to SharepointSearchQueryInitiated --
            despite being listed in the Exchange mailbox activities table, that
            Operation value is plainly a SharePoint operation, not an Exchange
            mailbox one. Excluded as very likely a documentation placement
            quirk rather than something that belongs in a mailbox-audit list.
      - PERMISSIONS: running this requires the operator's account to have the
        "View-Only Audit Logs" or "Audit Logs" role assigned in Microsoft
        Purview role groups, in addition to the usual "connect to Exchange
        Online" assumption every tool in this suite makes. This script can't
        grant or cheaply verify that role -- it's surfaced as a warning label,
        not silently assumed.
      - INGESTION DELAY: an action happening in Exchange doesn't mean its audit
        record is searchable yet -- Microsoft's own docs note it can take up to
        30 minutes or more for an entry to appear. A just-performed action
        legitimately might not show up in a search run immediately after.
      - PAGING: uses -SessionCommand ReturnLargeSet with a fresh SessionId per
        search, repeatedly calling Search-UnifiedAuditLog until it returns zero
        additional records or the documented ReturnLargeSet ceiling of 50,000
        records is reached. If that ceiling is hit, the status message says so
        explicitly rather than silently presenting a possibly-incomplete result
        set as complete.
      - SUBJECT FILTER: Search-UnifiedAuditLog has no dedicated "subject" search
        parameter. The Subject box here is a CLIENT-SIDE partial-match filter
        applied after parsing each record's AuditData -- the operator-facing
        "N results" count is post-filter, but how long the search takes (and
        whether it hits the paging ceiling) is governed by the date range and
        Operations filter only, not by what's typed in Subject.
      - AuditData PARSING (v1 scope): each record's nested AuditData JSON is
        parsed for the fields Microsoft's own reference documents as common
        across operation types -- Operation, CreationTime, UserId, ClientIP,
        ClientInfoString, MailboxOwnerUPN, ResultStatus, Item.Subject,
        Item.ParentFolder.Name. Several operation types carry additional,
        operation-specific fields (for example, rule-change records carry a
        Parameters array showing exactly what changed) that are NOT yet broken
        out into their own grid columns -- double-click any row to see that
        record's full raw AuditData JSON, so nothing is actually hidden, just
        not all reformatted into columns yet.
      - Write-CompletedLog fires once per search that's actually run (not once
        per result row) -- an audit-the-auditor entry, since running a mailbox
        audit search is itself worth a record of who ran it and against what.
        This still only logs Task+DateTime, same as every other tool -- it does
        NOT solve the "no operator identity in the log" gap flagged earlier;
        that remains a suite-wide fix for a future revision, not something
        bolted on to just this one script.
      - CONFIDENCE SCORING / CANDIDATE-POOL ATTRIBUTION is explicitly OUT OF
        SCOPE for this build -- discussed as a deliberately later, separately-
        designed phase (needs additional log sources like Entra sign-in data,
        not just mailbox audit data) and is not attempted here. Every result
        row shows exactly what Search-UnifiedAuditLog/AuditData reports, at
        face value.
      - Not yet built: self-validating that the resolved target is actually a
        Mailbox (the way GroupEditor.ps1 validates Type for groups). Given a
        mismatched object here just returns an empty/irrelevant result set
        rather than a raw cmdlet exception, this was judged lower priority than
        it would be for a tool that performs writes -- flagged as a cheap
        future addition, not done in this pass.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-18
```

- **L133** operator cancelled the prompt

## region Constants

- **L136-138** Verified against Microsoft Learn "Audit log activities" (Exchange mailbox activities table), July 2026 -- FriendlyName -> Operation value, in the same order as the source table. See header for the three deliberately excluded rows.
- **L172** Reverse lookup (Operation value -> FriendlyName) for display in the results grid.
- **L177-179** Documented ReturnLargeSet ceiling (Search-UnifiedAuditLog reference) -- used to detect and warn on a possibly-truncated result set rather than presenting it silently as complete.

## region Data layer

- **L185-188** Pages through Search-UnifiedAuditLog for one mailbox until either it returns no further records or the documented 50,000-record ReturnLargeSet ceiling is hit. Returns a PSCustomObject: Records (raw Search-UnifiedAuditLog rows, AuditData still a JSON string) + HitCeiling (bool).

## function Invoke-MailboxAuditSearch

- **L190-194** Cancellation checks $script:AuditCancelRequested directly (set by the Cancel button's click handler) rather than a [ref] parameter -- matches the established convention elsewhere in this suite (DLManager.ps1's member-count scan loop checks $script:CountScanCancelRequested the same way) instead of introducing a new pattern just for this script.
- **L206-209** 15 iterations x 5,000/page = 75,000 -- comfortably past the documented 50,000 ReturnLargeSet ceiling, purely as a defensive loop guard in case a future service change ever behaves unexpectedly; the ceiling check below is what actually governs normal completion.
- **L228** short page -- no more data
- **L236-238** Parses one raw Search-UnifiedAuditLog row's AuditData JSON into a flat, grid- ready object. Field names verified against Microsoft's "Learn about properties in Exchange audit records" reference -- see header.

## region Main window

- **L422-424** Popup showing one row's full parsed AuditData as pretty-printed JSON -- the escape hatch for every field this script doesn't yet break into its own column (see header "AuditData PARSING (v1 scope)").
