# Mailbox-Group-Inspector.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 1 -->

```text
.SYNOPSIS
    Looks up a mailbox or group by email address and displays/edits its
    EAC-equivalent properties (mailbox delegation + usage, or group
    delegation + membership).
.DESCRIPTION
    Single-target property inspector/editor. Operator pastes the local part
    of an email address (or a full address); the tool appends
    @expresspros.com if a domain wasn't typed, then auto-detects what kind
    of recipient it is and shows the matching panel:
      - Mailbox   : DisplayName, RecipientTypeDetails, HiddenFromAddressLists,
                    LitigationHoldEnabled (read-only sanity fields), mailbox
                    size, archive enabled + archive size, Full Access
                    grantees, Send As grantees. Full Access and Send As are
                    editable (grant/revoke one grantee at a time). An "Inbox
                    Rules" button (enabled only after a successful Mailbox
                    lookup) lists that mailbox's Inbox rules -- Name,
                    Description, Enabled -- in a table. Double-clicking a
                    row opens a separate window with the full description
                    and Enable / Disable / Delete buttons for that rule.
                    Inbox rules are read/written via Get-InboxRule /
                    Enable-InboxRule / Disable-InboxRule / Remove-InboxRule,
                    Exchange Online only, resolved through the same
                    Invoke-EXOCommand helper as everything else EXO-flavored
                    in this script (these cmdlet names also exist on-prem
                    and would otherwise collide the same way Get-Mailbox
                    did -- see FIELD-TESTED FOLLOW-UP below). Rules are
                    targeted by their own Identity property (not Name),
                    since Exchange doesn't guarantee rule names are unique
                    within a mailbox. Three more mailbox-action buttons,
                    gated the same way, all EXO cmdlets resolved through
                    Invoke-EXOCommand since every one also exists on-prem:
                    "Out of Office" (view/edit Automatic Replies via
                    Get/Set-MailboxAutoReplyConfiguration); "Calendar Perm"
                    (view/add/update/remove Calendar folder permissions via
                    Get/Add/Set/Remove-MailboxFolderPermission, common roles
                    Reviewer/Editor/Owner, calendar folder resolved via
                    Get-MailboxFolderStatistics rather than a hardcoded
                    ":\Calendar" path since that folder can be renamed, Add
                    vs. Set chosen automatically based on whether the user
                    already has an entry); "Health" (read-only -- see the
                    2026-07-09 revision note below for the full field list
                    [Potential Issues summary, delegates by name, DL
                    memberships, quota, forwarding, holds, inbox rules,
                    protocol enablement]; no writes, so no confirmation
                    dialogs or logging for this one -- has a Copy button
                    that puts the full text summary on the
                    clipboard for pasting into a ticket/email). "Message
                    Trace" (read-only mail-flow search via Get-MessageTraceV2
                    / Get-MessageTraceDetailV2, Exchange Online only, no
                    on-prem equivalent by this name so no collision risk --
                    still routed through Invoke-EXOCommand for consistency).
                    Sent/Received radio toggles the direction: Sent sets
                    SenderAddress to the looked-up mailbox and leaves
                    RecipientAddress blank (any recipient); Received is the
                    mirror image. An optional to/from box narrows the OTHER
                    side of the same query (e.g. Sent + "bob.smith@cnn.com"
                    in the box = messages from the selected mailbox TO
                    bob.smith@cnn.com). Both address parameters are
                    independently optional on Get-MessageTraceV2 -- verified
                    against Microsoft Learn before building this, since the
                    operator's own reference script always supplied both and
                    omitting one is new usage beyond what that script had
                    already validated (see standards Section 2). "Health"
                    now leads with a "Potential Issues" summary (forwarding
                    configured, litigation/retention hold, holds present,
                    suspicious inbox rules, account-disabled flag, DL
                    membership scan errors) so a tech has the likely-cause
                    checklist before escalating a ticket, followed by
                    Full Access and Send As delegates listed BY NAME (not
                    just a count), Distribution List memberships (reverse-
                    lookup via Get-MailboxDLMemberships -- iterates every DL
                    with Get-DistributionGroupMember since Exchange Online
                    has no supported single-query reverse membership lookup;
                    can be slow on tenants with many DLs), quota values next
                    to mailbox size, forwarding configuration, and the
                    existing last-logon/size/rule-count/protocol-enablement
                    fields. Date range picker enforces the same 30-day-back
                    / 10-day-window rule as the reference script. Selecting
                    a result row and double-clicking a result row calls
                    Get-MessageTraceDetailV2 exactly as the reference script
                    calls it and loads the detail box (same interaction as
                    the Inbox Rules grid -- no separate "open" button).
                    "Export to CSV" uses Export-Csv against the raw result
                    objects rather than the reference script's hand-rolled
                    CSV writer. "Release from Quarantine" looks up the
                    selected result via
                    Get-QuarantineMessage (matched by the trace result's
                    RFC822 MessageId + RecipientAddress, NOT MessageTraceId
                    -- those are different values) and, after a second
                    confirmation showing the quarantine Type/Policy/Received
                    time, releases it to all original recipients via
                    Release-QuarantineMessage -ReleaseToAll. ASSUMPTION not
                    verified against a live tenant: that Get-MessageTraceV2
                    results carry a populated MessageId property -- if not,
                    Find-QuarantineMatch throws a specific, clear error
                    rather than matching the wrong message (see comment
                    above that function).
      - Group     : DisplayName, GroupType (read-only classification -- see
                    note below), HiddenFromAddressLists, Owner(s), Members.
                    Owner and Members are editable. GroupType is NOT
                    editable -- there is no supported cmdlet to convert an
                    existing group's type after creation without direct
                    schema manipulation, so this tool deliberately does not
                    expose that as a control.
    Detection order: Get-Mailbox (EXO) -> Get-UnifiedGroup (EXO, Microsoft
    365 Group) -> Get-ADGroup (on-prem AD, Distribution / Mail-Enabled
    Security group, synced). First match wins.
    ARCHITECTURE NOTE -- why this script uses the ActiveDirectory module
    instead of an on-prem Exchange Management Shell (EMS) session for group
    work: Get-Mailbox, Get-DistributionGroup, and Get-Recipient all exist in
    BOTH on-prem Exchange and Exchange Online. If both an EMS session and an
    EXO session were connected at once (as the on-prem/EXO restore tools
    do), these ambiguous names would resolve to whichever module loaded
    last -- silently, with no error -- and this script calls nearly every
    mailbox cmdlet in the ambiguous set (Get-Mailbox, Get-MailboxPermission,
    Add/Remove-MailboxPermission, Get-MailboxStatistics). Rather than rely
    on session-load order, this script uses ONLY Exchange Online for
    mailboxes and Microsoft 365 Groups, and ONLY the ActiveDirectory (RSAT)
    module -- a completely separate cmdlet namespace with zero overlap --
    for on-prem-synced Distribution / Mail-Enabled Security groups. No
    on-prem Exchange session is used or required.
    FIELD-TESTED FOLLOW-UP (2026-07-08): not connecting an on-prem EMS
    session in THIS script does not guarantee an on-prem Exchange module is
    absent from the PowerShell host -- a leftover profile/autoload can still
    register on-prem Get-Mailbox/Get-UnifiedGroup-shaped commands, and a
    plain-name call can silently resolve to those instead of Exchange
    Online. Symptom seen in testing: a "couldn't be found on '<DC>'" error
    from what should be a cloud-only cmdlet, with a script-line number in
    the tens of thousands (an auto-generated remoting proxy, not this
    file). Fix: every EXO-flavored call in this script goes through
    Get-EXOCommand / Invoke-EXOCommand, which inspects every loaded copy of
    the cmdlet and picks the one actually sourced from the Exchange Online
    module -- never a plain unqualified cmdlet name. If it can't find an
    Exchange Online-sourced copy, it throws an error listing every copy it
    did find (name + module) so the mismatch is visible immediately instead
    of silently resolving to the wrong environment.
    KNOWN LIMITATION -- on-prem group Owner: this tool reads/writes Owner
    for on-prem-synced groups via the AD ManagedBy attribute (single-
    valued). Exchange's EAC "Group delegation > Owners" can show multiple
    co-owners for on-prem groups via msExchCoManagedByLink, which requires
    on-prem Exchange cmdlets to read or write and is therefore out of scope
    here (see architecture note above). If a group needs more than one
    owner, use on-prem Exchange Set-DistributionGroup -ManagedBy directly.
    Setting Owner in this tool REPLACES the single AD ManagedBy value; it
    does not add a co-owner.
    Lookup key: PrimarySmtpAddress (mail-enabled object exception, see
    standards Section 5). Local part is accepted alone and
    @expresspros.com is appended silently.
    Environment: Franchise or HQ, selectable. Governs which AD
    domain/forest on-prem group lookups target (Franchise DC dropdown,
    pinned; HQ forest-wide, no DC pinned). Exchange Online is single-tenant
    and is not affected by this selection -- mailbox and Microsoft 365
    Group data is identical regardless of which radio button is checked.
    This script does NOT change: mailbox type, mailbox forwarding, hidden-
    from-address-lists flags, litigation hold, or group type. It reads
    those as sanity-check context only. It does not create new mailboxes or
    groups -- lookup only, on existing objects.
    Error handling: every individual read (each mailbox/group property
    block) and every individual write (each grant/revoke/add/remove) is
    wrapped in its own try/catch. A failure in one field or one grantee is
    reported inline, by cmdlet name, script line number, exception type,
    and message -- it does not abort the rest of the lookup or a
    multi-grantee edit. See Format-ErrorDetail.
.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-07-08
    Revised         : 2026-07-10c (operator asked for both dialogs to support
                       BOTH edit triggers instead of one each. Mail Attribs
                       double-click was the only trigger; it now also has an
                       "Edit Attribute" button (12,616 Refresh / 122,616 Edit
                       Attribute / 552,616 Close), enabled only when the
                       selected grid row is editable (new
                       $grdMail.Add_SelectionChanged handler). All Attribs'
                       "Edit Attribute" button was the only trigger; its grid
                       now also supports double-click
                       ($grdAll.Add_CellDoubleClick). In both dialogs the
                       actual edit logic was refactored out of the
                       double-click/button handler bodies into one shared
                       function per dialog -- Invoke-MailAttribsEditSelectedRow
                       and Invoke-AllAttribsEditSelectedRow, both operating on
                       "whichever row is currently selected in the grid" --
                       so double-click and the button now call the exact same
                       code path rather than maintaining two copies of the
                       edit flow. Double-click sets the row's Selected state
                       first (matching the reference scripts' own
                       select-then-act idiom) before calling the shared
                       function, so both triggers behave identically down to
                       which row is treated as "the one being edited." No
                       change to what's editable in either dialog or to any
                       write/confirmation/BEFORE-AFTER/logging behavior --
                       this only changes how an edit is triggered.)
    Revised         : 2026-07-10b (added the second requested attribute-lookup
                       button, "All Attribs", right after "Mail Attribs" in
                       the same row (x=132). Ports the operator's separate
                       AD-UserAttributeEditor reference script -- a full
                       Get-ADUser -Properties * viewer/editor -- following the
                       same match-by-mail adaptation as Mail Attribs
                       (Get-AllAttribsADUser), but keeping THAT reference
                       script's own interaction model exactly, per explicit
                       instruction ("same as last -- if they are editable
                       here, they are editable in the final version"):
                         - Grid shows every populated attribute (sorted by
                           name, native column-header sort via BindingSource),
                           not a fixed list -- unlike Mail Attribs.
                         - Select a row, then click "Edit Attribute"
                           (enabled only when the selected row's attribute is
                           in the curated editable list) -- this was
                           initially the ONLY way to edit here, matching the
                           reference script's own model exactly. 2026-07-10c
                           added double-click as a second, equivalent trigger
                           (both call Invoke-AllAttribsEditSelectedRow), for
                           parity with Mail Attribs, which got the same
                           treatment in reverse (it started double-click-only
                           and gained an Edit Attribute button).
                         - Editable list copied verbatim from that reference:
                           Title, Department, Description, Company, Division,
                           Office, OfficePhone, MobilePhone, Fax, HomePhone,
                           StreetAddress, City, State, PostalCode, POBox,
                           EmployeeID, EmployeeNumber, Organization, Initials,
                           HomePage -- each a dedicated Set-ADUser string
                           parameter already Microsoft-Learn-verified there;
                           not re-derived here per standards Section 1(a).
                           Everything else (SamAccountName, UPN, DN,
                           ObjectGUID/SID, UAC, Manager, GivenName, Surname,
                           DisplayName, mail, proxyAddresses/Exchange-schema
                           attributes already covered by Mail Attribs, etc.)
                           is view-only, same as the reference.
                         - Blank new value clears the attribute
                           (Set-ADUser -Clear) rather than writing an empty
                           string, same as the reference.
                         - Format-AllAttribsValue (byte[]/SecurityIdentifier/
                           guid/datetime-aware) is a deliberately separate
                           function from Mail Attribs' simpler
                           Format-AttributeValue, since a full -Properties *
                           dump routinely returns those types (objectGUID,
                           objectSid, whenCreated, etc.) and the curated mail
                           list never does.
                       $script:AllAttribsCurrentADUser holds the loaded
                       AD-user object for this dialog instance -- same
                       cross-handler-state reasoning as
                       $script:MailAttribsCurrentADUser (read/reassigned from
                       Refresh, Edit Attribute, and the SelectionChanged
                       handler).)
    Revised         : 2026-07-10 (added a "Mail Attribs" mailbox-action
                       button -- the first of two requested attribute-lookup
                       buttons; "All Attribs" is a separate follow-up and not
                       part of this revision. Ports the on-prem AD attribute
                       grid and edit mechanics from the operator's separate
                       "Renaming Script" tool (on-prem AD name-change tool)
                       almost verbatim, per explicit instruction to match
                       that script's behavior exactly:
                         - Fixed attribute list/order/labels: Company
                           ("Office", read-only), DisplayName, name
                           (read-only -- CN/RDN, out of scope here), GivenName
                           ("FirstName"), Initials, Surname ("LastName"),
                           mail, mailNickname, msExchArchiveName,
                           PrimarySmtpAddress ("Primary SMTP" -- synthetic,
                           derived from proxyAddresses, not its own AD
                           property), proxyAddresses (read-only), targetAddress,
                           UserPrincipalName (read-only).
                         - Double-click a row to edit it, or select it and
                           click "Edit Attribute" (2026-07-10c added the
                           button alongside double-click -- both call the
                           same Invoke-MailAttribsEditSelectedRow); write
                           mechanism per attribute matches the reference
                           exactly -- dedicated Set-ADUser
                           parameters for DisplayName/GivenName/Initials/
                           Surname/mail(-EmailAddress); Set-ADUser -Replace
                           for the Exchange-schema attributes with no
                           dedicated parameter (mailNickname,
                           msExchArchiveName, targetAddress); PrimarySmtpAddress
                           rewrites proxyAddresses and, if a different primary
                           already exists, prompts Keep as Alias / Remove /
                           Cancel (Show-KeepOrRemoveOldPrimaryDialog) instead
                           of silently auto-demoting it. Blank-value rule
                           matches the reference: a populated field can't be
                           saved blank; name and PrimarySmtpAddress can never
                           be blank regardless of prior state.
                         - New "Add secondary SMTP address" box + button
                           below the grid, appending one smtp: entry to
                           proxyAddresses without touching the primary or any
                           other existing address -- same as the reference.
                         - Inline BEFORE/AFTER re-read after every write,
                           same confirmation-dialog-before-write pattern, and
                           every completed edit is logged via the existing
                           Write-CompletedLog helper (Task text prefixed
                           "Mail Attribs" so these are distinguishable from
                           this tool's other logged actions).
                       Target/AD-server come from this tool's own Mailbox
                       lookup rather than a separate UPN-entry/Environment UI
                       -- Get-MailAttribsADUser matches the AD account by
                       "mail" (Get-ADUser -Filter "mail -eq '$Smtp'"), the
                       exact pattern Get-MailboxDLMemberships already uses and
                       this file already treats as verified. Set-ADUser
                       parameter names/shapes and Rename-ADObject (kept for
                       parity but unreachable from this UI, same as the
                       reference -- 'name' is excluded from the editable list)
                       are the same calls already Microsoft-Learn-verified in
                       the reference script and/or already used elsewhere in
                       this file; not re-derived independently per standards
                       Section 1(a).
                       Cross-handler state (the currently loaded AD-user
                       object, read/reassigned from the Refresh button, the
                       grid double-click handler, and the Add Secondary
                       handler) is held in $script:MailAttribsCurrentADUser --
                       this file's established pattern for state that must
                       survive across separate event-handler closures (see
                       $script:CurrentRules / $script:CurrentTraceResults /
                       $script:DLScanCancelRequested) rather than a plain
                       local variable, which would not reliably propagate
                       between handlers.
                       Layout: added a second mailbox-actions button row
                       (y=344, "Mail Attribs" at x=12) below the existing row,
                       since that row is already flush with the form's right
                       margin. Everything from the Result box (lblOut)
                       downward -- lblOut, txtOut, the Inbox Rules grid,
                       Mailbox Permissions group, Group Delegation group --
                       shifted +30px, and the form height grew 1080->1110 to
                       match, so nothing overlaps.)
    Revised         : 2026-07-09e (Message Trace: (1) widened the window
                       780->1014 (30%) so the Fill-mode grid columns have
                       more room -- grdTrace/txtDetail/lblTraceStatus grew
                       740->974 to match, right margin held at 28px, and
                       btnTraceClose moved from x=650 to x=886 to stay
                       flush right. (2) Added a "Show time in" dropdown
                       (Eastern/Central/Mountain/Pacific/UTC, default
                       Central) that converts the grid's Received column,
                       the Detail box, and CSV export from UTC -- confirmed
                       via Microsoft Learn that Get-MessageTraceV2 output
                       time stamps are UTC -- to the selected zone. Changing
                       the dropdown re-renders already-fetched results
                       instantly via Update-TraceGridDisplay / cached
                       $script:CurrentTraceDetailRaw, no new Exchange query.
                       Uses .NET TimeZoneInfo with Windows time zone IDs
                       (e.g. "Central Standard Time", which despite the name
                       carries full DST rules) -- standard .NET framework
                       behavior, not an Exchange cmdlet, so not re-verified
                       against Microsoft Learn the way Get-MessageTraceV2
                       itself was. Every conversion forces DateTimeKind.Utc
                       explicitly first (ConvertTo-TraceDisplayTime) rather
                       than trusting the Kind on objects that crossed the
                       EXO remoting boundary, since that Kind isn't reliably
                       preserved. The Search From/Through date pickers are
                       UNCHANGED -- they still control the Get-MessageTraceV2
                       query window exactly as before; only how results are
                       DISPLAYED after the fact is affected.)
    Revised         : 2026-07-09d (Message Trace: added a Status dropdown
                       -- All Messages (default, no -Status filter applied)
                       plus the 7 values Get-MessageTraceV2 -Status actually
                       accepts: Delivered, Failed, Quarantined,
                       FilteredAsSpam, Pending, Expanded, GettingStatus.
                       Verified against Microsoft Learn's Get-MessageTraceV2
                       parameter reference before adding -- "All Messages"
                       is a UI-only convenience value, not a real API value;
                       selecting it just means the Trace click handler omits
                       -Status entirely rather than passing it through.
                       Invoke-MessageTrace gained a matching -Status
                       parameter. Placed in the empty space to the right of
                       the "To" field (below lblTraceNote, which only spans
                       down to y=58) so no other control on the form needed
                       to move.)
    Revised         : 2026-07-09c (operator asked whether EAC's "Group
                       membership" pane does a full-tenant scan the way the
                       2026-07-09b "DL Members" button did. It doesn't --
                       EAC (and Set-Mailbox/Get-User-style tooling generally)
                       reads Active Directory's own "memberOf" attribute on
                       the user object, a backlink AD maintains
                       automatically whenever a group's "member" attribute
                       changes, so it's a single indexed object read, not a
                       search across every group. Rewrote
                       Get-MailboxDLMemberships to do the same: one
                       Get-ADUser -Filter "mail -eq '$Smtp'" -Properties
                       memberOf call (matching the mailbox to its on-prem AD
                       account the same way Resolve-Target already matches
                       on-prem GROUPS by mail elsewhere in this script),
                       then one Get-ADGroup call per DN it returns to
                       resolve a friendly name -- cost is now bounded by how
                       many groups the user actually belongs to, not by how
                       many DLs exist in the tenant. Both cmdlets/parameters
                       verified against Microsoft Learn. Two scope notes
                       carried into the code comments: (1) memberOf is
                       DIRECT membership only, not nested/recursive -- same
                       as what EAC's pane and the on-prem AD "Member Of" tab
                       show; Get-ADPrincipalGroupMembership resolves nested
                       membership too but requires a Global Catalog and
                       behaves differently across domains in a forest, so it
                       was deliberately not used here. (2) a mailbox with no
                       matching on-prem AD account (never synced, cloud-only)
                       has no memberOf to read; Show-DLMembershipsForm now
                       reports that explicitly instead of showing an empty
                       "no memberships found" list. The "DL Members" button,
                       Scan/Cancel UI, and progress label from 2026-07-09b
                       are unchanged in shape -- only what happens when you
                       click Scan changed.
    Revised         : 2026-07-09b (field report: the Health button was
                       hanging / showing "Not Responding". Root cause was
                       the DL membership scan added earlier the same day
                       (see the 2026-07-09 entry below) -- it iterates
                       every DL in the tenant with Get-DistributionGroupMember
                       calls, entirely on the UI thread with no progress
                       updates, so a tenant with more than a couple dozen
                       DLs would make Windows flag the window as frozen even
                       though it was still working. Fix: pulled the DL scan
                       back OUT of Get-MailboxHealthSnapshot -- Health is
                       fast again (delegates, quota, forwarding, holds,
                       suspicious rules, protocol flags all remained, since
                       none of those were the slow part). DL membership is
                       now its own on-demand action: new "DL Members" button
                       next to Message Trace, calling the new
                       Show-DLMembershipsForm, which has its own Scan/Cancel/
                       Close buttons, a live "checking group X of Y: <name>"
                       progress label, and calls
                       [System.Windows.Forms.Application]::DoEvents() on
                       every group so the window stays responsive throughout
                       (and can be cancelled mid-scan, keeping whatever
                       partial results were already found). See
                       Get-MailboxDLMemberships's -ProgressAction /
                       -ShouldCancel parameters.
    Revised         : 2026-07-09 (two fixes, both verified against Microsoft
                       Learn before delivery -- see DEVIATIONS below for
                       full detail:
                       1. BUG: Enable/Disable/Remove-InboxRule were being
                       called with both -Identity and -Mailbox, which
                       Exchange rejects with "You can't use both the
                       Identity and Mailbox parameters to specify a
                       mailbox" whenever -Identity is already mailbox-scoped
                       (which it always is here, since it comes from
                       Get-InboxRule -Mailbox $Smtp). Fixed by dropping
                       -Mailbox from all three calls.
                       2. Health panel expanded: Full Access / Send As
                       delegates are now listed by name (not just counted);
                       added a Distribution List membership scan
                       (Get-MailboxDLMemberships); added quota, forwarding,
                       litigation/retention hold, in-place hold, and a
                       best-effort account-disabled field; added a
                       "Suspicious Inbox Rules" scan for enabled rules with
                       a forward/redirect/delete action; added a "Potential
                       Issues" summary at the top of the report for
                       ticket-escalation triage.)
    Revised         : 2026-07-08 2315 (Message Trace: replaced the "Get
                       Details" button with double-click-to-load-detail
                       [matches the Inbox Rules grid interaction]; added
                       "Release from Quarantine" -- Get-QuarantineMessage /
                       Release-QuarantineMessage, verified against Microsoft
                       Learn including full parameter sets and examples
                       before building, matched by RFC822 MessageId rather
                       than MessageTraceId since those are different values;
                       two-step confirm, the second one showing quarantine
                       Type/Policy/Received time before releasing)
    Revised         : 2026-07-08 2230 (added Message Trace button -- Sent/
                       Received toggle, optional to/from narrowing, subject
                       filter, date range, results grid, Get Details, Export
                       to CSV, via Get-MessageTraceV2/Get-MessageTraceDetailV2;
                       added a Copy button to the Health panel)
    Revised         : 2026-07-08 2145 (added Out of Office, Calendar Perm,
                       and Health buttons -- all gated on a successful
                       Mailbox lookup, all via Invoke-EXOCommand; added the
                       Inbox Rules grid + rule detail window with
                       Enable/Disable/Delete; added Get-EXOCommand/
                       Invoke-EXOCommand itself after live testing showed a
                       plain Get-Mailbox call resolving to an on-prem-sourced
                       command -- see header "FIELD-TESTED FOLLOW-UP" above)
    Run from        : Standard PowerShell with an Exchange Online
                       (ExchangeOnlineManagement) session connected. The
                       ActiveDirectory module must also be available (RSAT)
                       for on-prem group lookups -- no on-prem Exchange
                       Management Shell session is used.
    Required modules: ActiveDirectory; ExchangeOnlineManagement
    DEVIATIONS FROM THE STANDARD TEMPLATE (documented per standards intro --
    "apply every applicable section"):
      - Form is 780 x 920, not the standard 760 x 720. This tool has two
        extra editable-property group boxes (Mailbox Permissions, Group
        Delegation) that the batch Restore/Verify template doesn't need.
      - No Mode radio group (Restore/Verify). This is a single-purpose
        Lookup tool, not an action-vs-verify tool -- Run always reads and
        displays current state. Section 19 (Action/Verify split) doesn't
        apply because edits are scoped, single-grantee actions with their
        own confirmation dialogs, not a batch action mode.
      - Cmdlet syntax for the higher-risk permission/delegation cmdlets
        (Add/Remove-MailboxPermission, Add/Remove-RecipientPermission,
        Add/Remove-UnifiedGroupLinks, Set-ADGroup, Add/Remove-ADGroupMember)
        was verified against Microsoft Learn before delivery. Set-
        DistributionGroup's parameter list was checked specifically for a
        -Type conversion parameter and confirmed NOT to have one -- this is
        why GroupType is read-only in this tool.
      - Message Trace (Show-MessageTraceForm) deliberately does NOT copy the
        operator's reference script's System.Data.DataTable-bound grid. It
        uses the same manual DataGridView (Columns.Add/Rows.Add, FillWeight)
        already established for grdRules and grdCal. UNLIKE grdRules/grdCal
        it IS sortable on every column (operator request) -- sorting an
        unbound DataGridView reorders its Rows collection, so a hidden Key
        column (MessageTraceId|RecipientAddress) is added per row and
        double-click / Release-from-Quarantine look up the matching
        $script:CurrentTraceResults entry by that key (Find-TraceResultByKey)
        instead of by row position -- the same hidden-identity-column
        approach Inbox Rules already uses, adopted here once sorting made
        row-index matching unsafe. Export to CSV uses Export-Csv against the
        raw result objects instead of a hand-rolled writer.
        Get-MessageTraceV2's SenderAddress/RecipientAddress parameters were
        confirmed independently optional (both Mandatory: False) against
        Microsoft Learn before building the Sent/Received one-sided query
        behavior, since the operator's reference script always supplied
        both and this is new usage beyond what that script had already
        validated. Get-QuarantineMessage / Release-QuarantineMessage
        parameter sets and examples were likewise fetched in full from
        Microsoft Learn before building "Release from Quarantine" -- see
        the comment above Find-QuarantineMatch/Invoke-QuarantineRelease.
      - 2026-07-09 revision, cmdlet/property verification performed against
        Microsoft Learn before delivery: Enable-InboxRule / Disable-InboxRule
        / Remove-InboxRule syntax (confirming -Mailbox is optional and only
        needed to disambiguate a bare rule name, which is why dropping it
        fixes the reported error); Get-Mailbox output properties for the
        Quota, Delivery, Hold, and Retention property sets ("Property sets
        in Exchange Online PowerShell module cmdlets") backing
        ProhibitSendQuota/ProhibitSendReceiveQuota, ForwardingSmtpAddress/
        ForwardingAddress/DeliverToMailboxAndForward, LitigationHoldEnabled/
        InPlaceHolds, and RetentionHoldEnabled; Get-InboxRule/Set-InboxRule
        property names ForwardTo/ForwardAsAttachmentTo/RedirectTo/
        DeleteMessage/PermanentDelete/SoftDeleteMessage; and
        Get-ADUser's -Properties parameter and the memberOf backlink
        attribute for DL membership (Get-MailboxDLMemberships -- REWRITTEN
        2026-07-09c after a field report that the original
        Get-DistributionGroup/Get-DistributionGroupMember full-tenant-scan
        version froze the UI; the operator asked whether EAC's Group
        Membership pane does the same full scan, prompting the switch to
        reading AD's own memberOf backlink instead, the same mechanism EAC
        uses). ONE FIELD COULD NOT BE VERIFIED per
        standards Section 2 ("if verification is unavailable, clearly say
        so before delivery"): Get-Mailbox's AccountDisabled OUTPUT property.
        Microsoft Learn confirms -AccountDisabled as a Set-Mailbox
        parameter but this script found no Learn page confirming Get-Mailbox
        returns it back (Microsoft's own docs steer Exchange Online callers
        toward Get-EXOMailbox instead, whose documented property sets don't
        list it either). The Health panel still reads and displays this
        field, but labels it "unverified" inline -- confirm it against a
        live tenant before relying on it, and treat a blank value as
        "unknown" rather than "not disabled."
      - PowerShell was not available in the build environment to run the
        Section 15 AST parse check before delivery -- true for the original
        2026-07-08 build and still true for this 2026-07-09 revision. Run
        the following once before first use and confirm "No parse errors.":
          $tokens = $null; $errs = $null
          [System.Management.Automation.Language.Parser]::ParseFile(
              "<path to this file>", [ref]$tokens, [ref]$errs) | Out-Null
          if ($errs) { $errs } else { "No parse errors." }
```


## function Format-ErrorDetail

- **L563-565** Central formatter so every catch block reports the same level of detail: cmdlet, script line, exception type, message, category, and inner exception if present. Never just "$_.Exception.Message" alone.

## function Get-EXOCommand

- **L589-601** Resolves a cmdlet name to the Exchange Online-sourced command specifically. Field-tested finding (2026-07-08): even with no explicit on-prem EMS session connected, an on-prem Exchange module/snap-in can still be loaded in the host (leftover profile/autoload) and win a plain-name lookup -- the classic symptom is a "couldn't be found on '<DC>'" error from what should be a cloud-only cmdlet, or a script-line number in the thousands (an auto-generated remoting proxy, not this script). Every EXO-flavored call in this script MUST go through Invoke-EXOCommand below instead of calling the cmdlet name directly -- never call Get-Mailbox / Get-MailboxStatistics / Get-MailboxPermission / Add-MailboxPermission / Remove-MailboxPermission / Get-RecipientPermission / Add-RecipientPermission / Remove-RecipientPermission / Get-UnifiedGroup / Get-UnifiedGroupLinks / Add-UnifiedGroupLinks / Remove-UnifiedGroupLinks directly by name anywhere else in this file.

## function Invoke-EXOCommand

- **L618** Splat-friendly wrapper: Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }

## function Get-MailboxState

- **L698-699** Per Microsoft KB4486848, ArchiveStatus can read "None" even when an archive is active. ArchiveGuid is the reliable indicator.

## function Enable-TargetInboxRule

- **L808-817** BUGFIX (2026-07-09): do NOT also pass -Mailbox here. $RuleIdentity comes from Get-InboxRule -Mailbox $Smtp (see Update-RulesGrid), and Exchange returns that Identity already scoped to the mailbox it was queried against. Passing -Mailbox in addition to that already-scoped -Identity throws "You can't use both the Identity and Mailbox parameters to specify a mailbox" -- confirmed against the Enable-InboxRule syntax on Microsoft Learn, where -Mailbox exists only to disambiguate a bare rule NAME (per the Disable-InboxRule/Remove-InboxRule examples on the same page), which is not what's happening here. $Smtp is still accepted as a parameter and used by the caller for logging (Write-CompletedLog).

## function Disable-TargetInboxRule

- **L824-825** BUGFIX (2026-07-09): see the comment above Enable-TargetInboxRule -- same fix, same reason (this is the exact cmdlet/error the operator hit).

## function Remove-TargetInboxRule

- **L832-833** BUGFIX (2026-07-09): see the comment above Enable-TargetInboxRule -- same latent bug (not yet hit in the field, but identical cause).

## function Update-RulesGrid

- **L840-844** Repopulates the grdRules grid (defined in the WinForms UI region) for the given mailbox. Referencing $grdRules/$txtOut here without $script: works the same way Invoke-MailboxEdit/Invoke-GroupEdit already do -- they're script-scope (top-level) variables by the time any button handler calls into this function.

## function Format-GroupState

- **L1067** Unified Group (Microsoft 365 Group) edits -- EXO, resolved via Invoke-EXOCommand

## function Remove-UnifiedMember

- **L1092** On-prem group edits -- ActiveDirectory module, unambiguous cmdlets

## function Set-OnpremOwner

- **L1095** Replaces the single AD ManagedBy value -- see header "KNOWN LIMITATION".

## function Get-CalendarFolderIdentity

- **L1273-1275** Resolves the actual calendar folder rather than hardcoding ":\Calendar", since the folder can be renamed or (rarely) localized. Falls back to the first Calendar-scope folder if none matches the default root path.

## function Set-CalendarPermission

- **L1294-1297** Add-MailboxFolderPermission fails if the user already has an entry, so check first (best-effort) rather than parsing the resulting error text. If the check itself errors for any reason, fall through to Add-; a genuine problem will surface clearly from that call instead.

## function Remove-CalendarPermission

- **L1318-1319** Remove-MailboxFolderPermission removes ALL permissions for this user on the folder -- there's no selective/partial removal, per Microsoft Learn.

## function Get-MailboxDLMemberships

- **L1459-1506** REWRITTEN 2026-07-09c (field report: the original version iterated every DL in the tenant with Get-DistributionGroupMember and froze the UI -- see the two 2026-07-09 revision notes in the header. The operator then asked whether the EAC "Group membership" pane does the same full-tenant scan; it does not.) EAC's Group Membership pane, and Set-Mailbox et al under the hood, get this from Active Directory's own "memberOf" attribute on the user object -- a backlink that AD itself maintains automatically every time a group's "member" attribute changes, so reading it is a single indexed object lookup, not a search across every group. This function now does the same thing via Get-ADUser -Properties memberOf (both the -Properties parameter and its behavior confirmed against Microsoft Learn), matched to the mailbox via its "mail" attribute the same way Resolve-Target already matches on-prem GROUPS by mail elsewhere in this script. Cost is now one Get-ADUser call plus one Get-ADGroup call per group the user actually belongs to (typically a handful to a few dozen) -- bounded by the user's real membership count, not by how many DLs exist in the tenant. Two deliberate scope limits, both consistent with how this tool already treats on-prem groups elsewhere: - Direct membership only, not nested/recursive. memberOf lists groups the user is a DIRECT member of; if that group is itself nested inside another group, the outer group is not included. This is what EAC's Group Membership pane and the on-prem AD "Member Of" tab both show. Get-ADPrincipalGroupMembership resolves nested membership too but requires a Global Catalog and behaves differently in multi-domain forests (per Microsoft Learn) -- not used here to keep behavior simple and predictable across the Franchise/HQ environment split this tool already has. - On-prem-synced Distribution / Mail-Enabled Security groups only, via the ActiveDirectory module -- Microsoft 365 Group membership is a separate object type handled by this tool's own Group lookup path (Get-UnifiedGroupLinks), not included here. - A mailbox with no matching on-prem AD user (cloud-only mailbox, never synced) has no memberOf to read -- $result.NoADAccount is set rather than silently returning an empty list, so that case isn't mistaken for "this user is in zero groups." -ProgressAction, if supplied, is invoked once per group-DN being resolved as & $ProgressAction $currentIndex $totalCount $groupDN. -ShouldCancel, if supplied, is checked after each one; returning $true stops early and sets $result.Cancelled, keeping whatever partial results were already resolved. Both are vestigial now that this is a bounded, typically-fast operation, but kept since a handful of memberOf entries can still point at cross-domain groups that are slow to resolve individually.

## function Get-MailboxHealthSnapshot

- **L1557-1604** Read-only composite view. Reuses Get-InboxRulesForMailbox (Inbox Rules feature) for the rule list/count rather than querying it twice. Interpretation / verification notes (documented since these terms are not single, unambiguous Exchange fields, and per standards Section 2 every cmdlet/property used for a new capability is checked against Microsoft Learn before delivery): - "Protocol usage" is shown as Get-CASMailbox protocol ENABLEMENT flags (OWA/ActiveSync/POP/IMAP/MAPI/EWS), not actual traffic volume. - "Delegates" now lists Full Access and Send As grantees by name (each from its own cmdlet -- see Get-MailboxState), not just a combined count. Does NOT include Outlook's separate calendar/folder delegate list. - DL memberships are NOT part of this snapshot -- see Get-MailboxDLMemberships / Show-DLMembershipsForm / the "DL Memberships" button instead. That scan iterates every DL in the tenant and can take a while; running it automatically inside Health made the whole tool appear to hang (2026-07-09 field report), so it's now a separate on-demand action with its own progress display. - Quota (ProhibitSendQuota, ProhibitSendReceiveQuota), Forwarding (ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward), Litigation Hold (LitigationHoldEnabled, InPlaceHolds), and Retention Hold (RetentionHoldEnabled) are all confirmed Get-Mailbox output properties per Microsoft Learn's "Property sets in Exchange Online PowerShell module cmdlets" (Quota / Delivery / Hold / Retention property sets respectively). Quota values are shown as raw ByteQuantifiedSize text next to Mailbox Size for manual comparison rather than computed as a percentage -- the ByteQuantifiedSize/ Unlimited<T> numeric conversion API was not verified against Microsoft Learn, so this script does not guess at it. - AccountDisabled: Microsoft Learn confirms -AccountDisabled as a Set-Mailbox parameter but this script did NOT find Microsoft Learn confirmation that Get-Mailbox returns an AccountDisabled OUTPUT property (Microsoft's own Get-Mailbox page recommends Get-EXOMailbox instead, whose documented property sets don't list it either). Flagged here as UNVERIFIED per standards Section 2 ("if verification is unavailable, clearly say so before delivery") -- confirm this field is meaningful on first use; if the property doesn't exist on the returned object, PowerShell returns $null here rather than throwing, so a wrong/missing value fails quietly. Treat a blank value as "unknown," not as "account is enabled." - "Suspicious Inbox Rules" flags ENABLED rules with a forwarding, redirect, or delete action (ForwardTo/ForwardAsAttachmentTo/ RedirectTo/DeleteMessage/PermanentDelete/SoftDeleteMessage -- all confirmed Get-InboxRule/Set-InboxRule property names on Microsoft Learn). This is a flag for a human to review as a first-check item on compromised-account or missing-mail tickets, not a judgment that the rule is malicious.
- **L1713-1714** DL membership is intentionally NOT scanned here -- see the note above this function and the "DL Memberships" button (Show-DLMembershipsForm).

## function Show-MailboxHealthForm

- **L1747-1749** One-shot timer so "Copied!" reverts to "Copy" without needing a second click -- local to this form (closure-captured), same pattern as every other control reference in this function.

## function Update-HealthFromCurrent

- **L1763-1766** --- Potential Issues summary first -- the point of this button is "what do I need to know before I escalate this ticket", so lead with anything that looks like a likely cause rather than making the reader hunt for it further down.

## function Show-DLMembershipsForm

- **L1853-1860** Split out of the Health panel on 2026-07-09, then REWRITTEN 2026-07-09c to read AD's memberOf backlink instead of scanning every DL -- see the Get-MailboxDLMemberships header comment and the header .NOTES revision entries for the full story. Now a single Get-ADUser call plus one Get-ADGroup call per actual membership (typically fast), not a full-tenant scan. Scan/Cancel/progress-label plumbing is kept since a cross-domain group DN can still occasionally be slow to resolve, but this should complete quickly in the common case.

## region Actions - mail attributes

- **L1958-1999** "Mail Attribs" button (2026-07-10). Ports the on-prem AD attribute grid and edit mechanics from the operator's "Renaming Script" (on-prem AD name-change tool) almost verbatim -- same fixed attribute list, same derived PrimarySmtpAddress/proxyAddresses handling, same Keep-as-Alias/Remove prompt, same "Add secondary SMTP address" control -- per explicit instruction to match that script's behavior exactly. Differences are only what's required to fit this tool's existing conventions: - The target user comes from this tool's own Mailbox lookup (Smtp/ADServer already resolved by Resolve-Target / Get-SelectedADServer) instead of a separate UPN-entry/Parse/Environment UI -- Get-MailAttribsADUser matches the AD account by "mail" (Get-ADUser -Filter "mail -eq '$Smtp'"), the exact same match-by-mail pattern already used and verified by Get-MailboxDLMemberships just above. - The grid/dialog live inside one self-contained Show-MailAttribsForm function, matching how every other feature dialog in this script (Show-DLMembershipsForm, Show-MailboxHealthForm, etc.) is built. Because the current AD-user object needs to be read AND reassigned across several separate event handlers (Refresh, the shared edit function used by both double-click and the Edit Attribute button, and Add Secondary) -- not just one -- it is held in $script:MailAttribsCurrentADUser rather than a plain closure variable, for the same reason this file already uses $script:CurrentRules / $script:CurrentTraceResults / $script:DLScanCancelRequested for equivalent cross-handler state: a plain variable assignment inside one event-handler scriptblock does not reliably propagate back out to a sibling handler's scriptblock. - Function/variable names are new (MailAttribs-prefixed where the reference name would otherwise be too generic to safely reuse) to avoid any collision with this file's existing functions; the proxyAddresses helpers and Format-AttributeValue keep their original reference-script names since nothing else in this file defines them. - "All Attribs" (the operator's second requested button, a full every-attribute dump) is a separate follow-up request and is not part of this region. Cmdlets used here (Get-ADUser -Filter/-Properties/-Server, Set-ADUser -DisplayName/-GivenName/-Initials/-Surname/-EmailAddress/-Clear/-Replace, Rename-ADObject -NewName) are the exact same calls, in the exact same shapes, already used and Microsoft-Learn-verified in the reference script this was ported from, and/or already used elsewhere in this file (Get-MailboxDLMemberships, Resolve-Target's on-prem group path) -- not re-verified independently per standards Section 1(a) (prior verified usage stands).
- **L2011-2014** UserPrincipalName, Company ("Office"), name, and proxyAddresses are read-only/context-only, exactly as in the reference script -- name is the object's CN/RDN (out of scope for a mail-attributes tool) and proxyAddresses is only ever changed narrowly, via PrimarySmtpAddress and "Add Secondary".
- **L2020-2023** Matches the AD account to this mailbox by its "mail" attribute -- the same lookup Get-MailboxDLMemberships already uses just above, so a mailbox without a synced on-prem account throws the same style of message rather than silently showing an empty grid.

## function Get-MailAttribsADUser

- **L2047-2048** Reads the primary SMTP address out of proxyAddresses -- the entry prefixed "SMTP:" in uppercase (Exchange convention; lowercase "smtp:" is secondary).

## function Get-PrimarySmtpFromProxyAddresses

- **L2056-2060** Builds a new proxyAddresses array with $NewPrimary as the SMTP: entry. $KeepOldPrimaryAsAlias controls what happens to the previous primary (the operator's answer to Show-KeepOrRemoveOldPrimaryDialog): kept as a demoted smtp: entry, or dropped entirely. An existing secondary equal to the new primary is removed first so the address never appears twice.

## function Set-PrimarySmtpInProxyAddresses

- **L2088-2089** Appends $NewSecondary as an smtp: entry. Throws on a duplicate rather than silently creating one.

## function Add-SecondarySmtpToProxyAddresses

- **L2105-2106** Renders a single attribute value for grid display -- handles the multi-valued proxyAddresses array so it never throws on an odd property type.

## function Format-AttributeValue

- **L2117-2119** Builds the Attribute/Value DataTable, in the fixed order above. PrimarySmtpAddress has no real AD property of its own -- derived from proxyAddresses here, the single place that derivation happens.

## function Format-MailAttribsTable

- **L2141-2142** Comma operator required -- a bare "return $table" would unroll the DataTable into its DataRow objects instead of returning the table itself.
- **L2145-2151** Writes one attribute change. Branches by mechanism exactly as the reference script documents: dedicated Set-ADUser parameters, generic -Replace/-Clear for Exchange-schema attributes with no dedicated parameter, and the derived PrimarySmtpAddress field (rewrites proxyAddresses per $KeepOldPrimaryAsAlias). The 'name' case (Rename-ADObject) is kept for parity with the reference script but is unreachable from this UI, same as there -- 'name' is deliberately excluded from $script:MailAttribsEditableAttributes.

## function Set-MailAttribsAttribute

- **L2210-2211** Standalone write for the "Add secondary SMTP address" control -- appends one new smtp: entry without touching anything else in proxyAddresses.

## function Add-MailAttribsSecondaryProxyAddress

- **L2224-2227** Edit dialog for a single attribute -- ported near-verbatim from the reference script's Show-AttributeEditDialog, including the extra warning notes for 'name' and 'PrimarySmtpAddress' (the 'name' branch is kept for parity even though that row is never reachable from this UI, same as there).

## function Show-MailAttribsEditDialog

- **L2315-2318** Shown only when editing PrimarySmtpAddress AND an existing, different primary is about to be displaced -- ported verbatim from the reference script. Returns 'Keep', 'Remove', or 'Cancel' (Cancel, or closing via the X button, aborts the whole attribute edit -- nothing is written to AD).

## function Show-MailAttribsForm

- **L2451-2453** Added 2026-07-10c alongside double-click -- see the CellDoubleClick / SelectionChanged handlers below. Same position/size as All Attribs' Edit Attribute button for visual consistency between the two dialogs.
- **L2466-2472** Current AD-user object for this dialog instance. Script-scoped (not a plain closure variable) because it is reassigned from three separate event handlers below (Refresh, grid double-click edit, Add Secondary) -- see the region header comment for why that requires $script: scope in this file's established pattern. Only one Mail Attribs dialog can be open at a time (modal, ShowDialog($form)), so a single script-scope slot is safe, same as $script:CurrentRules / $script:CurrentTraceResults.

## function Update-MailAttribsFromCurrent

- **L2501-2508** Grid-row edit path -- reachable two ways as of 2026-07-10c: double-click a row (CellDoubleClick below), or select a row and click "Edit Attribute" (btnMailEdit below, enabled only for an editable row via SelectionChanged). Both call this same function against whatever row is currently selected in $grdMail, so the edit logic itself only exists once. Silently no-ops on a read-only row (defensive -- SelectionChanged already keeps btnMailEdit disabled for those, and double-click checks $script:MailAttribsEditableAttributes itself below).

## function Invoke-MailAttribsEditSelectedRow

- **L2599-2600** Double-click: select the row (matching the pattern the reference scripts use elsewhere), then run the shared edit function.
- **L2608** Selection drives Edit Attribute's enabled state, same as All Attribs.

## region Actions - all attributes

- **L2669-2713** "All Attribs" button (2026-07-10). Ports the operator's separate AD-UserAttributeEditor reference script (full Get-ADUser -Properties * viewer/editor with a curated editable subset) into this tool, following the same adaptation approach used for "Mail Attribs" just above: - Target/AD-server come from this tool's own Mailbox lookup, matched by "mail" (Get-ADUser -Filter "mail -eq '$Smtp'") -- same pattern as Get-MailAttribsADUser / Get-MailboxDLMemberships -- instead of the reference's separate UPN-entry/Parse/Environment UI. - Interaction model initially matched the REFERENCE SCRIPT exactly, per instruction ("same as last -- if they are editable here, they are editable in the final version"): select a row, then click the "Edit Attribute" button, enabled only when the selected row's hidden _Editable flag is 'Y' (grdAll's SelectionChanged handler). 2026-07-10c added double-click as a second, equivalent way to trigger the same edit (Invoke-AllAttribsEditSelectedRow), for parity with Mail Attribs -- both dialogs now support both triggers. - The editable-attribute list is copied verbatim from the reference (Title, Department, Description, Company, Division, Office, OfficePhone, MobilePhone, Fax, HomePhone, StreetAddress, City, State, PostalCode, POBox, EmployeeID, EmployeeNumber, Organization, Initials, HomePage) -- each already documented there as a verified, dedicated Set-ADUser string parameter (Microsoft Learn), never a generic -Replace guess. Everything else Get-ADUser -Properties * returns (SamAccountName, UserPrincipalName, DistinguishedName, ObjectGUID, ObjectSID, UserAccountControl, Manager, GivenName, Surname, DisplayName, mail, the whole proxyAddresses/Exchange-schema set already covered by Mail Attribs, etc.) is view-only here, same as the reference -- not re-derived independently per standards Section 1(a). - Get-ADUser -Properties * returns both ADUser's friendly property names (e.g. "Office") and, when populated, the matching raw LDAP attribute name (e.g. "physicalDeliveryOfficeName") as SEPARATE rows with the same value -- normal Get-ADUser behavior per the reference script's own header, not a bug introduced here. Only the friendly-named row is ever in $script:AllAttribsEditableAttributes, so only it is ever editable; the note label below the grid says so. - Function/variable names are AllAttribs-prefixed to avoid colliding with Mail Attribs' identically-purposed but differently-scoped functions (e.g. Format-AllAttribsValue here additionally handles byte[]/SID/ GUID/DateTime -- types the curated mail-attribute list never produces but a full -Properties * dump routinely does -- so it is intentionally NOT the same function as Mail Attribs' simpler Format-AttributeValue). - Same $script:-scope reasoning as Mail Attribs for the current AD-user object ($script:AllAttribsCurrentADUser), since it's read/reassigned from three separate event handlers (Refresh, Edit, and the SelectionChanged enable/disable check).
- **L2720-2721** Attribute-object bookkeeping properties added by the ADUser type itself -- never real directory attributes, always excluded from the display grid.
- **L2726-2728** Matches the AD account to this mailbox by its "mail" attribute -- same pattern as Get-MailAttribsADUser -- but requests every populated attribute (-Properties '*') rather than a fixed list.

## function Get-AllAttribsADUser

- **L2750-2754** Renders a single attribute value for grid display. Handles arrays, binary, GUID/SID, and DateTime -- ported verbatim from the reference script, since a full -Properties * dump routinely includes all of these (objectGUID, objectSid, whenCreated, lastLogon, etc.), unlike Mail Attribs' curated always-string attribute list.

## function Format-AllAttribsValue

- **L2772-2775** Builds the Attribute/Value/_Editable DataTable bound to the grid -- every populated property on the ADUser object (minus the bookkeeping properties above), sorted by name. Ported from the reference script's Format-UserAttributeTable.

## function Format-AllAttribsTable

- **L2792-2793** Comma operator required -- a bare "return $table" would unroll the DataTable into its DataRow objects instead of returning the table itself.
- **L2796-2799** Set-ADUser via the dedicated named parameter matching $AttributeName (must be a member of $script:AllAttribsEditableAttributes -- enforced by the caller). Blank new value clears the attribute instead of writing an empty string. Ported from the reference script's Set-OnPremADUserAttribute.

## function Set-AllAttribsAttribute

- **L2819-2822** Edit dialog for a single attribute -- ported near-verbatim from the reference script's Show-AttributeEditDialog (the simpler version, with no extra warning notes, since none of the curated attributes here have the special-case semantics PrimarySmtpAddress/name have in Mail Attribs).

## function Show-AllAttribsForm

- **L2934-2937** BindingSource is the Microsoft-recommended intermediary for binding a DataGridView to a DataTable -- same as Mail Attribs' grid. Column-header click-to-sort works natively once bound through it (SortMode defaults to Automatic), no custom sort handler needed -- matches the reference script.
- **L2978-2982** Current AD-user object for this dialog instance -- script-scoped for the same cross-handler reason as $script:MailAttribsCurrentADUser (read and reassigned from Refresh, Edit Attribute, and the SelectionChanged enable/disable check). Only one All Attribs dialog can be open at a time (modal, ShowDialog($form)), so a single script-scope slot is safe.

## function Update-AllAttribsFromCurrent

- **L3010** Selection drives Edit Attribute's enabled state.
- **L3017-3021** Grid-row edit path -- reachable two ways as of 2026-07-10c: select a row and click "Edit Attribute" (matches the reference script's own model), or double-click the row directly (added for parity with Mail Attribs). Both call this same function against whatever row is currently selected in $grdAll, so the edit logic itself only exists once.

## function Invoke-AllAttribsEditSelectedRow

- **L3069-3071** Double-click: select the row, then run the shared edit function. A double-click on a non-editable row is a silent no-op (the function itself checks $script:AllAttribsEditableAttributes).

## region Actions - message trace

- **L3084-3099** Get-MessageTraceV2 / Get-MessageTraceDetailV2 verified against Microsoft Learn (2026-07-08): both are "available only in the cloud-based service" (no on-prem equivalent by this name -- Get-MessageTrackingLog is the on-prem cmdlet, a completely different name), so there is no collision risk the way Get-Mailbox etc. have. SenderAddress and RecipientAddress on Get-MessageTraceV2 are BOTH independently optional (Mandatory: False on both), which is what makes the Sent/Received one-sided trace below possible with a single cmdlet call -- Sent omits RecipientAddress entirely (any recipient), Received omits SenderAddress entirely (any sender). Still routed through Invoke-EXOCommand for consistency with every other EXO-flavored call in this script, per the standing rule in Get-EXOCommand's header comment. Get-MessageTraceDetailV2 is called exactly the way the operator's own working "EXO Message Trace" reference script already calls it (-MessageTraceId, -RecipientAddress, -StartDate, -EndDate) -- that usage is treated as verified per standards Section 1(a) and is not re-derived here.

## function Invoke-MessageTrace

- **L3101-3105** -Status accepts any of the values Get-MessageTraceV2 documents for its own -Status parameter (verified against Microsoft Learn, 2026-07-09): Delivered, Expanded, Failed, FilteredAsSpam, GettingStatus, Pending, Quarantined. Omit it entirely (as the "All Messages" dropdown option does) to get the unfiltered default.

## function Invoke-MessageTraceDetail

- **L3144-3167** Quarantine release -- verified against Microsoft Learn (2026-07-08): - Get-QuarantineMessage / Release-QuarantineMessage are both "available only in the cloud-based service" (Exchange PowerShell module) -- same module family as Get-MessageTraceV2, routed through Invoke-EXOCommand. - Get-QuarantineMessage -MessageId expects the full RFC822 Message-ID header value, angle brackets included (Microsoft Learn Example 3: Get-QuarantineMessage -MessageID "<...@contoso.com>"). This is a DIFFERENT value than Get-MessageTraceV2's MessageTraceId (a GUID) -- the correlation below uses the trace result's MessageId property (the RFC822 header), NOT MessageTraceId. - ASSUMPTION, not verified against a live tenant: Get-MessageTraceV2 result objects expose a MessageId property carrying that RFC822 value (this has been true of the classic Get-MessageTrace schema and is what admins have always used to cross-reference quarantine, but the operator should confirm on first use -- Find-QuarantineMatch throws a clear, specific error if it's missing rather than silently matching the wrong thing). - Release-QuarantineMessage verified via Microsoft Learn Example 2: Release-QuarantineMessage -Identity <id> -ReleaseToAll releases a single quarantined message (by its own Identity, from Get-QuarantineMessage) to ALL of its original recipients. -Confirm $false is added the same way every other write cmdlet in this script suppresses the built-in prompt, since this tool has its own MessageBox confirmation instead.

## function Invoke-QuarantineRelease

- **L3195-3207** Time zone display for Message Trace results -- Microsoft Learn confirms "The time stamps on the output are in UTC time format" for Get-MessageTraceV2 (and by extension Get-MessageTraceDetailV2), so everything shown to the operator needs an explicit UTC-to-local conversion. Windows time zone IDs (not IANA names) are used with [System.TimeZoneInfo], which is correct for the "Standard PowerShell" Windows environment this script's header already requires -- these IDs carry their DST rules despite the "Standard Time" naming (.NET's own documented behavior, not Exchange-specific, so not re-verified against Microsoft Learn the way Exchange cmdlets are). This ONLY changes how already-fetched results are DISPLAYED (grid, Detail box, CSV export) -- it does not change how the Search From/Through date pickers are interpreted; those are passed to Get-MessageTraceV2 exactly as before.

## function Get-SelectedTraceTimeZone

- **L3222-3224** Falls back to UTC (unconverted) rather than throwing, so a missing time zone on this particular machine degrades gracefully instead of breaking the whole Trace window.

## function ConvertTo-TraceDisplayTime

- **L3229-3234** DateTime values coming back over Exchange Online's remoting connection do not reliably keep DateTimeKind.Utc through deserialization, so this forces Kind=Utc explicitly before converting rather than trusting whatever Kind the object already reports -- converting an already-local value as if it were UTC would silently shift every timestamp by several hours, which is worse than not converting at all.

## function Show-MessageTraceForm

- **L3260-3273** DESIGN NOTE: the operator's pasted reference script binds results to a System.Data.DataTable (auto-generated grid columns, hidden-column pairs for the trace-id/recipient needed by Get Details). This script instead follows the SAME manual-grid convention already used for grdRules and grdCal elsewhere in this file (Columns.Add + Rows.Add, FillWeight sizing). UNLIKE grdRules/grdCal, this grid IS sortable on every column (operator request) -- sorting an unbound DataGridView physically reorders its Rows collection, so row position can no longer be used to find the matching $script:CurrentTraceResults entry after a sort. Fix: a hidden Key column (MessageTraceId|RecipientAddress) on every row, looked up by value via Find-TraceResultByKey -- the same "hidden identity column" approach grdRules already uses for exactly this reason, just adopted here too once sorting made row-index matching unsafe.
- **L3277-3281** Width increased 30% (780 -> 1014) on request so the DataGridView's Fill-mode columns get more room to show full sender/recipient/subject text without truncation. Right margin (28px, matching the original 780-12-740 layout) preserved -- content-area controls below widen from 740 to 974 (740 + the 234px gained) to keep that same margin.
- **L3306-3310** Time zone display selector -- placed in the space opened up to the right of lblTraceNote by the 30%-wider form. Converts already-fetched results (grid, Detail box, CSV export) on selection change without re-querying Exchange Online -- see the ConvertTo-TraceDisplayTime / Get-SelectedTraceTimeZone header comment above Format-TraceDetail.
- **L3333-3339** Status filter -- placed in the empty space to the right of the "To" row (below lblTraceNote, which only spans y=12-58) so nothing else on the form needs to move. Values verified against Microsoft Learn (Get-MessageTraceV2 -Status): Delivered, Expanded, Failed, FilteredAsSpam, GettingStatus, Pending, Quarantined. "All Messages" is not a real API value -- it just means "don't pass -Status at all", handled in the Trace click handler below.
- **L3386-3387** 10-day-window business rule ported from the reference script: Through can never be more than 9 days after Search From, and never in the future.
- **L3434-3444** sorts chronologically, not alphabetically MessageTraceId|RecipientAddress -- see below Sortable on every visible column (unlike grdRules/grdCal elsewhere in this script, which are intentionally NotSortable) -- sorting reorders the underlying Rows collection in an unbound grid like this one, so row position can no longer be used to find the matching $script:CurrentTraceResults entry. The hidden Key column below is the fix: each row carries a MessageTraceId+RecipientAddress key that double-click and Release-from-Quarantine look up by value instead of by row index, the same "hidden identity column" approach grdRules already uses for exactly this reason.

## function Update-TraceGridDisplay

- **L3483-3487** Rebuilds the grid from the already-cached $script:CurrentTraceResults in whatever time zone cboTimeZone currently has selected. Called both right after a Trace run and whenever the time zone dropdown changes -- the latter needs no new Exchange query since the raw (UTC) data is already in hand.

## function Find-TraceResultByKey

- **L3555-3558** Looks up the raw Get-MessageTraceV2 object behind a grid row by its hidden Key column value (MessageTraceId|RecipientAddress) rather than by row position -- required now that the grid is sortable, since sorting an unbound DataGridView physically reorders Rows.
- **L3565-3567** Double-click a row to load its detail -- same interaction pattern as grdRules elsewhere in this script (double-click for detail, no separate "open" button).
- **L3635-3637** Received is exported in whatever zone the dropdown currently shows (matching the grid on screen), via a calculated property rather than the raw UTC value -- same conversion as the grid.

## region WinForms UI

- **L3688** Warning label -- first visible element
- **L3696** Environment radio group -- Franchise first, HQ second (Section 22)
- **L3713** DC dropdown -- enabled for Franchise, disabled/greyed for HQ

## function Update-EnvironmentUI

- **L3738** Input area
- **L3749** Parsed preview
- **L3764** Parse / Run / Close row (Section 13). No Mode group -- see header deviation note.
- **L3780-3783** Mailbox actions row -- enabled only after a successful Mailbox lookup. This is deliberately a separate row from Parse/Run/Close (Section 13) so more feature buttons can be added here later without re-flowing the core lookup controls above.
- **L3819-3824** Split out of the Health panel on 2026-07-09 -- see Get-MailboxDLMemberships / Show-DLMembershipsForm. Placed in the remaining space on the same row as the other mailbox-action buttons (516+130=646 through the row's right edge at 752, matching txtInput/txtOut's right edge) rather than a new row, since a new row would collide with lblOut (y=352) / txtOut (y=372) immediately below and require re-flowing every control after it.
- **L3831-3837** Second mailbox-actions row -- added 2026-07-10 for the new attribute-lookup buttons ("Mail Attribs" now; "All Attribs" is a separate follow-up request and is not added yet). The first action row (Inbox Rules..DL Members) is already flush with the form's right margin (650+102=752, matching txtInput's right edge at 12+740=752), so this is a new row rather than a squeezed-in extra button. Everything from lblOut downward is shifted +30px (and the form height +30px) to make room -- see those controls below.
- **L3844-3847** Second requested attribute-lookup button (2026-07-10, follow-up to Mail Attribs above) -- full Get-ADUser -Properties * viewer/editor, ported from the operator's separate AD-UserAttributeEditor reference script. Placed right after Mail Attribs in the same row (12+110=122, +10px gap=132).
- **L3854** Result box
- **L3868-3870** Inbox Rules grid -- populated by the "Inbox Rules" button, only after a successful Mailbox lookup. A plain TextBox can't have clickable rows, so this is a separate DataGridView rather than reusing the Result box above.
- **L3898** Mailbox Permissions edit group
- **L3935** Group Delegation edit group
- **L3977** ---- shared script state ----
- **L3980** result of Resolve-Target
- **L3982-3983** raw Get-InboxRule objects backing the grdRules grid set by Show-CalendarPermissionsForm while it's open

## function Update-ParsedPreview

- **L4082** ---- Mailbox permission edit handlers ----

## function Invoke-MailboxEdit

- **L4135** ---- Group delegation edit handlers ----

## function Invoke-GroupEdit

- **L4197** ---- Inbox Rules handlers ----
- **L4203-4207** (Test-MailboxLookupReady is defined further below with the other feature-button handlers, but PowerShell resolves function calls at invocation time, so any of these buttons can use it once the whole script has loaded -- this one is left inline since it was written first, functionally identical to the helper.)
- **L4228** ---- Out of Office / Calendar Perm / Health handlers ----
