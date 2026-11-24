<#
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
#>
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#region Parsing ---------------------------------------------------------------
$script:MailDomain = 'expresspros.com'
function Get-TargetSmtp {
    param([string]$Text)
    if (-not $Text) { return $null }
    $t = $Text.Trim()
    if ($t -eq '') { return $null }
    if ($t -like '*@*') { return $t }
    return "$t@$($script:MailDomain)"
}
#endregion
#region Helpers - error formatting --------------------------------------------
function Format-ErrorDetail {
    # Central formatter so every catch block reports the same level of
    # detail: cmdlet, script line, exception type, message, category, and
    # inner exception if present. Never just "$_.Exception.Message" alone.
    param(
        [Parameter(Mandatory)] $ErrorRecord,
        [string]$Context = ''
    )
    $ex   = $ErrorRecord.Exception
    $cmd  = if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.MyCommand) {
                $ErrorRecord.InvocationInfo.MyCommand.Name
            } else { '(cmdlet unknown)' }
    $line = if ($ErrorRecord.InvocationInfo) { $ErrorRecord.InvocationInfo.ScriptLineNumber } else { '?' }
    $sb = New-Object System.Text.StringBuilder
    $prefix = if ($Context) { "ERROR [$Context]" } else { 'ERROR' }
    [void]$sb.AppendLine("$prefix -- $cmd (script line $line)")
    [void]$sb.AppendLine("    Message  : $($ex.Message)")
    [void]$sb.AppendLine("    Type     : $($ex.GetType().FullName)")
    if ($ErrorRecord.CategoryInfo) {
        [void]$sb.AppendLine("    Category : $($ErrorRecord.CategoryInfo.Category) ($($ErrorRecord.CategoryInfo.Reason))")
    }
    if ($ex.InnerException) {
        [void]$sb.AppendLine("    Inner    : $($ex.InnerException.Message)")
    }
    return $sb.ToString().TrimEnd()
}
function Get-EXOCommand {
    # Resolves a cmdlet name to the Exchange Online-sourced command specifically.
    # Field-tested finding (2026-07-08): even with no explicit on-prem EMS session
    # connected, an on-prem Exchange module/snap-in can still be loaded in the host
    # (leftover profile/autoload) and win a plain-name lookup -- the classic symptom
    # is a "couldn't be found on '<DC>'" error from what should be a cloud-only
    # cmdlet, or a script-line number in the thousands (an auto-generated remoting
    # proxy, not this script). Every EXO-flavored call in this script MUST go
    # through Invoke-EXOCommand below instead of calling the cmdlet name directly --
    # never call Get-Mailbox / Get-MailboxStatistics / Get-MailboxPermission /
    # Add-MailboxPermission / Remove-MailboxPermission / Get-RecipientPermission /
    # Add-RecipientPermission / Remove-RecipientPermission / Get-UnifiedGroup /
    # Get-UnifiedGroupLinks / Add-UnifiedGroupLinks / Remove-UnifiedGroupLinks
    # directly by name anywhere else in this file.
    param([Parameter(Mandatory)][string]$Name)
    $all = @(Get-Command -Name $Name -All -ErrorAction SilentlyContinue)
    if ($all.Count -eq 0) {
        throw "No command named '$Name' exists in this session. Is Exchange Online connected (Connect-ExchangeOnline)?"
    }
    $cmd = $all | Where-Object { $_.ModuleName -like 'tmpEXO_*' -or $_.Source -like 'tmpEXO_*' } | Select-Object -First 1
    if (-not $cmd) {
        $cmd = $all | Where-Object { $_.ModuleName -eq 'ExchangeOnlineManagement' -or $_.Source -eq 'ExchangeOnlineManagement' } | Select-Object -First 1
    }
    if (-not $cmd) {
        $sources = ($all | ForEach-Object { "$($_.Name) [ModuleName=$($_.ModuleName); Source=$($_.Source)]" }) -join '  |  '
        throw "Found '$Name' but none of the loaded copies look like Exchange Online: $sources -- run 'Get-Command $Name -All | Select Name,ModuleName,Source' to confirm which one this should be; an on-prem Exchange module/snap-in may be shadowing it."
    }
    return $cmd
}
function Invoke-EXOCommand {
    # Splat-friendly wrapper: Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    $cmd = Get-EXOCommand -Name $Name
    & $cmd @Parameters
}
#endregion
#region Actions - resolution ---------------------------------------------------
function Resolve-Target {
    param(
        [string]$Smtp,
        [string]$ADServer
    )
    $result = [ordered]@{
        Type              = 'NotFound'
        Smtp              = $Smtp
        ADGroupDN         = $null
        ResolutionErrors  = @()
    }
    try {
        Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'Mailbox'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (EXO) resolution attempt')
    }
    try {
        Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'UnifiedGroup'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroup (EXO) resolution attempt')
    }
    try {
        $grp = @(Get-ADGroup -Filter "mail -eq '$Smtp'" -Server $ADServer `
                    -Properties mail, ManagedBy, GroupCategory, GroupScope, msExchHideFromAddressLists `
                    -ErrorAction Stop)
        if ($grp.Count -eq 1) {
            $result.Type      = 'OnpremGroup'
            $result.ADGroupDN = $grp[0].DistinguishedName
            return [PSCustomObject]$result
        } elseif ($grp.Count -gt 1) {
            $result.ResolutionErrors += "ERROR [Get-ADGroup (on-prem) resolution attempt] -- $($grp.Count) AD groups matched mail '$Smtp'. Expected 0 or 1. Resolve the duplicate mail attribute before using this tool."
        }
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADGroup (on-prem) resolution attempt')
    }
    return [PSCustomObject]$result
}
#endregion
#region Actions - mailbox -------------------------------------------------------
function Get-MailboxState {
    param([string]$Smtp)
    $state = [ordered]@{
        DisplayName          = $null
        PrimarySmtpAddress   = $null
        Alias                = $null
        RecipientTypeDetails = $null
        Hidden               = $null
        LitigationHold       = $null
        ArchiveEnabled       = $null
        ArchiveStatus        = $null
        MailboxSize          = $null
        MailboxItemCount     = $null
        ArchiveSize          = $null
        ArchiveItemCount     = $null
        FullAccess           = @()
        SendAs               = @()
        Errors               = @()
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.DisplayName          = $mbx.DisplayName
        $state.PrimarySmtpAddress   = $mbx.PrimarySmtpAddress
        $state.Alias                = $mbx.Alias
        $state.RecipientTypeDetails = $mbx.RecipientTypeDetails
        $state.Hidden               = $mbx.HiddenFromAddressListsEnabled
        $state.LitigationHold       = $mbx.LitigationHoldEnabled
        # Per Microsoft KB4486848, ArchiveStatus can read "None" even when an
        # archive is active. ArchiveGuid is the reliable indicator.
        $state.ArchiveEnabled = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        $state.ArchiveStatus  = $mbx.ArchiveStatus
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (core properties)')
    }
    try {
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.MailboxSize      = $stats.TotalItemSize
        $state.MailboxItemCount = $stats.ItemCount
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (mailbox size)')
    }
    if ($state.ArchiveEnabled) {
        try {
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $state.ArchiveSize      = $astats.TotalItemSize
            $state.ArchiveItemCount = $astats.ItemCount
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (archive size)')
        }
    }
    try {
        $fa = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.User -notlike 'NT AUTHORITY\SELF' -and
                $_.IsInherited -eq $false -and
                $_.Deny -eq $false -and
                ($_.AccessRights -contains 'FullAccess')
            }
        $state.FullAccess = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access)')
    }
    try {
        $sa = Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.Trustee -notlike 'NT AUTHORITY\SELF' -and
                $_.AccessControlType -eq 'Allow'
            }
        $state.SendAs = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As)')
    }
    return [PSCustomObject]$state
}
function Format-MailboxState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading (Mailbox) ---")
    [void]$sb.AppendLine("  Display Name        : $($State.DisplayName)")
    [void]$sb.AppendLine("  Primary SMTP        : $($State.PrimarySmtpAddress)")
    [void]$sb.AppendLine("  Alias               : $($State.Alias)")
    [void]$sb.AppendLine("  Mailbox Type        : $($State.RecipientTypeDetails)")
    [void]$sb.AppendLine("  Hidden From GAL     : $($State.Hidden)")
    [void]$sb.AppendLine("  Litigation Hold     : $($State.LitigationHold)")
    [void]$sb.AppendLine("  Mailbox Size        : $($State.MailboxSize)  ($($State.MailboxItemCount) items)")
    [void]$sb.AppendLine("  Archive Enabled     : $($State.ArchiveEnabled)  (ArchiveStatus: $($State.ArchiveStatus))")
    if ($State.ArchiveEnabled) {
        [void]$sb.AppendLine("  Archive Size        : $($State.ArchiveSize)  ($($State.ArchiveItemCount) items)")
    }
    [void]$sb.AppendLine("  Full Access grantees:")
    if ($State.FullAccess.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.FullAccess | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    [void]$sb.AppendLine("  Send As grantees    :")
    if ($State.SendAs.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.SendAs | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    if ($State.Errors.Count -gt 0) {
        [void]$sb.AppendLine("  --- Field errors (partial data above is still valid) ---")
        $State.Errors | ForEach-Object { [void]$sb.AppendLine("  $_") }
    }
    return $sb.ToString()
}
function Grant-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Grant-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
#endregion
#region Actions - inbox rules ---------------------------------------------------
function Get-InboxRulesForMailbox {
    param([string]$Smtp)
    try {
        $rules = @(Invoke-EXOCommand -Name 'Get-InboxRule' -Parameters @{ Mailbox = $Smtp; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ Rules = $rules; Error = $null }
    } catch {
        return [PSCustomObject]@{ Rules = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-InboxRule') }
    }
}
function Enable-TargetInboxRule {
    # BUGFIX (2026-07-09): do NOT also pass -Mailbox here. $RuleIdentity comes
    # from Get-InboxRule -Mailbox $Smtp (see Update-RulesGrid), and Exchange
    # returns that Identity already scoped to the mailbox it was queried
    # against. Passing -Mailbox in addition to that already-scoped -Identity
    # throws "You can't use both the Identity and Mailbox parameters to
    # specify a mailbox" -- confirmed against the Enable-InboxRule syntax on
    # Microsoft Learn, where -Mailbox exists only to disambiguate a bare rule
    # NAME (per the Disable-InboxRule/Remove-InboxRule examples on the same
    # page), which is not what's happening here. $Smtp is still accepted as a
    # parameter and used by the caller for logging (Write-CompletedLog).
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Enable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Disable-TargetInboxRule {
    # BUGFIX (2026-07-09): see the comment above Enable-TargetInboxRule --
    # same fix, same reason (this is the exact cmdlet/error the operator hit).
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Disable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-TargetInboxRule {
    # BUGFIX (2026-07-09): see the comment above Enable-TargetInboxRule --
    # same latent bug (not yet hit in the field, but identical cause).
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Remove-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Update-RulesGrid {
    # Repopulates the grdRules grid (defined in the WinForms UI region) for
    # the given mailbox. Referencing $grdRules/$txtOut here without $script:
    # works the same way Invoke-MailboxEdit/Invoke-GroupEdit already do --
    # they're script-scope (top-level) variables by the time any button
    # handler calls into this function.
    param([string]$Smtp)
    $grdRules.Rows.Clear()
    $result = Get-InboxRulesForMailbox -Smtp $Smtp
    if ($result.Error) {
        $txtOut.Text = $result.Error
        $script:CurrentRules = @()
        return
    }
    $script:CurrentRules = $result.Rules
    foreach ($r in $result.Rules) {
        $desc = if ($r.Description) { ($r.Description -replace '\s+', ' ').Trim() } else { '' }
        [void]$grdRules.Rows.Add($r.Name, $desc, $r.Enabled, $r.Identity.ToString())
    }
}
function Show-RuleDetailForm {
    param($Rule, [string]$Smtp)
    $ruleForm = New-Object System.Windows.Forms.Form
    $ruleForm.Text = "Inbox Rule: $($Rule.Name)"
    $ruleForm.Size = New-Object System.Drawing.Size(560, 430)
    $ruleForm.StartPosition = "CenterParent"
    $ruleForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = "Rule Name : $($Rule.Name)"
    $lblName.Location = New-Object System.Drawing.Point(12, 12)
    $lblName.AutoSize = $true
    $ruleForm.Controls.Add($lblName)
    $lblMailbox = New-Object System.Windows.Forms.Label
    $lblMailbox.Text = "Mailbox   : $Smtp"
    $lblMailbox.Location = New-Object System.Drawing.Point(12, 32)
    $lblMailbox.AutoSize = $true
    $ruleForm.Controls.Add($lblMailbox)
    $lblEnabled = New-Object System.Windows.Forms.Label
    $lblEnabled.Text = "Enabled   : $($Rule.Enabled)"
    $lblEnabled.Location = New-Object System.Drawing.Point(12, 52)
    $lblEnabled.AutoSize = $true
    $lblEnabled.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $ruleForm.Controls.Add($lblEnabled)
    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = "Full Description:"
    $lblDesc.Location = New-Object System.Drawing.Point(12, 78)
    $lblDesc.AutoSize = $true
    $ruleForm.Controls.Add($lblDesc)
    $txtDesc = New-Object System.Windows.Forms.TextBox
    $txtDesc.Multiline = $true
    $txtDesc.ReadOnly = $true
    $txtDesc.ScrollBars = "Vertical"
    $txtDesc.Location = New-Object System.Drawing.Point(12, 98)
    $txtDesc.Size = New-Object System.Drawing.Size(520, 200)
    $txtDesc.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtDesc.Text = $Rule.Description
    $ruleForm.Controls.Add($txtDesc)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.Location = New-Object System.Drawing.Point(12, 304)
    $lblStatus.Size = New-Object System.Drawing.Size(520, 34)
    $ruleForm.Controls.Add($lblStatus)
    $btnEnable = New-Object System.Windows.Forms.Button
    $btnEnable.Text = "Enable"
    $btnEnable.Location = New-Object System.Drawing.Point(12, 344)
    $btnEnable.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnEnable)
    $btnDisable = New-Object System.Windows.Forms.Button
    $btnDisable.Text = "Disable"
    $btnDisable.Location = New-Object System.Drawing.Point(120, 344)
    $btnDisable.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnDisable)
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "Delete"
    $btnDelete.Location = New-Object System.Drawing.Point(228, 344)
    $btnDelete.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnDelete)
    $btnRuleClose = New-Object System.Windows.Forms.Button
    $btnRuleClose.Text = "Close"
    $btnRuleClose.Location = New-Object System.Drawing.Point(432, 344)
    $btnRuleClose.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnRuleClose)
    $ruleIdentity = $Rule.Identity.ToString()
    $btnEnable.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Enable Inbox rule '$($Rule.Name)' on mailbox '$Smtp'?",
            "Confirm Enable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Enable-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Enabled successfully."
            $lblEnabled.Text = "Enabled   : True"
            Write-CompletedLog -Task "Enabled InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Enable-InboxRule'
        } finally {
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnDisable.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Disable Inbox rule '$($Rule.Name)' on mailbox '$Smtp'?",
            "Confirm Disable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Disable-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Disabled successfully."
            $lblEnabled.Text = "Enabled   : False"
            Write-CompletedLog -Task "Disabled InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Disable-InboxRule'
        } finally {
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnDelete.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Permanently delete Inbox rule '$($Rule.Name)' on mailbox '$Smtp'? This cannot be undone.",
            "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Remove-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            Write-CompletedLog -Task "Deleted InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
            [System.Windows.Forms.MessageBox]::Show("Rule deleted.", "Deleted", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            $ruleForm.Close()
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-InboxRule'
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnRuleClose.Add_Click({ $ruleForm.Close() })
    [void]$ruleForm.ShowDialog($form)
}
#endregion
#region Actions - group ---------------------------------------------------------
function Get-GroupState {
    param([string]$Type, [string]$Smtp, [string]$ADGroupDN, [string]$ADServer)
    $state = [ordered]@{
        DisplayName = $null
        PrimarySmtp = $Smtp
        GroupType   = $null
        Hidden      = $null
        Owners      = @()
        Members     = @()
        Errors      = @()
    }
    if ($Type -eq 'UnifiedGroup') {
        try {
            $g = Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
            $state.DisplayName = $g.DisplayName
            $state.GroupType   = 'Microsoft 365 Group (Unified)'
            $state.Hidden      = $g.HiddenFromAddressListsEnabled
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroup (core properties)')
        }
        try {
            $owners = Invoke-EXOCommand -Name 'Get-UnifiedGroupLinks' -Parameters @{ Identity = $Smtp; LinkType = 'Owners'; ErrorAction = 'Stop' }
            $state.Owners = @($owners | ForEach-Object { $_.PrimarySmtpAddress })
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroupLinks -LinkType Owners')
        }
        try {
            $members = Invoke-EXOCommand -Name 'Get-UnifiedGroupLinks' -Parameters @{ Identity = $Smtp; LinkType = 'Members'; ErrorAction = 'Stop' }
            $state.Members = @($members | ForEach-Object { $_.PrimarySmtpAddress })
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroupLinks -LinkType Members')
        }
    }
    elseif ($Type -eq 'OnpremGroup') {
        try {
            $g = Get-ADGroup -Identity $ADGroupDN -Server $ADServer `
                    -Properties mail, ManagedBy, GroupCategory, GroupScope, msExchHideFromAddressLists -ErrorAction Stop
            $state.DisplayName = $g.Name
            $state.PrimarySmtp = $g.mail
            $state.GroupType   = if ($g.GroupCategory -eq 'Security') { 'Mail-Enabled Security Group (on-prem, synced)' }
                                  else { 'Distribution Group (on-prem, synced)' }
            $state.Hidden      = [bool]$g.msExchHideFromAddressLists
            if ($g.ManagedBy) {
                try {
                    $owner = Get-ADUser -Identity $g.ManagedBy -Server $ADServer -ErrorAction Stop
                    $state.Owners = @($owner.SamAccountName)
                } catch {
                    $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADUser (resolving ManagedBy owner)')
                }
            }
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADGroup (core properties)')
        }
        try {
            $members = Get-ADGroupMember -Identity $ADGroupDN -Server $ADServer -ErrorAction Stop
            $state.Members = @($members | ForEach-Object { $_.SamAccountName })
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADGroupMember')
        }
    }
    return [PSCustomObject]$state
}
function Format-GroupState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading (Group) ---")
    [void]$sb.AppendLine("  Display Name        : $($State.DisplayName)")
    [void]$sb.AppendLine("  Primary SMTP        : $($State.PrimarySmtp)")
    [void]$sb.AppendLine("  Group Type          : $($State.GroupType)  (read-only -- see script header)")
    [void]$sb.AppendLine("  Hidden From GAL     : $($State.Hidden)")
    [void]$sb.AppendLine("  Owner(s)            :")
    if ($State.Owners.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.Owners | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    [void]$sb.AppendLine("  Members ($($State.Members.Count)) :")
    if ($State.Members.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.Members | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    if ($State.Errors.Count -gt 0) {
        [void]$sb.AppendLine("  --- Field errors (partial data above is still valid) ---")
        $State.Errors | ForEach-Object { [void]$sb.AppendLine("  $_") }
    }
    return $sb.ToString()
}
# Unified Group (Microsoft 365 Group) edits -- EXO, resolved via Invoke-EXOCommand
function Add-UnifiedOwner {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Owners'; Links = $Grantee; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-UnifiedOwner {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Owners'; Links = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Add-UnifiedMember {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Members'; Links = $Grantee; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-UnifiedMember {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Members'; Links = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
# On-prem group edits -- ActiveDirectory module, unambiguous cmdlets
function Set-OnpremOwner {
    param([string]$ADGroupDN, [string]$Grantee, [string]$ADServer)
    # Replaces the single AD ManagedBy value -- see header "KNOWN LIMITATION".
    Set-ADGroup -Identity $ADGroupDN -ManagedBy $Grantee -Server $ADServer -ErrorAction Stop
}
function Add-OnpremMember {
    param([string]$ADGroupDN, [string]$Grantee, [string]$ADServer)
    Add-ADGroupMember -Identity $ADGroupDN -Members $Grantee -Server $ADServer -ErrorAction Stop
}
function Remove-OnpremMember {
    param([string]$ADGroupDN, [string]$Grantee, [string]$ADServer)
    Remove-ADGroupMember -Identity $ADGroupDN -Members $Grantee -Server $ADServer -Confirm:$false -ErrorAction Stop
}
#endregion
#region Actions - out of office ---------------------------------------------------
function Get-OutOfOfficeConfig {
    param([string]$Smtp)
    Invoke-EXOCommand -Name 'Get-MailboxAutoReplyConfiguration' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
}
function Set-OutOfOfficeConfig {
    param(
        [string]$Smtp,
        [string]$AutoReplyState,
        [string]$ExternalAudience,
        [string]$InternalMessage,
        [string]$ExternalMessage,
        [Nullable[DateTime]]$StartTime,
        [Nullable[DateTime]]$EndTime
    )
    $params = @{
        Identity         = $Smtp
        AutoReplyState   = $AutoReplyState
        ExternalAudience = $ExternalAudience
        InternalMessage  = $InternalMessage
        ExternalMessage  = $ExternalMessage
        Confirm          = $false
        ErrorAction      = 'Stop'
    }
    if ($AutoReplyState -eq 'Scheduled') {
        if ($StartTime) { $params.StartTime = $StartTime }
        if ($EndTime)   { $params.EndTime   = $EndTime }
    }
    Invoke-EXOCommand -Name 'Set-MailboxAutoReplyConfiguration' -Parameters $params | Out-Null
}
function Show-OutOfOfficeForm {
    param([string]$Smtp)
    $oofForm = New-Object System.Windows.Forms.Form
    $oofForm.Text = "Out of Office: $Smtp"
    $oofForm.Size = New-Object System.Drawing.Size(560, 560)
    $oofForm.StartPosition = "CenterParent"
    $oofForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblState = New-Object System.Windows.Forms.Label
    $lblState.Text = "State:"
    $lblState.Location = New-Object System.Drawing.Point(12, 15)
    $lblState.AutoSize = $true
    $oofForm.Controls.Add($lblState)
    $cboState = New-Object System.Windows.Forms.ComboBox
    $cboState.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Disabled', 'Enabled', 'Scheduled') | ForEach-Object { [void]$cboState.Items.Add($_) }
    $cboState.Location = New-Object System.Drawing.Point(140, 12)
    $cboState.Size = New-Object System.Drawing.Size(150, 20)
    $oofForm.Controls.Add($cboState)
    $lblAudience = New-Object System.Windows.Forms.Label
    $lblAudience.Text = "External audience:"
    $lblAudience.Location = New-Object System.Drawing.Point(300, 15)
    $lblAudience.AutoSize = $true
    $oofForm.Controls.Add($lblAudience)
    $cboAudience = New-Object System.Windows.Forms.ComboBox
    $cboAudience.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('None', 'Known', 'All') | ForEach-Object { [void]$cboAudience.Items.Add($_) }
    $cboAudience.Location = New-Object System.Drawing.Point(420, 12)
    $cboAudience.Size = New-Object System.Drawing.Size(100, 20)
    $oofForm.Controls.Add($cboAudience)
    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = "Start (Scheduled only, e.g. 2026-07-10 08:00):"
    $lblStart.Location = New-Object System.Drawing.Point(12, 45)
    $lblStart.AutoSize = $true
    $oofForm.Controls.Add($lblStart)
    $txtStart = New-Object System.Windows.Forms.TextBox
    $txtStart.Location = New-Object System.Drawing.Point(280, 42)
    $txtStart.Size = New-Object System.Drawing.Size(240, 20)
    $txtStart.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtStart)
    $lblEnd = New-Object System.Windows.Forms.Label
    $lblEnd.Text = "End (Scheduled only):"
    $lblEnd.Location = New-Object System.Drawing.Point(12, 70)
    $lblEnd.AutoSize = $true
    $oofForm.Controls.Add($lblEnd)
    $txtEnd = New-Object System.Windows.Forms.TextBox
    $txtEnd.Location = New-Object System.Drawing.Point(280, 67)
    $txtEnd.Size = New-Object System.Drawing.Size(240, 20)
    $txtEnd.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtEnd)
    $lblInt = New-Object System.Windows.Forms.Label
    $lblInt.Text = "Internal message:"
    $lblInt.Location = New-Object System.Drawing.Point(12, 98)
    $lblInt.AutoSize = $true
    $oofForm.Controls.Add($lblInt)
    $txtInt = New-Object System.Windows.Forms.TextBox
    $txtInt.Multiline = $true
    $txtInt.ScrollBars = "Vertical"
    $txtInt.Location = New-Object System.Drawing.Point(12, 118)
    $txtInt.Size = New-Object System.Drawing.Size(520, 140)
    $txtInt.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtInt)
    $lblExt = New-Object System.Windows.Forms.Label
    $lblExt.Text = "External message:"
    $lblExt.Location = New-Object System.Drawing.Point(12, 264)
    $lblExt.AutoSize = $true
    $oofForm.Controls.Add($lblExt)
    $txtExt = New-Object System.Windows.Forms.TextBox
    $txtExt.Multiline = $true
    $txtExt.ScrollBars = "Vertical"
    $txtExt.Location = New-Object System.Drawing.Point(12, 284)
    $txtExt.Size = New-Object System.Drawing.Size(520, 140)
    $txtExt.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtExt)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.Location = New-Object System.Drawing.Point(12, 430)
    $lblStatus.Size = New-Object System.Drawing.Size(520, 40)
    $oofForm.Controls.Add($lblStatus)
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save"
    $btnSave.Location = New-Object System.Drawing.Point(12, 474)
    $btnSave.Size = New-Object System.Drawing.Size(100, 28)
    $oofForm.Controls.Add($btnSave)
    $btnOofClose = New-Object System.Windows.Forms.Button
    $btnOofClose.Text = "Close"
    $btnOofClose.Location = New-Object System.Drawing.Point(120, 474)
    $btnOofClose.Size = New-Object System.Drawing.Size(100, 28)
    $oofForm.Controls.Add($btnOofClose)
    function Update-OofFormFromCurrent {
        try {
            $cfg = Get-OutOfOfficeConfig -Smtp $Smtp
            $cboState.SelectedItem    = $cfg.AutoReplyState.ToString()
            $cboAudience.SelectedItem = $cfg.ExternalAudience.ToString()
            $txtStart.Text = if ($cfg.StartTime) { $cfg.StartTime.ToString() } else { '' }
            $txtEnd.Text   = if ($cfg.EndTime)   { $cfg.EndTime.ToString() }   else { '' }
            $txtInt.Text   = $cfg.InternalMessage
            $txtExt.Text   = $cfg.ExternalMessage
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Loaded current settings."
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxAutoReplyConfiguration'
        }
    }
    Update-OofFormFromCurrent
    $btnSave.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Save Out of Office settings for '$Smtp'? State: $($cboState.SelectedItem)",
            "Confirm Save", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnSave.Enabled = $false
        try {
            $st = $null; $et = $null
            if ($cboState.SelectedItem -eq 'Scheduled') {
                if ($txtStart.Text.Trim()) { $st = [DateTime]::Parse($txtStart.Text.Trim()) }
                if ($txtEnd.Text.Trim())   { $et = [DateTime]::Parse($txtEnd.Text.Trim()) }
            }
            Set-OutOfOfficeConfig -Smtp $Smtp -AutoReplyState $cboState.SelectedItem -ExternalAudience $cboAudience.SelectedItem `
                -InternalMessage $txtInt.Text -ExternalMessage $txtExt.Text -StartTime $st -EndTime $et
            Write-CompletedLog -Task "Set OutOfOffice $Smtp (EXO; state: $($cboState.SelectedItem))"
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Saved successfully."
            Update-OofFormFromCurrent
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-MailboxAutoReplyConfiguration'
        } finally {
            $btnSave.Enabled = $true
        }
    })
    $btnOofClose.Add_Click({ $oofForm.Close() })
    [void]$oofForm.ShowDialog($form)
}
#endregion
#region Actions - calendar permissions --------------------------------------------
function Get-CalendarFolderIdentity {
    # Resolves the actual calendar folder rather than hardcoding ":\Calendar",
    # since the folder can be renamed or (rarely) localized. Falls back to the
    # first Calendar-scope folder if none matches the default root path.
    param([string]$Smtp)
    $folders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{ Identity = $Smtp; FolderScope = 'Calendar'; ErrorAction = 'Stop' })
    if ($folders.Count -eq 0) { throw "No Calendar-scope folder found for $Smtp." }
    $primary = $folders | Where-Object { $_.FolderPath -eq '/Calendar' } | Select-Object -First 1
    if (-not $primary) { $primary = $folders | Select-Object -First 1 }
    return "$($Smtp):\$($primary.Name)"
}
function Get-CalendarPermissions {
    param([string]$Smtp)
    try {
        $calId = Get-CalendarFolderIdentity -Smtp $Smtp
        $perms = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $calId; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ CalendarIdentity = $calId; Permissions = $perms; Error = $null }
    } catch {
        return [PSCustomObject]@{ CalendarIdentity = $null; Permissions = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxFolderPermission (Calendar)') }
    }
}
function Set-CalendarPermission {
    # Add-MailboxFolderPermission fails if the user already has an entry, so
    # check first (best-effort) rather than parsing the resulting error text.
    # If the check itself errors for any reason, fall through to Add-; a
    # genuine problem will surface clearly from that call instead.
    param([string]$CalendarIdentity, [string]$Grantee, [string]$Role)
    $existing = @()
    try {
        $existing = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $CalendarIdentity; User = $Grantee; ErrorAction = 'Stop' })
    } catch {
        $existing = @()
    }
    if ($existing.Count -gt 0) {
        Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Updated'
    } else {
        Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Added'
    }
}
function Remove-CalendarPermission {
    # Remove-MailboxFolderPermission removes ALL permissions for this user on
    # the folder -- there's no selective/partial removal, per Microsoft Learn.
    param([string]$CalendarIdentity, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxFolderPermission' -Parameters @{
        Identity = $CalendarIdentity; User = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Show-CalendarPermissionsForm {
    param([string]$Smtp)
    $calForm = New-Object System.Windows.Forms.Form
    $calForm.Text = "Calendar Permissions: $Smtp"
    $calForm.Size = New-Object System.Drawing.Size(620, 480)
    $calForm.StartPosition = "CenterParent"
    $calForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblCalId = New-Object System.Windows.Forms.Label
    $lblCalId.Text = "Calendar folder:"
    $lblCalId.Location = New-Object System.Drawing.Point(12, 12)
    $lblCalId.AutoSize = $true
    $calForm.Controls.Add($lblCalId)
    $grdCal = New-Object System.Windows.Forms.DataGridView
    $grdCal.Location = New-Object System.Drawing.Point(12, 34)
    $grdCal.Size = New-Object System.Drawing.Size(580, 220)
    $grdCal.ReadOnly = $true
    $grdCal.AllowUserToAddRows = $false
    $grdCal.AllowUserToDeleteRows = $false
    $grdCal.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdCal.MultiSelect = $false
    $grdCal.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdCal.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdCal.Columns.Add('User', 'User')
    [void]$grdCal.Columns.Add('AccessRights', 'Access Rights')
    foreach ($c in $grdCal.Columns) { $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }
    $calForm.Controls.Add($grdCal)
    $lblGrantee = New-Object System.Windows.Forms.Label
    $lblGrantee.Text = "User (email or logon):"
    $lblGrantee.Location = New-Object System.Drawing.Point(12, 266)
    $lblGrantee.AutoSize = $true
    $calForm.Controls.Add($lblGrantee)
    $txtGrantee = New-Object System.Windows.Forms.TextBox
    $txtGrantee.Location = New-Object System.Drawing.Point(160, 263)
    $txtGrantee.Size = New-Object System.Drawing.Size(220, 20)
    $txtGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
    $calForm.Controls.Add($txtGrantee)
    $lblRole = New-Object System.Windows.Forms.Label
    $lblRole.Text = "Role:"
    $lblRole.Location = New-Object System.Drawing.Point(390, 266)
    $lblRole.AutoSize = $true
    $calForm.Controls.Add($lblRole)
    $cboRole = New-Object System.Windows.Forms.ComboBox
    $cboRole.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Reviewer', 'Editor', 'Owner') | ForEach-Object { [void]$cboRole.Items.Add($_) }
    $cboRole.SelectedIndex = 0
    $cboRole.Location = New-Object System.Drawing.Point(430, 263)
    $cboRole.Size = New-Object System.Drawing.Size(120, 20)
    $calForm.Controls.Add($cboRole)
    $lblCalStatus = New-Object System.Windows.Forms.Label
    $lblCalStatus.Text = ""
    $lblCalStatus.Location = New-Object System.Drawing.Point(12, 296)
    $lblCalStatus.Size = New-Object System.Drawing.Size(580, 40)
    $calForm.Controls.Add($lblCalStatus)
    $btnAddPerm = New-Object System.Windows.Forms.Button
    $btnAddPerm.Text = "Add / Update Permission"
    $btnAddPerm.Location = New-Object System.Drawing.Point(12, 340)
    $btnAddPerm.Size = New-Object System.Drawing.Size(170, 28)
    $calForm.Controls.Add($btnAddPerm)
    $btnRemovePerm = New-Object System.Windows.Forms.Button
    $btnRemovePerm.Text = "Remove Selected"
    $btnRemovePerm.Location = New-Object System.Drawing.Point(190, 340)
    $btnRemovePerm.Size = New-Object System.Drawing.Size(150, 28)
    $calForm.Controls.Add($btnRemovePerm)
    $btnCalClose = New-Object System.Windows.Forms.Button
    $btnCalClose.Text = "Close"
    $btnCalClose.Location = New-Object System.Drawing.Point(500, 340)
    $btnCalClose.Size = New-Object System.Drawing.Size(90, 28)
    $calForm.Controls.Add($btnCalClose)
    $script:CalendarIdentityForForm = $null
    function Update-CalGridFromCurrent {
        $grdCal.Rows.Clear()
        $result = Get-CalendarPermissions -Smtp $Smtp
        if ($result.Error) {
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = $result.Error
            $lblCalId.Text = "Calendar folder: (unresolved)"
            return
        }
        $script:CalendarIdentityForForm = $result.CalendarIdentity
        $lblCalId.Text = "Calendar folder: $($result.CalendarIdentity)"
        foreach ($p in $result.Permissions) {
            [void]$grdCal.Rows.Add($p.User.ToString(), ($p.AccessRights -join ', '))
        }
    }
    Update-CalGridFromCurrent
    $btnAddPerm.Add_Click({
        $grantee = $txtGrantee.Text.Trim()
        if (-not $grantee) { $lblCalStatus.Text = "Enter a user first."; return }
        if (-not $script:CalendarIdentityForForm) { $lblCalStatus.Text = "Calendar folder not resolved -- see status above."; return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Grant '$($cboRole.SelectedItem)' calendar access to '$grantee' on '$Smtp'?",
            "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnAddPerm.Enabled = $false; $btnRemovePerm.Enabled = $false
        try {
            $action = Set-CalendarPermission -CalendarIdentity $script:CalendarIdentityForForm -Grantee $grantee -Role $cboRole.SelectedItem
            Write-CompletedLog -Task "$action CalendarPermission $Smtp (EXO; user: $grantee; role: $($cboRole.SelectedItem))"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Black
            $lblCalStatus.Text = "$action permission for $grantee."
            Update-CalGridFromCurrent
        } catch {
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Calendar permission grant'
        } finally {
            $btnAddPerm.Enabled = $true; $btnRemovePerm.Enabled = $true
        }
    })
    $btnRemovePerm.Add_Click({
        if ($grdCal.SelectedRows.Count -eq 0) { $lblCalStatus.Text = "Select a row first."; return }
        $grantee = $grdCal.SelectedRows[0].Cells['User'].Value
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove all calendar permissions for '$grantee' on '$Smtp'?",
            "Confirm Remove", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnAddPerm.Enabled = $false; $btnRemovePerm.Enabled = $false
        try {
            Remove-CalendarPermission -CalendarIdentity $script:CalendarIdentityForForm -Grantee $grantee
            Write-CompletedLog -Task "Removed CalendarPermission $Smtp (EXO; user: $grantee)"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Black
            $lblCalStatus.Text = "Removed permission for $grantee."
            Update-CalGridFromCurrent
        } catch {
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-MailboxFolderPermission (Calendar)'
        } finally {
            $btnAddPerm.Enabled = $true; $btnRemovePerm.Enabled = $true
        }
    })
    $btnCalClose.Add_Click({ $calForm.Close() })
    [void]$calForm.ShowDialog($form)
}
#endregion
#region Actions - mailbox health ---------------------------------------------------
function Get-MailboxDLMemberships {
    # REWRITTEN 2026-07-09c (field report: the original version iterated
    # every DL in the tenant with Get-DistributionGroupMember and froze the
    # UI -- see the two 2026-07-09 revision notes in the header. The
    # operator then asked whether the EAC "Group membership" pane does the
    # same full-tenant scan; it does not.)
    #
    # EAC's Group Membership pane, and Set-Mailbox et al under the hood, get
    # this from Active Directory's own "memberOf" attribute on the user
    # object -- a backlink that AD itself maintains automatically every time
    # a group's "member" attribute changes, so reading it is a single
    # indexed object lookup, not a search across every group. This function
    # now does the same thing via Get-ADUser -Properties memberOf (both the
    # -Properties parameter and its behavior confirmed against Microsoft
    # Learn), matched to the mailbox via its "mail" attribute the same way
    # Resolve-Target already matches on-prem GROUPS by mail elsewhere in
    # this script. Cost is now one Get-ADUser call plus one Get-ADGroup call
    # per group the user actually belongs to (typically a handful to a few
    # dozen) -- bounded by the user's real membership count, not by how many
    # DLs exist in the tenant.
    #
    # Two deliberate scope limits, both consistent with how this tool
    # already treats on-prem groups elsewhere:
    #  - Direct membership only, not nested/recursive. memberOf lists groups
    #    the user is a DIRECT member of; if that group is itself nested
    #    inside another group, the outer group is not included. This is what
    #    EAC's Group Membership pane and the on-prem AD "Member Of" tab both
    #    show. Get-ADPrincipalGroupMembership resolves nested membership too
    #    but requires a Global Catalog and behaves differently in
    #    multi-domain forests (per Microsoft Learn) -- not used here to keep
    #    behavior simple and predictable across the Franchise/HQ environment
    #    split this tool already has.
    #  - On-prem-synced Distribution / Mail-Enabled Security groups only,
    #    via the ActiveDirectory module -- Microsoft 365 Group membership is
    #    a separate object type handled by this tool's own Group lookup
    #    path (Get-UnifiedGroupLinks), not included here.
    #  - A mailbox with no matching on-prem AD user (cloud-only mailbox,
    #    never synced) has no memberOf to read -- $result.NoADAccount is set
    #    rather than silently returning an empty list, so that case isn't
    #    mistaken for "this user is in zero groups."
    #
    # -ProgressAction, if supplied, is invoked once per group-DN being
    # resolved as & $ProgressAction $currentIndex $totalCount $groupDN.
    # -ShouldCancel, if supplied, is checked after each one; returning
    # $true stops early and sets $result.Cancelled, keeping whatever partial
    # results were already resolved. Both are vestigial now that this is a
    # bounded, typically-fast operation, but kept since a handful of memberOf
    # entries can still point at cross-domain groups that are slow to
    # resolve individually.
    param(
        [string]$Smtp,
        [string]$ADServer,
        [scriptblock]$ProgressAction,
        [scriptblock]$ShouldCancel
    )
    $result = [ordered]@{
        Memberships = @()
        ScanErrors  = @()
        FatalError  = $null
        Cancelled   = $false
        NoADAccount = $false
    }
    try {
        $adUsers = @(Get-ADUser -Filter "mail -eq '$Smtp'" -Properties memberOf -Server $ADServer -ErrorAction Stop)
    } catch {
        $result.FatalError = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADUser (DL membership lookup)')
        return [PSCustomObject]$result
    }
    if ($adUsers.Count -eq 0) {
        $result.NoADAccount = $true
        return [PSCustomObject]$result
    }
    if ($adUsers.Count -gt 1) {
        $result.FatalError = "$($adUsers.Count) on-prem AD users matched mail '$Smtp'. Expected 0 or 1 -- resolve the duplicate mail attribute before using this feature."
        return [PSCustomObject]$result
    }
    $dns = @($adUsers[0].memberOf)
    $found = @()
    $i = 0
    foreach ($dn in $dns) {
        $i++
        if ($ProgressAction) {
            try { & $ProgressAction $i $dns.Count $dn } catch { }
        }
        try {
            $g = Get-ADGroup -Identity $dn -Server $ADServer -ErrorAction Stop
            $found += $g.Name
        } catch {
            $result.ScanErrors += "$($dn): $($_.Exception.Message)"
        }
        if ($ShouldCancel -and (& $ShouldCancel)) {
            $result.Cancelled = $true
            break
        }
    }
    $result.Memberships = $found
    return [PSCustomObject]$result
}
function Get-MailboxHealthSnapshot {
    # Read-only composite view. Reuses Get-InboxRulesForMailbox (Inbox Rules
    # feature) for the rule list/count rather than querying it twice.
    #
    # Interpretation / verification notes (documented since these terms are
    # not single, unambiguous Exchange fields, and per standards Section 2
    # every cmdlet/property used for a new capability is checked against
    # Microsoft Learn before delivery):
    #  - "Protocol usage" is shown as Get-CASMailbox protocol ENABLEMENT
    #    flags (OWA/ActiveSync/POP/IMAP/MAPI/EWS), not actual traffic volume.
    #  - "Delegates" now lists Full Access and Send As grantees by name
    #    (each from its own cmdlet -- see Get-MailboxState), not just a
    #    combined count. Does NOT include Outlook's separate calendar/folder
    #    delegate list.
    #  - DL memberships are NOT part of this snapshot -- see
    #    Get-MailboxDLMemberships / Show-DLMembershipsForm / the "DL
    #    Memberships" button instead. That scan iterates every DL in the
    #    tenant and can take a while; running it automatically inside Health
    #    made the whole tool appear to hang (2026-07-09 field report), so
    #    it's now a separate on-demand action with its own progress display.
    #  - Quota (ProhibitSendQuota, ProhibitSendReceiveQuota), Forwarding
    #    (ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward),
    #    Litigation Hold (LitigationHoldEnabled, InPlaceHolds), and Retention
    #    Hold (RetentionHoldEnabled) are all confirmed Get-Mailbox output
    #    properties per Microsoft Learn's "Property sets in Exchange Online
    #    PowerShell module cmdlets" (Quota / Delivery / Hold / Retention
    #    property sets respectively). Quota values are shown as raw
    #    ByteQuantifiedSize text next to Mailbox Size for manual comparison
    #    rather than computed as a percentage -- the ByteQuantifiedSize/
    #    Unlimited<T> numeric conversion API was not verified against
    #    Microsoft Learn, so this script does not guess at it.
    #  - AccountDisabled: Microsoft Learn confirms -AccountDisabled as a
    #    Set-Mailbox parameter but this script did NOT find Microsoft Learn
    #    confirmation that Get-Mailbox returns an AccountDisabled OUTPUT
    #    property (Microsoft's own Get-Mailbox page recommends Get-EXOMailbox
    #    instead, whose documented property sets don't list it either).
    #    Flagged here as UNVERIFIED per standards Section 2 ("if
    #    verification is unavailable, clearly say so before delivery") --
    #    confirm this field is meaningful on first use; if the property
    #    doesn't exist on the returned object, PowerShell returns $null here
    #    rather than throwing, so a wrong/missing value fails quietly. Treat
    #    a blank value as "unknown," not as "account is enabled."
    #  - "Suspicious Inbox Rules" flags ENABLED rules with a forwarding,
    #    redirect, or delete action (ForwardTo/ForwardAsAttachmentTo/
    #    RedirectTo/DeleteMessage/PermanentDelete/SoftDeleteMessage -- all
    #    confirmed Get-InboxRule/Set-InboxRule property names on Microsoft
    #    Learn). This is a flag for a human to review as a first-check item
    #    on compromised-account or missing-mail tickets, not a judgment that
    #    the rule is malicious.
    param([string]$Smtp)
    $health = [ordered]@{
        LastLogonTime              = $null
        ItemCount                  = $null
        MailboxSize                = $null
        ProhibitSendQuota          = $null
        ProhibitSendReceiveQuota   = $null
        ArchiveEnabled             = $null
        ArchiveSize                = $null
        AccountDisabled            = $null
        ForwardingSmtpAddress      = $null
        ForwardingAddress          = $null
        DeliverToMailboxAndForward = $null
        LitigationHoldEnabled      = $null
        RetentionHoldEnabled       = $null
        InPlaceHoldsCount          = $null
        FullAccessDelegates        = @()
        SendAsDelegates            = @()
        RuleCount                  = $null
        SuspiciousRules            = @()
        OWAEnabled                 = $null
        ActiveSyncEnabled          = $null
        PopEnabled                 = $null
        ImapEnabled                = $null
        MAPIEnabled                = $null
        EwsEnabled                 = $null
        Errors                     = @()
    }
    try {
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.LastLogonTime = $stats.LastLogonTime
        $health.ItemCount     = $stats.ItemCount
        $health.MailboxSize   = $stats.TotalItemSize
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (health)')
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.ArchiveEnabled             = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        $health.AccountDisabled            = $mbx.AccountDisabled
        $health.ProhibitSendQuota          = $mbx.ProhibitSendQuota
        $health.ProhibitSendReceiveQuota   = $mbx.ProhibitSendReceiveQuota
        $health.ForwardingSmtpAddress      = $mbx.ForwardingSmtpAddress
        $health.ForwardingAddress          = $mbx.ForwardingAddress
        $health.DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
        $health.LitigationHoldEnabled      = $mbx.LitigationHoldEnabled
        $health.RetentionHoldEnabled       = $mbx.RetentionHoldEnabled
        $health.InPlaceHoldsCount          = if ($mbx.InPlaceHolds) { @($mbx.InPlaceHolds).Count } else { 0 }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (quota/forwarding/hold/account status)')
    }
    if ($health.ArchiveEnabled) {
        try {
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $health.ArchiveSize = $astats.TotalItemSize
        } catch {
            $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (health)')
        }
    }
    try {
        $fa = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.User -notlike 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false -and $_.Deny -eq $false -and ($_.AccessRights -contains 'FullAccess') })
        $health.FullAccessDelegates = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access delegates)')
    }
    try {
        $sa = @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.Trustee -notlike 'NT AUTHORITY\SELF' -and $_.AccessControlType -eq 'Allow' })
        $health.SendAsDelegates = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As delegates)')
    }
    try {
        $rules = Get-InboxRulesForMailbox -Smtp $Smtp
        if ($rules.Error) {
            $health.Errors += $rules.Error
        } else {
            $health.RuleCount = $rules.Rules.Count
            $health.SuspiciousRules = @(
                $rules.Rules | Where-Object {
                    $_.Enabled -and (
                        $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or
                        $_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage
                    )
                } | ForEach-Object {
                    $actions = @()
                    if ($_.ForwardTo -or $_.ForwardAsAttachmentTo) { $actions += 'Forward' }
                    if ($_.RedirectTo) { $actions += 'Redirect' }
                    if ($_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage) { $actions += 'Delete' }
                    "$($_.Name)  [$($actions -join ', ')]"
                }
            )
        }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-InboxRule (rule count / suspicious rule scan)')
    }
    try {
        $cas = Invoke-EXOCommand -Name 'Get-CASMailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.OWAEnabled        = $cas.OWAEnabled
        $health.ActiveSyncEnabled = $cas.ActiveSyncEnabled
        $health.PopEnabled        = $cas.PopEnabled
        $health.ImapEnabled       = $cas.ImapEnabled
        $health.MAPIEnabled       = $cas.MAPIEnabled
        $health.EwsEnabled        = $cas.EwsEnabled
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-CASMailbox (protocol access)')
    }
    # DL membership is intentionally NOT scanned here -- see the note above
    # this function and the "DL Memberships" button (Show-DLMembershipsForm).
    return [PSCustomObject]$health
}
function Show-MailboxHealthForm {
    param([string]$Smtp)
    $healthForm = New-Object System.Windows.Forms.Form
    $healthForm.Text = "Mailbox Health: $Smtp"
    $healthForm.Size = New-Object System.Drawing.Size(640, 660)
    $healthForm.StartPosition = "CenterParent"
    $healthForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $txtHealth = New-Object System.Windows.Forms.TextBox
    $txtHealth.Multiline = $true
    $txtHealth.ReadOnly = $true
    $txtHealth.ScrollBars = "Vertical"
    $txtHealth.Location = New-Object System.Drawing.Point(12, 12)
    $txtHealth.Size = New-Object System.Drawing.Size(600, 560)
    $txtHealth.Font = New-Object System.Drawing.Font("Consolas", 9)
    $healthForm.Controls.Add($txtHealth)
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 580)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnRefresh)
    $btnCopyHealth = New-Object System.Windows.Forms.Button
    $btnCopyHealth.Text = "Copy"
    $btnCopyHealth.Location = New-Object System.Drawing.Point(120, 580)
    $btnCopyHealth.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnCopyHealth)
    $btnHealthClose = New-Object System.Windows.Forms.Button
    $btnHealthClose.Text = "Close"
    $btnHealthClose.Location = New-Object System.Drawing.Point(228, 580)
    $btnHealthClose.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnHealthClose)
    # One-shot timer so "Copied!" reverts to "Copy" without needing a second
    # click -- local to this form (closure-captured), same pattern as every
    # other control reference in this function.
    $copyResetTimer = New-Object System.Windows.Forms.Timer
    $copyResetTimer.Interval = 1200
    $copyResetTimer.Add_Tick({
        $btnCopyHealth.Text = "Copy"
        $copyResetTimer.Stop()
    })
    function Update-HealthFromCurrent {
        $btnRefresh.Enabled = $false
        $txtHealth.Text = "Loading mailbox health..."
        $healthForm.Refresh()
        try {
            $h = Get-MailboxHealthSnapshot -Smtp $Smtp
            $sb = New-Object System.Text.StringBuilder
            # --- Potential Issues summary first -- the point of this button is
            # "what do I need to know before I escalate this ticket", so lead
            # with anything that looks like a likely cause rather than making
            # the reader hunt for it further down.
            $issues = @()
            if ($h.AccountDisabled -eq $true) { $issues += "Account sign-in is disabled (AccountDisabled=True) -- unverified as a Get-Mailbox field, confirm before relying on it." }
            if ($h.ForwardingSmtpAddress -or $h.ForwardingAddress) {
                $fwdTarget = if ($h.ForwardingSmtpAddress) { $h.ForwardingSmtpAddress } else { $h.ForwardingAddress }
                $fwdMode = if ($h.DeliverToMailboxAndForward) { "copy kept in mailbox AND forwarded" } else { "forwarded ONLY -- nothing kept in this mailbox" }
                $issues += "Mail forwarding is configured to '$fwdTarget' ($fwdMode)."
            }
            if ($h.LitigationHoldEnabled -eq $true) { $issues += "Litigation Hold is enabled." }
            if ($h.RetentionHoldEnabled -eq $true) { $issues += "Retention Hold is enabled." }
            if ($h.InPlaceHoldsCount -gt 0) { $issues += "$($h.InPlaceHoldsCount) In-Place/eDiscovery hold(s) applied." }
            if ($h.SuspiciousRules.Count -gt 0) { $issues += "$($h.SuspiciousRules.Count) enabled Inbox rule(s) forward, redirect, or delete mail -- see 'Suspicious Inbox Rules' below." }
            [void]$sb.AppendLine("Mailbox                : $Smtp")
            [void]$sb.AppendLine("=== Potential Issues (review before escalating) ===")
            if ($issues.Count -eq 0) { [void]$sb.AppendLine("  None of the checks below flagged anything.") }
            else { $issues | ForEach-Object { [void]$sb.AppendLine("  - $_") } }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Basics ---")
            [void]$sb.AppendLine("Last Logon             : $($h.LastLogonTime)")
            [void]$sb.AppendLine("Item Count              : $($h.ItemCount)")
            [void]$sb.AppendLine("Mailbox Size            : $($h.MailboxSize)")
            [void]$sb.AppendLine("ProhibitSendQuota       : $($h.ProhibitSendQuota)  (compare to Mailbox Size above)")
            [void]$sb.AppendLine("ProhibitSendReceiveQuota: $($h.ProhibitSendReceiveQuota)")
            [void]$sb.AppendLine("Archive Enabled         : $($h.ArchiveEnabled)")
            if ($h.ArchiveEnabled) { [void]$sb.AppendLine("Archive Size            : $($h.ArchiveSize)") }
            [void]$sb.AppendLine("Account Disabled        : $($h.AccountDisabled)  (unverified Get-Mailbox field -- see script header)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Forwarding & Holds ---")
            [void]$sb.AppendLine("Forwarding SMTP Address : $($h.ForwardingSmtpAddress)")
            [void]$sb.AppendLine("Forwarding Address      : $($h.ForwardingAddress)")
            [void]$sb.AppendLine("Deliver And Forward     : $($h.DeliverToMailboxAndForward)")
            [void]$sb.AppendLine("Litigation Hold         : $($h.LitigationHoldEnabled)")
            [void]$sb.AppendLine("Retention Hold          : $($h.RetentionHoldEnabled)")
            [void]$sb.AppendLine("In-Place Holds          : $($h.InPlaceHoldsCount)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Delegates ---")
            [void]$sb.AppendLine("Full Access grantees ($($h.FullAccessDelegates.Count)):")
            if ($h.FullAccessDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.FullAccessDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("Send As grantees ($($h.SendAsDelegates.Count)):")
            if ($h.SendAsDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.SendAsDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("  (Distribution List memberships are not scanned here -- use the 'DL Memberships' button, since a full-tenant scan is too slow to run on every Health click.)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Inbox Rules ---")
            [void]$sb.AppendLine("Inbox Rule Count        : $($h.RuleCount)")
            [void]$sb.AppendLine("Suspicious Inbox Rules (enabled, forward/redirect/delete):")
            if ($h.SuspiciousRules.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.SuspiciousRules | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Protocol Access (enabled/disabled -- not traffic volume) ---")
            [void]$sb.AppendLine("  OWA        : $($h.OWAEnabled)")
            [void]$sb.AppendLine("  ActiveSync : $($h.ActiveSyncEnabled)")
            [void]$sb.AppendLine("  POP        : $($h.PopEnabled)")
            [void]$sb.AppendLine("  IMAP       : $($h.ImapEnabled)")
            [void]$sb.AppendLine("  MAPI       : $($h.MAPIEnabled)")
            [void]$sb.AppendLine("  EWS        : $($h.EwsEnabled)")
            if ($h.Errors.Count -gt 0) {
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("--- Field errors (partial data above is still valid) ---")
                $h.Errors | ForEach-Object { [void]$sb.AppendLine($_) }
            }
            $txtHealth.Text = $sb.ToString()
        } catch {
            $txtHealth.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Health'
        } finally {
            $btnRefresh.Enabled = $true
        }
    }
    Update-HealthFromCurrent
    $btnRefresh.Add_Click({ Update-HealthFromCurrent })
    $btnCopyHealth.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetText($txtHealth.Text)
            $btnCopyHealth.Text = "Copied!"
            $copyResetTimer.Stop()
            $copyResetTimer.Start()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Copy to clipboard failed: $($_.Exception.Message)",
                "Copy", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
    })
    $btnHealthClose.Add_Click({ $healthForm.Close() })
    [void]$healthForm.ShowDialog($form)
}
function Show-DLMembershipsForm {
    # Split out of the Health panel on 2026-07-09, then REWRITTEN 2026-07-09c
    # to read AD's memberOf backlink instead of scanning every DL -- see the
    # Get-MailboxDLMemberships header comment and the header .NOTES revision
    # entries for the full story. Now a single Get-ADUser call plus one
    # Get-ADGroup call per actual membership (typically fast), not a
    # full-tenant scan. Scan/Cancel/progress-label plumbing is kept since a
    # cross-domain group DN can still occasionally be slow to resolve, but
    # this should complete quickly in the common case.
    param([string]$Smtp, [string]$ADServer)
    $dlForm = New-Object System.Windows.Forms.Form
    $dlForm.Text = "DL Memberships: $Smtp"
    $dlForm.Size = New-Object System.Drawing.Size(560, 520)
    $dlForm.StartPosition = "CenterParent"
    $dlForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads this mailbox's on-prem AD account (matched by mail attribute) and its memberOf attribute -- the same backlink EAC's Group Membership pane uses. Direct memberships only, classic DLs / mail-enabled security groups only (not Microsoft 365 Groups)."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(520, 46)
    $dlForm.Controls.Add($lblNote)
    $lblProgress = New-Object System.Windows.Forms.Label
    $lblProgress.Text = "Not started."
    $lblProgress.Location = New-Object System.Drawing.Point(12, 62)
    $lblProgress.Size = New-Object System.Drawing.Size(520, 20)
    $lblProgress.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $dlForm.Controls.Add($lblProgress)
    $txtDL = New-Object System.Windows.Forms.TextBox
    $txtDL.Multiline = $true
    $txtDL.ReadOnly = $true
    $txtDL.ScrollBars = "Vertical"
    $txtDL.Location = New-Object System.Drawing.Point(12, 88)
    $txtDL.Size = New-Object System.Drawing.Size(520, 340)
    $txtDL.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlForm.Controls.Add($txtDL)
    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = "Scan"
    $btnScan.Location = New-Object System.Drawing.Point(12, 436)
    $btnScan.Size = New-Object System.Drawing.Size(100, 28)
    $dlForm.Controls.Add($btnScan)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(120, 436)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 28)
    $btnCancel.Enabled = $false
    $dlForm.Controls.Add($btnCancel)
    $btnDLClose = New-Object System.Windows.Forms.Button
    $btnDLClose.Text = "Close"
    $btnDLClose.Location = New-Object System.Drawing.Point(432, 436)
    $btnDLClose.Size = New-Object System.Drawing.Size(100, 28)
    $dlForm.Controls.Add($btnDLClose)
    $script:DLScanCancelRequested = $false
    $btnScan.Add_Click({
        $btnScan.Enabled = $false
        $btnCancel.Enabled = $true
        $txtDL.Text = ""
        $script:DLScanCancelRequested = $false
        $lblProgress.Text = "Looking up on-prem AD account..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $progressCB = {
                param($i, $total, $dn)
                $lblProgress.Text = "Resolving group $i of $total`: $dn"
                [System.Windows.Forms.Application]::DoEvents()
            }
            $cancelCB = { return $script:DLScanCancelRequested }
            $result = Get-MailboxDLMemberships -Smtp $Smtp -ADServer $ADServer -ProgressAction $progressCB -ShouldCancel $cancelCB
            if ($result.FatalError) {
                $lblProgress.Text = "Failed."
                $txtDL.Text = $result.FatalError
            } elseif ($result.NoADAccount) {
                $lblProgress.Text = "No on-prem AD account found."
                $txtDL.Text = "No on-prem AD user account has a mail attribute matching '$Smtp'. This mailbox may be cloud-only (never synced from on-prem AD) -- DL membership can't be read this way for it. If you believe this account IS synced, double-check the Environment/Domain Controller selection at the top of the main window."
            } else {
                $status = if ($result.Cancelled) { "Cancelled -- partial results shown below." } else { "Done." }
                $lblProgress.Text = $status
                $sb = New-Object System.Text.StringBuilder
                [void]$sb.AppendLine("$status")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("Memberships ($($result.Memberships.Count)):")
                if ($result.Memberships.Count -eq 0) { [void]$sb.AppendLine("  (none found)") }
                else { $result.Memberships | Sort-Object | ForEach-Object { [void]$sb.AppendLine("  $_") } }
                if ($result.ScanErrors.Count -gt 0) {
                    [void]$sb.AppendLine("")
                    [void]$sb.AppendLine("Groups that could not be resolved ($($result.ScanErrors.Count)) -- membership above may be incomplete:")
                    $result.ScanErrors | ForEach-Object { [void]$sb.AppendLine("  $_") }
                }
                $txtDL.Text = $sb.ToString()
            }
        } catch {
            $lblProgress.Text = "Failed."
            $txtDL.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'DL Membership Lookup'
        } finally {
            $btnScan.Enabled = $true
            $btnCancel.Enabled = $false
        }
    })
    $btnCancel.Add_Click({
        $script:DLScanCancelRequested = $true
        $btnCancel.Enabled = $false
        $lblProgress.Text = "$($lblProgress.Text)  (cancelling after the current group...)"
    })
    $btnDLClose.Add_Click({ $dlForm.Close() })
    [void]$dlForm.ShowDialog($form)
}
#endregion
#region Actions - mail attributes --------------------------------------------
# "Mail Attribs" button (2026-07-10). Ports the on-prem AD attribute grid and
# edit mechanics from the operator's "Renaming Script" (on-prem AD name-change
# tool) almost verbatim -- same fixed attribute list, same derived
# PrimarySmtpAddress/proxyAddresses handling, same Keep-as-Alias/Remove prompt,
# same "Add secondary SMTP address" control -- per explicit instruction to
# match that script's behavior exactly. Differences are only what's required
# to fit this tool's existing conventions:
#   - The target user comes from this tool's own Mailbox lookup (Smtp/ADServer
#     already resolved by Resolve-Target / Get-SelectedADServer) instead of a
#     separate UPN-entry/Parse/Environment UI -- Get-MailAttribsADUser matches
#     the AD account by "mail" (Get-ADUser -Filter "mail -eq '$Smtp'"), the
#     exact same match-by-mail pattern already used and verified by
#     Get-MailboxDLMemberships just above.
#   - The grid/dialog live inside one self-contained Show-MailAttribsForm
#     function, matching how every other feature dialog in this script
#     (Show-DLMembershipsForm, Show-MailboxHealthForm, etc.) is built. Because
#     the current AD-user object needs to be read AND reassigned across
#     several separate event handlers (Refresh, the shared edit function used
#     by both double-click and the Edit Attribute button, and Add Secondary)
#     -- not just one -- it is held in $script:MailAttribsCurrentADUser
#     rather than a plain closure variable, for the same reason this file
#     already uses $script:CurrentRules / $script:CurrentTraceResults /
#     $script:DLScanCancelRequested for equivalent cross-handler state: a
#     plain variable assignment inside one event-handler scriptblock does not
#     reliably propagate back out to a sibling handler's scriptblock.
#   - Function/variable names are new (MailAttribs-prefixed where the
#     reference name would otherwise be too generic to safely reuse) to avoid
#     any collision with this file's existing functions; the proxyAddresses
#     helpers and Format-AttributeValue keep their original reference-script
#     names since nothing else in this file defines them.
#   - "All Attribs" (the operator's second requested button, a full
#     every-attribute dump) is a separate follow-up request and is not part of
#     this region.
#
# Cmdlets used here (Get-ADUser -Filter/-Properties/-Server, Set-ADUser
# -DisplayName/-GivenName/-Initials/-Surname/-EmailAddress/-Clear/-Replace,
# Rename-ADObject -NewName) are the exact same calls, in the exact same
# shapes, already used and Microsoft-Learn-verified in the reference script
# this was ported from, and/or already used elsewhere in this file
# (Get-MailboxDLMemberships, Resolve-Target's on-prem group path) -- not
# re-verified independently per standards Section 1(a) (prior verified usage
# stands).
$script:MailAttribsDisplayOrder = @(
    'Company', 'DisplayName', 'name', 'GivenName', 'Initials', 'Surname',
    'mail', 'mailNickname', 'msExchArchiveName', 'PrimarySmtpAddress',
    'proxyAddresses', 'targetAddress', 'UserPrincipalName'
)
$script:MailAttribsDisplayLabels = @{
    'Company'            = 'Office'
    'GivenName'          = 'FirstName'
    'Surname'            = 'LastName'
    'PrimarySmtpAddress' = 'Primary SMTP'
}
# UserPrincipalName, Company ("Office"), name, and proxyAddresses are
# read-only/context-only, exactly as in the reference script -- name is the
# object's CN/RDN (out of scope for a mail-attributes tool) and proxyAddresses
# is only ever changed narrowly, via PrimarySmtpAddress and "Add Secondary".
$script:MailAttribsEditableAttributes = @(
    'DisplayName', 'GivenName', 'Initials', 'Surname', 'mail', 'mailNickname',
    'msExchArchiveName', 'PrimarySmtpAddress', 'targetAddress'
)
$script:MailAttribsNeverBlank = @('name', 'PrimarySmtpAddress')
# Matches the AD account to this mailbox by its "mail" attribute -- the same
# lookup Get-MailboxDLMemberships already uses just above, so a mailbox
# without a synced on-prem account throws the same style of message rather
# than silently showing an empty grid.
function Get-MailAttribsADUser {
    param(
        [Parameter(Mandatory)] [string]$Smtp,
        [string]$ADServer
    )
    $safeSmtp = $Smtp.Replace("'", "''")
    $params = @{
        Filter      = "mail -eq '$safeSmtp'"
        Properties  = @('Company', 'DisplayName', 'GivenName', 'Initials', 'Surname', 'mail', 'mailNickname',
                         'msExchArchiveName', 'Name', 'proxyAddresses', 'targetAddress',
                         'UserPrincipalName', 'DistinguishedName', 'SamAccountName')
        ErrorAction = 'Stop'
    }
    if ($ADServer) { $params['Server'] = $ADServer }
    $results = @(Get-ADUser @params)
    if ($results.Count -eq 0) {
        throw "No on-prem AD user account has a mail attribute matching '$Smtp'. This mailbox may be cloud-only (never synced from on-prem AD), or the Environment/Domain Controller selection at the top of the main window may need to change."
    }
    if ($results.Count -gt 1) {
        throw "$($results.Count) on-prem AD users matched mail '$Smtp'. Expected 0 or 1 -- resolve the duplicate mail attribute before using this feature."
    }
    return $results[0]
}
# Reads the primary SMTP address out of proxyAddresses -- the entry prefixed
# "SMTP:" in uppercase (Exchange convention; lowercase "smtp:" is secondary).
function Get-PrimarySmtpFromProxyAddresses {
    param($ProxyAddresses)
    foreach ($addr in @($ProxyAddresses)) {
        if ("$addr" -cmatch '^SMTP:(.+)$') { return $Matches[1] }
    }
    return ''
}
# Builds a new proxyAddresses array with $NewPrimary as the SMTP: entry.
# $KeepOldPrimaryAsAlias controls what happens to the previous primary (the
# operator's answer to Show-KeepOrRemoveOldPrimaryDialog): kept as a demoted
# smtp: entry, or dropped entirely. An existing secondary equal to the new
# primary is removed first so the address never appears twice.
function Set-PrimarySmtpInProxyAddresses {
    param(
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewPrimary,
        [bool]$KeepOldPrimaryAsAlias = $true
    )
    $kept = New-Object System.Collections.Generic.List[string]
    $oldPrimary = $null
    foreach ($addr in @($CurrentProxyAddresses)) {
        $s = "$addr"
        if ($s -cmatch '^SMTP:(.+)$') {
            $oldPrimary = $Matches[1]
            continue
        }
        if ($s -match '^smtp:(.+)$' -and $Matches[1] -ieq $NewPrimary) {
            continue
        }
        [void]$kept.Add($s)
    }
    if ($KeepOldPrimaryAsAlias -and $oldPrimary -and $oldPrimary -ine $NewPrimary) {
        [void]$kept.Add("smtp:$oldPrimary")
    }
    $result = New-Object System.Collections.Generic.List[string]
    [void]$result.Add("SMTP:$NewPrimary")
    foreach ($k in $kept) { [void]$result.Add($k) }
    return $result.ToArray()
}
# Appends $NewSecondary as an smtp: entry. Throws on a duplicate rather than
# silently creating one.
function Add-SecondarySmtpToProxyAddresses {
    param(
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewSecondary
    )
    foreach ($addr in @($CurrentProxyAddresses)) {
        if ("$addr" -match '^smtp:(.+)$' -and $Matches[1] -ieq $NewSecondary) {
            throw "'$NewSecondary' already exists in proxyAddresses (as $addr)."
        }
    }
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($addr in @($CurrentProxyAddresses)) { [void]$result.Add("$addr") }
    [void]$result.Add("smtp:$NewSecondary")
    return $result.ToArray()
}
# Renders a single attribute value for grid display -- handles the
# multi-valued proxyAddresses array so it never throws on an odd property type.
function Format-AttributeValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = New-Object System.Collections.Generic.List[string]
        foreach ($item in $Value) { [void]$items.Add("$item") }
        return ($items -join '; ')
    }
    return "$Value"
}
# Builds the Attribute/Value DataTable, in the fixed order above.
# PrimarySmtpAddress has no real AD property of its own -- derived from
# proxyAddresses here, the single place that derivation happens.
function Format-MailAttribsTable {
    param($ADUserObject)
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Attribute', [string])
    [void]$table.Columns.Add('Value', [string])
    [void]$table.Columns.Add('_Editable', [string])
    [void]$table.Columns.Add('_AttributeKey', [string])
    foreach ($attrName in $script:MailAttribsDisplayOrder) {
        $rawValue = if ($attrName -eq 'PrimarySmtpAddress') {
            Get-PrimarySmtpFromProxyAddresses -ProxyAddresses $ADUserObject.proxyAddresses
        } else {
            $ADUserObject.$attrName
        }
        $label = if ($script:MailAttribsDisplayLabels.ContainsKey($attrName)) { $script:MailAttribsDisplayLabels[$attrName] } else { $attrName }
        $row = $table.NewRow()
        $row['Attribute']     = $label
        $row['Value']         = Format-AttributeValue -Value $rawValue
        $row['_Editable']     = if ($script:MailAttribsEditableAttributes -contains $attrName) { 'Y' } else { 'N' }
        $row['_AttributeKey'] = $attrName
        [void]$table.Rows.Add($row)
    }
    # Comma operator required -- a bare "return $table" would unroll the
    # DataTable into its DataRow objects instead of returning the table itself.
    return ,$table
}
# Writes one attribute change. Branches by mechanism exactly as the reference
# script documents: dedicated Set-ADUser parameters, generic -Replace/-Clear
# for Exchange-schema attributes with no dedicated parameter, and the derived
# PrimarySmtpAddress field (rewrites proxyAddresses per $KeepOldPrimaryAsAlias).
# The 'name' case (Rename-ADObject) is kept for parity with the reference
# script but is unreachable from this UI, same as there -- 'name' is
# deliberately excluded from $script:MailAttribsEditableAttributes.
function Set-MailAttribsAttribute {
    param(
        [Parameter(Mandatory)] [string]$SamAccountName,
        [Parameter(Mandatory)] [string]$DistinguishedName,
        [Parameter(Mandatory)] [string]$AttributeName,
        [string]$NewValue,
        $CurrentProxyAddresses,
        [bool]$KeepOldPrimaryAsAlias = $true,
        [string]$ADServer
    )
    $commonParams = @{ ErrorAction = 'Stop' }
    if ($ADServer) { $commonParams['Server'] = $ADServer }
    $isBlank = ("$NewValue".Trim() -eq '')
    switch ($AttributeName) {
        'DisplayName' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear displayName @commonParams }
            else { Set-ADUser -Identity $SamAccountName -DisplayName $NewValue @commonParams }
        }
        'GivenName' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear givenName @commonParams }
            else { Set-ADUser -Identity $SamAccountName -GivenName $NewValue @commonParams }
        }
        'Initials' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear initials @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Initials $NewValue @commonParams }
        }
        'Surname' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear sn @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Surname $NewValue @commonParams }
        }
        'mail' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear mail @commonParams }
            else { Set-ADUser -Identity $SamAccountName -EmailAddress $NewValue @commonParams }
        }
        'mailNickname' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear mailNickname @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Replace @{ mailNickname = $NewValue } @commonParams }
        }
        'msExchArchiveName' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear msExchArchiveName @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Replace @{ msExchArchiveName = $NewValue } @commonParams }
        }
        'targetAddress' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear targetAddress @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Replace @{ targetAddress = $NewValue } @commonParams }
        }
        'PrimarySmtpAddress' {
            if ($isBlank) { throw "PrimarySmtpAddress cannot be blank." }
            $newProxyAddresses = Set-PrimarySmtpInProxyAddresses -CurrentProxyAddresses $CurrentProxyAddresses -NewPrimary $NewValue.Trim() -KeepOldPrimaryAsAlias $KeepOldPrimaryAsAlias
            Set-ADUser -Identity $SamAccountName -Replace @{ proxyAddresses = $newProxyAddresses } @commonParams
        }
        'name' {
            if ($isBlank) { throw "The object's Name cannot be blank." }
            Rename-ADObject -Identity $DistinguishedName -NewName $NewValue @commonParams
        }
        default { throw "Attribute '$AttributeName' is not handled by Set-MailAttribsAttribute." }
    }
}
# Standalone write for the "Add secondary SMTP address" control -- appends
# one new smtp: entry without touching anything else in proxyAddresses.
function Add-MailAttribsSecondaryProxyAddress {
    param(
        [Parameter(Mandatory)] [string]$SamAccountName,
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewSecondary,
        [string]$ADServer
    )
    $commonParams = @{ ErrorAction = 'Stop' }
    if ($ADServer) { $commonParams['Server'] = $ADServer }
    $newProxyAddresses = Add-SecondarySmtpToProxyAddresses -CurrentProxyAddresses $CurrentProxyAddresses -NewSecondary $NewSecondary
    Set-ADUser -Identity $SamAccountName -Replace @{ proxyAddresses = $newProxyAddresses } @commonParams
}
# Edit dialog for a single attribute -- ported near-verbatim from the
# reference script's Show-AttributeEditDialog, including the extra warning
# notes for 'name' and 'PrimarySmtpAddress' (the 'name' branch is kept for
# parity even though that row is never reachable from this UI, same as there).
function Show-MailAttribsEditDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$AttributeName,
        [string]$CurrentValue
    )
    $warningNote = switch ($AttributeName) {
        'name' { 'This renames the AD OBJECT itself (its Common Name / RDN in the directory) via Rename-ADObject -- a structural directory change, distinct from DisplayName. Cannot be left blank.' }
        'PrimarySmtpAddress' { 'This rewrites proxyAddresses: sets this value as the SMTP: (primary) entry. If a different primary already exists, you will be prompted next to Keep it as a secondary alias or Remove it entirely. Does NOT update the separate "mail" attribute -- edit that too if it needs to match. Cannot be left blank.' }
        default { '' }
    }
    $dlgHeight = if ($warningNote -ne '') { 330 } else { 300 }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit Attribute -- $AttributeName"
    $dlg.Size = New-Object System.Drawing.Size(520, $dlgHeight)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblAttr = New-Object System.Windows.Forms.Label
    $lblAttr.Text = "Attribute: $AttributeName"
    $lblAttr.Location = New-Object System.Drawing.Point(12, 12)
    $lblAttr.AutoSize = $true
    $lblAttr.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblAttr)
    $nextY = 42
    if ($warningNote -ne '') {
        $lblWarnNote = New-Object System.Windows.Forms.Label
        $lblWarnNote.Text = $warningNote
        $lblWarnNote.Location = New-Object System.Drawing.Point(12, $nextY)
        $lblWarnNote.Size = New-Object System.Drawing.Size(480, 40)
        $lblWarnNote.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
        $lblWarnNote.ForeColor = [System.Drawing.Color]::DarkOrange
        $dlg.Controls.Add($lblWarnNote)
        $nextY += 46
    }
    $lblCurrent = New-Object System.Windows.Forms.Label
    $lblCurrent.Text = "Current Value:"
    $lblCurrent.Location = New-Object System.Drawing.Point(12, $nextY)
    $lblCurrent.AutoSize = $true
    $dlg.Controls.Add($lblCurrent)
    $nextY += 18
    $txtCurrent = New-Object System.Windows.Forms.TextBox
    $txtCurrent.Location = New-Object System.Drawing.Point(12, $nextY)
    $txtCurrent.Size = New-Object System.Drawing.Size(480, 40)
    $txtCurrent.Multiline = $true
    $txtCurrent.ReadOnly = $true
    $txtCurrent.ScrollBars = "Vertical"
    $txtCurrent.BackColor = [System.Drawing.Color]::WhiteSmoke
    $txtCurrent.Text = $CurrentValue
    $txtCurrent.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtCurrent)
    $nextY += 48
    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "New Value (blank only allowed if this field is already blank):"
    $lblNew.Location = New-Object System.Drawing.Point(12, $nextY)
    $lblNew.AutoSize = $true
    $dlg.Controls.Add($lblNew)
    $nextY += 18
    $txtNew = New-Object System.Windows.Forms.TextBox
    $txtNew.Location = New-Object System.Drawing.Point(12, $nextY)
    $txtNew.Size = New-Object System.Drawing.Size(480, 60)
    $txtNew.Multiline = $true
    $txtNew.ScrollBars = "Vertical"
    $txtNew.Text = $CurrentValue
    $txtNew.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtNew)
    $nextY += 68
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(320, $nextY)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(408, $nextY)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    $result = $dlg.ShowDialog($Owner)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $txtNew.Text
}
# Shown only when editing PrimarySmtpAddress AND an existing, different
# primary is about to be displaced -- ported verbatim from the reference
# script. Returns 'Keep', 'Remove', or 'Cancel' (Cancel, or closing via the X
# button, aborts the whole attribute edit -- nothing is written to AD).
function Show-KeepOrRemoveOldPrimaryDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$OldPrimary,
        [string]$NewPrimary
    )
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Previous Primary SMTP Address"
    $dlg.Size = New-Object System.Drawing.Size(460, 210)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $dlg.Tag = 'Cancel'
    $lblMsg = New-Object System.Windows.Forms.Label
    $lblMsg.Text = "Current primary SMTP address:`n$OldPrimary`n`nIt is being replaced by:`n$NewPrimary`n`nWhat should I do with the current PrimarySmtpAddress?"
    $lblMsg.Location = New-Object System.Drawing.Point(12, 12)
    $lblMsg.Size = New-Object System.Drawing.Size(430, 100)
    $dlg.Controls.Add($lblMsg)
    $btnKeep = New-Object System.Windows.Forms.Button
    $btnKeep.Text = "Keep as Alias"
    $btnKeep.Location = New-Object System.Drawing.Point(20, 130)
    $btnKeep.Size = New-Object System.Drawing.Size(130, 32)
    $btnKeep.Add_Click({ $dlg.Tag = 'Keep'; $dlg.Close() })
    $dlg.Controls.Add($btnKeep)
    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = "Remove"
    $btnRemove.Location = New-Object System.Drawing.Point(165, 130)
    $btnRemove.Size = New-Object System.Drawing.Size(110, 32)
    $btnRemove.Add_Click({ $dlg.Tag = 'Remove'; $dlg.Close() })
    $dlg.Controls.Add($btnRemove)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(290, 130)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnCancel.Add_Click({ $dlg.Tag = 'Cancel'; $dlg.Close() })
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    [void]$dlg.ShowDialog($Owner)
    return "$($dlg.Tag)"
}
function Show-MailAttribsForm {
    param([string]$Smtp, [string]$ADServer)
    $mailForm = New-Object System.Windows.Forms.Form
    $mailForm.Text = "Mail Attributes: $Smtp"
    $mailForm.Size = New-Object System.Drawing.Size(680, 700)
    $mailForm.StartPosition = "CenterParent"
    $mailForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads this mailbox's on-prem AD account (matched by mail attribute) -- the same fixed attribute set and edit mechanics as the Renaming Script tool. Double-click a row, or select it and click 'Edit Attribute' below, to edit it. 'proxyAddresses' is read-only here -- edit 'Primary SMTP' to change the primary address (you'll be prompted to Keep the old primary as an alias or Remove it), or use 'Add Secondary' below to add another address. 'name', 'UserPrincipalName', and 'Office' are read-only/context-only in this tool."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(640, 58)
    $mailForm.Controls.Add($lblNote)
    $grdMail = New-Object System.Windows.Forms.DataGridView
    $grdMail.Location = New-Object System.Drawing.Point(12, 74)
    $grdMail.Size = New-Object System.Drawing.Size(640, 340)
    $grdMail.ReadOnly = $true
    $grdMail.AllowUserToAddRows = $false
    $grdMail.AllowUserToDeleteRows = $false
    $grdMail.AllowUserToResizeRows = $false
    $grdMail.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdMail.MultiSelect = $false
    $grdMail.AutoGenerateColumns = $false
    $grdMail.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $grdMail.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $grdMail.RowHeadersVisible = $false
    $grdMail.BackgroundColor = [System.Drawing.Color]::White
    $grdMail.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $grdMail.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grdMail.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
    $colMailAttribute = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailAttribute.Name = 'Attribute'
    $colMailAttribute.HeaderText = 'Attribute'
    $colMailAttribute.DataPropertyName = 'Attribute'
    $colMailAttribute.Width = 160
    $colMailAttribute.ReadOnly = $true
    [void]$grdMail.Columns.Add($colMailAttribute)
    $colMailValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailValue.Name = 'Value'
    $colMailValue.HeaderText = 'Value'
    $colMailValue.DataPropertyName = 'Value'
    $colMailValue.ReadOnly = $true
    $colMailValue.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    [void]$grdMail.Columns.Add($colMailValue)
    $colMailEditableFlag = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailEditableFlag.Name = '_Editable'
    $colMailEditableFlag.DataPropertyName = '_Editable'
    $colMailEditableFlag.Visible = $false
    [void]$grdMail.Columns.Add($colMailEditableFlag)
    $colMailAttributeKey = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailAttributeKey.Name = '_AttributeKey'
    $colMailAttributeKey.DataPropertyName = '_AttributeKey'
    $colMailAttributeKey.Visible = $false
    [void]$grdMail.Columns.Add($colMailAttributeKey)
    $mailForm.Controls.Add($grdMail)
    $grdMail.DataSource = New-Object System.Windows.Forms.BindingSource
    $mailBindingSource = $grdMail.DataSource
    $lblAddSmtp = New-Object System.Windows.Forms.Label
    $lblAddSmtp.Text = "Add secondary SMTP address (appended as smtp: -- does not replace the primary or any other address):"
    $lblAddSmtp.Location = New-Object System.Drawing.Point(12, 422)
    $lblAddSmtp.Size = New-Object System.Drawing.Size(640, 16)
    $mailForm.Controls.Add($lblAddSmtp)
    $txtAddSmtp = New-Object System.Windows.Forms.TextBox
    $txtAddSmtp.Location = New-Object System.Drawing.Point(12, 440)
    $txtAddSmtp.Size = New-Object System.Drawing.Size(430, 22)
    $txtAddSmtp.Font = New-Object System.Drawing.Font("Consolas", 9)
    $mailForm.Controls.Add($txtAddSmtp)
    $btnAddSmtp = New-Object System.Windows.Forms.Button
    $btnAddSmtp.Text = "Add Secondary"
    $btnAddSmtp.Location = New-Object System.Drawing.Point(452, 439)
    $btnAddSmtp.Size = New-Object System.Drawing.Size(110, 26)
    $mailForm.Controls.Add($btnAddSmtp)
    $lblMailLog = New-Object System.Windows.Forms.Label
    $lblMailLog.Text = "Log:"
    $lblMailLog.Location = New-Object System.Drawing.Point(12, 474)
    $lblMailLog.AutoSize = $true
    $mailForm.Controls.Add($lblMailLog)
    $txtMailLog = New-Object System.Windows.Forms.TextBox
    $txtMailLog.Multiline = $true
    $txtMailLog.ReadOnly = $true
    $txtMailLog.ScrollBars = "Vertical"
    $txtMailLog.Location = New-Object System.Drawing.Point(12, 494)
    $txtMailLog.Size = New-Object System.Drawing.Size(640, 110)
    $txtMailLog.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtMailLog.BackColor = [System.Drawing.Color]::WhiteSmoke
    $mailForm.Controls.Add($txtMailLog)
    $btnMailRefresh = New-Object System.Windows.Forms.Button
    $btnMailRefresh.Text = "Refresh"
    $btnMailRefresh.Location = New-Object System.Drawing.Point(12, 616)
    $btnMailRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $mailForm.Controls.Add($btnMailRefresh)
    # Added 2026-07-10c alongside double-click -- see the CellDoubleClick /
    # SelectionChanged handlers below. Same position/size as All Attribs'
    # Edit Attribute button for visual consistency between the two dialogs.
    $btnMailEdit = New-Object System.Windows.Forms.Button
    $btnMailEdit.Text = "Edit Attribute"
    $btnMailEdit.Location = New-Object System.Drawing.Point(122, 616)
    $btnMailEdit.Size = New-Object System.Drawing.Size(140, 28)
    $btnMailEdit.Enabled = $false
    $btnMailEdit.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $mailForm.Controls.Add($btnMailEdit)
    $btnMailClose = New-Object System.Windows.Forms.Button
    $btnMailClose.Text = "Close"
    $btnMailClose.Location = New-Object System.Drawing.Point(552, 616)
    $btnMailClose.Size = New-Object System.Drawing.Size(100, 28)
    $mailForm.Controls.Add($btnMailClose)
    # Current AD-user object for this dialog instance. Script-scoped (not a
    # plain closure variable) because it is reassigned from three separate
    # event handlers below (Refresh, grid double-click edit, Add Secondary)
    # -- see the region header comment for why that requires $script: scope
    # in this file's established pattern. Only one Mail Attribs dialog can be
    # open at a time (modal, ShowDialog($form)), so a single script-scope slot
    # is safe, same as $script:CurrentRules / $script:CurrentTraceResults.
    $script:MailAttribsCurrentADUser = $null
    function Update-MailAttribsGrid {
        param($ADUserObject)
        $table = Format-MailAttribsTable -ADUserObject $ADUserObject
        $mailBindingSource.DataSource = $null
        $mailBindingSource.DataSource = $table
        $grdMail.Refresh()
    }
    function Update-MailAttribsFromCurrent {
        $btnMailRefresh.Enabled = $false
        $btnAddSmtp.Enabled = $false
        $txtMailLog.Text = "Loading mail attributes..."
        $mailForm.Refresh()
        try {
            $script:MailAttribsCurrentADUser = Get-MailAttribsADUser -Smtp $Smtp -ADServer $ADServer
            Update-MailAttribsGrid -ADUserObject $script:MailAttribsCurrentADUser
            $txtMailLog.Text = "Loaded mail attributes for SamAccountName '$($script:MailAttribsCurrentADUser.SamAccountName)'."
        } catch {
            $script:MailAttribsCurrentADUser = $null
            $mailBindingSource.DataSource = $null
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
        }
    }
    Update-MailAttribsFromCurrent
    $btnMailRefresh.Add_Click({ Update-MailAttribsFromCurrent })
    # Grid-row edit path -- reachable two ways as of 2026-07-10c: double-click
    # a row (CellDoubleClick below), or select a row and click "Edit
    # Attribute" (btnMailEdit below, enabled only for an editable row via
    # SelectionChanged). Both call this same function against whatever row is
    # currently selected in $grdMail, so the edit logic itself only exists
    # once. Silently no-ops on a read-only row (defensive -- SelectionChanged
    # already keeps btnMailEdit disabled for those, and double-click checks
    # $script:MailAttribsEditableAttributes itself below).
    function Invoke-MailAttribsEditSelectedRow {
        if ($grdMail.SelectedRows.Count -eq 0 -or $null -eq $script:MailAttribsCurrentADUser) { return }
        $row = $grdMail.SelectedRows[0]
        $attrLabel = "$($row.Cells['Attribute'].Value)"
        $attrName  = "$($row.Cells['_AttributeKey'].Value)"
        if ($script:MailAttribsEditableAttributes -notcontains $attrName) { return }
        $btnMailRefresh.Enabled = $false
        $btnMailEdit.Enabled = $false
        $btnAddSmtp.Enabled = $false
        try {
            $oldValue = "$($row.Cells['Value'].Value)"
            $newValue = Show-MailAttribsEditDialog -Owner $mailForm -AttributeName $attrLabel -CurrentValue $oldValue
            if ($null -eq $newValue) {
                $txtMailLog.Text = "Edit cancelled."
                return
            }
            if ($script:MailAttribsNeverBlank -contains $attrName -and $newValue.Trim() -eq '') {
                [void][System.Windows.Forms.MessageBox]::Show("'$attrLabel' cannot be blank.", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $oldIsBlank = ("$oldValue".Trim() -eq '')
            $newIsBlank = $newValue.Trim() -eq ''
            if ($newIsBlank -and -not $oldIsBlank) {
                [void][System.Windows.Forms.MessageBox]::Show("'$attrLabel' currently has a value and cannot be saved blank. Enter a replacement value instead.", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $keepOldPrimaryAsAlias = $true
            if ($attrName -eq 'PrimarySmtpAddress') {
                $oldPrimaryTrim = $oldValue.Trim()
                $newPrimaryTrim = $newValue.Trim()
                if ($oldPrimaryTrim -ne '' -and -not ($oldPrimaryTrim -ieq $newPrimaryTrim)) {
                    $choice = Show-KeepOrRemoveOldPrimaryDialog -Owner $mailForm -OldPrimary $oldPrimaryTrim -NewPrimary $newPrimaryTrim
                    if ($choice -ne 'Keep' -and $choice -ne 'Remove') {
                        $txtMailLog.Text = "Edit cancelled."
                        return
                    }
                    $keepOldPrimaryAsAlias = ($choice -eq 'Keep')
                }
            }
            $displayNew = if ($newValue.Trim() -eq '') { '(cleared)' } else { $newValue.Replace("`r`n", '; ') }
            $confirmExtra = switch ($attrName) {
                'PrimarySmtpAddress' {
                    if ($keepOldPrimaryAsAlias) { "`n`nThis rewrites the proxyAddresses list: sets this as SMTP: (primary) and demotes the previous primary to smtp: (secondary alias)." }
                    else { "`n`nThis rewrites the proxyAddresses list: sets this as SMTP: (primary) and REMOVES the previous primary address from proxyAddresses entirely." }
                }
                default { '' }
            }
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Change '$attrLabel' for $Smtp ($($script:MailAttribsCurrentADUser.SamAccountName)):`n`nFrom: $oldValue`nTo:   $displayNew$confirmExtra",
                'Confirm Attribute Edit',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtMailLog.Text = "Cancelled."
                return
            }
            Set-MailAttribsAttribute -SamAccountName $script:MailAttribsCurrentADUser.SamAccountName -DistinguishedName $script:MailAttribsCurrentADUser.DistinguishedName `
                -AttributeName $attrName -NewValue $newValue -CurrentProxyAddresses $script:MailAttribsCurrentADUser.proxyAddresses `
                -KeepOldPrimaryAsAlias $keepOldPrimaryAsAlias -ADServer $ADServer
            $refreshed = Get-MailAttribsADUser -Smtp $Smtp -ADServer $ADServer
            $script:MailAttribsCurrentADUser = $refreshed
            $afterValue = if ($attrName -eq 'PrimarySmtpAddress') {
                Get-PrimarySmtpFromProxyAddresses -ProxyAddresses $refreshed.proxyAddresses
            } else {
                Format-AttributeValue -Value ($refreshed.$attrName)
            }
            Update-MailAttribsGrid -ADUserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE ---")
            [void]$log.AppendLine("$attrLabel = $oldValue")
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine("$attrLabel = $afterValue")
            $txtMailLog.Text = $log.ToString().TrimEnd()
            $logVerb = if ($attrName -eq 'PrimarySmtpAddress') {
                "Edited Attribute $attrName ($(if($keepOldPrimaryAsAlias){'kept old primary as alias'}else{'removed old primary'}))"
            } else {
                "Edited Attribute $attrName"
            }
            Write-CompletedLog -Task "$logVerb $($script:MailAttribsCurrentADUser.SamAccountName) (Mail Attribs; $Smtp)"
        } catch {
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs edit'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
            $hasRow = ($grdMail.SelectedRows.Count -gt 0 -and $null -ne $script:MailAttribsCurrentADUser)
            $isEditable = $false
            if ($hasRow) { $isEditable = ("$($grdMail.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
            $btnMailEdit.Enabled = ($hasRow -and $isEditable)
        }
    }
    # Double-click: select the row (matching the pattern the reference
    # scripts use elsewhere), then run the shared edit function.
    $grdMail.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $grdMail.ClearSelection()
        $grdMail.Rows[$e.RowIndex].Selected = $true
        Invoke-MailAttribsEditSelectedRow
    })
    # Selection drives Edit Attribute's enabled state, same as All Attribs.
    $grdMail.Add_SelectionChanged({
        $hasRow = ($grdMail.SelectedRows.Count -gt 0 -and $null -ne $script:MailAttribsCurrentADUser)
        $isEditable = $false
        if ($hasRow) { $isEditable = ("$($grdMail.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
        $btnMailEdit.Enabled = ($hasRow -and $isEditable)
    })
    $btnMailEdit.Add_Click({ Invoke-MailAttribsEditSelectedRow })
    $btnAddSmtp.Add_Click({
        if ($null -eq $script:MailAttribsCurrentADUser) {
            [void][System.Windows.Forms.MessageBox]::Show("Load mail attributes first (click Refresh).", 'No User Loaded', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $btnMailRefresh.Enabled = $false
        $btnAddSmtp.Enabled = $false
        try {
            $newAddr = $txtAddSmtp.Text.Trim()
            if ($newAddr -eq '') {
                [void][System.Windows.Forms.MessageBox]::Show("Enter an address to add.", 'Missing Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            if ($newAddr -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                [void][System.Windows.Forms.MessageBox]::Show("'$newAddr' does not look like a valid email address (expected user@domain.tld).", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $beforeValue = Format-AttributeValue -Value $script:MailAttribsCurrentADUser.proxyAddresses
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Add 'smtp:$newAddr' as a secondary proxy address for $Smtp ($($script:MailAttribsCurrentADUser.SamAccountName))?`n`nThis does not change the primary address or remove any existing address.",
                'Confirm Add Secondary SMTP Address',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtMailLog.Text = "Cancelled."
                return
            }
            Add-MailAttribsSecondaryProxyAddress -SamAccountName $script:MailAttribsCurrentADUser.SamAccountName -CurrentProxyAddresses $script:MailAttribsCurrentADUser.proxyAddresses `
                -NewSecondary $newAddr -ADServer $ADServer
            $refreshed = Get-MailAttribsADUser -Smtp $Smtp -ADServer $ADServer
            $script:MailAttribsCurrentADUser = $refreshed
            $afterValue = Format-AttributeValue -Value $refreshed.proxyAddresses
            Update-MailAttribsGrid -ADUserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE (proxyAddresses) ---")
            [void]$log.AppendLine($beforeValue)
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine($afterValue)
            $txtMailLog.Text = $log.ToString().TrimEnd()
            Write-CompletedLog -Task "Added Secondary SMTP $($script:MailAttribsCurrentADUser.SamAccountName) (Mail Attribs; $Smtp; smtp:$newAddr)"
            $txtAddSmtp.Text = ''
        } catch {
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add Secondary SMTP'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
        }
    })
    $btnMailClose.Add_Click({ $mailForm.Close() })
    [void]$mailForm.ShowDialog($form)
}
#endregion
#region Actions - all attributes ----------------------------------------------
# "All Attribs" button (2026-07-10). Ports the operator's separate
# AD-UserAttributeEditor reference script (full Get-ADUser -Properties *
# viewer/editor with a curated editable subset) into this tool, following the
# same adaptation approach used for "Mail Attribs" just above:
#   - Target/AD-server come from this tool's own Mailbox lookup, matched by
#     "mail" (Get-ADUser -Filter "mail -eq '$Smtp'") -- same pattern as
#     Get-MailAttribsADUser / Get-MailboxDLMemberships -- instead of the
#     reference's separate UPN-entry/Parse/Environment UI.
#   - Interaction model initially matched the REFERENCE SCRIPT exactly, per
#     instruction ("same as last -- if they are editable here, they are
#     editable in the final version"): select a row, then click the "Edit
#     Attribute" button, enabled only when the selected row's hidden
#     _Editable flag is 'Y' (grdAll's SelectionChanged handler). 2026-07-10c
#     added double-click as a second, equivalent way to trigger the same
#     edit (Invoke-AllAttribsEditSelectedRow), for parity with Mail Attribs
#     -- both dialogs now support both triggers.
#   - The editable-attribute list is copied verbatim from the reference
#     (Title, Department, Description, Company, Division, Office,
#     OfficePhone, MobilePhone, Fax, HomePhone, StreetAddress, City, State,
#     PostalCode, POBox, EmployeeID, EmployeeNumber, Organization, Initials,
#     HomePage) -- each already documented there as a verified, dedicated
#     Set-ADUser string parameter (Microsoft Learn), never a generic -Replace
#     guess. Everything else Get-ADUser -Properties * returns (SamAccountName,
#     UserPrincipalName, DistinguishedName, ObjectGUID, ObjectSID,
#     UserAccountControl, Manager, GivenName, Surname, DisplayName, mail, the
#     whole proxyAddresses/Exchange-schema set already covered by Mail
#     Attribs, etc.) is view-only here, same as the reference -- not
#     re-derived independently per standards Section 1(a).
#   - Get-ADUser -Properties * returns both ADUser's friendly property names
#     (e.g. "Office") and, when populated, the matching raw LDAP attribute
#     name (e.g. "physicalDeliveryOfficeName") as SEPARATE rows with the same
#     value -- normal Get-ADUser behavior per the reference script's own
#     header, not a bug introduced here. Only the friendly-named row is ever
#     in $script:AllAttribsEditableAttributes, so only it is ever editable;
#     the note label below the grid says so.
#   - Function/variable names are AllAttribs-prefixed to avoid colliding with
#     Mail Attribs' identically-purposed but differently-scoped functions
#     (e.g. Format-AllAttribsValue here additionally handles byte[]/SID/
#     GUID/DateTime -- types the curated mail-attribute list never produces
#     but a full -Properties * dump routinely does -- so it is intentionally
#     NOT the same function as Mail Attribs' simpler Format-AttributeValue).
#   - Same $script:-scope reasoning as Mail Attribs for the current AD-user
#     object ($script:AllAttribsCurrentADUser), since it's read/reassigned
#     from three separate event handlers (Refresh, Edit, and the
#     SelectionChanged enable/disable check).
$script:AllAttribsEditableAttributes = @(
    'Title', 'Department', 'Description', 'Company', 'Division',
    'Office', 'OfficePhone', 'MobilePhone', 'Fax', 'HomePhone',
    'StreetAddress', 'City', 'State', 'PostalCode', 'POBox',
    'EmployeeID', 'EmployeeNumber', 'Organization', 'Initials', 'HomePage'
)
# Attribute-object bookkeeping properties added by the ADUser type itself --
# never real directory attributes, always excluded from the display grid.
$script:AllAttribsExcludedProperties = @(
    'PropertyNames', 'AddedProperties', 'RemovedProperties',
    'ModifiedProperties', 'PropertyCount', 'nTSecurityDescriptor'
)
# Matches the AD account to this mailbox by its "mail" attribute -- same
# pattern as Get-MailAttribsADUser -- but requests every populated attribute
# (-Properties '*') rather than a fixed list.
function Get-AllAttribsADUser {
    param(
        [Parameter(Mandatory)] [string]$Smtp,
        [string]$ADServer
    )
    $safeSmtp = $Smtp.Replace("'", "''")
    $params = @{
        Filter      = "mail -eq '$safeSmtp'"
        Properties  = '*'
        ErrorAction = 'Stop'
    }
    if ($ADServer) { $params['Server'] = $ADServer }
    $results = @(Get-ADUser @params)
    if ($results.Count -eq 0) {
        throw "No on-prem AD user account has a mail attribute matching '$Smtp'. This mailbox may be cloud-only (never synced from on-prem AD), or the Environment/Domain Controller selection at the top of the main window may need to change."
    }
    if ($results.Count -gt 1) {
        throw "$($results.Count) on-prem AD users matched mail '$Smtp'. Expected 0 or 1 -- resolve the duplicate mail attribute before using this feature."
    }
    return $results[0]
}
# Renders a single attribute value for grid display. Handles arrays, binary,
# GUID/SID, and DateTime -- ported verbatim from the reference script, since
# a full -Properties * dump routinely includes all of these (objectGUID,
# objectSid, whenCreated, lastLogon, etc.), unlike Mail Attribs' curated
# always-string attribute list.
function Format-AllAttribsValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [byte[]]) { return '<binary data>' }
    if ($Value -is [System.Security.Principal.SecurityIdentifier]) { return $Value.Value }
    if ($Value -is [guid]) { return $Value.ToString() }
    if ($Value -is [datetime]) { return ('{0:yyyy-MM-dd HH:mm:ss}' -f $Value) }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = New-Object System.Collections.Generic.List[string]
        foreach ($item in $Value) {
            if ($item -is [byte]) { return '<binary data>' }
            [void]$items.Add("$item")
        }
        return ($items -join '; ')
    }
    return "$Value"
}
# Builds the Attribute/Value/_Editable DataTable bound to the grid -- every
# populated property on the ADUser object (minus the bookkeeping properties
# above), sorted by name. Ported from the reference script's
# Format-UserAttributeTable.
function Format-AllAttribsTable {
    param($ADUserObject)
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Attribute', [string])
    [void]$table.Columns.Add('Value', [string])
    [void]$table.Columns.Add('_Editable', [string])
    $props = $ADUserObject.PSObject.Properties |
        Where-Object { $script:AllAttribsExcludedProperties -notcontains $_.Name } |
        Sort-Object Name
    foreach ($p in $props) {
        $row = $table.NewRow()
        $row['Attribute'] = $p.Name
        $row['Value']     = Format-AllAttribsValue -Value $p.Value
        $row['_Editable'] = if ($script:AllAttribsEditableAttributes -contains $p.Name) { 'Y' } else { 'N' }
        [void]$table.Rows.Add($row)
    }
    # Comma operator required -- a bare "return $table" would unroll the
    # DataTable into its DataRow objects instead of returning the table itself.
    return ,$table
}
# Set-ADUser via the dedicated named parameter matching $AttributeName (must
# be a member of $script:AllAttribsEditableAttributes -- enforced by the
# caller). Blank new value clears the attribute instead of writing an empty
# string. Ported from the reference script's Set-OnPremADUserAttribute.
function Set-AllAttribsAttribute {
    param(
        [Parameter(Mandatory)] [string]$SamAccountName,
        [Parameter(Mandatory)] [string]$AttributeName,
        [string]$NewValue,
        [string]$ADServer
    )
    $params = @{
        Identity    = $SamAccountName
        ErrorAction = 'Stop'
    }
    if ($ADServer) { $params['Server'] = $ADServer }
    if ("$NewValue".Trim() -eq '') {
        $params['Clear'] = $AttributeName
    } else {
        $params[$AttributeName] = $NewValue
    }
    Set-ADUser @params
}
# Edit dialog for a single attribute -- ported near-verbatim from the
# reference script's Show-AttributeEditDialog (the simpler version, with no
# extra warning notes, since none of the curated attributes here have the
# special-case semantics PrimarySmtpAddress/name have in Mail Attribs).
function Show-AllAttribsEditDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$AttributeName,
        [string]$CurrentValue
    )
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit Attribute -- $AttributeName"
    $dlg.Size = New-Object System.Drawing.Size(480, 240)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblAttr = New-Object System.Windows.Forms.Label
    $lblAttr.Text = "Attribute: $AttributeName"
    $lblAttr.Location = New-Object System.Drawing.Point(12, 12)
    $lblAttr.AutoSize = $true
    $lblAttr.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblAttr)
    $lblCurrent = New-Object System.Windows.Forms.Label
    $lblCurrent.Text = "Current Value:"
    $lblCurrent.Location = New-Object System.Drawing.Point(12, 42)
    $lblCurrent.AutoSize = $true
    $dlg.Controls.Add($lblCurrent)
    $txtCurrent = New-Object System.Windows.Forms.TextBox
    $txtCurrent.Location = New-Object System.Drawing.Point(12, 60)
    $txtCurrent.Size = New-Object System.Drawing.Size(440, 22)
    $txtCurrent.ReadOnly = $true
    $txtCurrent.BackColor = [System.Drawing.Color]::WhiteSmoke
    $txtCurrent.Text = $CurrentValue
    $txtCurrent.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtCurrent)
    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "New Value (leave blank to clear the attribute):"
    $lblNew.Location = New-Object System.Drawing.Point(12, 92)
    $lblNew.AutoSize = $true
    $dlg.Controls.Add($lblNew)
    $txtNew = New-Object System.Windows.Forms.TextBox
    $txtNew.Location = New-Object System.Drawing.Point(12, 110)
    $txtNew.Size = New-Object System.Drawing.Size(440, 22)
    $txtNew.Text = $CurrentValue
    $txtNew.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtNew)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(280, 150)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(368, 150)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    $result = $dlg.ShowDialog($Owner)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $txtNew.Text
}
function Show-AllAttribsForm {
    param([string]$Smtp, [string]$ADServer)
    $allForm = New-Object System.Windows.Forms.Form
    $allForm.Text = "All Attributes: $Smtp"
    $allForm.Size = New-Object System.Drawing.Size(680, 750)
    $allForm.StartPosition = "CenterParent"
    $allForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads every populated attribute (Get-ADUser -Properties *) on this mailbox's on-prem AD account (matched by mail attribute). Click a column header to sort. Double-click a row, or select it and click 'Edit Attribute' (enabled only for the curated editable set below), to edit it -- everything else is view-only."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(640, 40)
    $allForm.Controls.Add($lblNote)
    $grdAll = New-Object System.Windows.Forms.DataGridView
    $grdAll.Location = New-Object System.Drawing.Point(12, 58)
    $grdAll.Size = New-Object System.Drawing.Size(640, 440)
    $grdAll.ReadOnly = $true
    $grdAll.AllowUserToAddRows = $false
    $grdAll.AllowUserToDeleteRows = $false
    $grdAll.AllowUserToResizeRows = $false
    $grdAll.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdAll.MultiSelect = $false
    $grdAll.AutoGenerateColumns = $false
    $grdAll.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $grdAll.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $grdAll.RowHeadersVisible = $false
    $grdAll.BackgroundColor = [System.Drawing.Color]::White
    $grdAll.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $grdAll.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grdAll.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
    $colAllAttribute = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAllAttribute.Name = 'Attribute'
    $colAllAttribute.HeaderText = 'Attribute'
    $colAllAttribute.DataPropertyName = 'Attribute'
    $colAllAttribute.Width = 260
    $colAllAttribute.ReadOnly = $true
    [void]$grdAll.Columns.Add($colAllAttribute)
    $colAllValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAllValue.Name = 'Value'
    $colAllValue.HeaderText = 'Value'
    $colAllValue.DataPropertyName = 'Value'
    $colAllValue.ReadOnly = $true
    $colAllValue.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    [void]$grdAll.Columns.Add($colAllValue)
    $colAllEditableFlag = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAllEditableFlag.Name = '_Editable'
    $colAllEditableFlag.DataPropertyName = '_Editable'
    $colAllEditableFlag.Visible = $false
    [void]$grdAll.Columns.Add($colAllEditableFlag)
    $allForm.Controls.Add($grdAll)
    # BindingSource is the Microsoft-recommended intermediary for binding a
    # DataGridView to a DataTable -- same as Mail Attribs' grid. Column-header
    # click-to-sort works natively once bound through it (SortMode defaults to
    # Automatic), no custom sort handler needed -- matches the reference script.
    $grdAll.DataSource = New-Object System.Windows.Forms.BindingSource
    $allBindingSource = $grdAll.DataSource
    $lblDupeNote = New-Object System.Windows.Forms.Label
    $lblDupeNote.Text = "Note: some values appear twice -- once under a friendly name (e.g. Office) and once under the raw LDAP name (e.g. physicalDeliveryOfficeName). Only the friendly row is editable."
    $lblDupeNote.Location = New-Object System.Drawing.Point(12, 502)
    $lblDupeNote.Size = New-Object System.Drawing.Size(640, 16)
    $lblDupeNote.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 7.5, [System.Drawing.FontStyle]::Italic)
    $lblDupeNote.ForeColor = [System.Drawing.Color]::DimGray
    $allForm.Controls.Add($lblDupeNote)
    $btnAllRefresh = New-Object System.Windows.Forms.Button
    $btnAllRefresh.Text = "Refresh"
    $btnAllRefresh.Location = New-Object System.Drawing.Point(12, 522)
    $btnAllRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $allForm.Controls.Add($btnAllRefresh)
    $btnAllEdit = New-Object System.Windows.Forms.Button
    $btnAllEdit.Text = "Edit Attribute"
    $btnAllEdit.Location = New-Object System.Drawing.Point(122, 522)
    $btnAllEdit.Size = New-Object System.Drawing.Size(140, 28)
    $btnAllEdit.Enabled = $false
    $btnAllEdit.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $allForm.Controls.Add($btnAllEdit)
    $btnAllClose = New-Object System.Windows.Forms.Button
    $btnAllClose.Text = "Close"
    $btnAllClose.Location = New-Object System.Drawing.Point(552, 522)
    $btnAllClose.Size = New-Object System.Drawing.Size(100, 28)
    $allForm.Controls.Add($btnAllClose)
    $lblAllLog = New-Object System.Windows.Forms.Label
    $lblAllLog.Text = "Log:"
    $lblAllLog.Location = New-Object System.Drawing.Point(12, 558)
    $lblAllLog.AutoSize = $true
    $allForm.Controls.Add($lblAllLog)
    $txtAllLog = New-Object System.Windows.Forms.TextBox
    $txtAllLog.Multiline = $true
    $txtAllLog.ReadOnly = $true
    $txtAllLog.ScrollBars = "Vertical"
    $txtAllLog.Location = New-Object System.Drawing.Point(12, 578)
    $txtAllLog.Size = New-Object System.Drawing.Size(640, 110)
    $txtAllLog.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtAllLog.BackColor = [System.Drawing.Color]::WhiteSmoke
    $allForm.Controls.Add($txtAllLog)
    # Current AD-user object for this dialog instance -- script-scoped for the
    # same cross-handler reason as $script:MailAttribsCurrentADUser (read and
    # reassigned from Refresh, Edit Attribute, and the SelectionChanged
    # enable/disable check). Only one All Attribs dialog can be open at a
    # time (modal, ShowDialog($form)), so a single script-scope slot is safe.
    $script:AllAttribsCurrentADUser = $null
    function Update-AllAttribsGrid {
        param($ADUserObject)
        $table = Format-AllAttribsTable -ADUserObject $ADUserObject
        $allBindingSource.DataSource = $null
        $allBindingSource.DataSource = $table
        $grdAll.Refresh()
    }
    function Update-AllAttribsFromCurrent {
        $btnAllRefresh.Enabled = $false
        $btnAllEdit.Enabled = $false
        $txtAllLog.Text = "Loading all attributes..."
        $allForm.Refresh()
        try {
            $script:AllAttribsCurrentADUser = Get-AllAttribsADUser -Smtp $Smtp -ADServer $ADServer
            Update-AllAttribsGrid -ADUserObject $script:AllAttribsCurrentADUser
            $txtAllLog.Text = "Loaded attributes for SamAccountName '$($script:AllAttribsCurrentADUser.SamAccountName)'."
        } catch {
            $script:AllAttribsCurrentADUser = $null
            $allBindingSource.DataSource = $null
            $txtAllLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'All Attribs'
        } finally {
            $btnAllRefresh.Enabled = $true
        }
    }
    Update-AllAttribsFromCurrent
    $btnAllRefresh.Add_Click({ Update-AllAttribsFromCurrent })
    # Selection drives Edit Attribute's enabled state.
    $grdAll.Add_SelectionChanged({
        $hasRow = ($grdAll.SelectedRows.Count -gt 0 -and $null -ne $script:AllAttribsCurrentADUser)
        $isEditable = $false
        if ($hasRow) { $isEditable = ("$($grdAll.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
        $btnAllEdit.Enabled = ($hasRow -and $isEditable)
    })
    # Grid-row edit path -- reachable two ways as of 2026-07-10c: select a row
    # and click "Edit Attribute" (matches the reference script's own model),
    # or double-click the row directly (added for parity with Mail Attribs).
    # Both call this same function against whatever row is currently selected
    # in $grdAll, so the edit logic itself only exists once.
    function Invoke-AllAttribsEditSelectedRow {
        if ($grdAll.SelectedRows.Count -eq 0 -or $null -eq $script:AllAttribsCurrentADUser) { return }
        $row = $grdAll.SelectedRows[0]
        $attrName = "$($row.Cells['Attribute'].Value)"
        if ($script:AllAttribsEditableAttributes -notcontains $attrName) { return }
        $oldValue = "$($row.Cells['Value'].Value)"
        $btnAllEdit.Enabled = $false
        $btnAllRefresh.Enabled = $false
        try {
            $newValue = Show-AllAttribsEditDialog -Owner $allForm -AttributeName $attrName -CurrentValue $oldValue
            if ($null -eq $newValue) {
                $txtAllLog.Text = "Edit cancelled."
                return
            }
            $displayNew = if ($newValue.Trim() -eq '') { '(cleared)' } else { $newValue }
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Change '$attrName' for $Smtp ($($script:AllAttribsCurrentADUser.SamAccountName)):`n`nFrom: $oldValue`nTo:   $displayNew",
                'Confirm Attribute Edit',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtAllLog.Text = "Cancelled."
                return
            }
            Set-AllAttribsAttribute -SamAccountName $script:AllAttribsCurrentADUser.SamAccountName -AttributeName $attrName -NewValue $newValue -ADServer $ADServer
            $refreshed = Get-AllAttribsADUser -Smtp $Smtp -ADServer $ADServer
            $script:AllAttribsCurrentADUser = $refreshed
            $afterValue = Format-AllAttribsValue -Value ($refreshed.PSObject.Properties[$attrName].Value)
            Update-AllAttribsGrid -ADUserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE ---")
            [void]$log.AppendLine("$attrName = $oldValue")
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine("$attrName = $afterValue")
            $txtAllLog.Text = $log.ToString().TrimEnd()
            Write-CompletedLog -Task "Edited Attribute $attrName $($script:AllAttribsCurrentADUser.SamAccountName) (All Attribs; $Smtp)"
        } catch {
            $txtAllLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'All Attribs edit'
        } finally {
            $btnAllRefresh.Enabled = $true
            $hasRow = ($grdAll.SelectedRows.Count -gt 0 -and $null -ne $script:AllAttribsCurrentADUser)
            $isEditable = $false
            if ($hasRow) { $isEditable = ("$($grdAll.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
            $btnAllEdit.Enabled = ($hasRow -and $isEditable)
        }
    }
    $btnAllEdit.Add_Click({ Invoke-AllAttribsEditSelectedRow })
    # Double-click: select the row, then run the shared edit function. A
    # double-click on a non-editable row is a silent no-op (the function
    # itself checks $script:AllAttribsEditableAttributes).
    $grdAll.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $grdAll.ClearSelection()
        $grdAll.Rows[$e.RowIndex].Selected = $true
        Invoke-AllAttribsEditSelectedRow
    })
    $btnAllClose.Add_Click({ $allForm.Close() })
    [void]$allForm.ShowDialog($form)
}
#endregion
#region Actions - message trace ---------------------------------------------------
# Get-MessageTraceV2 / Get-MessageTraceDetailV2 verified against Microsoft Learn
# (2026-07-08): both are "available only in the cloud-based service" (no on-prem
# equivalent by this name -- Get-MessageTrackingLog is the on-prem cmdlet, a
# completely different name), so there is no collision risk the way Get-Mailbox
# etc. have. SenderAddress and RecipientAddress on Get-MessageTraceV2 are BOTH
# independently optional (Mandatory: False on both), which is what makes the
# Sent/Received one-sided trace below possible with a single cmdlet call -- Sent
# omits RecipientAddress entirely (any recipient), Received omits SenderAddress
# entirely (any sender). Still routed through Invoke-EXOCommand for consistency
# with every other EXO-flavored call in this script, per the standing rule in
# Get-EXOCommand's header comment.
#
# Get-MessageTraceDetailV2 is called exactly the way the operator's own working
# "EXO Message Trace" reference script already calls it (-MessageTraceId,
# -RecipientAddress, -StartDate, -EndDate) -- that usage is treated as verified
# per standards Section 1(a) and is not re-derived here.
function Invoke-MessageTrace {
    # -Status accepts any of the values Get-MessageTraceV2 documents for its
    # own -Status parameter (verified against Microsoft Learn, 2026-07-09):
    # Delivered, Expanded, Failed, FilteredAsSpam, GettingStatus, Pending,
    # Quarantined. Omit it entirely (as the "All Messages" dropdown option
    # does) to get the unfiltered default.
    param(
        [string]$SenderAddress,
        [string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [string]$Subject,
        [string]$Status
    )
    $params = @{
        StartDate   = $StartDate
        EndDate     = $EndDate
        ResultSize  = 5000
        ErrorAction = 'Stop'
    }
    if ($SenderAddress)    { $params.SenderAddress    = $SenderAddress }
    if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }
    if ($Subject) {
        $params.Subject           = $Subject
        $params.SubjectFilterType = 'Contains'
    }
    if ($Status) { $params.Status = $Status }
    Invoke-EXOCommand -Name 'Get-MessageTraceV2' -Parameters $params
}
function Invoke-MessageTraceDetail {
    param(
        [Parameter(Mandatory)][guid]$MessageTraceId,
        [Parameter(Mandatory)][string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )
    Invoke-EXOCommand -Name 'Get-MessageTraceDetailV2' -Parameters @{
        MessageTraceId   = $MessageTraceId
        RecipientAddress = $RecipientAddress
        StartDate        = $StartDate
        EndDate           = $EndDate
        ErrorAction       = 'Stop'
    }
}
# Quarantine release -- verified against Microsoft Learn (2026-07-08):
#  - Get-QuarantineMessage / Release-QuarantineMessage are both "available
#    only in the cloud-based service" (Exchange PowerShell module) -- same
#    module family as Get-MessageTraceV2, routed through Invoke-EXOCommand.
#  - Get-QuarantineMessage -MessageId expects the full RFC822 Message-ID
#    header value, angle brackets included (Microsoft Learn Example 3:
#    Get-QuarantineMessage -MessageID "<...@contoso.com>"). This is a
#    DIFFERENT value than Get-MessageTraceV2's MessageTraceId (a GUID) --
#    the correlation below uses the trace result's MessageId property
#    (the RFC822 header), NOT MessageTraceId.
#  - ASSUMPTION, not verified against a live tenant: Get-MessageTraceV2
#    result objects expose a MessageId property carrying that RFC822
#    value (this has been true of the classic Get-MessageTrace schema and
#    is what admins have always used to cross-reference quarantine, but
#    the operator should confirm on first use -- Find-QuarantineMatch
#    throws a clear, specific error if it's missing rather than silently
#    matching the wrong thing).
#  - Release-QuarantineMessage verified via Microsoft Learn Example 2:
#    Release-QuarantineMessage -Identity <id> -ReleaseToAll releases a
#    single quarantined message (by its own Identity, from
#    Get-QuarantineMessage) to ALL of its original recipients. -Confirm
#    $false is added the same way every other write cmdlet in this script
#    suppresses the built-in prompt, since this tool has its own
#    MessageBox confirmation instead.
function Find-QuarantineMatch {
    param(
        [string]$MessageId,
        [string]$RecipientAddress,
        [datetime]$StartReceivedDate,
        [datetime]$EndReceivedDate
    )
    if (-not $MessageId) {
        throw "This trace result has no Message-ID value to match against quarantine -- cannot look this message up automatically."
    }
    @(Invoke-EXOCommand -Name 'Get-QuarantineMessage' -Parameters @{
        MessageId         = $MessageId
        RecipientAddress  = $RecipientAddress
        StartReceivedDate = $StartReceivedDate
        EndReceivedDate   = $EndReceivedDate
        ErrorAction       = 'Stop'
    })
}
function Invoke-QuarantineRelease {
    param([Parameter(Mandatory)]$QuarantineIdentity)
    Invoke-EXOCommand -Name 'Release-QuarantineMessage' -Parameters @{
        Identity     = $QuarantineIdentity
        ReleaseToAll = $true
        Confirm      = $false
        ErrorAction  = 'Stop'
    } | Out-Null
}
# Time zone display for Message Trace results -- Microsoft Learn confirms
# "The time stamps on the output are in UTC time format" for
# Get-MessageTraceV2 (and by extension Get-MessageTraceDetailV2), so
# everything shown to the operator needs an explicit UTC-to-local
# conversion. Windows time zone IDs (not IANA names) are used with
# [System.TimeZoneInfo], which is correct for the "Standard PowerShell"
# Windows environment this script's header already requires -- these IDs
# carry their DST rules despite the "Standard Time" naming (.NET's own
# documented behavior, not Exchange-specific, so not re-verified against
# Microsoft Learn the way Exchange cmdlets are). This ONLY changes how
# already-fetched results are DISPLAYED (grid, Detail box, CSV export) -- it
# does not change how the Search From/Through date pickers are interpreted;
# those are passed to Get-MessageTraceV2 exactly as before.
$script:TraceTimeZoneIds = [ordered]@{
    'Eastern'  = 'Eastern Standard Time'
    'Central'  = 'Central Standard Time'
    'Mountain' = 'Mountain Standard Time'
    'Pacific'  = 'Pacific Standard Time'
    'UTC'      = 'UTC'
}
function Get-SelectedTraceTimeZone {
    param([string]$Label)
    $id = $script:TraceTimeZoneIds[$Label]
    if (-not $id) { $id = 'UTC' }
    try {
        return [System.TimeZoneInfo]::FindSystemTimeZoneById($id)
    } catch {
        # Falls back to UTC (unconverted) rather than throwing, so a missing
        # time zone on this particular machine degrades gracefully instead
        # of breaking the whole Trace window.
        return [System.TimeZoneInfo]::Utc
    }
}
function ConvertTo-TraceDisplayTime {
    # DateTime values coming back over Exchange Online's remoting connection
    # do not reliably keep DateTimeKind.Utc through deserialization, so this
    # forces Kind=Utc explicitly before converting rather than trusting
    # whatever Kind the object already reports -- converting an
    # already-local value as if it were UTC would silently shift every
    # timestamp by several hours, which is worse than not converting at all.
    param($UtcValue, [System.TimeZoneInfo]$TimeZone)
    if (-not $UtcValue) { return $UtcValue }
    try {
        $forced = [DateTime]::SpecifyKind([DateTime]$UtcValue, [DateTimeKind]::Utc)
        return [System.TimeZoneInfo]::ConvertTimeFromUtc($forced, $TimeZone)
    } catch {
        return $UtcValue
    }
}
function Format-TraceDetail {
    param($Details, [System.TimeZoneInfo]$TimeZone)
    $sb = New-Object System.Text.StringBuilder
    $rows = @($Details)
    if ($rows.Count -eq 0) {
        [void]$sb.AppendLine("(no detail records returned)")
        return $sb.ToString()
    }
    foreach ($d in $rows) {
        $displayDate = if ($TimeZone) { ConvertTo-TraceDisplayTime -UtcValue $d.Date -TimeZone $TimeZone } else { $d.Date }
        [void]$sb.AppendLine("$displayDate  [$($d.Event)]  $($d.Action)")
        if ($d.Detail) { [void]$sb.AppendLine("    $($d.Detail)") }
    }
    return $sb.ToString()
}
function Show-MessageTraceForm {
    # DESIGN NOTE: the operator's pasted reference script binds results to a
    # System.Data.DataTable (auto-generated grid columns, hidden-column pairs
    # for the trace-id/recipient needed by Get Details). This script instead
    # follows the SAME manual-grid convention already used for grdRules and
    # grdCal elsewhere in this file (Columns.Add + Rows.Add, FillWeight
    # sizing). UNLIKE grdRules/grdCal, this grid IS sortable on every column
    # (operator request) -- sorting an unbound DataGridView physically
    # reorders its Rows collection, so row position can no longer be used to
    # find the matching $script:CurrentTraceResults entry after a sort. Fix:
    # a hidden Key column (MessageTraceId|RecipientAddress) on every row,
    # looked up by value via Find-TraceResultByKey -- the same "hidden
    # identity column" approach grdRules already uses for exactly this
    # reason, just adopted here too once sorting made row-index matching
    # unsafe.
    param([string]$Smtp)
    $traceForm = New-Object System.Windows.Forms.Form
    $traceForm.Text = "Message Trace: $Smtp"
    # Width increased 30% (780 -> 1014) on request so the DataGridView's
    # Fill-mode columns get more room to show full sender/recipient/subject
    # text without truncation. Right margin (28px, matching the original
    # 780-12-740 layout) preserved -- content-area controls below widen from
    # 740 to 974 (740 + the 234px gained) to keep that same margin.
    $traceForm.Size = New-Object System.Drawing.Size(1014, 720)
    $traceForm.StartPosition = "CenterParent"
    $traceForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $grpDir = New-Object System.Windows.Forms.GroupBox
    $grpDir.Text = "Direction"
    $grpDir.Location = New-Object System.Drawing.Point(12, 12)
    $grpDir.Size = New-Object System.Drawing.Size(240, 50)
    $traceForm.Controls.Add($grpDir)
    $radSent = New-Object System.Windows.Forms.RadioButton
    $radSent.Text = "Sent (from this mailbox)"
    $radSent.Location = New-Object System.Drawing.Point(10, 16)
    $radSent.AutoSize = $true
    $radSent.Checked = $true
    $grpDir.Controls.Add($radSent)
    $radReceived = New-Object System.Windows.Forms.RadioButton
    $radReceived.Text = "Received (to this mailbox)"
    $radReceived.Location = New-Object System.Drawing.Point(10, 30)
    $radReceived.AutoSize = $true
    $grpDir.Controls.Add($radReceived)
    $lblTraceNote = New-Object System.Windows.Forms.Label
    $lblTraceNote.Text = "Read-only -- no objects are modified. Exchange Online only (Get-MessageTraceV2 / Get-MessageTraceDetailV2), routed through Invoke-EXOCommand."
    $lblTraceNote.Location = New-Object System.Drawing.Point(264, 12)
    $lblTraceNote.Size = New-Object System.Drawing.Size(488, 46)
    $traceForm.Controls.Add($lblTraceNote)
    # Time zone display selector -- placed in the space opened up to the
    # right of lblTraceNote by the 30%-wider form. Converts already-fetched
    # results (grid, Detail box, CSV export) on selection change without
    # re-querying Exchange Online -- see the ConvertTo-TraceDisplayTime /
    # Get-SelectedTraceTimeZone header comment above Format-TraceDetail.
    $lblTZ = New-Object System.Windows.Forms.Label
    $lblTZ.Text = "Show time in:"
    $lblTZ.Location = New-Object System.Drawing.Point(766, 14)
    $lblTZ.AutoSize = $true
    $traceForm.Controls.Add($lblTZ)
    $cboTimeZone = New-Object System.Windows.Forms.ComboBox
    $cboTimeZone.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Eastern', 'Central', 'Mountain', 'Pacific', 'UTC') | ForEach-Object { [void]$cboTimeZone.Items.Add($_) }
    $cboTimeZone.SelectedItem = 'Central'
    $cboTimeZone.Location = New-Object System.Drawing.Point(766, 34)
    $cboTimeZone.Size = New-Object System.Drawing.Size(200, 20)
    $traceForm.Controls.Add($cboTimeZone)
    $lblNarrow = New-Object System.Windows.Forms.Label
    $lblNarrow.Text = "To (optional, narrows recipient):"
    $lblNarrow.Location = New-Object System.Drawing.Point(12, 72)
    $lblNarrow.AutoSize = $true
    $traceForm.Controls.Add($lblNarrow)
    $txtNarrow = New-Object System.Windows.Forms.TextBox
    $txtNarrow.Location = New-Object System.Drawing.Point(270, 69)
    $txtNarrow.Size = New-Object System.Drawing.Size(220, 20)
    $txtNarrow.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtNarrow)
    # Status filter -- placed in the empty space to the right of the "To" row
    # (below lblTraceNote, which only spans y=12-58) so nothing else on the
    # form needs to move. Values verified against Microsoft Learn
    # (Get-MessageTraceV2 -Status): Delivered, Expanded, Failed,
    # FilteredAsSpam, GettingStatus, Pending, Quarantined. "All Messages" is
    # not a real API value -- it just means "don't pass -Status at all",
    # handled in the Trace click handler below.
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Status:"
    $lblStatus.Location = New-Object System.Drawing.Point(510, 72)
    $lblStatus.AutoSize = $true
    $traceForm.Controls.Add($lblStatus)
    $cboStatus = New-Object System.Windows.Forms.ComboBox
    $cboStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('All Messages', 'Delivered', 'Failed', 'Quarantined', 'FilteredAsSpam', 'Pending', 'Expanded', 'GettingStatus') |
        ForEach-Object { [void]$cboStatus.Items.Add($_) }
    $cboStatus.SelectedIndex = 0
    $cboStatus.Location = New-Object System.Drawing.Point(562, 69)
    $cboStatus.Size = New-Object System.Drawing.Size(180, 20)
    $traceForm.Controls.Add($cboStatus)
    $lblSubject = New-Object System.Windows.Forms.Label
    $lblSubject.Text = "Subject contains (optional):"
    $lblSubject.Location = New-Object System.Drawing.Point(12, 100)
    $lblSubject.AutoSize = $true
    $traceForm.Controls.Add($lblSubject)
    $txtSubject = New-Object System.Windows.Forms.TextBox
    $txtSubject.Location = New-Object System.Drawing.Point(270, 97)
    $txtSubject.Size = New-Object System.Drawing.Size(220, 20)
    $txtSubject.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtSubject)
    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = "Search From:"
    $lblStart.Location = New-Object System.Drawing.Point(12, 129)
    $lblStart.AutoSize = $true
    $traceForm.Controls.Add($lblStart)
    $dtpStart = New-Object System.Windows.Forms.DateTimePicker
    $dtpStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtpStart.Location = New-Object System.Drawing.Point(100, 125)
    $dtpStart.Size = New-Object System.Drawing.Size(120, 20)
    $dtpStart.MinDate = (Get-Date).Date.AddDays(-30)
    $dtpStart.MaxDate = (Get-Date).Date
    $dtpStart.Value   = (Get-Date).Date.AddDays(-7)
    $traceForm.Controls.Add($dtpStart)
    $lblEnd = New-Object System.Windows.Forms.Label
    $lblEnd.Text = "Through:"
    $lblEnd.Location = New-Object System.Drawing.Point(240, 129)
    $lblEnd.AutoSize = $true
    $traceForm.Controls.Add($lblEnd)
    $dtpEnd = New-Object System.Windows.Forms.DateTimePicker
    $dtpEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtpEnd.Location = New-Object System.Drawing.Point(300, 125)
    $dtpEnd.Size = New-Object System.Drawing.Size(120, 20)
    $traceForm.Controls.Add($dtpEnd)
    # 10-day-window business rule ported from the reference script: Through
    # can never be more than 9 days after Search From, and never in the future.
    $dtpStart.Add_ValueChanged({
        $maxEnd = $dtpStart.Value.Date.AddDays(9)
        $today  = (Get-Date).Date
        if ($maxEnd -gt $today) { $maxEnd = $today }
        $dtpEnd.MinDate = $dtpStart.Value.Date
        $dtpEnd.MaxDate = $maxEnd
        if ($dtpEnd.Value -lt $dtpStart.Value.Date) { $dtpEnd.Value = $dtpStart.Value.Date }
        if ($dtpEnd.Value -gt $maxEnd) { $dtpEnd.Value = $maxEnd }
    })
    $dtpEnd.MinDate = $dtpStart.Value.Date
    $initialMaxEnd = $dtpStart.Value.Date.AddDays(9)
    if ($initialMaxEnd -gt (Get-Date).Date) { $initialMaxEnd = (Get-Date).Date }
    $dtpEnd.MaxDate = $initialMaxEnd
    $dtpEnd.Value   = $initialMaxEnd
    $btnTraceRun = New-Object System.Windows.Forms.Button
    $btnTraceRun.Text = "Trace"
    $btnTraceRun.Location = New-Object System.Drawing.Point(440, 123)
    $btnTraceRun.Size = New-Object System.Drawing.Size(90, 26)
    $traceForm.Controls.Add($btnTraceRun)
    $lblTraceStatus = New-Object System.Windows.Forms.Label
    $lblTraceStatus.Text = ""
    $lblTraceStatus.Location = New-Object System.Drawing.Point(12, 156)
    $lblTraceStatus.Size = New-Object System.Drawing.Size(974, 40)
    $traceForm.Controls.Add($lblTraceStatus)
    $grdTrace = New-Object System.Windows.Forms.DataGridView
    $grdTrace.Location = New-Object System.Drawing.Point(12, 200)
    $grdTrace.Size = New-Object System.Drawing.Size(974, 260)
    $grdTrace.ReadOnly = $true
    $grdTrace.AllowUserToAddRows = $false
    $grdTrace.AllowUserToDeleteRows = $false
    $grdTrace.AllowUserToResizeRows = $false
    $grdTrace.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdTrace.MultiSelect = $false
    $grdTrace.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdTrace.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdTrace.Columns.Add('Received', 'Received')
    [void]$grdTrace.Columns.Add('Sender', 'Sender Address')
    [void]$grdTrace.Columns.Add('Recipient', 'Recipient Address')
    [void]$grdTrace.Columns.Add('Subject', 'Subject')
    [void]$grdTrace.Columns.Add('Status', 'Status')
    [void]$grdTrace.Columns.Add('Key', 'Key')
    $grdTrace.Columns['Received'].FillWeight  = 15
    $grdTrace.Columns['Sender'].FillWeight    = 20
    $grdTrace.Columns['Recipient'].FillWeight = 20
    $grdTrace.Columns['Subject'].FillWeight   = 30
    $grdTrace.Columns['Status'].FillWeight    = 15
    $grdTrace.Columns['Received'].ValueType   = [DateTime]   # sorts chronologically, not alphabetically
    $grdTrace.Columns['Key'].Visible          = $false        # MessageTraceId|RecipientAddress -- see below
    # Sortable on every visible column (unlike grdRules/grdCal elsewhere in
    # this script, which are intentionally NotSortable) -- sorting reorders
    # the underlying Rows collection in an unbound grid like this one, so
    # row position can no longer be used to find the matching
    # $script:CurrentTraceResults entry. The hidden Key column below is the
    # fix: each row carries a MessageTraceId+RecipientAddress key that
    # double-click and Release-from-Quarantine look up by value instead of
    # by row index, the same "hidden identity column" approach grdRules
    # already uses for exactly this reason.
    foreach ($col in $grdTrace.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
    $traceForm.Controls.Add($grdTrace)
    $lblDetail = New-Object System.Windows.Forms.Label
    $lblDetail.Text = "Detail (double-click a row above to load it here):"
    $lblDetail.Location = New-Object System.Drawing.Point(12, 468)
    $lblDetail.AutoSize = $true
    $traceForm.Controls.Add($lblDetail)
    $txtDetail = New-Object System.Windows.Forms.TextBox
    $txtDetail.Multiline = $true
    $txtDetail.ReadOnly = $true
    $txtDetail.ScrollBars = "Vertical"
    $txtDetail.BackColor = [System.Drawing.Color]::WhiteSmoke
    $txtDetail.Location = New-Object System.Drawing.Point(12, 488)
    $txtDetail.Size = New-Object System.Drawing.Size(974, 130)
    $txtDetail.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtDetail)
    $btnRelease = New-Object System.Windows.Forms.Button
    $btnRelease.Text = "Release from Quarantine"
    $btnRelease.Location = New-Object System.Drawing.Point(12, 628)
    $btnRelease.Size = New-Object System.Drawing.Size(160, 28)
    $btnRelease.Enabled = $false
    $traceForm.Controls.Add($btnRelease)
    $btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Text = "Export to CSV"
    $btnExportCsv.Location = New-Object System.Drawing.Point(180, 628)
    $btnExportCsv.Size = New-Object System.Drawing.Size(110, 28)
    $btnExportCsv.Enabled = $false
    $traceForm.Controls.Add($btnExportCsv)
    $btnTraceClose = New-Object System.Windows.Forms.Button
    $btnTraceClose.Text = "Close"
    $btnTraceClose.Location = New-Object System.Drawing.Point(886, 628)
    $btnTraceClose.Size = New-Object System.Drawing.Size(100, 28)
    $traceForm.Controls.Add($btnTraceClose)
    $script:CurrentTraceResults    = @()
    $script:CurrentTraceDetailRaw  = $null
    $script:TraceSearchStart       = $null
    $script:TraceSearchEnd         = $null
    function Update-TraceGridDisplay {
        # Rebuilds the grid from the already-cached $script:CurrentTraceResults
        # in whatever time zone cboTimeZone currently has selected. Called
        # both right after a Trace run and whenever the time zone dropdown
        # changes -- the latter needs no new Exchange query since the raw
        # (UTC) data is already in hand.
        $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
        $grdTrace.Columns['Received'].HeaderText = "Received ($($cboTimeZone.SelectedItem))"
        $grdTrace.Rows.Clear()
        foreach ($r in $script:CurrentTraceResults) {
            $rowKey = "$($r.MessageTraceId)|$($r.RecipientAddress)"
            $displayReceived = ConvertTo-TraceDisplayTime -UtcValue $r.Received -TimeZone $tz
            [void]$grdTrace.Rows.Add($displayReceived, $r.SenderAddress, $r.RecipientAddress, $r.Subject, $r.Status, $rowKey)
        }
    }
    $cboTimeZone.Add_SelectedIndexChanged({
        Update-TraceGridDisplay
        if ($script:CurrentTraceDetailRaw) {
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $txtDetail.Text = Format-TraceDetail -Details $script:CurrentTraceDetailRaw -TimeZone $tz
        }
    })
    function Update-NarrowLabel {
        if ($radSent.Checked) { $lblNarrow.Text = "To (optional, narrows recipient):" }
        else                  { $lblNarrow.Text = "From (optional, narrows sender):" }
    }
    Update-NarrowLabel
    $radSent.Add_CheckedChanged({ Update-NarrowLabel })
    $radReceived.Add_CheckedChanged({ Update-NarrowLabel })
    $btnTraceRun.Add_Click({
        $btnTraceRun.Enabled = $false
        $btnRelease.Enabled = $false
        $btnExportCsv.Enabled = $false
        $txtDetail.Text = ""
        $grdTrace.Rows.Clear()
        $script:CurrentTraceResults = @()
        $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
        $lblTraceStatus.Text = "Running trace..."
        try {
            $narrow = $txtNarrow.Text.Trim()
            $subject = $txtSubject.Text.Trim()
            $sd = $dtpStart.Value.Date
            $ed = $dtpEnd.Value.Date.AddDays(1).AddSeconds(-1)
            $traceParams = @{ StartDate = $sd; EndDate = $ed }
            if ($radSent.Checked) {
                $traceParams.SenderAddress = $Smtp
                if ($narrow) { $traceParams.RecipientAddress = $narrow }
            } else {
                $traceParams.RecipientAddress = $Smtp
                if ($narrow) { $traceParams.SenderAddress = $narrow }
            }
            if ($subject) { $traceParams.Subject = $subject }
            if ($cboStatus.SelectedItem -and $cboStatus.SelectedItem -ne 'All Messages') {
                $traceParams.Status = $cboStatus.SelectedItem
            }
            $results = @(Invoke-MessageTrace @traceParams)
            $script:CurrentTraceResults = $results
            $script:CurrentTraceDetailRaw = $null
            $script:TraceSearchStart = $sd
            $script:TraceSearchEnd   = $ed
            Update-TraceGridDisplay
            $btnExportCsv.Enabled = ($results.Count -gt 0)
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $statusNote = if ($traceParams.Status) { "  Status: $($traceParams.Status)." } else { "" }
            $lblTraceStatus.Text = "$($results.Count) result(s). $sd through $ed.$statusNote"
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MessageTraceV2'
        } finally {
            $btnTraceRun.Enabled = $true
        }
    })
    function Find-TraceResultByKey {
        # Looks up the raw Get-MessageTraceV2 object behind a grid row by its
        # hidden Key column value (MessageTraceId|RecipientAddress) rather
        # than by row position -- required now that the grid is sortable,
        # since sorting an unbound DataGridView physically reorders Rows.
        param([string]$RowKey)
        $script:CurrentTraceResults | Where-Object { "$($_.MessageTraceId)|$($_.RecipientAddress)" -eq $RowKey } | Select-Object -First 1
    }
    $grdTrace.Add_SelectionChanged({
        $btnRelease.Enabled = ($grdTrace.SelectedRows.Count -gt 0)
    })
    # Double-click a row to load its detail -- same interaction pattern as
    # grdRules elsewhere in this script (double-click for detail, no
    # separate "open" button).
    $grdTrace.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $rowKey = $grdTrace.Rows[$e.RowIndex].Cells['Key'].Value
        $rec = Find-TraceResultByKey -RowKey $rowKey
        if (-not $rec) {
            $txtDetail.Text = "Could not match that row back to a trace result -- run Trace again."
            return
        }
        $txtDetail.Text = "Loading..."
        try {
            $details = Invoke-MessageTraceDetail -MessageTraceId $rec.MessageTraceId -RecipientAddress $rec.RecipientAddress `
                -StartDate $script:TraceSearchStart -EndDate $script:TraceSearchEnd
            $script:CurrentTraceDetailRaw = $details
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $txtDetail.Text = Format-TraceDetail -Details $details -TimeZone $tz
        } catch {
            $script:CurrentTraceDetailRaw = $null
            $txtDetail.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MessageTraceDetailV2'
        }
    })
    $btnRelease.Add_Click({
        if ($grdTrace.SelectedRows.Count -eq 0) { return }
        $rowKey = $grdTrace.SelectedRows[0].Cells['Key'].Value
        $rec = Find-TraceResultByKey -RowKey $rowKey
        if (-not $rec) {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = "Could not match that row back to a trace result -- run Trace again."
            return
        }
        $confirmSearch = [System.Windows.Forms.MessageBox]::Show(
            "Look up quarantine entries for this message?`nSender: $($rec.SenderAddress)`nRecipient: $($rec.RecipientAddress)`nSubject: $($rec.Subject)",
            "Confirm Quarantine Lookup", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirmSearch -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnRelease.Enabled = $false
        try {
            $matches = @(Find-QuarantineMatch -MessageId $rec.MessageId -RecipientAddress $rec.RecipientAddress `
                -StartReceivedDate $script:TraceSearchStart -EndReceivedDate $script:TraceSearchEnd)
            if ($matches.Count -eq 0) {
                $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
                $lblTraceStatus.Text = "No quarantine entry found for this message (it may not have been quarantined, may already be released, or may have expired -- quarantine data is only retained for a limited time)."
                return
            }
            foreach ($qmsg in $matches) {
                $confirmRelease = [System.Windows.Forms.MessageBox]::Show(
                    "Quarantine match found:`nType: $($qmsg.Type)   Policy: $($qmsg.PolicyName)`nReceived: $($qmsg.ReceivedTime)   Release status: $($qmsg.ReleaseStatus)`n`nRelease this message to ALL original recipients?",
                    "Confirm Release", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($confirmRelease -ne [System.Windows.Forms.DialogResult]::Yes) { continue }
                Invoke-QuarantineRelease -QuarantineIdentity $qmsg.Identity
                Write-CompletedLog -Task "Released QuarantineMessage $($rec.RecipientAddress) ($(Get-SelectedEnvLabel); sender: $($rec.SenderAddress); subject: $($rec.Subject))"
            }
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $lblTraceStatus.Text = "$($matches.Count) quarantine match(es) processed."
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Quarantine release'
        } finally {
            $btnRelease.Enabled = $true
        }
    })
    $btnExportCsv.Add_Click({
        if ($script:CurrentTraceResults.Count -eq 0) { $lblTraceStatus.Text = "No results to export."; return }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $sfd.FileName = "MessageTrace_$($Smtp -replace '[^a-zA-Z0-9]','_')_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($traceForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        try {
            # Received is exported in whatever zone the dropdown currently
            # shows (matching the grid on screen), via a calculated property
            # rather than the raw UTC value -- same conversion as the grid.
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $script:CurrentTraceResults |
                Select-Object @{Name = "Received ($($cboTimeZone.SelectedItem))"; Expression = { ConvertTo-TraceDisplayTime -UtcValue $_.Received -TimeZone $tz }},
                              SenderAddress, RecipientAddress, Subject, Status, MessageTraceId |
                Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $lblTraceStatus.Text = "Exported $($script:CurrentTraceResults.Count) row(s) to $($sfd.FileName)."
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = "Export failed: $($_.Exception.Message)"
        }
    })
    $btnTraceClose.Add_Click({ $traceForm.Close() })
    [void]$traceForm.ShowDialog($form)
}
#endregion
#region Logging -----------------------------------------------------------------
function Write-CompletedLog {
    param([string]$Task)
    $folder = "E:\Scripts\M365\Logs\Completed"
    $file   = Join-Path $folder "CompletedTasks.csv"
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $file)) {
        "Task,DateTime" | Out-File -FilePath $file -Encoding utf8
    }
    $dt = Get-Date -Format "yyyy-MM-dd HHmm"
    if ($Task -and $Task.Trim() -ne "") {
        Add-Content -Path $file -Value ('"{0}","{1}"' -f $Task.Trim(), $dt)
    }
}
#endregion
#region WinForms UI ---------------------------------------------------------
$script:FranchiseDCs = @(
    'fran-vm-dc01.franchises.expresspersonnel.com'
    'fran-vm-dc02.franchises.expresspersonnel.com'
    'fran-vm-dc03.franchises.expresspersonnel.com'
    'fran-vm-dc04.franchises.expresspersonnel.com'
    'fran-vm-dc05.franchises.expresspersonnel.com'
    'fran-vm-dc06.franchises.expresspersonnel.com'
    'fran-vm-dc07.franchises.expresspersonnel.com'
    'fran-vm-dc08.franchises.expresspersonnel.com'
)
$script:HQDomain = 'expresspersonnel.com'
$form = New-Object System.Windows.Forms.Form
$form.Text = "Mailbox / Group Inspector (EAC-style Lookup & Edit)"
$form.Size = New-Object System.Drawing.Size(780, 1110)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
# Warning label -- first visible element
$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Connect to Exchange Online before running. On-prem-synced group lookups use the ActiveDirectory (RSAT) module against the selected environment below. Having an on-prem Exchange session ALSO connected in this window is fine -- every EXO call is resolved to the Exchange Online-sourced cmdlet explicitly, regardless of what else is loaded."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(740, 48)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)
# Environment radio group -- Franchise first, HQ second (Section 22)
$grpEnv = New-Object System.Windows.Forms.GroupBox
$grpEnv.Text = "Environment (governs on-prem AD group lookups only -- EXO data is single-tenant)"
$grpEnv.Location = New-Object System.Drawing.Point(12, 62)
$grpEnv.Size = New-Object System.Drawing.Size(740, 50)
$form.Controls.Add($grpEnv)
$radFranchise = New-Object System.Windows.Forms.RadioButton
$radFranchise.Text = "Franchise"
$radFranchise.Location = New-Object System.Drawing.Point(10, 20)
$radFranchise.AutoSize = $true
$radFranchise.Checked = $true
$grpEnv.Controls.Add($radFranchise)
$radHQ = New-Object System.Windows.Forms.RadioButton
$radHQ.Text = "HQ"
$radHQ.Location = New-Object System.Drawing.Point(120, 20)
$radHQ.AutoSize = $true
$grpEnv.Controls.Add($radHQ)
# DC dropdown -- enabled for Franchise, disabled/greyed for HQ
$lblDC = New-Object System.Windows.Forms.Label
$lblDC.Text = "Domain Controller (Franchise only; HQ is forest-wide):"
$lblDC.Location = New-Object System.Drawing.Point(12, 116)
$lblDC.AutoSize = $true
$form.Controls.Add($lblDC)
$cboDC = New-Object System.Windows.Forms.ComboBox
$cboDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:FranchiseDCs | ForEach-Object { [void]$cboDC.Items.Add($_) }
$cboDC.SelectedIndex = 0
$cboDC.Location = New-Object System.Drawing.Point(320, 113)
$cboDC.Size = New-Object System.Drawing.Size(400, 20)
$cboDC.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($cboDC)
function Update-EnvironmentUI {
    if ($radHQ.Checked) {
        $cboDC.Enabled = $false
        $lblWarn.Text = "Connect to Exchange Online before running. HQ AD lookups use the ActiveDirectory (RSAT) module, forest-wide (no DC pinned) -- allow for replication before re-running Lookup. An on-prem Exchange session may also be connected; EXO calls are resolved explicitly and won't be affected."
    } else {
        $cboDC.Enabled = $true
        $lblWarn.Text = "Connect to Exchange Online before running. Franchise AD lookups use the ActiveDirectory (RSAT) module, pinned to the selected DC below. An on-prem Exchange session may also be connected; EXO calls are resolved explicitly and won't be affected."
    }
}
$radFranchise.Add_CheckedChanged({ Update-EnvironmentUI })
$radHQ.Add_CheckedChanged({ Update-EnvironmentUI })
# Input area
$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Enter email address (@expresspros.com -- optional):"
$lblInput.Location = New-Object System.Drawing.Point(12, 146)
$lblInput.AutoSize = $true
$form.Controls.Add($lblInput)
$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(12, 166)
$txtInput.Size = New-Object System.Drawing.Size(740, 20)
$txtInput.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtInput)
# Parsed preview
$lblParsed = New-Object System.Windows.Forms.Label
$lblParsed.Text = "Parsed Preview:"
$lblParsed.Location = New-Object System.Drawing.Point(12, 194)
$lblParsed.AutoSize = $true
$form.Controls.Add($lblParsed)
$txtParsed = New-Object System.Windows.Forms.TextBox
$txtParsed.Multiline = $true
$txtParsed.ReadOnly = $true
$txtParsed.ScrollBars = "Vertical"
$txtParsed.BackColor = [System.Drawing.Color]::WhiteSmoke
$txtParsed.Location = New-Object System.Drawing.Point(12, 214)
$txtParsed.Size = New-Object System.Drawing.Size(740, 40)
$txtParsed.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtParsed)
# Parse / Run / Close row (Section 13). No Mode group -- see header deviation note.
$btnParse = New-Object System.Windows.Forms.Button
$btnParse.Text = "Parse"
$btnParse.Location = New-Object System.Drawing.Point(12, 262)
$btnParse.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnParse)
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run (Lookup)"
$btnRun.Location = New-Object System.Drawing.Point(112, 262)
$btnRun.Size = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnRun)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(232, 262)
$btnClose.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)
# Mailbox actions row -- enabled only after a successful Mailbox lookup.
# This is deliberately a separate row from Parse/Run/Close (Section 13) so
# more feature buttons can be added here later without re-flowing the core
# lookup controls above.
$lblMbxActions = New-Object System.Windows.Forms.Label
$lblMbxActions.Text = "Mailbox actions (enabled after a Mailbox lookup):"
$lblMbxActions.Location = New-Object System.Drawing.Point(12, 296)
$lblMbxActions.AutoSize = $true
$form.Controls.Add($lblMbxActions)
$btnRules = New-Object System.Windows.Forms.Button
$btnRules.Text = "Inbox Rules"
$btnRules.Location = New-Object System.Drawing.Point(12, 316)
$btnRules.Size = New-Object System.Drawing.Size(120, 28)
$btnRules.Enabled = $false
$form.Controls.Add($btnRules)
$btnOOF = New-Object System.Windows.Forms.Button
$btnOOF.Text = "Out of Office"
$btnOOF.Location = New-Object System.Drawing.Point(140, 316)
$btnOOF.Size = New-Object System.Drawing.Size(130, 28)
$btnOOF.Enabled = $false
$form.Controls.Add($btnOOF)
$btnCalPerm = New-Object System.Windows.Forms.Button
$btnCalPerm.Text = "Calendar Perm"
$btnCalPerm.Location = New-Object System.Drawing.Point(278, 316)
$btnCalPerm.Size = New-Object System.Drawing.Size(130, 28)
$btnCalPerm.Enabled = $false
$form.Controls.Add($btnCalPerm)
$btnHealth = New-Object System.Windows.Forms.Button
$btnHealth.Text = "Health"
$btnHealth.Location = New-Object System.Drawing.Point(416, 316)
$btnHealth.Size = New-Object System.Drawing.Size(90, 28)
$btnHealth.Enabled = $false
$form.Controls.Add($btnHealth)
$btnTrace = New-Object System.Windows.Forms.Button
$btnTrace.Text = "Message Trace"
$btnTrace.Location = New-Object System.Drawing.Point(516, 316)
$btnTrace.Size = New-Object System.Drawing.Size(130, 28)
$btnTrace.Enabled = $false
$form.Controls.Add($btnTrace)
# Split out of the Health panel on 2026-07-09 -- see Get-MailboxDLMemberships /
# Show-DLMembershipsForm. Placed in the remaining space on the same row as the
# other mailbox-action buttons (516+130=646 through the row's right edge at
# 752, matching txtInput/txtOut's right edge) rather than a new row, since a
# new row would collide with lblOut (y=352) / txtOut (y=372) immediately below
# and require re-flowing every control after it.
$btnDLMembers = New-Object System.Windows.Forms.Button
$btnDLMembers.Text = "DL Members"
$btnDLMembers.Location = New-Object System.Drawing.Point(650, 316)
$btnDLMembers.Size = New-Object System.Drawing.Size(102, 28)
$btnDLMembers.Enabled = $false
$form.Controls.Add($btnDLMembers)
# Second mailbox-actions row -- added 2026-07-10 for the new attribute-lookup
# buttons ("Mail Attribs" now; "All Attribs" is a separate follow-up request
# and is not added yet). The first action row (Inbox Rules..DL Members) is
# already flush with the form's right margin (650+102=752, matching
# txtInput's right edge at 12+740=752), so this is a new row rather than a
# squeezed-in extra button. Everything from lblOut downward is shifted +30px
# (and the form height +30px) to make room -- see those controls below.
$btnMailAttribs = New-Object System.Windows.Forms.Button
$btnMailAttribs.Text = "Mail Attribs"
$btnMailAttribs.Location = New-Object System.Drawing.Point(12, 344)
$btnMailAttribs.Size = New-Object System.Drawing.Size(110, 28)
$btnMailAttribs.Enabled = $false
$form.Controls.Add($btnMailAttribs)
# Second requested attribute-lookup button (2026-07-10, follow-up to Mail
# Attribs above) -- full Get-ADUser -Properties * viewer/editor, ported from
# the operator's separate AD-UserAttributeEditor reference script. Placed
# right after Mail Attribs in the same row (12+110=122, +10px gap=132).
$btnAllAttribs = New-Object System.Windows.Forms.Button
$btnAllAttribs.Text = "All Attribs"
$btnAllAttribs.Location = New-Object System.Drawing.Point(132, 344)
$btnAllAttribs.Size = New-Object System.Drawing.Size(110, 28)
$btnAllAttribs.Enabled = $false
$form.Controls.Add($btnAllAttribs)
# Result box
$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = "Result:"
$lblOut.Location = New-Object System.Drawing.Point(12, 382)
$lblOut.AutoSize = $true
$form.Controls.Add($lblOut)
$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.Location = New-Object System.Drawing.Point(12, 402)
$txtOut.Size = New-Object System.Drawing.Size(740, 180)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOut)
# Inbox Rules grid -- populated by the "Inbox Rules" button, only after a
# successful Mailbox lookup. A plain TextBox can't have clickable rows, so
# this is a separate DataGridView rather than reusing the Result box above.
$lblRules = New-Object System.Windows.Forms.Label
$lblRules.Text = "Inbox Rules (double-click a row to open):"
$lblRules.Location = New-Object System.Drawing.Point(12, 588)
$lblRules.AutoSize = $true
$form.Controls.Add($lblRules)
$grdRules = New-Object System.Windows.Forms.DataGridView
$grdRules.Location = New-Object System.Drawing.Point(12, 608)
$grdRules.Size = New-Object System.Drawing.Size(740, 150)
$grdRules.ReadOnly = $true
$grdRules.AllowUserToAddRows = $false
$grdRules.AllowUserToDeleteRows = $false
$grdRules.AllowUserToResizeRows = $false
$grdRules.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grdRules.MultiSelect = $false
$grdRules.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grdRules.Font = New-Object System.Drawing.Font("Consolas", 9)
[void]$grdRules.Columns.Add('Name', 'Rule Name')
[void]$grdRules.Columns.Add('Description', 'Description')
[void]$grdRules.Columns.Add('Enabled', 'Enabled')
[void]$grdRules.Columns.Add('Identity', 'Identity')
$grdRules.Columns['Name'].FillWeight        = 25
$grdRules.Columns['Description'].FillWeight = 55
$grdRules.Columns['Enabled'].FillWeight     = 12
$grdRules.Columns['Identity'].FillWeight    = 8
$grdRules.Columns['Identity'].Visible = $false
foreach ($col in $grdRules.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }
$form.Controls.Add($grdRules)
# Mailbox Permissions edit group
$grpMbx = New-Object System.Windows.Forms.GroupBox
$grpMbx.Text = "Mailbox Permissions (enabled after a successful Mailbox lookup)"
$grpMbx.Location = New-Object System.Drawing.Point(12, 766)
$grpMbx.Size = New-Object System.Drawing.Size(740, 90)
$grpMbx.Enabled = $false
$form.Controls.Add($grpMbx)
$lblMbxGrantee = New-Object System.Windows.Forms.Label
$lblMbxGrantee.Text = "Grantee (email or logon):"
$lblMbxGrantee.Location = New-Object System.Drawing.Point(10, 25)
$lblMbxGrantee.AutoSize = $true
$grpMbx.Controls.Add($lblMbxGrantee)
$txtMbxGrantee = New-Object System.Windows.Forms.TextBox
$txtMbxGrantee.Location = New-Object System.Drawing.Point(170, 22)
$txtMbxGrantee.Size = New-Object System.Drawing.Size(280, 20)
$txtMbxGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
$grpMbx.Controls.Add($txtMbxGrantee)
$btnFAGrant = New-Object System.Windows.Forms.Button
$btnFAGrant.Text = "Grant Full Access"
$btnFAGrant.Location = New-Object System.Drawing.Point(10, 55)
$btnFAGrant.Size = New-Object System.Drawing.Size(140, 26)
$grpMbx.Controls.Add($btnFAGrant)
$btnFARevoke = New-Object System.Windows.Forms.Button
$btnFARevoke.Text = "Revoke Full Access"
$btnFARevoke.Location = New-Object System.Drawing.Point(155, 55)
$btnFARevoke.Size = New-Object System.Drawing.Size(140, 26)
$grpMbx.Controls.Add($btnFARevoke)
$btnSAGrant = New-Object System.Windows.Forms.Button
$btnSAGrant.Text = "Grant Send As"
$btnSAGrant.Location = New-Object System.Drawing.Point(305, 55)
$btnSAGrant.Size = New-Object System.Drawing.Size(140, 26)
$grpMbx.Controls.Add($btnSAGrant)
$btnSARevoke = New-Object System.Windows.Forms.Button
$btnSARevoke.Text = "Revoke Send As"
$btnSARevoke.Location = New-Object System.Drawing.Point(450, 55)
$btnSARevoke.Size = New-Object System.Drawing.Size(140, 26)
$grpMbx.Controls.Add($btnSARevoke)
# Group Delegation edit group
$grpGrp = New-Object System.Windows.Forms.GroupBox
$grpGrp.Text = "Group Delegation (enabled after a successful Group lookup)"
$grpGrp.Location = New-Object System.Drawing.Point(12, 864)
$grpGrp.Size = New-Object System.Drawing.Size(740, 120)
$grpGrp.Enabled = $false
$form.Controls.Add($grpGrp)
$lblGrpGrantee = New-Object System.Windows.Forms.Label
$lblGrpGrantee.Text = "Grantee (email or logon):"
$lblGrpGrantee.Location = New-Object System.Drawing.Point(10, 25)
$lblGrpGrantee.AutoSize = $true
$grpGrp.Controls.Add($lblGrpGrantee)
$txtGrpGrantee = New-Object System.Windows.Forms.TextBox
$txtGrpGrantee.Location = New-Object System.Drawing.Point(170, 22)
$txtGrpGrantee.Size = New-Object System.Drawing.Size(280, 20)
$txtGrpGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
$grpGrp.Controls.Add($txtGrpGrantee)
$lblOwnerNote = New-Object System.Windows.Forms.Label
$lblOwnerNote.Text = "On-prem groups: Set Owner REPLACES the single owner (see script header). M365 Groups: Add/Remove Owner requires the grantee to already be a Member."
$lblOwnerNote.Location = New-Object System.Drawing.Point(10, 48)
$lblOwnerNote.Size = New-Object System.Drawing.Size(720, 30)
$grpGrp.Controls.Add($lblOwnerNote)
$btnOwnerSet = New-Object System.Windows.Forms.Button
$btnOwnerSet.Text = "Set/Add Owner"
$btnOwnerSet.Location = New-Object System.Drawing.Point(10, 82)
$btnOwnerSet.Size = New-Object System.Drawing.Size(140, 26)
$grpGrp.Controls.Add($btnOwnerSet)
$btnOwnerRemove = New-Object System.Windows.Forms.Button
$btnOwnerRemove.Text = "Remove Owner"
$btnOwnerRemove.Location = New-Object System.Drawing.Point(155, 82)
$btnOwnerRemove.Size = New-Object System.Drawing.Size(140, 26)
$grpGrp.Controls.Add($btnOwnerRemove)
$btnMemberAdd = New-Object System.Windows.Forms.Button
$btnMemberAdd.Text = "Add Member"
$btnMemberAdd.Location = New-Object System.Drawing.Point(305, 82)
$btnMemberAdd.Size = New-Object System.Drawing.Size(140, 26)
$grpGrp.Controls.Add($btnMemberAdd)
$btnMemberRemove = New-Object System.Windows.Forms.Button
$btnMemberRemove.Text = "Remove Member"
$btnMemberRemove.Location = New-Object System.Drawing.Point(450, 82)
$btnMemberRemove.Size = New-Object System.Drawing.Size(140, 26)
$grpGrp.Controls.Add($btnMemberRemove)
# ---- shared script state ----
$script:Parsed      = $null
$script:ParsedFrom  = $null
$script:Target      = $null   # result of Resolve-Target
$script:LastSmtp    = $null
$script:CurrentRules = @()    # raw Get-InboxRule objects backing the grdRules grid
$script:CalendarIdentityForForm = $null   # set by Show-CalendarPermissionsForm while it's open
function Get-SelectedADServer {
    if ($radHQ.Checked) { return $script:HQDomain }
    return $cboDC.SelectedItem
}
function Get-SelectedEnvLabel {
    if ($radHQ.Checked) { return 'HQ' } else { return 'Franchise' }
}
function Update-ParsedPreview {
    $smtp = Get-TargetSmtp -Text $txtInput.Text
    $script:Parsed     = $smtp
    $script:ParsedFrom = $txtInput.Text
    if ($smtp) {
        $txtParsed.Lines = @("Target SMTP : $smtp", "Environment : $(Get-SelectedEnvLabel)")
    } else {
        $txtParsed.Lines = @("No email address entered.")
    }
}
$btnParse.Add_Click({
    $btnParse.Enabled = $false
    $btnRun.Enabled   = $false
    try { Update-ParsedPreview }
    catch { $txtOut.Text = "ERROR: $($_.Exception.Message)" }
    finally { $btnParse.Enabled = $true; $btnRun.Enabled = $true }
})
$btnRun.Add_Click({
    $btnParse.Enabled = $false
    $btnRun.Enabled   = $false
    $grpMbx.Enabled = $false
    $grpGrp.Enabled = $false
    $btnRules.Enabled = $false
    $btnOOF.Enabled = $false
    $btnCalPerm.Enabled = $false
    $btnHealth.Enabled = $false
    $btnTrace.Enabled = $false
    $btnDLMembers.Enabled = $false
    $btnMailAttribs.Enabled = $false
    $btnAllAttribs.Enabled = $false
    $grdRules.Rows.Clear()
    $script:CurrentRules = @()
    try {
        if (-not $script:Parsed -or $txtInput.Text -ne $script:ParsedFrom) { Update-ParsedPreview }
        if (-not $script:Parsed) {
            $txtOut.Text = "Cannot run: enter an email address first."
            return
        }
        $smtp = $script:Parsed
        $adServer = Get-SelectedADServer
        $script:Target   = Resolve-Target -Smtp $smtp -ADServer $adServer
        $script:LastSmtp = $smtp
        $out = New-Object System.Text.StringBuilder
        [void]$out.AppendLine("Environment queried : $(Get-SelectedEnvLabel)  (AD server/domain: $adServer)")
        [void]$out.AppendLine("Target SMTP         : $smtp")
        [void]$out.AppendLine("Detected type        : $($script:Target.Type)")
        [void]$out.AppendLine("")
        switch ($script:Target.Type) {
            'Mailbox' {
                $state = Get-MailboxState -Smtp $smtp
                [void]$out.Append((Format-MailboxState -Heading "CURRENT STATE" -State $state))
                $grpMbx.Enabled = $true
                $btnRules.Enabled = $true
                $btnOOF.Enabled = $true
                $btnCalPerm.Enabled = $true
                $btnHealth.Enabled = $true
                $btnTrace.Enabled = $true
                $btnDLMembers.Enabled = $true
                $btnMailAttribs.Enabled = $true
                $btnAllAttribs.Enabled = $true
            }
            'UnifiedGroup' {
                $state = Get-GroupState -Type 'UnifiedGroup' -Smtp $smtp
                [void]$out.Append((Format-GroupState -Heading "CURRENT STATE" -State $state))
                $grpGrp.Enabled = $true
            }
            'OnpremGroup' {
                $state = Get-GroupState -Type 'OnpremGroup' -Smtp $smtp -ADGroupDN $script:Target.ADGroupDN -ADServer $adServer
                [void]$out.Append((Format-GroupState -Heading "CURRENT STATE" -State $state))
                $grpGrp.Enabled = $true
            }
            default {
                [void]$out.AppendLine("Not found as a Mailbox, Microsoft 365 Group, or on-prem AD group.")
                if ($script:Target.ResolutionErrors.Count -gt 0) {
                    [void]$out.AppendLine("")
                    [void]$out.AppendLine("--- Resolution attempt errors (for troubleshooting) ---")
                    $script:Target.ResolutionErrors | ForEach-Object { [void]$out.AppendLine($_) }
                }
            }
        }
        $txtOut.Text = $out.ToString()
    }
    catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Run (Lookup)'
    }
    finally {
        $btnParse.Enabled = $true
        $btnRun.Enabled   = $true
    }
})
$btnClose.Add_Click({ $form.Close() })
# ---- Mailbox permission edit handlers ----
function Invoke-MailboxEdit {
    param([string]$ActionLabel, [string]$Verb, [scriptblock]$Action)
    $grantee = $txtMbxGrantee.Text.Trim()
    if (-not $script:LastSmtp) { $txtOut.Text = "Run a Lookup first."; return }
    if (-not $grantee) { $txtOut.Text = "Enter a grantee (email or logon) first."; return }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "$ActionLabel for '$grantee' on mailbox '$($script:LastSmtp)'?",
        "Confirm $ActionLabel",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
    $grpMbx.Enabled = $false
    $btnRun.Enabled = $false
    try {
        $before = Get-MailboxState -Smtp $script:LastSmtp
        $out = New-Object System.Text.StringBuilder
        [void]$out.Append((Format-MailboxState -Heading "BEFORE" -State $before))
        & $Action
        [void]$out.AppendLine("$ActionLabel : SUCCESS ($grantee)")
        $after = Get-MailboxState -Smtp $script:LastSmtp
        [void]$out.Append((Format-MailboxState -Heading "AFTER" -State $after))
        $txtOut.Text = $out.ToString()
        Write-CompletedLog -Task "$Verb $($script:LastSmtp) ($(Get-SelectedEnvLabel); grantee: $grantee)"
    }
    catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context $ActionLabel
    }
    finally {
        $grpMbx.Enabled = $true
        $btnRun.Enabled = $true
    }
}
$btnFAGrant.Add_Click({
    Invoke-MailboxEdit -ActionLabel "Grant Full Access" -Verb "Granted FullAccess" -Action {
        Grant-FullAccessPermission -Smtp $script:LastSmtp -Grantee $txtMbxGrantee.Text.Trim()
    }
})
$btnFARevoke.Add_Click({
    Invoke-MailboxEdit -ActionLabel "Revoke Full Access" -Verb "Revoked FullAccess" -Action {
        Revoke-FullAccessPermission -Smtp $script:LastSmtp -Grantee $txtMbxGrantee.Text.Trim()
    }
})
$btnSAGrant.Add_Click({
    Invoke-MailboxEdit -ActionLabel "Grant Send As" -Verb "Granted SendAs" -Action {
        Grant-SendAsPermission -Smtp $script:LastSmtp -Grantee $txtMbxGrantee.Text.Trim()
    }
})
$btnSARevoke.Add_Click({
    Invoke-MailboxEdit -ActionLabel "Revoke Send As" -Verb "Revoked SendAs" -Action {
        Revoke-SendAsPermission -Smtp $script:LastSmtp -Grantee $txtMbxGrantee.Text.Trim()
    }
})
# ---- Group delegation edit handlers ----
function Invoke-GroupEdit {
    param([string]$ActionLabel, [string]$Verb, [scriptblock]$Action)
    $grantee = $txtGrpGrantee.Text.Trim()
    if (-not $script:LastSmtp -or -not $script:Target) { $txtOut.Text = "Run a Lookup first."; return }
    if (-not $grantee) { $txtOut.Text = "Enter a grantee (email or logon) first."; return }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "$ActionLabel for '$grantee' on group '$($script:LastSmtp)'?",
        "Confirm $ActionLabel",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
    $grpGrp.Enabled = $false
    $btnRun.Enabled = $false
    $adServer = Get-SelectedADServer
    try {
        $before = Get-GroupState -Type $script:Target.Type -Smtp $script:LastSmtp -ADGroupDN $script:Target.ADGroupDN -ADServer $adServer
        $out = New-Object System.Text.StringBuilder
        [void]$out.Append((Format-GroupState -Heading "BEFORE" -State $before))
        & $Action
        [void]$out.AppendLine("$ActionLabel : SUCCESS ($grantee)")
        $after = Get-GroupState -Type $script:Target.Type -Smtp $script:LastSmtp -ADGroupDN $script:Target.ADGroupDN -ADServer $adServer
        [void]$out.Append((Format-GroupState -Heading "AFTER" -State $after))
        $txtOut.Text = $out.ToString()
        Write-CompletedLog -Task "$Verb $($script:LastSmtp) ($(Get-SelectedEnvLabel); grantee: $grantee)"
    }
    catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context $ActionLabel
    }
    finally {
        $grpGrp.Enabled = $true
        $btnRun.Enabled = $true
    }
}
$btnOwnerSet.Add_Click({
    Invoke-GroupEdit -ActionLabel "Set/Add Owner" -Verb "Set Owner" -Action {
        $grantee = $txtGrpGrantee.Text.Trim()
        if ($script:Target.Type -eq 'UnifiedGroup') { Add-UnifiedOwner -Smtp $script:LastSmtp -Grantee $grantee }
        else { Set-OnpremOwner -ADGroupDN $script:Target.ADGroupDN -Grantee $grantee -ADServer (Get-SelectedADServer) }
    }
})
$btnOwnerRemove.Add_Click({
    Invoke-GroupEdit -ActionLabel "Remove Owner" -Verb "Removed Owner" -Action {
        $grantee = $txtGrpGrantee.Text.Trim()
        if ($script:Target.Type -eq 'UnifiedGroup') { Remove-UnifiedOwner -Smtp $script:LastSmtp -Grantee $grantee }
        else { throw "On-prem groups have a single ManagedBy owner. Use 'Set/Add Owner' with a new value to replace it -- there is no separate remove action for the AD ManagedBy attribute." }
    }
})
$btnMemberAdd.Add_Click({
    Invoke-GroupEdit -ActionLabel "Add Member" -Verb "Added Member" -Action {
        $grantee = $txtGrpGrantee.Text.Trim()
        if ($script:Target.Type -eq 'UnifiedGroup') { Add-UnifiedMember -Smtp $script:LastSmtp -Grantee $grantee }
        else { Add-OnpremMember -ADGroupDN $script:Target.ADGroupDN -Grantee $grantee -ADServer (Get-SelectedADServer) }
    }
})
$btnMemberRemove.Add_Click({
    Invoke-GroupEdit -ActionLabel "Remove Member" -Verb "Removed Member" -Action {
        $grantee = $txtGrpGrantee.Text.Trim()
        if ($script:Target.Type -eq 'UnifiedGroup') { Remove-UnifiedMember -Smtp $script:LastSmtp -Grantee $grantee }
        else { Remove-OnpremMember -ADGroupDN $script:Target.ADGroupDN -Grantee $grantee -ADServer (Get-SelectedADServer) }
    }
})
# ---- Inbox Rules handlers ----
$btnRules.Add_Click({
    if (-not $script:LastSmtp -or -not $script:Target -or $script:Target.Type -ne 'Mailbox') {
        $txtOut.Text = "Run a Mailbox lookup first."
        return
    }
    # (Test-MailboxLookupReady is defined further below with the other
    # feature-button handlers, but PowerShell resolves function calls at
    # invocation time, so any of these buttons can use it once the whole
    # script has loaded -- this one is left inline since it was written
    # first, functionally identical to the helper.)
    $btnRules.Enabled = $false
    try {
        Update-RulesGrid -Smtp $script:LastSmtp
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Inbox Rules lookup'
    } finally {
        $btnRules.Enabled = $true
    }
})
$grdRules.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $identity = $grdRules.Rows[$e.RowIndex].Cells['Identity'].Value
    $rule = $script:CurrentRules | Where-Object { $_.Identity.ToString() -eq $identity } | Select-Object -First 1
    if ($rule) {
        Show-RuleDetailForm -Rule $rule -Smtp $script:LastSmtp
    } else {
        $txtOut.Text = "Could not match that row back to a rule -- click 'Inbox Rules' again to refresh."
    }
})
# ---- Out of Office / Calendar Perm / Health handlers ----
function Test-MailboxLookupReady {
    if (-not $script:LastSmtp -or -not $script:Target -or $script:Target.Type -ne 'Mailbox') {
        $txtOut.Text = "Run a Mailbox lookup first."
        return $false
    }
    return $true
}
$btnOOF.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-OutOfOfficeForm -Smtp $script:LastSmtp }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Out of Office' }
})
$btnCalPerm.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-CalendarPermissionsForm -Smtp $script:LastSmtp }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Calendar Permissions' }
})
$btnHealth.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-MailboxHealthForm -Smtp $script:LastSmtp }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Health' }
})
$btnTrace.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-MessageTraceForm -Smtp $script:LastSmtp }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Message Trace' }
})
$btnDLMembers.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-DLMembershipsForm -Smtp $script:LastSmtp -ADServer (Get-SelectedADServer) }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'DL Memberships' }
})
$btnMailAttribs.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-MailAttribsForm -Smtp $script:LastSmtp -ADServer (Get-SelectedADServer) }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs' }
})
$btnAllAttribs.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-AllAttribsForm -Smtp $script:LastSmtp -ADServer (Get-SelectedADServer) }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'All Attribs' }
})
[void]$form.ShowDialog()
#endregion
