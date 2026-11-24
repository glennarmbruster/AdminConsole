# MailboxOneDriveAccess.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 1 -->

```text
.SYNOPSIS
    WinForms tool to grant or remove Exchange mailbox access (Full Access,
    Send-As) and/or OneDrive site collection admin access -- merged from the
    operator's two separate tools, Set-MailboxAccess.ps1 and
    onedriveaccess.ps1, into a single console.
.DESCRIPTION
    REVISED 2026-07-24 -- merge of two previously separate tools:
      - Set-MailboxAccess.ps1 (Full Access / Send-As delegation on Exchange
        Online mailboxes, many-to-many mailbox x user)
      - onedriveaccess.ps1 (SharePoint Online site collection admin rights
        on a mailbox's OneDrive, many-to-many user x custodian)
    Both tools shared the same many-to-many shape (one set of targets, one
    set of grantees, one action, apply to every combination) and the same
    two-box-plus-radio-plus-Parse/Run-preview UX, so merging them into one
    console with a single, wider set of action choices removes the need to
    separately track/run two tools for what is operationally one workflow:
    "give this person access to that account's stuff."

    Intake boxes, unified from the two source tools' four boxes down to two:
      Mailbox(es)        -- was "Mailbox(es)" in Set-MailboxAccess.ps1 AND
                            "Disabled User(s)" in onedriveaccess.ps1. Same
                            concept in both: the target account. The
                            OneDrive site URL is derived from THIS box's
                            address(es), exactly the way onedriveaccess.ps1
                            derived it from its own "Disabled User(s)" box.
      User(s)/Manager(s) -- was "User(s)" in Set-MailboxAccess.ps1 AND
                            "Custodian(s)" in onedriveaccess.ps1. Same
                            concept in both: the grantee(s) receiving
                            access. Every User/Manager in this box receives
                            the selected access on every Mailbox in the
                            other box (same many-to-many cross-product both
                            source tools used).

    Access Type (mutually exclusive radio buttons, NOT independent
    checkboxes -- see DESIGN NOTE below):
      Full Access        -- Add-/Remove-MailboxPermission, AccessRights
                            FullAccess, AutoMapping on (not configurable
                            from this form, matching Set-MailboxAccess.ps1's
                            original behavior).
      Send-As            -- Add-/Remove-RecipientPermission, AccessRights
                            SendAs. On Grant, also sets
                            MessageCopyForSentAs = $true on the mailbox
                            (once per mailbox) so Send-As mail lands in Sent
                            Items -- the original tool only did this on the
                            combined Full+SendAs path since that was the
                            only path that included Send-As; this merge
                            applies it any time Send-As is being granted,
                            standalone or combined, since the flag's whole
                            purpose is tied to Send-As itself, not to Full
                            Access. Per the original tool's own documented
                            rule, MessageCopyForSentAs is NEVER cleared on
                            Remove, in any combination.
      OneDrive           -- Set-SPOUser -IsSiteCollectionAdmin, on the site
                            derived from the Mailbox box (see
                            onedriveaccess.ps1's original URL construction,
                            moved here unchanged: local-part before @, dots
                            -> underscores, appended to
                            https://expresspros-my.sharepoint.com/personal/).
      Full & Send-As     -- both Full Access and Send-As operations above,
                            against the same mailbox/user pairs, in one run.
      Full & OneDrive    -- both Full Access and OneDrive operations above.

    NOT offered, by design, per the operator's explicit rule ("there should
    never be a Send-As, OneDrive need"): Send-As + OneDrive, and Full +
    Send-As + OneDrive, are not selectable combinations. This isn't a
    validation check that blocks an otherwise-selectable state -- Send-As
    and OneDrive together is simply never one of the five radio options, so
    it's structurally impossible to request, not just discouraged.

    DESIGN NOTE on "checkboxes" vs radio buttons: the operator described
    these five options as checkboxes. They're built here as mutually
    exclusive RadioButtons instead, for two reasons: (1) both source tools
    already modeled their own single-action-per-run choice as a RadioButton
    group ("Action (mutually exclusive radio buttons)" in both originals'
    own headers), and this merge combines those two groups into one wider
    group of the same kind, not a new interaction model; (2) the five
    listed options are themselves five complete, mutually exclusive end
    states (exactly one is performed per run), not five independent toggles
    that could sensibly all be checked at once -- modeling them as actual
    checkboxes would either need the same mutual-exclusion logic layered
    back on top (checking one unchecking the others, functionally identical
    to radio buttons but a less standard control for that behavior), or
    would need to reject nonsensical states like "Full Access" and "Full &
    Send-As" both checked simultaneously. If independent, freely-combinable
    checkboxes (Full Access / Send-As / OneDrive as three separate,
    optional toggles) are what's actually wanted instead of five preset
    combinations, that's a straightforward follow-up change -- flag it and
    it can be redone that way, with Send-As/OneDrive mutual exclusion
    enforced by disabling whichever one isn't compatible with the other's
    current check state.

    Grant vs Remove is a second, separate two-way radio group (defaults to
    Grant). Combined with the five Access Type options above, Remove is now
    more precise than either original tool's own Remove:
    Set-MailboxAccess.ps1's single "Remove All Access" always removed BOTH
    FullAccess and Send-As regardless of what was actually granted; here,
    Remove + "Full Access" removes only Full Access, Remove + "Send-As"
    removes only Send-As, Remove + "OneDrive" removes only OneDrive admin,
    and the two combined options remove both of their respective pieces --
    same SKIP-if-not-present logic as both original tools (removing
    something that isn't there is reported, not treated as an error).

    Pre-flight: Get-Mailbox resolution for every Mailbox and User/Manager
    runs ONLY when the selected Access Type touches Exchange (Full Access
    and/or Send-As) -- matching Set-MailboxAccess.ps1's original pre-flight
    exactly. A pure OneDrive run skips it entirely, matching
    onedriveaccess.ps1's original behavior of not requiring an Exchange
    Online connection at all for a SharePoint-only operation. If any
    identity fails to resolve during a pre-flight that DOES run, no changes
    are made at all (Exchange OR OneDrive), same all-or-nothing guarantee
    the original Set-MailboxAccess.ps1 had.

    ADDED 2026-07-24 -- OneDrive link surfaced first: per operator request,
    whenever the run includes a OneDrive grant (Access Type OneDrive or
    Full & OneDrive, Mode = Grant) and at least one User/Manager succeeded
    on a given Mailbox's OneDrive, the very first line(s) of the Result box
    read "OneDrive can be accessed via the following link: <url>" -- one
    such line per Mailbox with a successful grant, ahead of the rest of the
    run log. This only appears on a successful Grant (not Remove, and not
    for a Mailbox where every attempted grant failed), so the message is
    never shown for access that doesn't actually exist.

    What this script does NOT change (carried over from both originals):
      - MessageCopyForSentAs is never cleared, even on Remove.
      - AutoMapping is always enabled on a Full Access grant, not
        configurable from this form.
      - No on-prem Active Directory attributes are modified (both source
        tools were cloud-only; this merge stays cloud-only too).
      - No other SharePoint site permissions, sharing settings, or document
        library permissions are touched; no OneDrive files/folders are
        moved, deleted, or shared; IsSiteCollectionAdmin on other existing
        admins is never cleared.

    Environment: Exchange Online AND/OR SharePoint Online, depending on
    Access Type -- see the dynamic warning label, which updates with the
    Access Type selection to say exactly which connection(s) the current
    selection needs. Connections are assumed already established; this
    script never calls Connect-/Disconnect- for either service, same as
    both originals.
.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-06-18 (Set-MailboxAccess.ps1), 2026-07-01
                       (onedriveaccess.ps1)
    Revised         : 2026-07-24 (merged into MailboxOneDriveAccess.ps1 --
                       see REVISED note above)
    Run from        : Standard PowerShell with ExchangeOnlineManagement
                       and/or Microsoft.Online.SharePoint.PowerShell loaded,
                       depending on Access Type
    Required modules: ExchangeOnlineManagement (Full Access / Send-As),
                       Microsoft.Online.SharePoint.PowerShell (OneDrive)
```


## region Access Type mapping

- **L206-210** Single source of truth for what each of the five Access Type radio buttons actually does -- both the pre-flight gate (Full -or SendAs) and the per-pair action loop read these same three flags, so there is no separate list anywhere else of "which type does what" that could drift out of sync with the radio buttons themselves.

## function Get-AccessFlags

- **L212** 'Full' | 'SendAs' | 'OneDrive' | 'FullSendAs' | 'FullOneDrive'

## function Get-ParsedInputs

- **L242** 'Grant' | 'Remove'
- **L247-249** Remove any user that is also listed as a mailbox -- a mailbox granting itself access (Exchange or OneDrive) never makes sense. Same dedup Set-MailboxAccess.ps1 originally applied.

## function Invoke-MailboxOneDriveAccess

- **L286-289** --- Pre-flight: only when Exchange is actually touched -- matches Set-MailboxAccess.ps1's original behavior. A pure OneDrive run needs no Exchange Online connection at all, so it skips this entirely, matching onedriveaccess.ps1's original behavior.
- **L315-316** Mailboxes where at least one OneDrive grant succeeded -- drives the "OneDrive can be accessed" lines prepended to the output below.
- **L321-323** Set MessageCopyForSentAs once per mailbox, any time Send-As is being granted (standalone or combined) -- see header note. Never touched on Remove, in any combination.
- **L419** --- Verification ---
- **L468** --- Logging: one row per mailbox (target object), per Section 21 ---
- **L479-481** --- ADDED 2026-07-24: OneDrive link, prepended as the very first line(s) of the output -- see header ADDED note. Built last, then stitched in front of everything above.

## region WinForms UI

- **L498** Shared state between Parse and Run
- **L502** --- Form ---
- **L511-513** --- Warning label (Section 4) -- text is dynamic, see Update-ConnectionWarning below: it updates with the Access Type selection to name exactly which connection(s) (Exchange Online and/or SharePoint Online) that choice needs.
- **L521** --- Mailbox label + box ---
- **L534** Enter adds newline; does not trigger any button
- **L537** --- User/Manager label + box ---
- **L550** Enter adds newline; does not trigger any button
- **L553** --- Parsed preview label + box ---
- **L570** --- Mode group: Grant / Remove ---
- **L590-593** --- Access Type group: the five mutually exclusive options -- see header DESIGN NOTE for why these are RadioButtons, not checkboxes, and why Send-As + OneDrive is not one of the five (structurally impossible to select, not merely validated against).
- **L631** --- Buttons: Parse / Run / Close (right-aligned) ---
- **L650-651** No AcceptButton set -- pressing Enter in the multiline boxes adds a newline, not trigger Run. Clicking Run is the only way to execute.
- **L653** --- Result label + box ---
- **L669** --- Helpers: return selected Access Type / Mode strings ---

## function Get-SelectedMode

- **L682-685** --- Dynamic warning label -- names exactly which connection(s) the currently selected Access Type needs, instead of a generic "connect before running" that doesn't distinguish EXO-only, SPO-only, and EXO+SPO selections.

## function Update-ConnectionWarning

- **L701** --- Helper: populate the parsed preview box ---

## function Update-ParsedPreview

- **L717** --- Parse button: show preview without executing ---
- **L748** --- Run button: auto-parses if stale, confirms, then executes ---
- **L757** Silently re-parse if input changed since last Parse click
- **L808** --- Close button ---
- **L811-816** ADDED 2026-07-24 -- open bottom-left, so a tool never covers the Admin Console (which opens top-left). Done inline rather than via Common.ps1's Set-FormScreenCorner because this tool is deliberately standalone and does not dot-source Common.ps1 -- see the header. WorkingArea (not Bounds) so the taskbar is respected, and clamped so an over-tall form can't lose its title bar off the top of the screen.
