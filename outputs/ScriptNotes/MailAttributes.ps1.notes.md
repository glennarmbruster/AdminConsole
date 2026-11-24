# MailAttributes.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Standalone Mail Attributes tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\MailAttributes.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the Mail Attributes window.

    Fully functional from AdminConsole.ps1 too:
        .\MailAttributes.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Mail Attributes
    window for that mailbox.

    Every function below is MOVED, UNCHANGED, from the original monolithic
    Mailbox-Group-Inspector.ps1 (original lines 2000-2666, region "Actions -
    mail attributes"), including every design/verification comment. The only
    edit anywhere in this file that touches original tool behavior is
    mechanical: Show-MailAttribsForm's closing call changed from
        [void]$mailForm.ShowDialog($form)
    to
        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
        # Console, which opens top-left. Helper is in Common.ps1.
        Set-FormScreenCorner -Form $mailForm -Corner 'BottomLeft'
        [void]$mailForm.ShowDialog()
    because there is no longer a parent $form to own it -- this script IS
    the top-level window now. (The two internal dialogs it opens,
    Show-MailAttribsEditDialog and Show-KeepOrRemoveOldPrimaryDialog, already
    took their owner as an explicit -Owner parameter -- $mailForm -- so those
    calls are unchanged.) Every other line, including all comments, is
    byte-for-byte the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
```

- **L50** operator cancelled the prompt

## region Actions - mail attributes (moved unchanged from Mailbox-Group-Inspector.ps1 lines 2000-2666)

- **L53-94** "Mail Attribs" button (2026-07-10). Ports the on-prem AD attribute grid and edit mechanics from the operator's "Renaming Script" (on-prem AD name-change tool) almost verbatim -- same fixed attribute list, same derived PrimarySmtpAddress/proxyAddresses handling, same Keep-as-Alias/Remove prompt, same "Add secondary SMTP address" control -- per explicit instruction to match that script's behavior exactly. Differences are only what's required to fit this tool's existing conventions: - The target user comes from this tool's own Mailbox lookup (Smtp/ADServer already resolved by Resolve-Target / Get-SelectedADServer) instead of a separate UPN-entry/Parse/Environment UI -- Get-MailAttribsADUser matches the AD account by "mail" (Get-ADUser -Filter "mail -eq '$Smtp'"), the exact same match-by-mail pattern already used and verified by Get-MailboxDLMemberships just above. - The grid/dialog live inside one self-contained Show-MailAttribsForm function, matching how every other feature dialog in this script (Show-DLMembershipsForm, Show-MailboxHealthForm, etc.) is built. Because the current AD-user object needs to be read AND reassigned across several separate event handlers (Refresh, the shared edit function used by both double-click and the Edit Attribute button, and Add Secondary) -- not just one -- it is held in $script:MailAttribsCurrentADUser rather than a plain closure variable, for the same reason this file already uses $script:CurrentRules / $script:CurrentTraceResults / $script:DLScanCancelRequested for equivalent cross-handler state: a plain variable assignment inside one event-handler scriptblock does not reliably propagate back out to a sibling handler's scriptblock. - Function/variable names are new (MailAttribs-prefixed where the reference name would otherwise be too generic to safely reuse) to avoid any collision with this file's existing functions; the proxyAddresses helpers and Format-AttributeValue keep their original reference-script names since nothing else in this file defines them. - "All Attribs" (the operator's second requested button, a full every-attribute dump) is a separate follow-up request and is not part of this region. Cmdlets used here (Get-ADUser -Filter/-Properties/-Server, Set-ADUser -DisplayName/-GivenName/-Initials/-Surname/-EmailAddress/-Clear/-Replace, Rename-ADObject -NewName) are the exact same calls, in the exact same shapes, already used and Microsoft-Learn-verified in the reference script this was ported from, and/or already used elsewhere in this file (Get-MailboxDLMemberships, Resolve-Target's on-prem group path) -- not re-verified independently per standards Section 1(a) (prior verified usage stands).
- **L106-109** UserPrincipalName, Company ("Office"), name, and proxyAddresses are read-only/context-only, exactly as in the reference script -- name is the object's CN/RDN (out of scope for a mail-attributes tool) and proxyAddresses is only ever changed narrowly, via PrimarySmtpAddress and "Add Secondary".
- **L115-118** Matches the AD account to this mailbox by its "mail" attribute -- the same lookup Get-MailboxDLMemberships already uses just above, so a mailbox without a synced on-prem account throws the same style of message rather than silently showing an empty grid.

## function Get-MailAttribsADUser

- **L142-143** Reads the primary SMTP address out of proxyAddresses -- the entry prefixed "SMTP:" in uppercase (Exchange convention; lowercase "smtp:" is secondary).

## function Get-PrimarySmtpFromProxyAddresses

- **L151-155** Builds a new proxyAddresses array with $NewPrimary as the SMTP: entry. $KeepOldPrimaryAsAlias controls what happens to the previous primary (the operator's answer to Show-KeepOrRemoveOldPrimaryDialog): kept as a demoted smtp: entry, or dropped entirely. An existing secondary equal to the new primary is removed first so the address never appears twice.

## function Set-PrimarySmtpInProxyAddresses

- **L183-184** Appends $NewSecondary as an smtp: entry. Throws on a duplicate rather than silently creating one.

## function Add-SecondarySmtpToProxyAddresses

- **L200-201** Renders a single attribute value for grid display -- handles the multi-valued proxyAddresses array so it never throws on an odd property type.

## function Format-AttributeValue

- **L212-214** Builds the Attribute/Value DataTable, in the fixed order above. PrimarySmtpAddress has no real AD property of its own -- derived from proxyAddresses here, the single place that derivation happens.

## function Format-MailAttribsTable

- **L236-237** Comma operator required -- a bare "return $table" would unroll the DataTable into its DataRow objects instead of returning the table itself.
- **L240-246** Writes one attribute change. Branches by mechanism exactly as the reference script documents: dedicated Set-ADUser parameters, generic -Replace/-Clear for Exchange-schema attributes with no dedicated parameter, and the derived PrimarySmtpAddress field (rewrites proxyAddresses per $KeepOldPrimaryAsAlias). The 'name' case (Rename-ADObject) is kept for parity with the reference script but is unreachable from this UI, same as there -- 'name' is deliberately excluded from $script:MailAttribsEditableAttributes.

## function Set-MailAttribsAttribute

- **L305-306** Standalone write for the "Add secondary SMTP address" control -- appends one new smtp: entry without touching anything else in proxyAddresses.

## function Add-MailAttribsSecondaryProxyAddress

- **L319-322** Edit dialog for a single attribute -- ported near-verbatim from the reference script's Show-AttributeEditDialog, including the extra warning notes for 'name' and 'PrimarySmtpAddress' (the 'name' branch is kept for parity even though that row is never reachable from this UI, same as there).

## function Show-MailAttribsEditDialog

- **L410-413** Shown only when editing PrimarySmtpAddress AND an existing, different primary is about to be displaced -- ported verbatim from the reference script. Returns 'Keep', 'Remove', or 'Cancel' (Cancel, or closing via the X button, aborts the whole attribute edit -- nothing is written to AD).

## function Show-MailAttribsForm

- **L546-548** Added 2026-07-10c alongside double-click -- see the CellDoubleClick / SelectionChanged handlers below. Same position/size as All Attribs' Edit Attribute button for visual consistency between the two dialogs.
- **L561-567** Current AD-user object for this dialog instance. Script-scoped (not a plain closure variable) because it is reassigned from three separate event handlers below (Refresh, grid double-click edit, Add Secondary) -- see the region header comment for why that requires $script: scope in this file's established pattern. Only one Mail Attribs dialog can be open at a time (modal, ShowDialog($form)), so a single script-scope slot is safe, same as $script:CurrentRules / $script:CurrentTraceResults.

## function Update-MailAttribsFromCurrent

- **L596-603** Grid-row edit path -- reachable two ways as of 2026-07-10c: double-click a row (CellDoubleClick below), or select a row and click "Edit Attribute" (btnMailEdit below, enabled only for an editable row via SelectionChanged). Both call this same function against whatever row is currently selected in $grdMail, so the edit logic itself only exists once. Silently no-ops on a read-only row (defensive -- SelectionChanged already keeps btnMailEdit disabled for those, and double-click checks $script:MailAttribsEditableAttributes itself below).

## function Invoke-MailAttribsEditSelectedRow

- **L694-695** Double-click: select the row (matching the pattern the reference scripts use elsewhere), then run the shared edit function.
- **L703** Selection drives Edit Attribute's enabled state, same as All Attribs.
- **L760-765** CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the monolith's shared parent window. This script has no parent window, so it shows itself as the top-level dialog instead. (Show-MailAttribsEditDialog and Show-KeepOrRemoveOldPrimaryDialog above already take their owner as an explicit -Owner parameter -- $mailForm -- so those calls are unchanged.) Nothing else in this function changed.
