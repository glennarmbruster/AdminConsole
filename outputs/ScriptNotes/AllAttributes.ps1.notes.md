# AllAttributes.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Standalone All Attributes tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\AllAttributes.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the All Attributes window.

    Fully functional from AdminConsole.ps1 too:
        .\AllAttributes.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the All Attributes
    window for that mailbox.

    Every function below is MOVED, UNCHANGED, from the original monolithic
    Mailbox-Group-Inspector.ps1 (original lines 2714-3080, region "Actions -
    all attributes"), including every design/verification comment. The only
    edit anywhere in this file that touches original tool behavior is
    mechanical: Show-AllAttribsForm's closing call changed from
        [void]$allForm.ShowDialog($form)
    to
        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
        # Console, which opens top-left. Helper is in Common.ps1.
        Set-FormScreenCorner -Form $allForm -Corner 'BottomLeft'
        [void]$allForm.ShowDialog()
    because there is no longer a parent $form to own it -- this script IS
    the top-level window now. (Show-AllAttribsEditDialog already takes its
    owner as an explicit -Owner parameter -- $allForm -- so that call is
    unchanged.) Every other line, including all comments, is byte-for-byte
    the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
```

- **L49** operator cancelled the prompt

## region Actions - all attributes (moved unchanged from Mailbox-Group-Inspector.ps1 lines 2714-3080)

- **L52-96** "All Attribs" button (2026-07-10). Ports the operator's separate AD-UserAttributeEditor reference script (full Get-ADUser -Properties * viewer/editor with a curated editable subset) into this tool, following the same adaptation approach used for "Mail Attribs" just above: - Target/AD-server come from this tool's own Mailbox lookup, matched by "mail" (Get-ADUser -Filter "mail -eq '$Smtp'") -- same pattern as Get-MailAttribsADUser / Get-MailboxDLMemberships -- instead of the reference's separate UPN-entry/Parse/Environment UI. - Interaction model initially matched the REFERENCE SCRIPT exactly, per instruction ("same as last -- if they are editable here, they are editable in the final version"): select a row, then click the "Edit Attribute" button, enabled only when the selected row's hidden _Editable flag is 'Y' (grdAll's SelectionChanged handler). 2026-07-10c added double-click as a second, equivalent way to trigger the same edit (Invoke-AllAttribsEditSelectedRow), for parity with Mail Attribs -- both dialogs now support both triggers. - The editable-attribute list is copied verbatim from the reference (Title, Department, Description, Company, Division, Office, OfficePhone, MobilePhone, Fax, HomePhone, StreetAddress, City, State, PostalCode, POBox, EmployeeID, EmployeeNumber, Organization, Initials, HomePage) -- each already documented there as a verified, dedicated Set-ADUser string parameter (Microsoft Learn), never a generic -Replace guess. Everything else Get-ADUser -Properties * returns (SamAccountName, UserPrincipalName, DistinguishedName, ObjectGUID, ObjectSID, UserAccountControl, Manager, GivenName, Surname, DisplayName, mail, the whole proxyAddresses/Exchange-schema set already covered by Mail Attribs, etc.) is view-only here, same as the reference -- not re-derived independently per standards Section 1(a). - Get-ADUser -Properties * returns both ADUser's friendly property names (e.g. "Office") and, when populated, the matching raw LDAP attribute name (e.g. "physicalDeliveryOfficeName") as SEPARATE rows with the same value -- normal Get-ADUser behavior per the reference script's own header, not a bug introduced here. Only the friendly-named row is ever in $script:AllAttribsEditableAttributes, so only it is ever editable; the note label below the grid says so. - Function/variable names are AllAttribs-prefixed to avoid colliding with Mail Attribs' identically-purposed but differently-scoped functions (e.g. Format-AllAttribsValue here additionally handles byte[]/SID/ GUID/DateTime -- types the curated mail-attribute list never produces but a full -Properties * dump routinely does -- so it is intentionally NOT the same function as Mail Attribs' simpler Format-AttributeValue). - Same $script:-scope reasoning as Mail Attribs for the current AD-user object ($script:AllAttribsCurrentADUser), since it's read/reassigned from three separate event handlers (Refresh, Edit, and the SelectionChanged enable/disable check).
- **L103-104** Attribute-object bookkeeping properties added by the ADUser type itself -- never real directory attributes, always excluded from the display grid.
- **L109-111** Matches the AD account to this mailbox by its "mail" attribute -- same pattern as Get-MailAttribsADUser -- but requests every populated attribute (-Properties '*') rather than a fixed list.

## function Get-AllAttribsADUser

- **L133-137** Renders a single attribute value for grid display. Handles arrays, binary, GUID/SID, and DateTime -- ported verbatim from the reference script, since a full -Properties * dump routinely includes all of these (objectGUID, objectSid, whenCreated, lastLogon, etc.), unlike Mail Attribs' curated always-string attribute list.

## function Format-AllAttribsValue

- **L155-158** Builds the Attribute/Value/_Editable DataTable bound to the grid -- every populated property on the ADUser object (minus the bookkeeping properties above), sorted by name. Ported from the reference script's Format-UserAttributeTable.

## function Format-AllAttribsTable

- **L175-176** Comma operator required -- a bare "return $table" would unroll the DataTable into its DataRow objects instead of returning the table itself.
- **L179-182** Set-ADUser via the dedicated named parameter matching $AttributeName (must be a member of $script:AllAttribsEditableAttributes -- enforced by the caller). Blank new value clears the attribute instead of writing an empty string. Ported from the reference script's Set-OnPremADUserAttribute.

## function Set-AllAttribsAttribute

- **L202-205** Edit dialog for a single attribute -- ported near-verbatim from the reference script's Show-AttributeEditDialog (the simpler version, with no extra warning notes, since none of the curated attributes here have the special-case semantics PrimarySmtpAddress/name have in Mail Attribs).

## function Show-AllAttribsForm

- **L317-320** BindingSource is the Microsoft-recommended intermediary for binding a DataGridView to a DataTable -- same as Mail Attribs' grid. Column-header click-to-sort works natively once bound through it (SortMode defaults to Automatic), no custom sort handler needed -- matches the reference script.
- **L361-365** Current AD-user object for this dialog instance -- script-scoped for the same cross-handler reason as $script:MailAttribsCurrentADUser (read and reassigned from Refresh, Edit Attribute, and the SelectionChanged enable/disable check). Only one All Attribs dialog can be open at a time (modal, ShowDialog($form)), so a single script-scope slot is safe.

## function Update-AllAttribsFromCurrent

- **L393** Selection drives Edit Attribute's enabled state.
- **L400-404** Grid-row edit path -- reachable two ways as of 2026-07-10c: select a row and click "Edit Attribute" (matches the reference script's own model), or double-click the row directly (added for parity with Mail Attribs). Both call this same function against whatever row is currently selected in $grdAll, so the edit logic itself only exists once.

## function Invoke-AllAttribsEditSelectedRow

- **L452-454** Double-click: select the row, then run the shared edit function. A double-click on a non-editable row is a silent no-op (the function itself checks $script:AllAttribsEditableAttributes).
- **L463-468** CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the monolith's shared parent window. This script has no parent window, so it shows itself as the top-level dialog instead. (Show-AllAttribsEditDialog above already takes its owner as an explicit -Owner parameter -- $allForm -- so that call is unchanged.) Nothing else in this function changed.
