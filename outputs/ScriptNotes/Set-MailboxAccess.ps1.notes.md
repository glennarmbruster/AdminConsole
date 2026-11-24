# Set-MailboxAccess.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 1 -->

```text
.SYNOPSIS
    WinForms tool to grant or remove mailbox delegation for multiple mailboxes and users in a single run.

.DESCRIPTION
    Manages FullAccess and Send-As permissions across all combinations of one or more mailboxes
    and one or more users (many-to-many). Every user in the Users box receives the selected
    permission on every mailbox in the Mailboxes box.

    Lookup key: PrimarySmtpAddress � documented exception to the sAMAccountName default.
    Target objects are shared mailboxes and user mailboxes identified by email address, not
    by sAMAccountName. Local-part alone is accepted; the script silently appends @expresspros.com.

    Actions (mutually exclusive radio buttons):
      Full & Send-As � Grants FullAccess (AutoMapping on) + Send-As. Sets
                       MessageCopyForSentAs = $true on each mailbox so Send-As mail
                       lands in the mailbox Sent Items.
      Full Access    � Grants FullAccess only (AutoMapping on).
      Remove All     � Removes FullAccess and Send-As if present.

    What this script does NOT change:
      - MessageCopyForSentAs is NEVER cleared, even on Remove � other delegates may rely on it.
      - AutoMapping is always enabled on Add; it is not configurable from this form.
      - No on-prem Active Directory attributes are modified.

    Pre-flight: every mailbox and user must resolve via Get-Mailbox before any change is made.
    If any identity fails to resolve, no changes are made and the unresolved names are listed.

    Environment: Exchange Online only.
    Connection is assumed already established � no Connect/Disconnect inside this script.

.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-06-18
    Revised         : 2026-07-01 0000
    Run from        : Standard PowerShell with ExchangeOnlineManagement module loaded
    Required modules: ExchangeOnlineManagement
```


## function Get-ParsedInputs

- **L98** Remove any user that is also listed as a mailbox

## function Invoke-MailboxAccess

- **L117** "FullSend" | "Full" | "Remove"
- **L134** --- Pre-flight: all identities must resolve before any change is made ---
- **L158** --- Apply (per mailbox x user pair; failures are isolated) ---
- **L165** Set MessageCopyForSentAs once per mailbox on the Full+Send-As path
- **L205** FullAccess
- **L219** Send-As
- **L249** --- Verification (per mailbox) ---
- **L286-287** --- Logging: one row per mailbox (target object), per Section 21 --- Grammar: <Verb> MailboxAccess <action> <mailbox> (users: u1+u2)

## region WinForms UI

- **L306** Shared state between Parse and Run
- **L310** --- Form ---
- **L319** --- Warning label (always first � red bold, Section 4) ---
- **L328** --- Mailbox label + box ---
- **L341** Enter adds newline; does not trigger any button
- **L344** --- User label + box ---
- **L357** Enter adds newline; does not trigger any button
- **L360** --- Parsed preview label + box ---
- **L377** --- Action group ---
- **L403** --- Buttons: Parse / Run / Cancel (right-aligned) ---
- **L422-423** No AcceptButton set � pressing Enter in the multiline boxes adds a newline, not trigger Run. Clicking Run is the only way to execute.
- **L425** --- Result label + box ---
- **L441** --- Helper: return selected action string ---

## function Get-SelectedAction

- **L448** --- Helper: populate the parsed preview box ---

## function Update-ParsedPreview

- **L467** --- Parse button: show preview without executing ---
- **L497** --- Run button: auto-parses if stale, confirms, then executes ---
- **L505** Silently re-parse if input changed since last Parse click
- **L560** --- Cancel button: close form ---
