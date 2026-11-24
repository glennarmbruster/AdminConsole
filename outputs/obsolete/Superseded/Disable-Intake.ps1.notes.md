# Disable-Intake.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Unified intake for HQ and Franchise user-disable requests.
 
.DESCRIPTION
    Single WinForms intake that replaces the two former (near-identical) intake
    scripts. One HQ request at a time plus a dedicated Manager field; the
    Franchise box accepts multiple blocks.
 
      - HQ section   : paste one SNOW description (Employee Name / UPN / Last Working Day)
                        plus an optional HQ Manager field.
                        -> writes disablelist.csv -> calls disable-new-hq.ps1
 
      - Franchise box : blocks formatted as Logon Name / Delegate / OneDrive
                        -> writes disablelist.csv -> calls disable-new.ps1
 
    Either section may be left empty; the other is still processed.
    When both contain requests, Franchise runs first, then HQ.
 
    Lookup key:
      - Franchise : sAMAccountName (the Logon Name). Validated alphanumeric.
                    The Franchise backend uses Get-ADUser -Identity, which accepts
                    only a sAMAccountName - never a UPN/email.
      - HQ        : UPN. The HQ backend uses Exchange cmdlets only (no Get-ADUser),
                    which resolve a UPN directly.
 
    HQ manager: entered in the Manager field as "bob smith" or "bob.smith".
    Both resolve to bob.smith@expresspros.com. This address grants BOTH delegate
    (email) and OneDrive access. If the field is left empty, no access is granted
    (the mailbox is still converted to Shared and hidden from the GAL by the backend).
 
    What this script does NOT do:
      - It does not modify any mailbox, AD object, or SharePoint site itself.
        All changes are made by the backend scripts it calls.
      - It does not connect or disconnect any session. Connections are assumed
        already established (see the on-form warning and "Run from", below).
      - It does not alter CN, UPN, or sAMAccountName.
 
    Fail-fast note: the backend scripts call exit() in their catch block. If the
    first environment's run errors, the second will not run; re-run with the
    remaining section.
 
.NOTES
    Name            : disable-intake.ps1
    Author          : Glenn Armbruster
    Created         : 2026-06-18
    Revised         : 2026-07-02 0900
    Run from        : on-prem Exchange Management Shell, with Exchange Online,
                      SharePoint Online, and ActiveDirectory connected in the
                      same session (the backends use all of these).
    Required modules: ExchangeOnlineManagement; Microsoft.Online.SharePoint.PowerShell;
                      ActiveDirectory  (on-prem RemoteMailbox cmdlets via the EMS host)
```


## function Normalize-Manager

- **L83-85** Accepts "bob smith", "bob.smith", or "bob.smith@expresspros.com". Returns the local part with spaces converted to dots (e.g. "bob.smith"). Returns "" for blank input. Returns $null if the result fails validation.
- **L89** Strip domain if a full email was pasted
- **L91** "bob smith" -> "bob.smith"
- **L93** Must be letters, digits, dots, underscores, or hyphens

## function Test-Identity

- **L105** Accepts a plain alphanumeric alias/logon OR an email/UPN.
- **L109-110** alias / logon email / UPN

## function New-Record

- **L142-144** Each parser returns an object: { Ok = <bool>; Records = <array> }. Ok=$false means a validation error was shown and the caller should stop. Ok=$true with an empty Records array means the section was simply empty.

## function Parse-HQ

- **L174** other labels -> informational, ignored
- **L177** Colon-less lines (e.g. SNOW narrative) are ignored; manager is entered separately.

## function Parse-Franchise

- **L222** Email:, etc. -> informational, ignored

## region WinForms UI

- **L308** Connection warning
- **L318** --- HQ block ---
- **L344** --- Franchise block ---
- **L359** Note
- **L377** error already shown; keep form open
- **L379** Resolve manager from the dedicated field (empty = no delegation/OneDrive)
- **L391** error already shown; keep form open
- **L407** keep form open
- **L430-432** Post-dialog: write each CSV and hand off, one environment at a time. Each block is independent - an empty section is simply skipped. Franchise first, then HQ. (Fail-fast caveat in header.)
