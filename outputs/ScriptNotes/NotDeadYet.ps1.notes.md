# NotDeadYet.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Restores up to 15 returning (previously termed) Franchise users in a single batch run.

.DESCRIPTION
    Batch intake tool: paste up to 15 logon IDs (one per line, or comma/semicolon
    separated), Franchise only. For each logon ID, performs:

      1. Unhide the on-prem mailbox from address lists (Set-RemoteMailbox
         -HiddenFromAddressListsEnabled $false).
      2. Ensure no forwarding is set, cleared in Exchange Online only
         (ForwardingAddress / ForwardingSmtpAddress). Set-RemoteMailbox has
         no ForwardingAddress/ForwardingSmtpAddress parameters -- confirmed
         against Microsoft Learn -- forwarding is a mailbox property that
         lives in the cloud, so it is never touched on-prem.
      3. Ensure the mailbox type is Regular (not Shared), on-prem and in
         Exchange Online.
      4. Re-add the user to the "E5 License Fran - Full" AD group in the
         franchises.expresspersonnel.com domain, if not already a member.

    All target mailboxes are confirmed to already exist in the cloud
    (hybrid), so the Exchange Online lookup by PrimarySmtpAddress is
    expected to always resolve.

    Lookup key is the pasted Logon value (sAMAccountName) for every target.

    This script deliberately does NOT touch: CN, display name, UPN, password,
    OU location, or any group membership other than "E5 License Fran - Full".

    Environment: Franchise only. There is no environment selector by design
    (operator requirement) -- this tool never touches HQ.

    Domain controller: pinned to fran-vm-dc03.franchises.expresspersonnel.com
    for the entire session (operator requirement). There is no DC dropdown --
    the DC is fixed and shown read-only in the form so the operator can
    confirm it before running.

    Requires BOTH an on-prem Exchange (Franchise) session AND an Exchange
    Online session connected in the same PowerShell host before running.

    Mode:
      - Restore  : performs the four actions above for every logon in the
                   list, with an inline BEFORE/AFTER snapshot per user.
      - Verify   : read-only. Re-queries current state for every logon in
                   the list without changing anything. Use this after
                   replication has had time to propagate, or any time later
                   to confirm current state.

.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-07-07
    Revised         : 2026-07-07 1600
    Run from        : PowerShell host with an on-prem Exchange Management
                       Shell (Franchise) session AND an Exchange Online
                       (EXO) session both connected.
    Required modules: ActiveDirectory; on-prem Exchange (EMS); ExchangeOnlineManagement (EXO)
```


## function Initialize-ADSession

- **L98** Pins on-prem Exchange / AD resolution to DC3 for the whole call.

## function Get-OnPremState

- **L103-106** Forwarding is intentionally excluded here: Set-RemoteMailbox has no ForwardingAddress/ForwardingSmtpAddress parameters (confirmed against Microsoft Learn) -- forwarding is a cloud mailbox property, read from Get-CloudState instead.

## function Get-CloudState

- **L116** Get-Mailbox (EXO) -- ambiguous name, reserved for cloud only in this script (see header).

## function Restore-User

- **L193** ---- BEFORE ----
- **L204-205** ---- ACTIONS (on-prem) ---- Set-RemoteMailbox (on-prem): unhide
- **L208-211** Set-RemoteMailbox (on-prem): force Regular type. NOTE: Set-RemoteMailbox has no ForwardingAddress/ForwardingSmtpAddress parameters (confirmed against Microsoft Learn) -- forwarding is a cloud mailbox property and is cleared below via Set-Mailbox (EXO) only.
- **L214** Re-add to E5 License group if not already a member (AD, on-prem, DC3-pinned)
- **L225-226** ---- ACTIONS (EXO) ---- Set-Mailbox (EXO): force Regular type, clear forwarding
- **L236** ---- AFTER ----

## region WinForms UI

- **L280** Warning label -- first visible element
- **L289** DC info -- fixed, no selector by design
- **L305** Input area
- **L320** Parsed preview
- **L337** Mode group (Restore / Verify) + Parse / Run / Close buttons on one row
- **L375** Result
- **L391** ---- shared parse state ----
