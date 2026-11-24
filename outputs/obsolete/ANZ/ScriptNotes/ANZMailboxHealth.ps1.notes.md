# ANZMailboxHealth.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    ANZ (cloud-only, expresspros.com.au) Mailbox Health tool.
.DESCRIPTION
    Fully functional on its own:
        .\ANZMailboxHealth.ps1
    prompts for a mailbox via ANZCommon.ps1's Show-TargetPromptDialog, then
    opens the Health window.

    Fully functional from ANZAdminConsole.ps1 too:
        .\ANZMailboxHealth.ps1 -SmtpAddress user@expresspros.com.au
    skips the prompt entirely and opens straight to the Health window for
    that mailbox.

    Built the same way as SRGMailboxHealth.ps1 -- both are ported from the
    original suite's MailboxHealth.ps1. Get-MailboxHealthSnapshot and
    Show-MailboxHealthForm are unchanged (every check they run is EXO-only:
    Get-MailboxStatistics, Get-Mailbox, Get-MailboxPermission,
    Get-RecipientPermission, Get-InboxRule, Get-CASMailbox -- there was
    never any on-prem dependency here to remove). The only edits vs. the
    original are the same mechanical ones every cloud-only tool script in
    this family needs:
      - Dot-sources ANZCommon.ps1, not Common.ps1 or SRGCommon.ps1.
      - No -Environment / -DomainController parameters -- this suite has no
        on-prem concept at all (see ANZCommon.ps1 header).
      - Resolve-ToolTarget is called with only -SmtpAddress.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21 (ported from MailboxHealth.ps1, same way as
              SRGMailboxHealth.ps1)
```

- **L43** operator cancelled the prompt

## function Get-MailboxHealthSnapshot

- **L47-68** Read-only composite view. Reuses Get-InboxRulesForMailbox (ANZCommon.ps1) for the rule list/count rather than querying it twice. Interpretation / verification notes carried over from the original: - "Protocol usage" is shown as Get-CASMailbox protocol ENABLEMENT flags (OWA/ActiveSync/POP/IMAP/MAPI/EWS), not actual traffic volume. - "Delegates" lists Full Access and Send As grantees by name (each from its own cmdlet), not just a combined count. Does NOT include Outlook's separate calendar/folder delegate list. - DL memberships are NOT part of this snapshot -- no DL Memberships tool exists yet in this suite; add one the same way the original suite's DLMemberships.ps1 was built if that's needed later. - Quota, Forwarding, Litigation Hold, and Retention Hold are all confirmed Get-Mailbox output properties per Microsoft Learn's "Property sets in Exchange Online PowerShell module cmdlets." - AccountDisabled: Microsoft Learn does NOT confirm Get-Mailbox returns this as an output property (only as a Set-Mailbox parameter). Flagged UNVERIFIED, same as the original -- treat a blank value as "unknown," not as "account is enabled." - "Suspicious Inbox Rules" flags ENABLED rules with a forwarding, redirect, or delete action -- a flag for human review, not a judgment that the rule is malicious.

## function Show-MailboxHealthForm

- **L209-210** One-shot timer so "Copied!" reverts to "Copy" without needing a second click -- local to this form (closure-captured).
