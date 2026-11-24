# MailboxHealth.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Standalone Mailbox Health tool -- Layer 2 reference template for the
    launcher/tool-script architecture refactor.
.DESCRIPTION
    Fully functional on its own:
      .\MailboxHealth.ps1
    prompts for mailbox / Environment / Domain Controller via
    Common.ps1's Show-TargetPromptDialog, then opens the Health window.

    Fully functional from AdminConsole.ps1 too:
      .\MailboxHealth.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Health window for
    that mailbox.

    Get-MailboxHealthSnapshot and Show-MailboxHealthForm below are MOVED,
    UNCHANGED, from the original monolithic Mailbox-Group-Inspector.ps1
    (original lines 1556-1851). The only edit anywhere in this file that
    touches original tool behavior is mechanical: Show-MailboxHealthForm's
    closing call changed from

        [void]$healthForm.ShowDialog($form)

    to

        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin

        # Console, which opens top-left. Helper is in Common.ps1.

        Set-FormScreenCorner -Form $healthForm -Corner 'BottomLeft'

        [void]$healthForm.ShowDialog()

    because there is no longer a parent $form to own it -- this script IS
    the top-level window now. Every other line of both functions,
    including all comments, is byte-for-byte the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 --
              see REFACTOR-DESIGN.md, "Phase 2: extract MailboxHealth.ps1")
```

- **L55** operator cancelled the prompt

## function Get-MailboxHealthSnapshot

- **L59-105** Read-only composite view. Reuses Get-InboxRulesForMailbox (now in Common.ps1 -- see that file's header for why) for the rule list/count rather than querying it twice. Interpretation / verification notes (documented since these terms are not single, unambiguous Exchange fields, and per standards Section 2 every cmdlet/property used for a new capability is checked against Microsoft Learn before delivery): - "Protocol usage" is shown as Get-CASMailbox protocol ENABLEMENT flags (OWA/ActiveSync/POP/IMAP/MAPI/EWS), not actual traffic volume. - "Delegates" now lists Full Access and Send As grantees by name (each from its own cmdlet), not just a combined count. Does NOT include Outlook's separate calendar/folder delegate list. - DL memberships are NOT part of this snapshot -- see DLMemberships.ps1 instead. That scan iterates every DL in the tenant and can take a while; running it automatically inside Health made the whole tool appear to hang (2026-07-09 field report), so it's a separate on-demand tool with its own progress display. - Quota (ProhibitSendQuota, ProhibitSendReceiveQuota), Forwarding (ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward), Litigation Hold (LitigationHoldEnabled, InPlaceHolds), and Retention Hold (RetentionHoldEnabled) are all confirmed Get-Mailbox output properties per Microsoft Learn's "Property sets in Exchange Online PowerShell module cmdlets" (Quota / Delivery / Hold / Retention property sets respectively). Quota values are shown as raw ByteQuantifiedSize text next to Mailbox Size for manual comparison rather than computed as a percentage -- the ByteQuantifiedSize/ Unlimited<T> numeric conversion API was not verified against Microsoft Learn, so this script does not guess at it. - AccountDisabled: Microsoft Learn confirms -AccountDisabled as a Set-Mailbox parameter but this script did NOT find Microsoft Learn confirmation that Get-Mailbox returns an AccountDisabled OUTPUT property (Microsoft's own Get-Mailbox page recommends Get-EXOMailbox instead, whose documented property sets don't list it either). Flagged here as UNVERIFIED per standards Section 2 ("if verification is unavailable, clearly say so before delivery") -- confirm this field is meaningful on first use; if the property doesn't exist on the returned object, PowerShell returns $null here rather than throwing, so a wrong/missing value fails quietly. Treat a blank value as "unknown," not as "account is enabled." - "Suspicious Inbox Rules" flags ENABLED rules with a forwarding, redirect, or delete action (ForwardTo/ForwardAsAttachmentTo/ RedirectTo/DeleteMessage/PermanentDelete/SoftDeleteMessage -- all confirmed Get-InboxRule/Set-InboxRule property names on Microsoft Learn). This is a flag for a human to review as a first-check item on compromised-account or missing-mail tickets, not a judgment that the rule is malicious.
- **L214** DL membership is intentionally NOT scanned here -- see DLMemberships.ps1.

## function Show-MailboxHealthForm

- **L247-249** One-shot timer so "Copied!" reverts to "Copy" without needing a second click -- local to this form (closure-captured), same pattern as every other control reference in this function.

## function Update-HealthFromCurrent

- **L263-266** --- Potential Issues summary first -- the point of this button is "what do I need to know before I escalate this ticket", so lead with anything that looks like a likely cause rather than making the reader hunt for it further down.
- **L350-353** CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the monolith's shared parent window. This script has no parent window, so it shows itself as the top-level dialog instead. Nothing else in this function changed.
