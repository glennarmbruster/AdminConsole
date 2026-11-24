# M365_MailboxHealth.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Tenant-agnostic WinForms Mailbox Health tool -- now a pure UI layer over
    M365Logic.psm1.
.DESCRIPTION
    Fully functional on its own:
        .\M365_MailboxHealth.ps1
    prompts for a mailbox via M365Common.ps1's Show-TargetPromptDialog, then
    opens the Health window.

    Also launchable with a target already known:
        .\M365_MailboxHealth.ps1 -SmtpAddress user@anydomain.example
    skips the prompt and opens straight to the Health window.

    REVISED 2026-07-24 -- converted to the "logic module + UI layer" model.
    Get-MailboxHealthSnapshot and the report-building code that used to live
    in this file have MOVED to M365Logic.psm1 (as
    Get-MailboxHealthSnapshot and Format-MailboxHealthReport), unchanged.
    This file kept ONLY the WinForms window: build controls, call those two
    functions, put the returned text in a TextBox.

    WHAT THIS BUYS, concretely -- the identical two calls this window makes:
        $h = Get-MailboxHealthSnapshot -Smtp $Smtp
        Format-MailboxHealthReport -Smtp $Smtp -H $h
    are now the same two calls M365_MailboxHealth_Web.ps1 makes for its
    browser page, and the same two a scheduled job or an ad-hoc console
    one-liner would make. One implementation, four consumers -- so a fix or
    a new health field lands everywhere at once instead of being copied into
    each tool and drifting. That was the whole point of the split.

    This tool remains cloud-only and has no -Environment /
    -DomainController parameters -- see M365Logic.psm1's header on the
    on-prem module boundary.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24 (ported from ANZMailboxHealth.ps1, made
              tenant-agnostic)
    Revised : 2026-07-24, later -- logic extracted to M365Logic.psm1; this
              file is now UI only. See REVISED note above.
```

- **L49-52** Dot-sourcing M365Common.ps1 also imports M365Logic.psm1 -- see that file's header. So Get-MailboxHealthSnapshot / Format-MailboxHealthReport / Format-ErrorDetail are all available below without importing anything else here.
- **L56** operator cancelled the prompt

## function Show-MailboxHealthForm

- **L89-90** One-shot timer so "Copied!" reverts to "Copy" without needing a second click -- local to this form (closure-captured).
- **L98-100** The whole body of this function is now two module calls plus error handling -- everything that used to be ~120 lines of snapshot and report building here lives in M365Logic.psm1.
