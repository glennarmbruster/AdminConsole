# PermissionsAudit.ps1

Full Access, Send As and Send on Behalf delegations across every mailbox.
Shared report conventions in `ARCHITECTURE.md` §16.

## Slow by nature, not by implementation

`Get-MailboxPermission` and `Get-RecipientPermission` are both per-mailbox
cmdlets with no bulk form, so cost scales directly with mailbox count. Worth
knowing before anyone tries to "optimise" it.

**Send on Behalf** comes from `GrantSendOnBehalfTo`, already retrieved with the
mailbox, so it costs nothing extra. It is a distinct permission people routinely
forget exists and it appears in neither of the two cmdlets above.

## Exclusions are exclusions, not an allowlist

`NT AUTHORITY\SELF` and friends are identities Exchange reports on every
mailbox. They are the platform's own and would bury the real delegations, so
they are **excluded rather than flagged** — unlike an allowlist entry, these are
never interesting.

## Ordering

Grouped by **delegate**, not by mailbox. One person holding access to many
mailboxes is the pattern worth seeing, and it is invisible when sorted the other
way.

## PARTIAL

`PARTIAL` is the keyword `CompletedTasks.csv` is scanned for. The count alone
did not match it, so incomplete runs of this report were invisible to that scan
until it was added.
