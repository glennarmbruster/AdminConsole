# ExternalForwarding.ps1

Mail leaving the tenant, from both mailbox settings and user inbox rules.
Shared report conventions in `ARCHITECTURE.md` §16.

## Two halves, deliberately

**Mailbox-level forwarding** is one bulk query — cheap, always runs. Two
distinct properties and both matter:

- `ForwardingAddress` points at a directory recipient. Usually internal, but a
  MailContact or MailUser can point anywhere, so it is resolved and checked
  rather than assumed safe.
- `ForwardingSmtpAddress` is a free-text external address.

**Inbox rules** are one mailbox at a time — `Get-InboxRule` has no bulk form.
This is the slow half and the important half: a mailbox-level forward is usually
an admin action, whereas *a rule quietly forwarding to a personal address is the
classic exfiltration pattern*, and the one an attacker sets after compromising
an account.

## Allowlisting sorts, it does not hide

Domains that are external but expected are still reported, flagged
`Allowlisted`, and sorted down. **Silently dropping a destination is how a real
one gets missed.**

A mailbox that cannot be read is recorded rather than skipped: *"no rules
found"* and *"could not look"* must not appear the same.

Real findings sort first — not allowlisted, no copy kept, rule-based.

## The PARTIAL keyword

The literal string `PARTIAL` is what makes `CompletedTasks.csv` greppable for
incomplete runs. The detail follows it; without the keyword the entry reads as a
clean run to anything scanning the log. (Since the v2 schema this is also a
`Result` column value — see §20 — but the keyword stays for the older rows.)
