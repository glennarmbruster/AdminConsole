# StaleAccounts.ps1

Enabled, licensed accounts that have not signed in recently. AdminConsole only.
Shared report conventions in `ARCHITECTURE.md` §16.

## A different failure from DisabledLicensed

Disabled accounts are that report's job. This one is about accounts that are
still **live, still licensed, and unused** — which is a different problem:
somebody left and was never disabled at all.

## Buckets, not one threshold

100 days and 1,000 days are both "stale" but they are not the same
conversation, so the report buckets rather than applying a single cutoff.

A licensed account created very recently and never used is probably a starter
who has not begun yet, so there is a grace period. Past it, never-logged-on
counts as stale.

Never-used first, then longest dormant.
