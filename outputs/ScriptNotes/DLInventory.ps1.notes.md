# DLInventory.ps1

Mail-enabled groups inventoried from on-prem AD for migration planning. Shared
report conventions in `ARCHITECTURE.md` §16.

## Read from AD, not Exchange — deliberately

`Get-DistributionGroup` exists in **both** the on-prem shell and Exchange
Online, and which one answers depends on module load order (§2). This report is
specifically about the on-prem groups being migrated, so guessing wrong would
inventory the wrong side. `Get-ADGroup` is unambiguous.

Mail-enabled **security** groups are included on purpose: they carry both
membership and permissions, so they are the awkward ones to migrate and the ones
most likely to be forgotten.

## Nested groups without N round trips

Nested-group detection needs to know whether each member DN is itself a group.
Asking AD per member would be one call per member — a 500-member group would
cost 500 round trips. Instead every group DN in the forest is fetched **once**
as a lookup set, turning the test into an O(1) set membership check.

If that index cannot be built the run continues: `NestedGroups` reports 0 rather
than failing, and the summary flags that the figure is unavailable.

Owner DNs repeat heavily — one manager typically owns many groups — so each
distinct owner is resolved once.

## Ordering is migration order

The awkward ones first: nested groups, then no owner, then size. Those are the
ones needing a decision rather than a script.

`whenChanged` drives the stale threshold and is a **weak signal** — it moves for
reasons unrelated to use — but it is the only age indicator AD offers without
message tracking. Noted in the report summary so nobody over-reads it.
