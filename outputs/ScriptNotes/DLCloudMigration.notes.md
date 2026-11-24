# DLCloudMigration

Re-anchoring on-premises distribution groups as cloud-native groups.

`Projects\DLMigration\`

> **Renamed and relocated, 2026-08-13.** This was `Projects\Express\DLCloudMigration\`.
> The three scripts it described were never run and are parked in
> `obsolete\DLCloudMigration-v1_2026-08-13\`; the work restarted in
> `Projects\DLMigration\` with `Inventory-AllGroups.ps1`.
>
> The `Express` level came out of `Projects\` entirely -- one tenant per
> toolkit build makes it redundant, and it was a folder to click through on
> the way to everything. `Projects\<ProjectName>\` is the convention, and
> always was (ARCHITECTURE section 24).
>
> The reasoning below still holds. The phasing and the risks did not change
> when the folder did.

**Not the same project as DLtoSharedMbx.** That one deletes franchise DLs
on-premises and creates shared mailboxes in the cloud. This one keeps the group
as a group and moves where it is mastered. Different scripts, different folder,
different risks.

## Phasing

| Phase | Scope | Count | Notes |
|---|---|---|---|
| 1 | HQ | ~240 in scope (299 estimated) | Two weekends: small pilot, then the rest |
| 2 | Franchise | ~4,778 | Largely via the DLtoSharedMbx process instead |

The franchise half is resource-bound rather than technically hard: permissions
assigned that way do not set automapping on the shared mailboxes, so each
conversion needs hands. A year or more is the working assumption.

## Scope

**In:** distribution groups, dynamic distribution groups, room lists.
**Out:** mail-enabled security groups, security groups.

Filtering is on `RecipientTypeDetails`, not `RecipientType`. `RecipientType`
reports every mail-enabled group as `MailUniversalDistributionGroup` regardless
of whether it is a security group, so filtering on it would pull the entire
out-of-scope population into the migration.

Out-of-scope groups are still exported where they are **members** of an in-scope
group, flagged as blockers. A cloud distribution group cannot hold a
mail-enabled security group that does not exist in the cloud, so that membership
disappears at migration -- silently, unless something records it.

## Scripts

    1-Export-Groups.ps1        read-only, on-premises Exchange, 8 CSVs
    2-Export-Usage.ps1         read-only, Exchange Online message trace
    3-Assess-PilotCandidates.ps1  read-only, reads the CSVs, writes one more

All three are read-only. Nothing in this project writes to a directory yet.

### Decisions taken

**Source of truth is on-premises Exchange, not AD.** The groups are mastered
on-premises, and most of what has to be recreated -- moderation, restrictions,
`LegacyExchangeDN` -- is either absent from raw AD or stored as DNs that mean
nothing in the cloud. The existing `AD-DLInventory.ps1` is a planning inventory
and is not sufficient for a rebuild.

**Every Exchange call goes through `Invoke-OnPremCommand`.** Both Exchanges are
normally connected in the same process, and a bare `Get-DistributionGroup`
binds to whichever module imported last. That exact bug produced a cloud
inventory when an on-premises one was wanted, earlier the same day this project
started. The export is the last place it could be afforded.

**One timestamp per run.** All eight CSVs share it, and scripts 2 and 3 load by
that timestamp rather than by "newest file". Attributes from one run and members
from another produce an answer that looks authoritative and describes a set of
groups that never existed.

**Empty CSVs are written, not skipped.** A missing file cannot be distinguished
from a failed step.

## What the export captures, and why

Beyond the obvious attributes:

- **`LegacyExchangeDN`** -- the single most important column. It must become an
  `X500:` proxy on the new cloud group. Without it, cached Outlook autocomplete
  entries and reply-all on existing threads bounce. The symptom appears days
  later and does not look related to the migration.
- **Send-As, read from the AD ACL** via `Get-ADPermission`. It is not a property
  of the group and does not appear in `Get-DistributionGroup` output at all, so
  an export built only from the group object silently loses every Send-As
  delegate. `GrantSendOnBehalfTo` is a *different* right and is captured
  separately.
- **`EmailAddressPolicyEnabled`** -- where true, addresses are policy-generated
  on-premises and must be set explicitly in the cloud.
- **Transport rule references.** A rule scoped to a group keeps matching the old
  object; recreate the group and the rule stops applying with no error anywhere.
- **Restriction and moderation lists resolved to SMTP**, not left as DNs.
- **Every member with a `Reproducible` flag and a reason.**

## OPEN -- the residual cloud object (KB 4096314)

**This project has the same landmine that has already failed once on
DLtoSharedMbx, and it has not been solved there either.**

Removing a synced group from sync scope does not remove its cloud copy. Exchange
Online falls back to the `@...onmicrosoft.com` address added at first sync, and
the object stays in the address book accepting mail and dropping it.

On DLtoSharedMbx the mitigation -- hide on-premises, let it sync, then
mail-disable -- was tested on 2026-08-07 against two groups and **failed**: the
mail-disable strips the whole Exchange attribute set including the hidden flag,
and both groups returned to the GAL after the next sync.

This project cannot assume that problem is solved. The order of operations for
the cutover is unresolved and must be settled before the pilot, not during it.
Untested options: deleting the cloud object in Entra and forcing a re-sync, or
`Set-ADObject -Replace` on the hidden flag *after* the mail-disable.

## OPEN -- data quality found on the first run

The first HQ export surfaced Exchange validation warnings that had been
scrolling past unrecorded. They are now captured to `DataQuality_*.csv`.

Two real examples from the first run:

- `AR.CoppellTX` -- `WindowsEmailAddress` is `AR,CoppellTX@ExpressPros.com`.
  A comma in the local part is not a legal SMTP address. That group cannot be
  recreated in the cloud until it is corrected on-premises.
- `Team Midtown Dallas` -- `MemberJoinRestriction` is `ApprovalRequired` with no
  manager set. Inconsistent state.

Exchange emits these in runs: a header naming the object, then detail lines
naming only the property. The detail lines are the actionable half and carry no
identity, so the parser carries the last-seen object forward. Without that, the
useful half of every warning arrives orphaned.

The file carries an `Environment` column because these warnings are raised
during the **forest-wide** read that happens before the environment filter, so
an HQ run legitimately reports franchise objects. That is a free defect list for
phase 2, but only because the file says which is which.

## Pilot selection

`3-Assess-PilotCandidates.ps1` scores for safety, then proposes a set by
**greedy feature coverage** rather than by top-N score.

Twelve highest-scoring groups would be twelve small, plain, unmoderated lists.
That pilot succeeds and proves only that easy groups migrate. The proposed set
instead takes the highest-scoring group that exercises a feature nothing chosen
so far covers -- moderation, Send-As, restrictions, room list, existing X500 --
and only falls back to plain high scorers once coverage is exhausted. Features
that appear ONLY on blocked or low-scoring groups are reported as not covered,
because meeting them for the first time in week three is the risk worth naming.

Scoring is written as subtractions from 100 rather than a weighted sum, so every
number in the `Why` column traces back to one named reason.

**A pilot on a dead list proves nothing** -- nobody would report a failure, so a
clean result is indistinguishable from no result. Groups with no traffic in the
measured window are scored down for that reason.

## Usage evidence and its limit

`2-Export-Usage.ps1` counts delivery per group over up to ten days.

The verdict column says `NoTrafficInWindow`, deliberately not `Dead`. Message
trace reaches back about ten days. Monthly reports, quarterly notices and
seasonal lists all look dead in that window. Traffic proves a group is live;
silence proves nothing, and a column called `Dead` would be acted on as though
it did.

## Still to build

- Migration scripts. Nothing here writes to a directory yet.
- The cutover order of operations, blocked on the residual-object question above.
- A comparison script: export the rebuilt cloud group and diff it against the
  on-premises snapshot, so "it worked" is a checked claim rather than an
  impression.
