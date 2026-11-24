# DL Migration to Microsoft 365 — project record

Source-anchoring on-premises distribution groups to Microsoft 365.

Author: Glenn Armbruster · Started 2026-08-13 · **Living document**

> **What this is.** The technical record for the project: decisions and the
> reasoning behind them, real numbers, cutover mechanics, and what is still
> open. The management-facing plan is a separate document and deliberately
> says less.
>
> **What this is not.** A runbook. Individual scripts document themselves;
> this says why they exist and how they fit together.

---

## 1. Driver

Exchange decommission, and modernisation of group administration.

An object mastered in on-premises AD must be managed on-premises — a synced
distribution group is read-only in Exchange Online. All ~5,100 groups are in
that state, so administering them requires on-premises Exchange management.
Moving them removes one category of object from that dependency.

**One category, not all of them.** While user accounts still sync from AD,
some on-premises recipient management remains necessary regardless. This
project is a prerequisite on the decommission path, not the thing that retires
the server. Say so plainly when asked — the expectation gap is the risk here,
not the work.

Not drivers: licensing, tenant consolidation. Confirmed 2026-08-14.

---

## 2. The estate

Full inventory, 2026-08-13, via `Inventory-AllGroups.ps1`.

| Directory | Total | Distribution | Mail-enabled security |
|---|---:|---:|---:|
| HQ | 298 | 241 | 54 |
| Franchise | 4,781 | ~3,194 | 1,587 |
| **Total** | **~5,079** | **~3,435** | **1,641** |

HQ also holds **2 `MailNonUniversalGroup`** objects — legacy, pre-universal
scope. See section 6.

Franchise is ~800 offices.

---

## 3. Scope

**In:** all distribution groups, both directories, moved to cloud-mastered.
Plus the *mail function* of mail-enabled security groups.

**Out:**

- Security groups themselves — stay on-prem, AD team owns them
- Room lists — move with resource mailboxes, separate work
- Microsoft 365 Groups (`GroupMailbox`) — already cloud-mastered
- Retiring unused groups — a business decision, deliberately excluded so it
  cannot delay the migration

### Division of labour

Glenn owns the Exchange objects. The AD team owns whatever security groups
remain on-premises. This is a boundary agreement, not a work request — nothing
is being asked of them.

### The 1,641 mail-enabled security groups

They do two jobs — deliver mail, grant access — and those jobs belong in
different places now. Split them: security group stays on-prem untouched, a new
cloud distribution group carries the mail.

**Every security group stays, whether or not it grants anything.** Determining
which actually appear in an ACL would need a file/system permissions audit that
is neither available nor necessary. The asymmetry decides it: a needless
security group left in place costs clutter; a needed one removed breaks access
on a Monday morning.

**Known open question — membership drift.** Two objects, one membership,
nothing keeping them aligned. The obvious fix is nesting the on-prem security
group inside the cloud DL, so there is one membership and two objects.
**Probably does not work** — a DL's members must be mail-enabled recipients,
and mail-disabling the on-prem group stops it being one. *Verify before
designing around it*, because the answer decides whether this is permanent dual
maintenance. Glenn's position (2026-08-14): drift is acceptable, he owns
Exchange objects only.

---

## 4. Group types, verified against Microsoft Learn

`Get-DistributionGroup` returns exactly four subtypes:

| RecipientTypeDetails | Disposition |
|---|---|
| `MailUniversalDistributionGroup` | migrate |
| `MailNonUniversalGroup` | migrate (legacy, 2 at HQ) |
| `MailUniversalSecurityGroup` | split |
| `RoomList` | out of scope |

`DynamicDistributionGroup` comes from a **separate cmdlet** and is not in that
result set — omitting the second call under-reports the estate with no sign
that it has. `GroupMailbox` is returned by neither and is cloud-mastered
already.

Between them: no gap, no double-count.

**`-ResultSize Unlimited` is load-bearing.** Default is 1000 and it truncates
*silently* — no warning, just a short list that looks complete.

---

## 5. Cutover mechanics

### The sequence, per batch

1. Capture — addresses, `legacyExchangeDN`, full membership, office code
2. Mail-disable on-premises
3. **Wait for sync** to release the address
4. Create the cloud DL with the same primary address
5. Stamp the old `legacyExchangeDN` as `X500:`
6. Verify before starting the next batch

### `legacyExchangeDN` is the whole ballgame

Every group carries one. Outlook autocomplete and replies-to-old-messages
resolve against it. Create the cloud DL without carrying it across as an X500
and Monday brings NDRs from anyone replying to Friday's mail — at franchise
scale, thousands of them, with no obvious cause.

**Capture before touching anything.** Once the on-prem object is gone the value
is gone with it.

Same trap sits behind `Disable-DistributionGroup` + `Enable-DistributionGroup`,
which drops `legacyExchangeDN` and proxy addresses. Do not reach for that pair
as a quick fix for anything.

### What sets the schedule

Not the group count, and not headcount:

- **The sync wait** between removing on-prem and creating in cloud. Mandatory,
  per batch, and the reason this is weekend work rather than an evening.
- **EXO write throttling.** ~3,200 new DLs plus 1,587 splits is a lot of write
  cmdlets, and this determines batch size.

Both get **measured in the pilot**, not assumed. Time one group end to end and
multiply — that number times batch size is the real schedule.

### Batch by office, never alphabetically

Alphabetical failure = a stripe across all 800 offices. By office = a handful
that can be named and phoned. Hence the office code in the capture file.

### Failure handling

At 298 you can eyeball what broke. At 4,781 you cannot. Anything failing
mid-batch goes to the **retry queue** — the pattern already built for
`Disable-User` — not a transcript nobody reads at 02:00.

### Rollback

There isn't one, once an address is released. That is *why* batches stay small
enough to verify before the next starts.

---

## 6. Exceptions

### The 2 `MailNonUniversalGroup` objects

Must be Universal to be mail-enabled. `Set-Group -Identity X -Universal` is the
only supported Exchange cmdlet, and Microsoft prescribes it for exactly this.

**It is not a bypass of what ADUC refuses.** Same AD constraints apply:

- Global → Universal: blocked if the group is a *member of* another Global group
- DomainLocal → Universal: blocked if it *contains* another DomainLocal group

Un-nest, convert, re-nest.

**Undocumented behaviour to check after:** whether `msExchRecipientTypeDetails`
re-stamps. `Get-Group` reads `groupType` from AD and shows Universal
immediately; `Get-Recipient` may still report `MailNonUniversalGroup`. Check
both. Do **not** fix a non-restamp with disable/re-enable — see section 5.

Attribute mapping, verified: `0 = Closed, 1 = Open, 2 = ApprovalRequired` for
both `msExchGroupJoinRestriction` and `msExchGroupDepartRestriction`. Note 0/1
are *not* in the order you would guess.

### Directory data quality

Two malformed objects found and fixed 2026-08-13:

- An HQ DL with `MemberJoinRestriction = ApprovalRequired` and no `ManagedBy`
  — approval requests routed nowhere
- A franchise user with `AR,CoppellTX@ExpressPros.com` — comma instead of a
  dot, so `WindowsEmailAddress` would not convert to an SMTP address

**Two is not the total.** Both surfaced during a `Get-RemoteMailbox -ResultSize 1`
health probe that stops at the first match — a sample, not a survey. A full
enumeration will find more, and broken *users* remain entirely unsurveyed.
Worth a validator sweep before each phase.

---

## 7. Phasing

| Phase | Content | Duration |
|---|---|---|
| Pilot | A few low-risk HQ groups, one-offs | 1 weekend |
| 1 | All HQ (298) | 1 weekend |
| 2 | All franchise (4,781), batched by office | 4–5 weekends |
| Close | Verification, reconciliation | 1 week |

HQ is "pilot and big bang" — the population is small and internal.

Franchise blast radius is different in kind, not just size: 800 offices who do
not know you exist and will call someone else.

---

## 8. Scripts

Live in `Projects\DLMigration\`. Buttons in the **Reporting** section.

| Script | Status | Does |
|---|---|---|
| `Inventory-AllGroups.ps1` | **built** | Every mail-enabled group, 4 fields, CSV + grid |
| Capture script | *next* | Per-group: addresses, `legacyExchangeDN`, membership, office code |
| Cutover script | not started | Mail-disable, create cloud, stamp X500 |
| Verify script | not started | Reconcile a batch against its capture file |

### `Inventory-AllGroups.ps1`

Name, PrimarySmtpAddress, RecipientTypeDetails, OU. Deliberately no member
counts, owners or nesting — those cost a directory call per group and this is
meant to be quick.

`OU` is a **label, not a path**: `Fran` or `HQ`, from whether the container
path contains `Franchise`. Falls back to `DistinguishedName` when
`OrganizationalUnit` is empty, because the latter is **undocumented as an
output property** — Microsoft documents it only as an input parameter. Both
forms contain the franchise container name, so the substring test works either
way. Unreadable path defaults to `HQ`, so a directory fault inflates a visible
count rather than dropping rows silently.

Sets forest-wide scope explicitly every run. Exchange's default recipient scope
is the Exchange domain, so without this the row count depends on what ran
earlier in the session — unreconcilable.

### Capture script — requirements

Must bank, **before anything is touched**:

- `LegacyExchangeDN`
- Every entry in `EmailAddresses` — all proxies, and any X500s already present
- Full membership
- **Office code** (`3467-F41`) — the batch key, available in the full OU path
  that the inventory collapses to Fran/HQ

Keep the full OU path here. Cheap to keep, annoying to re-derive across 4,781
groups later.

---

## 9. Superseded work

`obsolete\DLMigration-v1_2026-08-13\` — scoped to one directory and enriched
every row (member counts, nested-group detection, resolved owners). Replaced
same day, never run against real data.

**Why:** Exchange searches the whole forest to answer `Get-DistributionGroup`
regardless, so filtering afterwards saved nothing at query time and cost an AD
read per group. If the expensive part happens anyway, pull everything once and
slice it afterwards.

Worth recovering if enrichment is wanted later: `Get-GroupDnLookup` (chunked
LDAP nesting detection, 50 DNs per filter, routed per directory) and
`Get-OwnerText` (ManagedBy resolved and cached).

`obsolete\DLCloudMigration-v1_2026-08-13\` — three earlier scripts, never run.
`2-Export-Usage.ps1` counted messages actually *delivered* to each group rather
than assuming a group is in use. That instinct is right for retire-vs-migrate
if that ever comes into scope.

`Projects\DLtoSharedMbx\` is a **different project** — deletes franchise DLs
on-prem and creates cloud shared mailboxes. Not this.

---

## 10. Open questions

1. **Does nesting the on-prem security group inside the cloud DL work?**
   Probably not — DL members must be mail-enabled recipients. Decides whether
   the split means permanent dual membership maintenance. *Verify.*
2. **Decommission target date?** Determines whether franchise runs comfortably
   or compressed.
3. **Franchise comms owner** for Phase 2.
4. **Does `Set-Group -Universal` re-stamp `msExchRecipientTypeDetails`?**
   Undocumented. Check both `Get-Group` and `Get-Recipient` after the two HQ
   conversions.
5. **Validator sweep** across users and groups — how many more malformed
   objects are there?

---

## 11. Decisions log

| Date | Decision |
|---|---|
| 2026-08-13 | Phase by directory then group type; HQ first |
| 2026-08-13 | Inventory unscoped rather than filtered — the expensive part happens regardless |
| 2026-08-13 | `OU` column is a Fran/HQ label, not a path |
| 2026-08-14 | Split mail-enabled security groups; every security group stays on-prem |
| 2026-08-14 | No permissions audit — asymmetry of risk decides it |
| 2026-08-14 | Drift between split objects accepted; Glenn owns Exchange objects only |
| 2026-08-14 | Batch franchise by office, not alphabetically |
| 2026-08-14 | Driver confirmed: Exchange decommission + modernisation |
