# SetDLMembership.ps1

Bulk distribution list membership: every user in the user block, against every
DL in the DL block, added or removed together.

Built 2026-08-03 for the DL cloud migration -- a three-month project where
membership gets moved in batches rather than one person at a time.

**AdminConsole only, on-premises Exchange only.** Declares
`# TOOLKIT-REQUIRES: OwnProcess`, so the console runs it in its own process
whether or not **Own window** is ticked (§14e, §14h). This is not a preference --
sharing a process with Exchange Online silently routes its writes to the cloud.

---

## The shape

Two paste boxes and a pair of radio buttons. The operation is the **cross
product**: N users × M DLs. Ten users across four lists is forty operations,
and the preview says so before anything runs, because the count is the only
thing on screen that conveys how big the action actually is.

`Check` resolves everything and shows what will happen. `Run` re-checks if the
input changed since, so the preview can never describe a different batch from
the one that executes.

## Everything is resolved before anything is written

A typo'd DL fails once **per user**. A typo'd user fails once **per DL**. Ten
users against four lists with one bad entry produces ten or four identical
errors that bury whatever else went wrong.

So both lists are resolved first, and unresolved entries are named individually
with the reason. The run then proceeds with what resolved, and the confirmation
says how many are being skipped -- rather than either failing the whole batch or
silently dropping them.

Same reasoning as the disable pipeline's preflight (§21): one clear message
beats thirty copies of the same misleading one.

## The outer loop is the user

A user is finished across every list before the next one starts. An
interruption therefore leaves earlier people fully done and later people
untouched -- a state describable in one sentence.

Looping DL-first would leave everyone half-done across a partial set of lists,
which is the state nobody can reason about afterwards. It also makes one audit
row per user honest, because by the time that row is written the user's work
really is complete.

## Logging

**One row per user**, not per operation. Ten users across four lists is ten
rows, not forty. `Target` is the user, `Detail` lists the DLs, and `Result` is
per user:

- `Success` -- every list took, or every list was already correct
- `Partial` -- some took, some were already in the desired state
- `Failed` -- at least one list errored

`Target` holds the person because that is what LogViewer and the monthly
reporting filter on. During the migration the DL is arguably the more
interesting subject, and if that turns out to be how the work is actually
reviewed, the alternative is one row per DL with the users in `Detail` -- a
small change, but it should follow real use rather than a guess.

## Accepting what people actually paste

**DLs** are passed through as typed. `Get-DistributionGroup -Identity` accepts
a display name, an alias or an address, and forcing one form would reject two
things that work.

**Users** are tried as typed first, then retried with the mail domain appended
if there was no `@`. Both forms turn up in tickets. `Get-TargetSmtp` does the
completion, so the rule lives in one place rather than being restated here.

Entries are de-duplicated case-insensitively, so the same list pasted twice is
not processed twice.

## Other decisions

**No cap on the batch size.** `ManageOwnerGroups.ps1` caps at ten because it
serves a ticket workflow with ten-ish people. This serves a migration, where
large batches are the point. Instead of a cap, anything at or above 100
operations gets an extra line in the confirmation -- Remove across a wide list
is not undoable from here.

**`-BypassSecurityGroupManagerCheck` on every call**, both directions, through
one splat. Same reasoning as `ManageOwnerGroups.ps1`: a managed group appearing
later must not make removals fail where adds succeeded.

**"Already a member" and "not a member" are Skipped, not Failed.** Exchange
returns them as errors, but they mean the desired state already holds. Matched
on English message text, which is fragile in principle and fine here.

**Neither radio is preselected.** Add and Remove are opposite and equally
plausible, so a default is a trap -- you confirm a dialog that agrees with your
intent while the radio says otherwise.

## No domain controller is pinned (a real flaw, but not the one that bit)

**First live run failed**, 2026-08-06. An HQ distribution list, one user,
Remove:

```
DC     : fran-vm-dc01.franchises.expresspersonnel.com
ERROR  : The operation on Identity "# E-Mail Team" failed because it's out of
         the current user's write scope
```

It reads like a permissions problem. It is not one.

`franchises.expresspersonnel.com` is a **child domain** of
`expresspersonnel.com`. Pinning the shell to a franchise controller with
`Set-ADServerSettings -PreferredServer` means writes go to a DC that is not
authoritative for HQ objects. **Reads still worked** -- the DL resolved by name
and appeared correctly in the preview -- so the failure surfaced only at the
write, wearing the wrong label.

The cause was a design error carried over from `ManageOwnerGroups.ps1`. When the
DC selector was hidden across the toolkit, the control was hidden but its
**franchise default was still applied**. That is correct for
`ManageOwnerGroups`, which only ever touches franchise owner groups. It is wrong
here, because this tool has to reach HQ and Franchise in the same batch.

**No `-PreferredServer` at all now.** Exchange binds to a controller that can
serve each object, which is what the console itself does.

`-ViewEntireForest $true` is **asserted** at run time rather than assumed.
`Connect-ExSE.ps1` sets it, but a session established another way may not have,
and without it a lookup silently sees only the local domain -- which reads as
"no such group" rather than as a scoping problem. Failure to set it is
non-fatal and reported, since some RBAC roles cannot.

A `write scope` error now appends what it usually means, so the next person does
not spend the hour this cost.

## What actually broke it: two Exchanges in one process

Three fixes preceded this one. The DC pin (real, §14f) and the Identity
round-trip (real, §14g) were both genuine defects and neither was the cause. The
symptom never moved, which should have been read as "wrong diagnosis" after the
first one.

The tool ran with **Own window unticked**, so it loaded into the console's
process -- where Exchange Online was already connected. Both Exchanges export
`Get-DistributionGroup` and `Remove-DistributionGroupMember`; last import wins.
The banner said on-premises and was telling the truth about what was *loaded*,
not about what was *answering*.

Exchange Online then refused to modify a synced group:

> ...out of the current user's write scope, because the object is being
> synchronized from your on-premises organization

The log box had `WordWrap = $false`. Everything after "write sco" was off the
right edge, and the entire diagnosis was in the part nobody could see.

**Three changes, in order of how much they matter:**

1. `# TOOLKIT-REQUIRES: OwnProcess` in the header. The console reads it and
   forces a fresh process regardless of the checkbox. The tool no longer depends
   on the operator remembering.
2. `Test-ExoInProcess` gates Check and Run. Both refuse outright -- a tool that
   can write to the wrong directory should not be one confirmation away from it.
3. `WordWrap = $true`. A log you have to scroll sideways to read is a log nobody
   reads sideways.

## The bind key, which was a real bug but not the one that bit

The DC-pin story below is real and the fix stands, but **it was not what broke
the first three runs.**

`Resolve-DLEntry` stored `$g.Identity` and passed it to
`Remove-DistributionGroupMember`. For this group that string is `# E-Mail Team`
-- the AD CN, kept from before the group was renamed to "M365 Team". A bare name
fed back into `-Identity` is a fresh forest-wide lookup, not a bind, so it can
resolve somewhere other than the object that was just found.

It failed with **"out of the current user's write scope"**, which sent the
diagnosis toward domain controllers and RBAC for two rounds. What settled it was
Glenn running the cmdlet by hand:

```powershell
Remove-DistributionGroupMember m365team@expresspros.com -Member glenn.armbruster@expresspros.com -BypassSecurityGroupManagerCheck
```

Identical in every respect except the identity string -- and it worked. The diff
between a working invocation and a failing one is worth more than any amount of
reasoning about the error text.

**Now binds on `Guid`**, falling back to `DistinguishedName` then
`PrimarySmtpAddress`. `Identity` is last resort only. The run log prints which
key was used, because when a write fails that is the first thing worth seeing.

The user side already bound on `PrimarySmtpAddress` and is left alone -- unique
org-wide, and it never once misfired.

## A PowerShell 5.1 trap hit while writing this

The input-change signature originally used `` `u{1} `` to join the two boxes
with a non-printing separator. **`` `u{...} `` is PowerShell 6+.** In 5.1 it is
not an escape at all -- the literal characters end up in the string, the
signature never matches, and `Run` would silently re-check every time.

Replaced with `[char]1`. Cheap to fix, invisible if it had shipped, and worth
recording because nothing about the syntax looks version-specific.

## Adding the button

ToolManager → **Add**:

| Field | Value |
|---|---|
| Name | `DL Membership` |
| Script | `.\SetDLMembership.ps1` |
| Section | `OnPrem Scripts` |
| Own window | ticked |
| ObjectTypes | empty -- it does not act on the console's looked-up target |

## Still open

- No undo. A Remove across a wide list is recoverable only by re-adding, and
  the audit log is what tells you who was in what.
- Membership is not read back afterwards. The per-operation result says what
  Exchange reported; it does not re-query the list to confirm. On a replicating
  directory an immediate re-read can disagree with itself, which would be worse
  than not looking.
