# DLtoSharedMbx

`Projects\DLtoSharedMbx\` — 1-ONPREM-Export, 2-ONPREM-Decomm, 3-CLOUD-CheckSync,
4-CLOUD-create.

Turns an on-premises distribution list into a cloud shared mailbox plus an
access group whose members are the old DL's members. Payroll lists are renamed
and kept; everything else is deleted.

---

# OPEN — the residual cloud object. Tested 2026-08-07, mitigation FAILED.

**Do not treat the hide step in `1-ONPREM-Export.ps1` as working. It does not.**

## What was tried

Mail-disabling a synced DL does not remove its cloud copy. The AD object keeps
syncing, and Exchange Online falls back to the `@...onmicrosoft.com` address it
appended at first sync — so the list stays in the address book as a recipient
that accepts mail and silently drops it. That is
[KB 4096314](https://learn.microsoft.com/en-us/troubleshoot/entra/entra-id/user-prov-sync/sg-mail-enabled-sync),
not a fault in Step 2.

The theory was: hide it on-premises first, let a sync carry the hidden flag up,
*then* mail-disable. Step 1 was changed to set
`HiddenFromAddressListsEnabled` on every payroll row, and to record before/after
state in `payrollhide_<date>.csv` precisely so this question could be answered
from data.

## What happened

Tested against `Payroll_E64` and `Payroll_I69`.

1. Step 1 hid both. Confirmed in the cloud after sync — **the hide itself works.**
2. Step 2 mail-disabled them.
3. After the next sync, **both were visible in the GAL again.**

## Why — and it is worse than the attribute being cleared

`msExchHideFromAddressLists` **does not exist on the object any more.**
Mail-disabling strips the entire Exchange attribute set, so there is no
attribute left to hold the flag. On-premises these are now plain security
groups with no Exchange attributes at all.

The cloud copy never followed. It is still stuck mail-enabled, holding the
onmicrosoft address. The two directories disagree and the cloud one wins the
address book.

So hiding before the mail-disable can never work, by construction. It was not a
timing problem or a sync problem — the flag is destroyed by the very operation
it was meant to survive.

**This was flagged as unverified when it was built.** The `Set-DistributionGroup`
documentation says the cmdlet removes "the email attributes required by
Exchange" without enumerating them, and `msExchHideFromAddressLists` is an
Exchange attribute. The code comment in Step 1 named that risk and said the CSV
existed to settle it. It settled it: the risk was real.

## What to do next, in order

**1. Ten-minute test, subjects already exist.**
`Set-ADObject -Replace @{msExchHideFromAddressLists=$true}` on one of the two
stripped groups. The attribute is writable on a plain security group even with
no other Exchange attributes present. Unknown: whether Entra syncs it, and
whether Exchange Online honours it on an object it still considers mail-enabled.
If it works this becomes a one-line addition to Step 2 **after** the
mail-disable, and the whole problem is automated away.

**2. If that fails, the documented remedy.** Delete the cloud object in the
Entra portal, then re-sync. The on-premises group is untouched and returns as a
plain security group, absent from the GAL. KB 4096314 lists only two
workarounds and the other one — moving to a non-synced OU — is unavailable here.

**3. Either way, remove the hide from Step 1.** As it stands it hides a
still-functioning DL from the address book for the thirty minutes before it is
disabled, which helps nobody and could send someone looking for a live list.
Worse, its help block claims a benefit it does not deliver. A tool that
describes a mitigation it does not provide is the failure this project spends
most of its time hunting.

Keep the `payrollhide_*.csv` output if the Set-ADObject route wins — before and
after state is worth having either way.

---

# OPEN — the destroy/keep decision is invisible

`NewName` blank means **the group is deleted at Step 2**. That is the most
consequential rule in the project, and it is encoded in whether a spreadsheet
cell is empty.

```
name,newname
Invoice.UnionMO,          <- this row will be DELETED
Jobs.UnionMO,             <- so will this one
Payroll_F26,Payroll.ColoradoSpringsCO
```

Nothing in the file says so. The column reads as a rename field, because that is
what it is called. A trailing comma is doing the work of a `Delete` flag.

The rule is documented in five places across four scripts — all of them in
`.ps1` headers. The person most likely to get it wrong is editing the CSV, and
the CSV is silent.

**This is the most likely reason a second engineer will not touch this project.**
Jeff is at least as senior as Glenn and still avoids these scripts. That is
evidence about the tool, not about Jeff: Step 2 is irreversible, and whether a
given group survives depends on an unlabelled blank cell. Declining to operate
an unmarked destructive switch is correct behaviour.

Cheap fixes, none requiring a rewrite:

- An explicit `Action` column (`Keep` / `Delete`) — redundant with `NewName`, but
  visible where the decision is actually made.
- A `README.md` in the project folder stating the end state. Reading all four
  scripts still does not tell you that a DL becomes a shared mailbox plus an
  access group; that only emerges from Step 4's internals.
- Split Step 2's two behaviours. One script called "Decomm" that either keeps or
  destroys depending on a blank cell is one name for two operations.

---

# OPEN — the bounce window is larger than it needs to be

Between Step 2 stripping the address and Step 4 re-adding it, mail to these
lists bounces. Listservs unsubscribe addresses that bounce; **one group has
already been lost that way.**

Most of Step 4 does not need to be in that window. For a payroll row it builds:

- `$MailboxName = $NewName` — a name that does not exist yet
- `$Primary = "$NewName@ExpressPros.com"` — an address that does not exist yet

Neither collides with the live DL. Nor do the access group, the FullAccess and
SendAs grants, the ManagedBy resolution, or the member additions. **Only the
alias loop collides**, because that is where the old addresses are re-added.

So for payroll rows the order could be:

1. Export
2. Create mailbox, access group, permissions, members — new address only, mail
   still flowing
3. Mail-disable on-premises
4. After sync, add the old addresses as aliases — one call per group

That reduces exposure to sync latency plus one `Set-Mailbox`.

The stronger argument is not the shorter clock but *what leaves the window*:
member and manager resolution and the permission grants are where Step 4
actually fails — three fallback lookups per member, `SilentlyContinue` on the
adds. Today a failure there happens while the address is dead. Reordered, it
happens while mail still works.

**Only applies to payroll rows.** Non-payroll rows set
`$MailboxName = $OriginalName` and reuse the old primary address, so they
genuinely must follow the deletion.

---

# OPEN — no test harness

The destructive path can only be tested by destroying something. A test that can
be run once is a rehearsal; one that can be run twenty times is a harness — and
it is the prerequisite for evaluating any rewrite against current behaviour.

Wanted:

- `New-TestDLs.ps1` — throwaway DLs in a nominated OU, prefixed `ZZTest-`, with
  members, owners and extra aliases so member/manager resolution and the alias
  loop are actually exercised. Some rows with a NewName, some without, to cover
  both paths.
- `Remove-TestDLs.ps1` — teardown that refuses to touch anything without the
  prefix, so it cannot be aimed at production.
- A generated `DL-Input.csv` to match.

It cannot simulate the failure that started this: a listserv unsubscribing on
bounce. It tests mechanics and timing, not consequences.

---

## Settled — what is known to work

- **Step 3's address check.** Object-gone is the right test for rows being
  deleted. For payroll rows the object is kept on purpose and never disappears,
  so that test answered "not ready" for ever and had to be ignored. It now asks
  whether the ADDRESSES Step 4 needs are free, mirroring Step 4's
  `Normalize-Domain` exactly. Verified 2026-08-07: Step 4 succeeded, aliases
  attached.
- **Guid binding must be a string.** `-Identity $dg.Guid` passes a
  `[System.Guid]` object and `DistributionGroupIdParameter` refuses to convert
  from that type. `-Identity "$($dg.Guid)"` works. Bind on Guid rather than a
  round-tripped name (§14g) — but hand it over as text.
- **The write happens only after a clean export comparison.** The two exports
  exist to prove the directory is not shifting mid-run; writing before that
  check would make the script the thing that changes it.
