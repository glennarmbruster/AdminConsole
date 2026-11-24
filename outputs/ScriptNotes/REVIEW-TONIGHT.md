# For your review tonight -- 2026-09-18

Two new tools, both on the Mailbox Editor pattern, both wired into Unified
Lookup and the console. Everything below is a decision I made without being
able to ask you.

---

## What to try first

**Room Editor** -- open a room you know is behaving oddly. The Overview tab
gives a verdict before you read anything else; if it says nothing is wrong and
the room still misbehaves, that is worth telling me.

**Group Editor** -- open a synced distribution list and check the amber banner
says ON-PREMISES with a controller named. Then a Microsoft 365 group: Mail
Flow and Send As should grey out with a reason, not fail.

---

## Decisions I made on my own

**1. The Group Editor is `Group-Editor.ps1`, with a hyphen.**
`GroupEditor.ps1` already exists and is the weakest of the three group tools,
but retiring a live tool in the same change that replaces it breaks the
standing rule. So the new file is hyphenated, its console button is called
**Group Editor**, and the old one is now **Group Editor (old)**.

*If you want the tidier name*, rename `Group-Editor.ps1` to `GroupEditor.ps1`
when you retire the old three -- that is one rename plus the Tools.json
Script value plus the Unified Lookup hand-off rows.

**2. Nothing was retired.** `Group-Manager.ps1`, `Group-Manager-Editor.ps1`,
`GroupEditor.ps1`, `CalendarProcessing.ps1`, `Set-BookingPolicy.ps1`,
`RoomList.ps1`, `RoomPlace.ps1`, `RoomRegional.ps1` and
`CalendarPermissions.ps1` are all still installed and still on their buttons.
Same rule as the fourteen mailbox tools: a fortnight of use first.

**3. The Mailbox Editor no longer claims rooms.** Its Tools.json ObjectTypes
dropped RoomMailbox and EquipmentMailbox, and the Unified Lookup hand-off row
for rooms now goes to the Room Editor. A room opened from the console button
still works in either.

**4. Four things stayed out of the Room Editor**, because folding them in
would have made them worse: `Protocols.ps1` (tenant-wide report),
`ResourceDuplicator.ps1` (creates an object across two sessions),
`CalendarReset.ps1` (destructive bulk flow with a Preview gate), and room-list
*creation* (a tenant-wide object that belongs to no room).

**5. Five things stayed out of the Group Editor**: creation, deletion, the
bulk many-to-many membership tools, the tenant-wide reports, and
`DLtoSharedMB.ps1`.

---

## Things I could not verify, because I cannot run PowerShell

These are the ones to be suspicious of:

- **`Get-EXOWritableParameter` returns a hashtable, not a list.** I had this
  wrong in the first draft and it made the entire Booking tab and the entire
  Place tab silently read-only -- every row grey, every edit refused. Caught
  in review and fixed, but it is the shape of mistake that a run would find in
  five seconds and a read might not. **Please open Booking and Place on a real
  room and confirm you can actually edit a value.**
- **`Send As` versus `Send-As`.** The toolkit is inconsistent:
  `Group-Manager-Editor.ps1` uses `'Send-As'` (hyphen) and says the two are
  not interchangeable; `ManageMailObject.ps1:778` uses `'Send As'` (space).
  I followed the hyphen, being the newer and more carefully reasoned file.
  **One of those two files is wrong and I cannot tell you which.**
- **Set-Place read-back.** The Place tab writes one property, reads it back
  and compares. If that reports "did not stick" on a field you expect to work,
  tell me -- it may be my comparison rather than Exchange.
- **The room-list scan** asks every room list in the tenant for its members.
  On a tenant with hundreds that is slow; it warns above sixty. If it is
  unusable, that tab needs a different approach.

---

## Also changed while I was in there

- `Unified-Lookup.ps1` -- new hand-off rows for Room Editor and Group Editor,
  placed **above** the Mailbox Editor row, because the first match wins and a
  RoomMailbox contains the word "Mailbox".
- `Common.ps1` -- the tab layout and tick-box grid helpers moved in from
  MailboxEditor.ps1 so all three editors share one copy.
- `Config/Tools.json` -- Room Editor added under Mailbox and Access, Group
  Editor under Groups and DLs (NewProcess = true), old Group Editor renamed.
  **Config does not ship**, so you will need to add these two buttons on your
  own install with Tool Manager, or take the Tools.json from here by hand.

Thirteen defects were found by review and fixed before staging. The five worst
were: the hashtable one above; `ConvertTo-CmdletParameterValue` called with
parameters it does not declare, which defeated the boolean guard; Cancel on an
edit dialog being read as "set it to empty", which would have cleared the
setting; `-BypassSecurityGroupManagerCheck` sent to a cmdlet that does not
take it, failing every dynamic-group mail-flow Apply; and the Group Editor
reading its detail from the cloud after writing on-premises, so the tab
repainted with the old values.
