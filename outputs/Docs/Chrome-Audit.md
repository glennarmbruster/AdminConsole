# Chrome consistency audit

69 forms scanned. Read-only — nothing was changed.

Re-run any time: `python3 probes/audit_chrome.py AdminConsole`

---

## What I threw out before showing you

The first pass produced 19 overlap findings and 13 edge findings. Most were the
scanner being wrong, so they were fixed rather than reported:

- **Five "overlapping buttons" were a deliberate swap-in-place.** AuditSMB,
  AuditUser, MailboxAudit, SharedMailboxAudit and DLManager each stack a Cancel
  button exactly over the Run button and toggle `Visible`. That's correct
  WinForms. The scanner now skips any control built with `Visible = $false`.
- **Eleven "overlaps" in ToolManager were one variable, two dialogs.**
  `$dlg` is declared twice, 670 lines apart, for unrelated windows. Treating
  them as one parent made every control of the first collide with every control
  of the second. Containers are now bucketed by which declaration precedes them.
- **Edge findings under ~12px are now "verify", not "fail".** `Form.Size`
  includes the border and title bar, and how much depends on border style,
  theme and DPI. 16x39 is right for a sizable form at 100% but it is an
  estimate, and asserting a 5px clip as fact is how a check earns a reputation
  for crying wolf.

---

## Worth fixing before the demo

### 1. InboxRules — two buttons are clipped (high confidence)

`Refresh` and `Close` sit at y=428 and are 28px tall, so they end at 456. The
form is 480 tall, giving roughly 441 of client area. **About 15px of each
28px button is below the visible area** — you'd see the top half.

For comparison, MailboxHealth puts its Close button 13px clear of the bottom.
That's the intended spacing.

Fix: form height 480 → 510. One line.

### 2. MailboxOneDriveAccess uses a different font

68 forms use `Microsoft Sans Serif, 9`. This one uses **`Segoe UI, 9`**. It's
the single most noticeable difference in the set, because every label, button
and grid header renders differently — the "what's different about that screen?"
question, in font form.

Fix: one line.

### 3. Three tools have no Close button

`Dashboard`, `FranUser_Office`, `Test-Toolkit` — title-bar X only. Every other
tool has an explicit Close. FranUser_Office is one I wrote and missed.

Fix: one button each, bottom-right, matching the others.

---

## Worth doing, but not before the demo

### 4. Five tools have no "?" button

`DLManager`, `MailboxAudit`, `PowerShellConsole`, `SharedMailboxAudit`,
`UnifiedGroupMembership`.

**These are not one-line fixes.** None of the five has a `TOOLKIT-HELP` block,
so adding the button alone gives an empty help window — worse than no button,
because it looks broken rather than absent. Each needs its help text written
first.

DLManager is the one that matters most: it's a substantial tool, and the "?"
is the strongest thing in your demo. But writing five help blocks well is an
hour or two, and rushing them produces the thin filler text that undermines the
claim the button is making.

Suggest: write DLManager's before the demo if you have a quiet hour, leave the
other four until after.

---

## No action needed

### Size and ClientSize both set (4 files)

`AdminConsole`, `Dashboard`, `M365Reports`, `TenantManager` set both.
`ClientSize` comes later in all four, so it wins and the `Size` line is dead
code. Harmless — but it's exactly what hid the Dashboard Configure block from
you, so the dead lines are worth deleting eventually just so the next person
reading doesn't trust the wrong one.

### Near-edge controls (11)

5–10px past the estimated client edge across AD-GroupEditor, GroupEditor,
MailboxPermissions, DL_SendToAs, MailboxOneDriveAccess, RoomPlace. Within the
margin of error. Glance at the bottom edge of those six windows during the dry
run; if nothing looks cut off, they're fine.

### ToolManager: 1px overlap

`$lblSpoHint` and `$txtExs` overlap by 330x1px. Invisible.

### Quote style

40 forms declare the font with single quotes, 28 with double. No rendering
difference whatsoever. Mentioned only so it doesn't get re-reported later as a
finding.

---

## What this audit could NOT check

Eight files position most controls dynamically, so their geometry was not
verified: `ChangeMeetingOrganizer` (0 of 11 controls checkable), `TenantManager`
(1 of 15), `M365Reports` (3 of 7), `ToolStatus` (4 of 9), `Dashboard` (6 of 17),
`ConfigBackup` (7 of 12), `MailboxOneDriveAccess` (10 of 21),
`ManageMailObject` (14 of 25).

Dynamic layout is usually the better approach — it's what makes the console's
button grid adapt — but it means these windows can only be checked by opening
them. Worth including all eight in the dry run for that reason.
