# Dashboard.ps1

The start-of-day page. Six panels: Environment, Toolkit Health, Today, Retry
Queue, Top Tools, Failure Hotspots.

Written 2026-08-07. Read-only — nothing on it changes a mailbox, a group or a
file, and the buttons open other tools rather than acting.

---

## It was not built when it was asked for

The request covered six sections. Three of them — Today, Top Tools, Failure
Hotspots — read the completed-task log, and at the time the log was **two files
with two schemas**.

Three tools still used a private v1 logger writing `Task,DateTime` to
`CompletedTasks.csv`, while everything else wrote the ten-column format to
`CompletedTasks-v2.csv`. `Mailbox-Group-Inspector.ps1` had **12 logging call
sites, more than any other tool in the toolkit**, and it was on the v1 side.

v1 has no Result column. So on that data:

- "Today: 0 failures" would have been a **false statement of health**
- Top Tools would have ranked the most-logged tool at zero
- Failure Hotspots could never have shown those tools at all

That is the same failure this project keeps meeting: a check that passes
because it cannot see the thing that is wrong. Building the page first would
have produced something confidently wrong rather than usefully incomplete.

So the logging was unified first (see `Versioning.notes.md`), then this.

---

## Unknown is a fourth count, not a rounding error

Rows converted from the old log have a blank Result, because it was never
recorded. Today shows **Success / Partial / Failed / Unknown**, and the four sum
to Actions.

The alternatives were both worse. Excluding them makes "Actions" disagree with
the sum of its parts, which invites the question every morning. Counting them as
Success puts an invented fact on screen — the exact reassuring-direction error
the whole design is trying to avoid.

The Unknown count shrinks to zero on its own as those rows age past the window.
Nothing has to be remembered or cleaned up.

Same reasoning for the tool name: converted rows have no Script, so they group
under `(not recorded)` and the panel footnote says how many. Dropping them would
quietly understate every total.

---

## It calls the console's functions, it does not repeat them

Environment uses `Get-CapabilityStateMap`, `Get-ConnectedTenantDomain`,
`Test-TenantDomainMismatch`, `Get-PimStatusLine` and `Get-PimStatusColour`.
Toolkit Health uses `Test-ToolkitIntegrity`.

Every one of those is what the console header and footer already call. **Two
places computing the same state will drift**, and today proved it: the console's
private copy of `Find-TargetEnvironment` had gone stale against Common's and was
still pinning a domain controller that had been removed after the write-scope
bug. A dashboard that disagrees with the header about whether you are connected
is worse than no dashboard, because now there are two answers.

`-SkipSpo` on the first pass, for the reason the console skips it: `Get-SPOTenant`
is a network round trip, and this page should open instantly. Refresh does the
full probe.

---

## Panels fail independently

Each `Update-` function has its own `try`. One unreadable CSV must not blank the
other five panels — that is the difference between a page that is partly useful
and a page that is empty. Anything unresolvable reads **Unknown** rather than
throwing.

`Get-DashLogRows` returns rows *and* a reason, so a missing or old-format log
produces a sentence naming the fix ("run Convert-LogHistory.ps1 once") instead
of a blank panel or a stack trace.

---

## Two traps avoided by having met them earlier the same day

**Tools launch in a NEW process.** This form runs inside `ShowDialog()`, and a
modal form disables every other window on its thread. A WinForms tool started
with `&` from here would draw and then refuse input — §14d, which cost days the
first time. `Start-Process` sidesteps it, and every path is quoted, because
`Start-Process -ArgumentList` does not quote for you and one space in a folder
name turns a path into two arguments.

**The ToolTip provider is created above its first use.** The console footer had
it the other way round earlier today: `SetToolTip` on a `$null` threw, the
`catch` fired, and the footer reported "version unavailable" — blaming the
version code for a tooltip failure.

**Layout was measured, not eyeballed.** `Add-ToolHelpButton` places the `?` at
`ClientSize-40`, which computed to x880 y8 — straight through the Toolkit Health
group box. Every panel is now below a 30px clear strip. Checked by computing the
rectangles, the same way the popup in `1-ONPREM-Export.ps1` was checked (§14l).

---

## What it deliberately does not do

No operator ranking, no productivity scoring, no charts. Failure Hotspots exists
to find scripts worth fixing, not people worth blaming — it groups by script and
has no operator column at all, which is a design choice rather than an omission.

Nothing destructive, and no retries are executed from here. The retry panel
reads `disable-retry.csv` and offers to open `Disable-User.ps1`; acting on a
queue is that tool's job.

---

## Known limits

**The retry panel knows about one queue.** `disable-retry.csv` is specific to
`Disable-User.ps1` — it is the only retry queue that exists. If a second tool
grows one, retry should become a Common concept rather than this panel learning
a second file path.

**Top Tools depends on the Script column**, which only the current log format
records. Until the converted rows age out of the 30-day window, the
`(not recorded)` bucket may be the largest one. That is honest but not useful,
and it resolves itself with time rather than with work.

**Connected is not the same as working.** `Get-CapabilityStateMap` says nothing
is obviously missing; a stale Exchange Online session reports as connected and
fails on the next real call. The dashboard inherits that limit and cannot fix
it.
