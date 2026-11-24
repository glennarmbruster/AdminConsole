# Versioning and integrity

`Version.json` · `Publish-Toolkit.ps1` · the version region in `Common.ps1` ·
the console footer · the integrity section in `Test-Toolkit.ps1`

Written 2026-08-07, in answer to: *"I'm thinking a version 1.x at the bottom of
admin console — but how do I make sure that every script aligns with the master
copy I have here?"*

---

## Those are two questions, and only one of them is about a version number

**"What version are you on?"** wants a label. A number at the bottom of the
console answers it.

**"Does your copy match mine?"** does not. And it is the one that actually gets
asked, usually in this form: *"it's not doing that for me."*

A version string cannot answer the second question, because **a version string
is a claim.** Someone edits a script and forgets to bump it. Now the toolkit
says 1.4 and is not 1.4, and every conversation that starts from that number
starts from a false premise. The failure is silent, it is the *normal* human
outcome, and it gets worse the more people touch the code.

A hash is a fact. It is computed from the bytes, so it cannot be wrong by
omission — and forgetting to regenerate it makes the check **fail**, not pass.
That is the right direction to fail in. A check that goes quiet when neglected
is worse than no check, because it also carries authority.

So `Version.json` holds both: one version number, and a SHA-256 for every
script.

Verified before shipping, not assumed: a manifest built over the 70 scripts,
then **one extra space** added to a comment in `HeaderAnalyzer.ps1`. The check
named that file. Restored, it went quiet again. A version string would have
said 1.4 throughout.

### Verify the other two paths with something you were discarding anyway

That test covered *modified*. `missing` and `extra` were still only believed.

When `New-VersionManifest.ps1` was retired, the obvious advice was to delete it
before the first stamp so it never entered the manifest. Glenn's call was
better: stamp **with** it, then delete it, and watch the console go amber and
the Tester list it under MISSING. Then re-stamp. A file already condemned is a
free test subject, and it exercises the one path nobody had actually run.

Generalises. Anything being removed on purpose is an opportunity to confirm the
detector still detects. The cost is one extra stamp.

The one risk is forgetting the re-stamp and shipping a permanent phantom
`1 missing` — which is precisely the "train everyone to ignore the line"
failure this design exists to avoid. The amber footer nags until it is fixed,
which is the mechanism working as intended.

---

## What is hashed, and the one exclusion that matters

Covered: every `.ps1` at the toolkit root and everything under `Projects\`.

**`Config\` is deliberately excluded.** `Tools.json`, `Settings.json`,
`QuickLinks.json` and the rest are *meant* to differ between engineers — Jeff's
sections are not mine. Hashing them would report drift on every machine, every
launch. Within about a week nobody reads the line, and then it is not a check,
it is decoration. The value of an integrity check is entirely in how boring it
is when things are fine.

Also excluded: reports, logs, exports, and `Version.json` itself.

`Publish-Toolkit.ps1` appears in the manifest it generates. That is
correct — it is a script that ships, and an edited copy of the stamping tool is
worth knowing about.

---

## Version.json lives at the toolkit root, not in Config\

Deliberate, and it looks inconsistent with everything else.

Every other JSON is per-engineer configuration and belongs in `Config\`.
`Version.json` is the opposite: it is a statement about the *shipped* code, made
by me, that nobody downstream should edit. Putting it in `Config\` would file it
with the things engineers overwrite with their own — and it would eventually be
overwritten, quietly, by exactly the kind of copy that most needs checking.

Root also means the exclusion rule is one sentence and stays true: *config is
not hashed; nothing at root is config except this.*

---

## The workflow

On the **master copy only**, after changes, before shipping:

```powershell
.\Publish-Toolkit.ps1
```

It shows what changed, then offers the next number:

```
  Current: v1.03

    [1]  minor  ->  v1.04
    [2]  major  ->  v2.0
    [3]  cancel

  Which? [1]:
```

`-Version 1.07` overrides, for correcting a bad stamp or jumping deliberately.

### Why the number is offered rather than typed

Glenn's idea, and it removes a class of problem instead of guarding against
one: a number never typed cannot be mistyped or reversed. The backwards check
below still runs, but it is now a net under the `-Version` path rather than the
main defence.

Two things it does **not** do. It cannot stop a wrong keypress at the menu —
but `2` gives `2.0`, which is loud and obvious, where a typed `1.02` over
`1.03` is silent and wrong. And "major" carries no definition here: asked what
should qualify, Glenn's answer was his own judgment, so the prompt states the
two numbers and does not lecture. The change list sits directly above it, which
is the part that actually informs the call.

### Computing the next number without destroying the padding

Not done with `[version]` arithmetic. `[version]'1.03'` is **1.3**, so
incrementing that way yields 1.4 and silently ends the 1.0x style.

Instead the previous string is split, the minor field incremented as an
integer, and re-padded to the width it already had — so whatever convention is
in use survives without the code knowing what it is. Padding is applied with
`PadLeft` only, never truncation: at `1.99 -> 1.100` the number outgrows the
padding, and the number matters more than the alignment.

Walked forward and checked as `[version]` at each step:

```
1.03 -> 1.04 -> ... -> 1.09 -> 1.10 -> 1.11   strictly increasing
1.11 -> 2.0 (major) -> 2.1 -> 2.2             strictly increasing
```

**A major resets the padding**, as specified: 1.03 becomes 2.0, and minors from
there run 2.1, 2.2 rather than 2.01, 2.02. Ordering is unaffected. If the
padded style is wanted throughout, the major branch should produce `2.00`
instead — a one-line change.

Running it on someone else's copy re-stamps **their** edits as official, which
is exactly backwards. The help block says so in those words.

**That is the only command, and the only thing that writes `Version.json`.**

There was briefly a separate `New-VersionManifest.ps1` holding the stamping
half, on the theory that "stamp" and "prepare to ship" are two ideas worth
keeping apart. Glenn's response was *"why do I have both?"*, and he was right.
The separation had no use case — unblocking is idempotent, so stamping without
it was never a thing anyone would want — and it was actively harmful: a second
entry point that skipped the confirmation prompt and the reused-version
warning. **Two doors, one of which bypasses the guardrails, is worse than one
door.**

The genuinely reusable logic was never in that file anyway. It is in
`Common.ps1` — `Get-ToolkitCodeFile`, `Test-ToolkitIntegrity`,
`Get-FileHashHex`. `New-VersionManifest.ps1` was just a second way to call the
same functions. Worth remembering the next time a "separation of concerns"
argument comes up for a hundred-line script: check whether the shared part has
already been extracted somewhere else, because if it has, the split is
decoration.

---

## Why publishing and unblocking are the same command

Glenn already ran `Unblock-File` over the folder every time he shipped —
downloaded scripts carry a mark-of-the-web tag and PowerShell refuses to run
them. Attaching the stamp to a habit that already exists is far more reliable
than adding a step that has to be remembered on its own. That was his idea, and
it is the right one.

**Unblock runs before hashing.** `Unblock-File` should only strip the
`Zone.Identifier` alternate data stream, leaving the primary stream
`Get-FileHash` reads untouched — so in principle the order does not matter. But
that belief was not tested, and if it is wrong in the other order the manifest
would describe the files in their *pre-ship* state. Unblock-then-hash makes the
question moot: whatever clearing the tag does, the manifest describes the bytes
actually handed over. Ordering chosen so the assumption cannot cost anything.

**The blocked list is printed, not swallowed.** A blocked file is a file that
came from off the machine, so that list should match what was just downloaded.
Anything unexpected on it is worth looking at before shipping. A silent
`Unblock-File *` would have thrown away the only interesting thing it knows.

**Unblock failure stops the run.** A manifest is a statement that this copy is
ready to hand over; a file PowerShell will refuse to execute is not ready to
hand over, so nothing gets stamped.

---

## The prompt is the feature

`Publish-Toolkit` prints what changed since the last stamp and then stops for a
`y/N`. `-Force` skips it, and is deliberately not the default.

The change list reuses `Test-ToolkitIntegrity` rather than reimplementing it —
the same three lists, relabelled for where they are being read. On a copy they
mean drift; on the master they mean:

| on a copy | on the master |
|---|---|
| Modified | **changed** since that stamp |
| Extra | **new** since that stamp |
| Missing | **removed** since that stamp |

One function, two audiences. If these were ever written twice they would drift,
and the master would stop agreeing with the copies about what counts as a file.

---

## A shipped version number is spent

Stamp 1.4, ship it, fix something, re-stamp 1.4 — and two different toolkits
now both answer "1.4". The hashes still catch the difference, but they report
it as drift with no way to tell which 1.4 is the real one. And the number is
the thing said out loud on a call.

`Publish-Toolkit` warns when the version being stamped matches the previous one
*and* files have changed. It warns rather than blocks, because re-stamping the
same number is legitimate while iterating on the master before anything ships.

### Going backwards stops the run

Re-using a number at least announces itself. Stamping 1.02 over 1.03 is worse:
a stale number that looks completely legitimate, on a toolkit reporting an
older version than code someone already has. It is almost always a typo, so it
**blocks** rather than warns. `-Force` overrides.

### Compare as [version], never as strings

`"1.10" -lt "1.9"` is **true** when compared as text — the tenth release would
be judged older than the ninth, and the guard would fire on a perfectly correct
bump. Checked rather than assumed, over the real numbering:

| previous | new | `[version]` | as strings |
|---|---|---|---|
| 1.02 | 1.03 | forward | forward |
| 1.03 | 1.02 | **backwards** | backwards |
| 1.9 | 1.10 | forward | **backwards** — wrong |
| 1.2 | 1.10 | forward | **backwards** — wrong |
| 1.01 | 1.1 | **same** | forward — wrong |

Three of five disagree once double digits appear. Since this ships as 1.0x, the
divergence would not have surfaced until 1.10 — long after the check had been
trusted.

### Leading zeros are cosmetic

`[System.Version]` parses each part with `Int32.Parse`, so `1.01` and `1.1` are
**the same version**. The padding is display only. `Version.json` stores the
string as typed, so the console still shows `v1.03`, but any comparison
collapses it.

Consequence for the 1.0x scheme: it works fine, and `1.09 -> 1.10` is a correct
forward step. The only real hazard is switching padding style mid-stream —
stamping `1.1` after `1.01` reads as a new release and is not one. That case
warns explicitly rather than passing silently.

The generator **refuses to write a partial manifest.** If any file fails to
hash, nothing is written. A file quietly absent from the list is a file the
check will never complain about — a hole shaped like the thing you were trying
to detect.

It then **verifies the manifest it just wrote.** If the generator and the
checker disagree about which files are covered, that is a bug in my code, and I
would rather find it here than have it surface as phantom drift on Jeff's
machine, where it looks like his problem.

---

## What the console shows

At the bottom, one line, three states:

| | |
|---|---|
| **green** | `v1.4 — 70 files verified — released 2026-08-07` |
| **amber** | `v1.4 — 2 modified, 0 missing, 1 extra — run Tester for the list` |
| **grey** | `unversioned copy (no Version.json)` |

Amber names the count; `Test-Toolkit` names the files, under MODIFIED /
MISSING / EXTRA. The headline and the detail live in different places on
purpose — the console is glanced at, the Tester is read.

Grey is the state a working copy is in before it has been stamped, which is not
an error and does not warrant a warning colour.

---

## Two bugs found while verifying this, both of the same shape

**1. The footer would have said "version unavailable" on every launch.**

`$toolTip.SetToolTip($lblVersion, ...)` sat inside the `try`, eight lines above
the line that creates `$toolTip`. Calling a method on `$null` throws; the
`catch` caught it and wrote *"version unavailable"* — a message about the
version having failed, when what actually failed was a tooltip.

The bug is not the ordering, which is trivial. It is that **a `catch` reported
the wrong cause with total confidence**, and would have sent me looking at the
hashing code. The SetToolTip call is now outside the `try`: a tooltip that fails
to attach must not be able to overwrite a version line that was worked out
correctly.

Neither the parser nor a brace check can see this. What found it was reading the
code in the order it executes. (§14j, again.)

**2. `Stale` was a false alarm generator.**

I had the console go amber when any `.ps1` was newer than the manifest. But a
file that is newer *and hashes the same* is byte-for-byte identical — touched,
not changed. Any copy method that resets timestamps would have turned a perfect
copy amber.

Which is the exact failure I had written a warning about, two hundred lines
higher in the same file, as the reason for excluding `Config\`.

`Stale` now never fails a clean check. It only annotates a failing one, where it
is genuinely useful: *"the manifest is older than the code — re-stamp the
master."* That is an explanation, not a verdict.

Worth keeping in mind for anything added here later: **the hashes are the
authority.** Any signal that can contradict them while they say "clean" is
noise, however reasonable it sounded when it was added.

---

## What this does not do

It detects difference, not tampering. Anyone who can edit a script can also run
`Publish-Toolkit.ps1`. This answers *"is your copy current?"*, which is the
question that actually comes up, and it answers it exactly.

It also does not tell you *which direction* — a modified file might be an
engineer's local fix that ought to come back to the master. Amber means "these
differ", not "these are wrong". Worth asking rather than assuming, since one of
those files may be a fix I want.
