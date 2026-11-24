# ACCEPTANCE CHECKLIST — first 30 minutes on a Windows machine

*For the operator receiving the build, not for the agent that wrote it.*

The build was written on a platform that **cannot run it** — no WinForms, no
tenant, no AD, and a more permissive parser. Everything below is therefore
unverified until you run it, and the order is chosen so the cheapest checks that
can invalidate everything come first.

Run these in a **fresh PowerShell 5.1 window** with nothing connected.

---

## Gate 1 — does it load at all? (2 minutes)

The build environment's parser accepts syntax 5.1 rejects. This is the most
likely single failure and it is free to check.

```powershell
$fail = 0
Get-ChildItem *.ps1 | ForEach-Object {
    $t = $null; $e = $null
    [System.Management.Automation.Language.Parser]::ParseFile($_.FullName, [ref]$t, [ref]$e) | Out-Null
    if ($e) { $fail++; Write-Host "$($_.Name):" -ForegroundColor Red; $e | ForEach-Object { "   line $($_.Extent.StartLineNumber): $($_.Message)" } }
}
"`n$fail file(s) failed to parse."
```

**Any failure here stops the review.** Send it back with the output — do not
start fixing, because a 7-only construct in one file usually means the same
construct in ten.

Then confirm encoding, which is equally cheap and equally fatal:

```powershell
Get-ChildItem *.ps1 | ForEach-Object {
    $b = [IO.File]::ReadAllBytes($_.FullName)
    $bom = ($b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
    $dbl = ($bom -and $b[3] -eq 0xEF -and $b[4] -eq 0xBB -and $b[5] -eq 0xBF)
    if (-not $bom -or $dbl) { "$($_.Name): BOM=$bom DOUBLE=$dbl" }
}
```

A double BOM is the dangerous one — invisible, and it silently breaks
`#Requires`.

---

## Gate 2 — the verifier (3 minutes)

```powershell
.\Test-Toolkit.ps1
```

Expect a clean run. Then **prove the verifier can fail**, because a clean run
from a check that cannot return "no" is worth nothing:

```powershell
'$x = $a ?? $b' | Out-File .\_scratch.ps1 -Encoding UTF8
.\Test-Toolkit.ps1        # must report PS7002
Remove-Item .\_scratch.ps1
```

If `PS7002` does not fire, the whole delivery is unverified regardless of what
Gate 1 said — that rule is the only substitute for a 5.1 parser on the build
side.

---

## Gate 3 — cold start with no config (2 minutes)

Rename `Config\Tenant.json` aside and open the console.

**Expected:** it opens, reports every capability off, and says why. **Not
expected:** an exception, or a silent failure to start. A toolkit that will not
start cannot tell you what is wrong.

Put the file back.

---

## Gate 4 — the console (5 minutes)

Open `AdminConsole.ps1` with nothing connected.

- [ ] Opens **top-left**, does not overrun the screen, title bar reachable
- [ ] Open the console AND a tool together: the tool must be **bottom-left** and
      must not cover the connection indicators. This is invisible with one
      window open, which is why it is worth doing deliberately
- [ ] Every button visible; unavailable ones greyed **with a tooltip saying why**
- [ ] No two controls overlapping — look, do not just trust the linter, which is
      blind to controls that position themselves at run time
- [ ] Tenant label and connection indicators present and honest about being
      disconnected

Then connect and confirm the indicators change colour. **Red must stay red** —
if theming can disguise a "not connected" as decoration, that is a defect, not a
cosmetic issue.

---

## Gate 5 — one read tool end to end (5 minutes)

Mailbox Health, against a mailbox you know well.

- [ ] Every field populated, or explicitly says why not. A blank meaning "not
      applicable" and a blank meaning "the lookup failed" must not look alike
- [ ] Archive status agrees with reality — this is `ArchiveGuid` vs
      `ArchiveStatus` (C9), and the wrong one reports `None` on an archive that
      exists
- [ ] An audit row was written, with the **primary SMTP address** as Target
- [ ] The **`?` help button** is present, top-right, and shows real text rather
      than "no overview has been written"
- [ ] Action tools are the same width as the console; only viewers are wider
- [ ] Now do it against a **Microsoft 365 Group** and a **shared mailbox**. The
      group is the interesting case: `GroupMailbox` contains the substring
      "mailbox" and is not one (C12)

---

## Gate 6 — error handling (3 minutes)

Deliberately break things. This is the gate most likely to find something.

- [ ] Look up an address that does not exist → a legible message, not a CLR
      crash dialog with no line number
- [ ] Disconnect mid-session and click something → refused clearly
- [ ] Confirm the failure produced an audit row. **An error the operator
      dismissed is otherwise the one failure that leaves no trace**
- [ ] Check the output box wraps. A diagnosis off the right edge is a diagnosis
      nobody reads — the second clause of a message has hidden a real cause
      before

---

## Gate 7 — one write, on something disposable (5 minutes)

Use a test mailbox. Mailbox Permissions is the mildest write in the set.

- [ ] Confirmation names exactly what will change before it happens
- [ ] The change actually took effect — **read it back independently in a plain
      shell.** Several Exchange parameters are accepted and silently ignored
      (C10), so "no exception" is not evidence
- [ ] Audit row on success, with primary SMTP as target
- [ ] Cancel does nothing at all
- [ ] Now force a failure — a target with no permission to change — and confirm
      it is reported and logged rather than swallowed

---

## Gate 8 — read the handover (5 minutes)

- [ ] `UNVERIFIED.md` exists and is **not empty**. An empty one, after a build
      that could not be executed, means the assumptions were not noticed rather
      than that there were none
- [ ] `HANDOVER.md` states plainly that nothing was run against a tenant
- [ ] Spot-check two or three `# Learn <date>:` citations at call sites against
      the actual documentation — particularly any parameter the agent claims
      works in Exchange Online

---

## What "acceptable first pass" means

**Acceptable:** everything parses, the console opens, delivered tools work,
writes are correct and audited, and the gaps are written down.

**Not acceptable, send it back:** parse failures, stubs or `TODO`s, empty
`catch` blocks, a write that cannot be confirmed, an empty `UNVERIFIED.md`, or
any claim that something was "tested" when nothing was executed.

**Fine, and expected:** rough layout, missing polish, fewer tools than planned,
open questions in `UNVERIFIED.md`, and conventions that need a pass for
consistency. That is refinement. The line between the two lists is whether it
*runs*.

---

## After the first pass — accepting a stage

Gates 1–8 above are for the **first pass**: eight tools. Everything after that
arrives in the lettered stages of `09-FULL-TOOL-CATALOGUE.md`, and each stage
gets a shorter version of the same treatment.

**Every stage, without exception:**

- [ ] **Gate 1 and Gate 2 again** — parse, encoding, verifier. Two minutes, and
      still the cheapest thing that can invalidate everything
- [ ] **Count the buttons in each section that grew.** `Rows × ButtonsPerRow`
      must still cover them, a group counting as one. **Over-capacity does not
      error — the surplus button is simply missing**, and Stage A alone adds six
      to a section that starts with ten slots and eight tools
- [ ] Every new button opens something, and every tool's `?` shows real text
- [ ] One tool from the stage exercised end to end against a real object
- [ ] `UNVERIFIED.md` grew — a stage that added ten tools and no assumptions did
      not notice its assumptions

**Additionally, once per toolkit — the checks that only fail at scale:**

- [ ] **Console height.** With the finished `Sections.json`, the console fits on
      a 1920×1080 screen **with the taskbar showing**, and the output box is
      still visible. Buttons drawn below the bottom edge are invisible, not
      obviously broken
- [ ] **Section capacity.** Each section's button count ≤ `Rows × ButtonsPerRow`,
      **counting a group as one button**. Over-capacity does not error — the
      surplus button is simply absent
- [ ] **The group chooser** (Stage E). Click `Rooms`: five named entries, and a
      member whose capability is unavailable is greyed **in the chooser** with a
      reason — not silently enabled and then failing
- [ ] **Search still reaches a grouped tool.** Type a room tool's name in the
      console's find box; it should surface directly, without going through the
      group
- [ ] **A stale audit `Target`.** Spot-check rows from Toolkit-section tools:
      they log a scope name (`Toolkit`, `Config`, `SharePoint Online`), not a
      fabricated address. An invented SMTP address in the log is worse than a
      plain label
- [ ] **Stage F, before anything else in it runs:** confirm the `OwnProcess`
      tools actually launch in their own process — check the window, not the
      config checkbox. There are three in the finished toolkit: Group Manager
      Editor (first pass), Manage Mail Object and Bulk DL Membership (Stage F)

**The one that matters most, and only works once:** the first time Resource
Duplicator is used (Stage F), it should **stop cleanly** and tell you directory
sync has not published the mailbox yet. If it throws something obscure instead,
the two-stage flow was not built and the tool will fail on every run.
