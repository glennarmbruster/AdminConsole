# DEVIN BUILD PACK — M365 / Exchange hybrid admin toolkit

**THIS FILE IS TEN SEPARATE DOCUMENTS CONCATENATED.**

Treat each numbered section below as a standalone document with its own
authority and its own audience. They are combined into one file only because of
a transfer constraint, not because they are one argument. Cross-references
between them use the original filenames, which are given in each heading.

---

# ▶ STEP 0 — DO THIS BEFORE READING ANYTHING ELSE

You have been given **three files** in a directory:

| File | What it is |
|---|---|
| `DEVIN-PACK.md` | This file. Ten instruction documents. |
| `DEVIN-REFERENCE-CODE.md` | Sixteen working files: the shared layer, the verifier, the console, core tools and config. |
| `DEVIN-REFERENCE-CODE-2.md` | Eight more working tools, one for each remaining tool in the first-pass scope. |

> **Scope, in one line.** `DOCUMENT 1` §5 is the **first pass** — eight tools.
> `DOCUMENT 8` is the **whole toolkit** — 55 buttons over 65 files, in eight
> further stages. Build the first pass, stop for review, then work the stages.
> Where an early document says something is "out of scope", `DOCUMENT 8` §0 says
> whether that means *deferred* or *never*.

**Your first action is to extract the two reference-code files into real files
on disk.** Every instruction in this pack refers to those scripts by path, and
those paths only resolve after you have done this.

Extract them into a `reference/` folder at the repo root, preserving the exact
names in each section heading:

```
reference/
    Common.ps1              Test-Toolkit.ps1        AdminConsole.ps1
    ToolManager.ps1         Connect-ExO.ps1         Connect-ExSE.ps1
    MailboxHealth.ps1       MessageTrace.ps1        Unified-Lookup.ps1
    Group-Manager.ps1       Group-Manager-Editor.ps1  activate.ps1
    MailboxPermissions.ps1  CalendarPermissions.ps1 OOFEditor.ps1
    InboxRules.ps1          MailAttributes.ps1      Protocols.ps1
    ExternalForwarding.ps1
    Config/
        Tenant.TEMPLATE.json    Settings.json
        Sections.json           Tools.json
        Themes.TEMPLATE.json
```

Each section in those files is delimited by a heading of the form
`# ◆ FILE n OF m — \`Name.ps1\`` followed by the content in a fenced code
block. Extract the block, not the heading or the commentary above it.

**Confirm the extraction before proceeding:** report how many files you wrote
and their line counts. The counts are stated in each section heading, so a
mismatch means the extraction is wrong — and a subtly truncated `Common.ps1`
would poison everything built on top of it.

Then read `DOCUMENT 1` below and follow it.

**Do not write any toolkit code until Step 0 is confirmed and you have read
Documents 1 and 2 in full.**

---

<br>

==============================================================================

# ◆ DOCUMENT 1 OF 10 — `00-BUILD-PROMPT.md`

> The instruction. Role, environment reality, phases, gates, failure modes.
>
> *Standalone document. Referred to elsewhere as `00`.*

==============================================================================

# BUILD PROMPT — M365 / Exchange hybrid admin toolkit

*Paste as the opening instruction. Keep `01`–`09` and `reference/` in the repo
for the whole build.*

---

## 1. Your role, and what "done" means

You are building a **Windows PowerShell 5.1 WinForms admin toolkit** for a
Microsoft 365 tenant in **Exchange hybrid** configuration with **one
on-premises AD domain**. It is operated by IT administrators from a Windows
desktop, against production, by people who will trust what it tells them.

That last clause is the design constraint that matters. **A tool that reports
success on a change it did not make is worse than one that crashes.**

**The target for the first pass is a toolkit that runs and needs refinement —
not one that still needs fixing.** Concretely, on first delivery:

- the console opens on a Windows machine and every button either works or is
  greyed with a reason;
- each delivered tool opens, performs its read, and closes without an error
  dialog;
- writes are confirmed, audited, and correct;
- nothing is a stub, a placeholder, or a `TODO`.

Fewer tools that work beats more tools that nearly work. **If you are running
short, cut scope — never quality.** Say what you cut.

---

## 2. Two environments — do not confuse them

This is the single most important section in this document, because getting it
wrong produces code that passes all your checks and fails on the first machine
it reaches.

| | **Target** (where it runs) | **Yours** (where you build) |
|---|---|---|
| PowerShell | Windows PowerShell **5.1** | PowerShell **7.x** on Linux |
| UI | WinForms | **Not available at all** |
| Exchange | On-prem EMS + Exchange Online | **No tenant, no modules** |
| AD | A live domain | **None** |

### What follows from this

**You cannot run this toolkit. Do not pretend otherwise.** You cannot open a
form, connect to a tenant, or execute a single Exchange cmdlet. Any claim that a
tool "was tested" is false. Say "parsed and linted" — that is what you did.

**Your parser is more permissive than the target's.** PowerShell 7 accepts
syntax that Windows PowerShell 5.1 rejects outright at parse time. If you use
it, the file will lint clean for you and fail to load for the operator. This is
the highest-probability way for this build to fail, so it is a hard rule.

**Banned — 7-only syntax:**

| Construct | Example |
|---|---|
| Null-coalescing | `$a ?? $b`, `$a ??= $b` |
| Ternary | `$c ? $x : $y` |
| Pipeline chain operators | `cmd1 && cmd2`, `cmd1 \|\| cmd2` |
| Null-conditional member access | `${a}?.b` |
| `ForEach-Object -Parallel` | |
| `Sort-Object -Stable` | |
| `ConvertFrom-Json -AsHashtable` / `-Depth` | `-Depth` on **ConvertFrom** is 6.2+ |
| `Join-String`, `Test-Json`, `Get-Error` | |
| `Split-Path -LeafBase` | |
| `$PSStyle` | |
| `$IsWindows` / `$IsMacOS` / `$IsLinux` | `$null` in 5.1 — see constraint A5 |

Use `if`/`else` instead of ternaries, and `if ($null -eq $x) { }` instead of
`??`. It is more verbose and it loads.

**Two 5.1 behaviours that differ silently rather than failing:**

- `ConvertTo-Json` defaults to **`-Depth 2`**. Nested config serialises as
  `System.Object[]` with no error. Always pass an explicit `-Depth` (6 is
  enough for the schemas here).
- `ConvertFrom-Json` in 5.1 returns `PSCustomObject` and has no `-AsHashtable`.
  Write property access accordingly, and use a `Get-Prop`-style helper for
  optional properties rather than assuming they exist.

### What you CAN verify, and must

1. **Parse every file with the real parser**, via `pwsh`:
   ```powershell
   $t=$null; $e=$null
   [System.Management.Automation.Language.Parser]::ParseFile($p,[ref]$t,[ref]$e)
   ```
   This catches genuine syntax errors. It does **not** catch 7-only syntax, so
   the linter rule below is not optional.
2. **Run `reference/Test-Toolkit.ps1`**, extended with a rule that rejects every
   banned construct above. This is your substitute for a 5.1 parser and it is
   the only thing standing between you and an undeliverable build.
3. **Verify cmdlet surfaces against Microsoft Learn** — parameter names,
   accepted values, and whether a parameter exists in the *cloud* edition as
   opposed to the on-premises one. They differ, frequently, and the
   documentation is explicit about it.

Structure the verifier so its rule engine does **not** require WinForms. Parsing
a file that contains `Add-Type -AssemblyName System.Windows.Forms` is fine;
executing it is not. Never let the linter dot-source the scripts it checks.

---

## 3. Read before you write

`01-HARD-CONSTRAINTS.md` is authoritative and non-negotiable. **Every item in it
describes code that parses cleanly, runs, and produces a wrong result or a
misleading error.** They were each paid for once against a live tenant. You will
not rediscover them by testing — you have no tenant, and most of them do not
fail visibly even when you do.

Read it fully before writing code, and re-read section C before touching
anything that calls Exchange.

`08-COMMON-CONTRACT.md` is the other one to read early. It fixes the
function names every tool calls. Invent your own and nothing fits together —
and because tools are written one at a time, the mismatch surfaces late and
everywhere at once.

`09-FULL-TOOL-CATALOGUE.md` is the destination — every tool in the finished
toolkit, the final console layout, and the stages that get there. **Read its §0
now** so you know what §5 below is the first eighth of. Do not build from it
until the first pass is accepted.

---

## 4. The three rules that override everything

**1. Build the verifier first.** Before any tool, before the console, get
`reference/Test-Toolkit.ps1` running and extend it per `03-SELF-TEST-SPEC.md`
plus the 7-only-syntax rule. Then run it after every change and fix what it
reports before moving on. This is the gate between every step, not a phase at
the end.

**2. Never assume a cmdlet's surface. Verify it.** Parameter names, accepted
values, and whether a parameter exists at all vary by module version, by tenant,
and between the identically-named on-premises and cloud editions. Guessing a
parameter fails the whole call. When you cannot verify, **stop and record it**
(§7) rather than writing something plausible.

**3. When a fix does not move the symptom, the diagnosis is wrong — not
incomplete.** Do not refine a failed fix. Go back, reproduce, and diff the
working case against the failing one: arguments first, then environment. In the
source material this rule was learned three separate times, each more expensive
than the last, because a *real* defect near the traceback absorbed all the
attention.

---

## 4A. Verifying a cmdlet against Microsoft Learn

**Do this for every cmdlet and every parameter you have not personally
confirmed.** Cmdlet invention is the highest-probability failure mode in this
build, and you have no tenant to catch it.

The reference is `learn.microsoft.com/en-us/powershell/module/exchange/<cmdlet>`
(and `.../module/microsoft.online.sharepoint.powershell/...` for SharePoint).

### The procedure

**1. Confirm the cmdlet exists in the edition you are calling.** The page header
lists which products it applies to. On-premises Exchange and Exchange Online
export many identically-named cmdlets with different surfaces — see constraint
C1, which is the highest-consequence hazard in this codebase.

**2. Read the parameter's own description — not just the applicability list.**

This is the step that matters, because **the two contradict each other and the
body text is the one that is true.**

The documented example, which cost real production time:

> `Set-MailUser -PrimarySmtpAddress` lists **Exchange Online** among its
> applicable products. The body of the parameter description says
> *"This parameter is functional only in on-premises Exchange."*
>
> The parameter is accepted in the cloud. It does nothing. No error.

So: **when the applicability list and the parameter body disagree, believe the
body.** Look specifically for phrases like *"available only in on-premises
Exchange"*, *"functional only in on-premises Exchange"*, and *"this parameter is
reserved for internal Microsoft use"*.

**3. Check the parameter combination rules.** Learn documents mutual exclusions
explicitly, and they are easy to miss. `EmailAddresses` and
`PrimarySmtpAddress` cannot be used in the same command — stated on all four of
the relevant cmdlets, and the cause of constraint C11's X500 loss.

**4. Check whether it is a filter or a property.** `RecipientTypeDetails` is a
filter parameter and **not** a returnable property; requesting it in
`-Properties` fails the whole call (C13).

**5. Prefer the cmdlet's own declaration where you can reach it.** Documentation
describes a version; the installed module *is* one. A
`Get-CmdletParameterNames`-style helper that asks `Get-Command` what a cmdlet
declares, and passes nothing that is absent, is more reliable than any page —
and it is the pattern already used in the reference toolkit.

### Recording it

For every cmdlet you verify this way, put a one-line comment at the call site:

```powershell
# Learn 2026-08-22: Set-DistributionGroup accepts -PrimarySmtpAddress in EXO.
```

If you **cannot** verify something — the page is ambiguous, the behaviour is
undocumented, or the parameter is documented but its effect is not — do not
write plausible code. Add it to `UNVERIFIED.md` with what you assumed and what
would settle it, and say so in your progress report.

Two things in this toolkit's history were genuinely undocumented rather than
merely hard to find: **AutoMapping has no documented read path** in Exchange
Online, and `Add-MailboxPermission -GroupMailbox` exists but Microsoft documents
nothing about the resulting behaviour. Encountering a real gap is normal. Saying
so is the correct response; inventing the missing half is not.

---

## 5. Scope of the first pass

Build exactly this. Do not add tools beyond it without asking.

**Infrastructure (all required):**

- `Test-Toolkit.ps1` — verifier
- `Common.ps1` — shared layer (contract in `02`)
- `AdminConsole.ps1` — launcher
- `ToolManager.ps1` — edits the JSON config
- `Connect-ExO.ps1`, `Connect-ExSE.ps1` — connection scripts
- `activate.ps1` — PIM activation. **Belongs with the console, not among the
  tools**: this tenant's admin rights are *eligible, not standing*, so until an
  operator activates they have no admin access and every other tool fails in
  ways that look like faults. It also drives the countdown in the console
  header.
- `Config/` — `Tenant.json`, `Settings.json`, `Tools.json`, `Sections.json`,
  `Themes.json` (**take the shipped `Themes.TEMPLATE.json` as-is** -- one working
  palette is in scope; the ThemeManager editor is not)

**Tools, in this order** (specs in `05-TOOL-SPECS.md`):

1. Mailbox Health *(read-only)*
2. Message Trace *(read-only)*
3. Unified Lookup *(read-only, most complex)*
4. Mailbox Permissions *(first writes)*
5. Calendar Permissions
6. Out-of-Office Editor
7. Inbox Rules
8. Group Manager + Group Manager Editor *(most exposed to the cmdlet collision)*

**Deferred out of the first pass — NOT cancelled:** provisioning, reporting
suite, SharePoint tools, theming beyond one working palette. **All of these are
specified in `09-FULL-TOOL-CATALOGUE.md` and all of them get built**, in
lettered stages that begin only after the first pass has been run on Windows and
accepted. They come after something is running; they do not come after nothing.

**Genuinely out of scope, permanently:** programmatic PIM self-activation (see
the Activate spec in `05` and the exclusion table in `09` §6), and anything
requiring Microsoft Graph consent (C20).

> **The finished toolkit is 55 console buttons over 65 `.ps1` files.** This
> section is the first eight tools of it. `09` is the other forty-eight, and the
> destination `Sections.json`. Read `09` §0 before you build anything beyond
> this list — but build **this list first**, and stop at Phase 3 as instructed.

---

## 6. Phases and gates

Do not begin a phase until its gate passes. State clearly when you pass one.

**Phase 0 — verifier.**
Get `Test-Toolkit.ps1` running under `pwsh`. Add the 7-only-syntax rule and any
missing rules from `03`.
*Gate:* runs clean over `reference/`, **and** you have demonstrated at least
three rules firing by deliberately injecting the defects they target — including
the 7-only-syntax rule. A check never shown to fail has not been tested.

**Phase 1 — config and shared layer.**
`Config/*.json` per `02`, then `Common.ps1` **to the contract in `08`**.

> **Budget: roughly 1,000–1,600 lines.** The reference `Common.ps1` is 7,200
> lines and 164 functions, grown against a tenant with more moving parts than
> this one. Build **Tier 1 of `08` and nothing else**; add Tier 2 when a tool
> actually needs it. A shared layer that consumes the whole budget and leaves no
> tools is the most likely way this build fails while looking productive.

*Gate:* verifier clean; no tenant-specific string in any `.ps1`; every Tier 1
name in `08` exists with the documented shape; and `Get-TenantConfig` returns a
usable object when the config file is **absent** — it must never throw, because
a toolkit that will not start cannot tell you why.

**Phase 2 — console and registry.**
`AdminConsole.ps1`, `Tools.json`, `Sections.json`, `ToolManager.ps1`,
`activate.ps1` and the PIM countdown in the console header.
*Gate:* verifier clean, including the control-overlap rule; every registered
tool resolves to a file that exists.

**Phase 3 — first tool, then STOP.**

Build **Mailbox Health only** (spec in `05`, shape in `06`). Then stop and hand
over for review before writing anything else.

This checkpoint is not optional and not a formality. Every convention in the
toolkit — how a tool guards its connection, lays out a form, names things, logs,
reports partial results — is visible in the first tool and is about to be
repeated seven times. **A convention error costs minutes here and hours after
tool eight.** It is also the first point at which a human can run the thing on
Windows, which is the only complete test that exists (see `07`).

*Gate:* the operator confirms the console opens and Mailbox Health works.
Do not proceed on your own judgement.

**Phase 4 — remaining tools, in the §5 order.**
*Gate per tool:* verifier clean; handler wrapping verified structurally; audit
row on success and failure; the tool's specific constraints from `05`
demonstrably applied.

**Stages A–H — the rest of the toolkit.**
Specified in `09-FULL-TOOL-CATALOGUE.md` §5. Forty-eight further tools in eight
lettered stages, each with its own stop. **Lettered, not numbered, so that
"Phase 5" below never becomes ambiguous.** Do not begin Stage A until the first
pass has been run on Windows and accepted per `07`.

**Phase 5 — handover.**
Produce `HANDOVER.md` per §9. Do this at the end of the first pass **and** again
at the end of Stage H.

---

## 7. Working method

**Commit per tool, not per phase.** One reviewable commit per file with a
message saying what it does and which constraints it applies. A single
30-file commit cannot be reviewed and will not be.

**Report at each gate**, briefly: what you built, what the verifier says, what
you could not verify, what you decided that you would like checked.

**Stop and ask when:**

- a constraint in `01` conflicts with a spec in `05` (the constraint wins — but
  say so);
- you cannot verify a cmdlet's behaviour and the wrong guess would write to a
  mailbox;
- a spec is ambiguous in a way that changes what gets written;
- you find yourself about to build something not in §5 **and not in
  `09-FULL-TOOL-CATALOGUE.md`**. Work specified in either document is
  authorised; work specified in neither is not.

**Do not stop and ask** for naming, layout, colours, or anything the reference
files already answer by example. Match them and move on.

**Keep `UNVERIFIED.md`** in the repo. Every time you write something you could
not confirm, add a line: file, line, what you assumed, and what would confirm
it. This is the highest-value artefact you produce after the code itself,
because it tells the operator exactly where to look first. An empty
`UNVERIFIED.md` at the end of a build you could not test is not a good sign — it
means the assumptions were not noticed.

---

## 8. How this specific build goes wrong

Read these as failure modes to watch for in your own output.

**Inventing a cmdlet surface.** The most likely single failure. Parameters that
sound right often do not exist, or exist only in the on-premises edition. See
constraints C10 and C13 for two that are documented as available and are not.

**Writing a conditional branch for a path that cannot occur here.** A single AD
domain means no cross-directory routing. A branch nobody exercises rots silently
and fails on the day it matters. **Prefer absence to a dead branch.**

**Reaching for Microsoft Graph.** It is the modern answer and it will feel like
the right one — for licence state, sign-in status, directory objects. **Do not**
(constraint C20). Almost every permission involved needs admin consent this
tenant may not have granted, and `Connect-MgGraph` authenticates fine without it
and then refuses every call. That looks like a bug in your tool, not a missing
grant, and you cannot detect the difference from where you are building.

**"Improving" a convention mid-build.** Consistency across the toolkit is worth
more than any individual improvement. If a reference convention seems wrong,
say so — do not diverge quietly, and above all do not diverge in *some* files.

**Stubs, placeholders, `TODO`, and `throw "not implemented"`.** None of these
ship. If a tool cannot be finished, deliver fewer tools and say which.

**Swallowing an error to make a run look clean.** An empty `catch {}` is a
defect. Every failure path either surfaces to the operator or writes an audit
row — ideally both.

**Reporting success from the absence of an exception.** Several Exchange
parameters are accepted and silently ignored (C10). Where practical, read back
what you wrote and compare. Where that is impossible, say in the UI that the
value could not be confirmed rather than showing an unverified success.

**Writing a zero-row CSV or an unlabelled partial result.** Both read as "found
nothing" rather than "the run failed" (D6).

**Believing your own checks.** A verification that inspects the thing just
written, in the shape it was written, will agree with itself. For any new check,
ask which defect would make it fail.

---

## 9. Handover

Produce `HANDOVER.md` containing:

1. **What was built**, per file, one line each.
2. **What was verified and how** — parsed with which parser, which linter rules
   passed. State plainly that nothing was executed against a tenant and no form
   was rendered.
3. **`UNVERIFIED.md` summary** — the assumptions most likely to be wrong, ranked
   by what they would cost.
4. **First-run instructions** — what to fill into `Tenant.json` before the
   console will do anything useful.
5. **What was cut**, and what you would build next.
6. **Anything in the reference files you did not understand and therefore did
   not port.** This is genuinely useful information, not an admission.

---

## 10. Definition of done, per script

- Parses under the real parser; contains no banned 7-only construct.
- UTF-8 with a single BOM, CRLF line endings.
- Dot-sources `Common.ps1` by upward search, never a counted `..`.
- Every WinForms handler wrapped, errors surfaced legibly (B1), word wrap on
  (B6).
- Every write audited, with the **primary SMTP address** as target — never a
  display name. With several mailboxes in play, a display name does not identify
  what changed. Where the tool acts on no mail-enabled object at all, a stable
  literal scope name is correct and a fabricated address is not (`09` §4).
- Every enumerating Exchange call passes `ResultSize = 'Unlimited'`; every
  deliberately limited call silences its warning (C3).
- No tenant-specific string anywhere in the file.
- Verifier clean.

---

## 11. Tone of the code

Comments explain **why**, at the line where someone would otherwise reintroduce
a known bug. A short guardrail earns its place —
`# -Identity ONLY: the cloud rejects Identity and Mailbox together` — because
the next person edits that line. Essays belong in notes, not in the script.

Where you apply a constraint from `01`, cite its code in the comment (`# C11:
carry X500 across or replies to historical mail break`). It makes the reasoning
findable and tells the next reader the rule is not arbitrary.

<br>

==============================================================================

# ◆ DOCUMENT 2 OF 10 — `01-HARD-CONSTRAINTS.md`

> The traps. Authoritative and non-negotiable.
>
> *Standalone document. Referred to elsewhere as `01`.*

==============================================================================

# HARD CONSTRAINTS — PowerShell 5.1 WinForms admin tooling

**Status: authoritative. These are not preferences.**

Every item below was paid for once, in production, on a live tenant. Each one
describes code that *parses cleanly, runs, and produces a wrong result or a
misleading error*. That is the category they share and the reason they are worth
more than a style guide: a linter for unbalanced braces would have caught none
of them.

When you are about to violate one of these, stop and say so rather than doing it
and mentioning it afterwards.

---

## A. PowerShell 5.1 language traps

### A1. `return @()` arrives at the caller as `$null`

PowerShell unrolls a collection on return. Two elements arrive as two objects,
one arrives as a bare scalar, **none arrives as nothing at all**.

This becomes a bug the moment `$null` already means something else — typically
"the user cancelled". The two cases become indistinguishable:

```powershell
if ($dlg.ShowDialog() -ne 'OK') { return $null }   # cancelled
if ($nothingSelected)           { return @() }     # unrolls to $null too
```

The caller's `if ($null -ne $result)` guard then treats "empty" as "cancelled"
and the empty case can never be acted on. **There is no error.** Nothing throws,
nothing logs; it presents as a control that does not work.

**Rule:** a function that returns a collection which can legitimately be empty
must not also use `$null` as a sentinel. Return an object:

```powershell
return [PSCustomObject]@{ Cancelled = $false; Values = @($picked) }
```

`,@()` also survives the return and is fine where emptiness is the only
question — but it protects the value without documenting why, so prefer the
object wherever a second meaning is in play.

### A2. `[void]<expr> | Out-Null` fails at COMPILE time

```powershell
[void]$list.Add($x) | Out-Null      # "Argument type cannot be System.Void"
```

Not a runtime error — the file will not compile. Inside a WinForms handler this
surfaces as an unhandled .NET dialog **with no file name and no line number**,
and only when the handler is first invoked. Pick one discard mechanism, never
both.

### A3. `System.Collections.Generic.List[object]` through `@()` can throw

5.1 routes it via `PSToObjectArrayBinder`, which throws *"Argument types do not
match"* from deep inside `Expression.Condition` — including on an **empty**
list. Use `System.Collections.ArrayList` wherever in-place mutation is needed.

Related: `+=` inside a scriptblock only updates a local copy, so any collection
appended from inside a handler or `ForEach-Object` must be an ArrayList.
Remember `[void]` on `.Add()`, which returns the new index (see A2 — `[void]`
**or** `| Out-Null`, never both).

### A4. `[int]` rounds, half to even. It does not truncate

```powershell
[int](239 / 60)     # 4, not 3
```

Anywhere a total is split into units — seconds into minutes, minutes into
hours — use `[Math]::Floor`. The failure mode is a *plausible* number, off by
one unit, in a display nobody cross-checks.

### A5. PowerShell 7 syntax does not load on 5.1 — and your parser may not care

If you are writing on PowerShell 7 and shipping to Windows PowerShell 5.1, the
environment you verify in **accepts syntax the target rejects at parse time**.
The file lints clean where it is written and fails to load where it runs. This
is the single highest-probability way to deliver an unusable build.

Banned outright:

| Construct | Example |
|---|---|
| Null-coalescing | `$a ?? $b`, `$a ??= $b` |
| Ternary | `$c ? $x : $y` |
| Pipeline chain operators | `cmd1 && cmd2`, `cmd1 \|\| cmd2` |
| Null-conditional access | `${a}?.b`, `${a}?[0]` |
| `ForEach-Object -Parallel` | |
| `Sort-Object -Stable` | |
| `ConvertFrom-Json -AsHashtable` / `-Depth` | `-Depth` on **ConvertFrom** is 6.2+ |
| `Join-String`, `Test-Json`, `Get-Error` | |
| `Split-Path -LeafBase` | |
| `$PSStyle` | |

Use `if`/`else` rather than a ternary, and `if ($null -eq $x) { }` rather than
`??`. More verbose, and it loads.

**Two 5.1 behaviours that differ silently instead of failing:**

- `ConvertTo-Json` defaults to **`-Depth 2`**. Anything deeper serialises as
  `System.Object[]` with no error at all. Always pass `-Depth` explicitly.
- `ConvertFrom-Json` returns a `PSCustomObject` and has no `-AsHashtable`. Use a
  `Get-Prop`-style helper for optional properties rather than assuming they
  exist.

**And the automatic variables:** `$IsWindows` / `$IsMacOS` / `$IsLinux` are
PS7-only. In 5.1 they are `$null`, so `if ($IsWindows)` is silently false —
which is the worst of both worlds, since it neither works nor complains. Guard
with `$PSVersionTable.PSVersion.Major -lt 6` first.

This whole class is enforced mechanically by rule `PS7002` in the self-test
spec, because it is not something to remember.

### A6. A block comment cannot contain its own terminator

There is no escape sequence — not a backtick, not doubling. A comment that
spells out the close-comment marker **ends there**, and the rest of the
paragraph becomes code. Describe the marker in words instead.

Detection: a terminator with prose after it on the same line is nearly always
this mistake.

### A7. A doubled quote inside a string swallows the code after it

```powershell
$lbl.Text = "...start the countdown.""
```

The string closes, a new one opens, and everything to the next quote is
consumed. The parser then reports errors on lines that are nowhere near the
fault. Every structural check passes, because they all strip strings before
counting.

### A8. Encoding is load-bearing

All `.ps1` / `.psm1` / `.psd1` files: **UTF-8 with a single BOM, CRLF line
endings.**

- LF-only files render as one continuous line in ISE and Notepad.
- **A double BOM is invisible and destructive.** A reader that strips one leaves
  the second as a zero-width character *inside* the line, so `^\s*#requires` no
  longer matches and the directive is silently dropped.
- `#Requires` is a parser directive, not a comment. It must survive comment
  stripping.

**Write ASCII in your own scripts.** Use two hyphens rather than an em-dash, and
an asterisk rather than a bullet character, in comments and UI strings alike. It
renders identically where this code is read, and it removes an entire class of
encoding damage permanently: a file containing no byte above 0x7F cannot be
corrupted by a reader that guesses the wrong code page. The reference files
follow this. Twenty-nine characters across four adopted scripts were once
destroyed exactly this way, and repairing them to ASCII is what stopped it
recurring.

### A9. Check the bytes of an adopted script BEFORE reading it

A script written elsewhere is very often **CP1252 with no BOM**. Reading it as
UTF-8 destroys every byte above 0x7F — bullets, en-dashes, smart quotes — and a
lenient decoder replaces each with U+FFFD, unrecoverably. If a file has no BOM
and any high byte, it is ANSI. Assuming UTF-8 is not a neutral default; it is a
destructive one.

### A10. Splatting needs a variable

```powershell
& $path @params      # correct
& $path @($params)   # array subexpression: passes the whole dictionary as ONE positional arg
```

### A11. `Start-Process -ArgumentList` does not quote for you

Every path and value in a constructed child command line must be quoted
explicitly, with internal apostrophes doubled. Valid PowerShell can build an
invalid command line for a *different* program, and no PowerShell-level check
inspects that string.

---

## B. WinForms traps

### B1. Wrap EVERY event handler in try/catch

An exception inside a handler is caught by .NET rather than PowerShell, so the
operator gets a CLR crash dialog — a hundred lines of internals and the loaded
assembly list, **no script name, no line number**. Provide one
`Show-HandlerError` helper that converts it back to message / script / line /
statement and writes a failure row to the audit log. An error the operator
dismissed is otherwise the one failure that leaves no trace.

### B2. A modal `ShowDialog()` blocks any window raised on the same thread

If the console ends in `$form.ShowDialog()` and launches tools with `&` (same
process, same thread), any window a tool raises is created into a thread whose
input the console owns. It draws and then ignores all input.

Affects: embedded sign-in prompts, `Get-Credential`, brokers. **Not** affected:
anything in a separate process (a default-browser sign-in works fine).

Rule: a tool that raises an external window it does not own cannot run inline
from a console button. It must run after the modal loop exits — record the
choice, close the form, do the work below `ShowDialog()`, relaunch with `&`.

Diagnostic worth ten seconds: close the console and run the same script from a
bare prompt. If it works, the console is the problem, not the script.

### B3. Colour and control rendering

- `Button.BackColor` is **ignored unless `FlatStyle` is `Flat`**.
- A RadioButton/CheckBox glyph is drawn by the system and **cannot be
  coloured**. On a dark theme the selected dot is dark-on-dark and the control
  reads as though nothing is selected. Signal selection through the **label
  text** instead.
- A `Label` carries exactly one `ForeColor`. Two independent states need two
  labels.
- `Control.DoubleBuffered` is `protected` — set it by reflection or gradient
  forms flicker on resize.
- **Semantic colour must survive theming.** Tag controls whose colour carries
  meaning and skip them; apply status colouring *after* theming so it wins. A
  theme must never be able to disguise a red "not connected" as decoration.

### B3a. Window placement and width are conventions, not preferences

**The console owns TOP-LEFT. Every other window opens BOTTOM-LEFT.**

A tool that also opens top-left lands squarely on the console and hides the
connection indicators -- which are exactly what you may need to read *while the
tool is running*. In the source toolkit **63 of 65 forms** use bottom-left; the
two exceptions are the two consoles.

This is invisible until somebody has both windows open at once, which is why it
is a rule rather than something to notice later. Three tools broke it in one day
and it was found by looking, not by any check.

**Width follows what the tool is FOR:**

| Kind | Width | Why |
|---|---|---|
| **Action tools** -- anything that takes input and changes something | **Match the console** (840) | A window wider than the console that launched it reads as a different application. One shipped at 1160 and was rejected on sight. |
| **Viewers** -- log and report browsers | Deliberately wider | Comparing many columns across many rows IS the feature; constraining them forces horizontal scrolling on the one screen where seeing everything at once is the point. |

**Taller is cheap; wider is not.** Where a column will not fit, it belongs in a
detail pane below the grid rather than pushing everything else off screen.

### B3b. Every tool carries a Help button

`Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title '<name>'` on
**every** tool, with the `TOOLKIT-HELP-BEGIN`/`END` block filled in. No
exceptions -- a tool without one is not finished.

It costs one line and it is the only in-product answer to *"what does this
actually do, and is it the right tool for my ticket?"* An operator who reaches
for the wrong tool concludes the data is missing rather than that they are in
the wrong place, and that misdiagnosis is expensive and completely avoidable.

**The end-of-form sequence is fixed, and every step has a reason:**

```powershell
$bar = Add-ToolTargetBar -Form $form ...        # 1. FIRST -- it shifts every
                                                #    existing control DOWN, so
                                                #    anything added after it is
                                                #    left at the old offset
Set-ConsoleTheme     -Form $form                # 2.
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'   # 3. AFTER any resize
Add-ToolHelpButton   -Form $form -ScriptPath $PSCommandPath -Title '<name>' | Out-Null
                                                # 4. LAST -- after the theme so
                                                #    it is not restyled into an
                                                #    ordinary button
[void]$form.ShowDialog()
```

Two of those orderings are not obvious. **The target bar goes first** because it
pushes existing controls down and a control added afterwards keeps the old
coordinates. **The help button goes last** because the theme would otherwise
restyle it into a normal button -- and it calls `BringToFront` on itself
internally, because WinForms draws *earlier-added* controls on top.

### B4. Geometry

- `Form.Size` is not `Form.ClientSize`. Decide which one your layout maths
  means.
- Use `Screen.PrimaryScreen.WorkingArea`, **not** `Bounds` — `Bounds` ignores
  the taskbar and pushes bottom-aligned windows underneath it.
- Clamp any corner placement so an over-tall form cannot lose its title bar
  off-screen and become undraggable.
- **Two controls at identical coordinates both exist, both wire up, and one is
  invisible underneath the other.** Only a screenshot or an overlap check finds
  it.
- **A control that computes its own position at run time is invisible to a
  static overlap check**, and those are precisely the ones most likely to
  collide. Include their computed positions explicitly.
- A control that is empty at layout time and populated later measures as
  zero-width and gets "avoided" — of nothing. Give it the space it will need or
  put it where the dynamic control never reaches.

### B5. Handlers created in a loop

Attach per-item state via the control's `.Tag`. Do **not** use
`.GetNewClosure()` — it captures a snapshot and the divergence between what you
captured and what you meant is invisible at the call site.

### B6. Never truncate an error in a log window

`WordWrap = $true` everywhere. A real diagnosis was once entirely in the second
clause of a message, off the right edge of a window with wrapping disabled, and
three rounds of fixes went to the wrong layer. The cost of a horizontal
scrollbar is that nobody scrolls.

---

## C. Exchange Online / on-premises Exchange / AD

### C1. The on-prem / cloud cmdlet collision is the highest-consequence hazard

`Get-Mailbox`, `Set-Mailbox`, `Get-MailboxPermission` and dozens more exist in
**both** the on-premises Exchange Management Shell and the Exchange Online
module. With both loaded, **which one runs depends on module load order** — not
on anything visible at the call site. A tool meant to modify an on-premises
group can silently operate against the cloud.

Two defences, use both:

1. Force resolution to the temporary `tmpEXO_*` module that
   `Connect-ExchangeOnline` creates, via a `Get-EXOCommand` / `Invoke-EXOCommand`
   wrapper, so a cloud call is guaranteed to be one.
2. Prefer cmdlets that cannot collide. `*-RemoteMailbox` exists only
   on-premises, so a tool using only those is safe beside an open cloud session.

### C2. "EMS is available" is not "EMS will answer"

A probe can confirm the on-premises cmdlets are loaded and still be wrong about
which ones will *run* (see C1). A tool that **writes** on-premises must check
both "are on-prem cmdlets present" and "is a cloud session live in this
process", and **refuse** — not warn — when both are true.

Let a tool declare its own process requirement in a header marker the console
obeys, e.g. a `TOOLKIT-REQUIRES: OwnProcess` line in the first 40 lines.
Requiring an operator to remember a checkbox for **correctness** is a design
defect.

> Implementation warning learned the hard way: if you detect that marker by
> matching text in the first N lines, a *comment about* the marker is
> indistinguishable from the marker. Do not write its literal name in prose near
> the top of a file that must not declare it.

### C3. `ResultSize` has two opposite traps

| | Cause | Symptom | Fix |
|---|---|---|---|
| A | No `ResultSize`; the default (1000) is hit | Warning **and a silently short answer** | `ResultSize = 'Unlimited'` |
| B | A deliberately small `ResultSize` (a probe) | Warning only; answer is correct | `-WarningAction SilentlyContinue` |

Any call that **enumerates** must pass `Unlimited` explicitly. Any call that
**deliberately limits** will warn every time and must silence it. Treating these
as one problem costs three attempts.

Not universal: `Get-SPOSite` with `-Limit`/`-Filter` stops populating
`SharingCapability` and `StorageUsageCurrent`; `Get-SPOExternalUser` caps
`PageSize` at 50. Ask the cmdlet, do not generalise.

### C4. Never round-trip a resolved object's `Identity` back into `-Identity`

`Get-DistributionGroup -Identity someone@contoso.com` returns an object whose
`Identity` renders as the AD CN — no domain, no OU, no address. Passing that
back is **not a bind**; it is a fresh forest-wide name lookup that can resolve
to a different object.

| Key | Unique | Accepted by |
|---|---|---|
| `Guid` | forest-wide | on-prem and cloud — **use this** |
| `DistinguishedName` | forest-wide | on-prem only |
| `PrimarySmtpAddress` | org-wide | both, but may be a secondary elsewhere |
| `Identity` | **not unique** | never round-trip |

Resolve once, then bind on a key that cannot be re-interpreted. **Log which key
was used** — when a write fails, that is the first thing worth seeing.

### C5. "Out of the current user's write scope" almost never means permissions

That message has at least three unrelated causes:

- a domain controller pinned to a domain that is not authoritative for the
  object (C6);
- an ambiguous `Identity` that resolved to a different object (C4);
- a cloud session refusing to modify a directory-synchronised object — where
  **the entire diagnosis is in the second clause** of the message.

Read the whole message before theorising.

### C6. Pinning a domain controller breaks cross-domain writes

In a multi-domain forest, a controller in one domain is not authoritative for an
object in another. Pinning lets **reads** succeed forest-wide while **writes**
fail with the message in C5.

Rule: pin only in a tool scoped to one domain by definition. A general-purpose
tool must not pin at all. Assert `-ViewEntireForest $true` rather than assuming
it — without it a cross-domain lookup returns nothing, which reads as "no such
object" rather than "wrong scope".

**Removing a control does not remove its effect.** A hidden DC selector kept
applying its default for weeks.

### C7. `Set-ADServerSettings` scope is process-wide and cannot be cleared

It outlives the tool that set it. `-PreferredServer` is a mandatory FQDN and
`-ViewEntireForest $true` is documented to clear only `RecipientViewRoot`, so
there is no documented way to unset a pin. One tool's pin is inherited by the
next tool in the same window.

Consequence: pass `-DomainController` per cmdlet call rather than relying on
ambient session scope. A per-call parameter overrides the session preference.

### C8. Exchange's `-DomainController` rejects a domain name

Exchange's `-DomainController` and `-PreferredServer` both run a suitability
check and reject a domain with *"Active Directory error 0x52 … A local error
occurred"*. The AD module's `-Server` accepts a domain quite happily.

**The rule is by module, not by parameter name.** An AD cmdlet may be handed a
domain; an Exchange cmdlet must be handed a real controller, or nothing at all.

### C9. `ArchiveStatus` lies — use `ArchiveGuid`

`ArchiveStatus` is a cloud-side property. On an on-premises remote mailbox it
reflects what has replicated back, which lags by minutes to hours, so reading it
seconds after enabling an archive returns `None` for an archive that exists.
`ArchiveGuid` is stamped locally as part of the enable. **GUID present = archive
enabled.**

### C10. Primary SMTP address in Exchange Online

- `Set-Mailbox -PrimarySmtpAddress` is **on-premises only**. It does not exist
  in the cloud.
- `Set-MailUser -PrimarySmtpAddress` is worse: it is **accepted and inert**. The
  parameter lists Exchange Online as applicable, and the body of the
  documentation says it functions only on-premises.
- `Set-MailContact -PrimarySmtpAddress` is on-premises only.
- `Set-DistributionGroup` and `Set-UnifiedGroup` **do** accept it.

The supported cloud route for a mailbox is to rewrite the whole
`EmailAddresses` collection with exactly one uppercase `SMTP:` entry.

### C11. A bare `EmailAddresses` list REPLACES the whole proxy collection

Everything not restated is destroyed — including **X500 addresses**, whose loss
silently breaks replies to historical mail. Read the existing collection, carry
every non-primary entry across, and write the whole thing back.

Documented on all four cmdlets: you cannot use `EmailAddresses` and
`PrimarySmtpAddress` in the same command.

### C12. `GroupMailbox` contains the substring "mailbox" and is not one

It is a Microsoft 365 Group. Any `-match 'mailbox'` type test will classify it
wrongly. This single substring caused three separate bugs in one session. Match
recipient types exactly; never by substring.

### C13. Ask the cmdlet what it declares — do not assume

The module surface varies by version and tenant.

- `RecipientTypeDetails` is a **filter parameter, not a returnable property**;
  requesting it in `-Properties` fails the whole call.
- `-PageSize` does not exist on every build of `Get-EXOCasMailbox`.

Provide a `Get-CmdletParameterNames` helper and pass nothing that is not
declared. Reading the documentation is not sufficient.

### C14. Properties that are `$null` meaning "inherit"

`EwsEnabled` is frequently `$null`, meaning *inherit the organisation default* —
not `False`. Resolve against `Get-OrganizationConfig` and report the effective
value. Same class of trap as C9.

### C15. Bulk queries: shape the request, do not retry it

The same cmdlet that never fails for one identity will die with *"The underlying
connection was closed"* when asked for 5,000 in one call. **Retrying reproduces
the oversized response** — backoff helps with throttling and does nothing for a
payload too large to survive the trip. Slice the request instead (by recipient
type, then by first letter of alias) and deduplicate.

Guard against silent-wrong: if a filter were quietly dropped, each pass would
return everything, and you would ship a plausible CSV with sevenfold duplicated,
mislabelled rows.

### C16. A stale cloud session presents as a transport error, not an auth error

*"The underlying connection was closed"* with no other detail is worth one
reconnect in a fresh window before debugging code.

### C17. Connections are per-process

`Connect-ExchangeOnline` builds an in-memory temporary module inside the calling
process. **A child process cannot inherit it.** Any launcher that spawns a child
per tool will re-prompt for login every time, and this cannot be fixed by better
plumbing — it is wrong at the design level. Launch tools with `&` in the same
process, or accept that a child must connect for itself.

Corollary: the cloud session's implicit remoting is **not thread-safe**.
Serialise.

### C18. Directory-synchronised objects

Anything with `IsDirSynced` true is mastered elsewhere and the cloud will refuse
most changes. Say so *before* the attempt rather than surfacing the refusal.

### C19. Deep links need no Graph consent

`Get-Mailbox` returns `ExternalDirectoryObjectId`, which **is** the Entra object
GUID. Portal deep links can be built from it with no Graph permission and no
admin consent.

Prefer `go.microsoft.com/fwlink/p/?linkid=` redirects over portal hash routes —
Microsoft maintains the former and changes the latter without notice. Mark which
kind each link is, so a button that lands somewhere unexpected is already
labelled as the fragile sort.

### C19a. Escape any operator-typed value before putting it in a filter

Every lookup builds a filter around something a person typed. Two escapers
exist and they are **not interchangeable**:

| Filter kind | Escaper | Escapes |
|---|---|---|
| Exchange `-Filter` (OPATH) | `ConvertTo-OpathValue` | the single quote, by doubling |
| LDAP / `Get-AD*` `-LDAPFilter` | `ConvertTo-LdapFilterValue` | backslash, parens, asterisk, NUL, as hex |

Using the wrong one produces a query that is **wrong rather than one that
fails** -- the worst available outcome, because nothing reports it.

**One apostrophe is enough.** A search for an Irish surname, or an office with
`O'` in its name, breaks an unescaped OPATH filter. It is a correctness problem
rather than a security one -- these filters read, and the operator already has
rights to read -- but a tool that silently finds nobody named `O'Brien` gets
reported as "search is broken", and the cause is nowhere near the symptom.

Take the reference implementations verbatim. The hand-rolled version was written
twice, identically, and **threw every time it ran**: `[char]0` in the escape
table makes the first argument a char, `String.Replace` has a `(char, char)`
overload as well as `(string, string)`, PowerShell binds to the char one, and
then cannot convert a four-character replacement string. Enforced by `FLT001`.

---

### C20. Do NOT use Microsoft Graph. Consent is the blocker, not capability

**No `Connect-MgGraph`, no `Microsoft.Graph` module, no `Mg*` cmdlets, no direct
calls to `graph.microsoft.com`, unless explicitly told otherwise.**

This is not a claim that Graph is the wrong technology. For several things it is
the *only* documented answer, and where that is true this toolkit does without
the feature rather than without the consent.

**The blocker is organisational, not technical.** Nearly every permission this
toolkit would want from Graph is an **admin-consent delegated permission** on the
Microsoft Graph PowerShell application. Whether that consent exists is somebody
else's decision, it is frequently refused, and you cannot find out from here.

**The failure mode is why this is a hard rule.** `Connect-MgGraph` authenticates
perfectly against a tenant that has not consented. The sign-in succeeds, the
session looks healthy, and then every call returns a permission error — which
reads as a bug in the tool rather than as a missing grant. You would be shipping
something that looks finished and works nowhere.

**What to use instead, all of which need no consent beyond what the operator
already has:**

| Instead of Graph for… | Use |
|---|---|
| Licence state | **AD group membership.** Where licensing is group-based, the licence someone holds is a fact about `memberOf` — and it reports the group that drives the assignment, which is the thing an operator can actually act on |
| Portal deep links | `ExternalDirectoryObjectId` from `Get-Mailbox` **is** the Entra object GUID (C19). No permission, no consent |
| Mailbox, group, recipient data | Exchange Online cmdlets |
| Directory objects | The AD module, against the on-premises directory |

**If you believe a requirement genuinely cannot be met without Graph:** stop,
say so, name the exact permission and whether it needs admin consent, and record
it in `UNVERIFIED.md`. Do not implement it and leave a note. Do not add the
module "just in case" — `Microsoft.Graph` is large and slow to import, and an
unused dependency in a toolkit that runs from a share is a cost with no benefit.

Rule `GRF001` in the verifier enforces this.

---

## D. Verification discipline

### D1. Brace counting is not a syntax check

Comments and string literals contribute, so a file can balance while being
unparseable and can fail to balance while being perfectly valid. An unterminated
block comment once swallowed an entire function for two days — and the braces
balanced, because a stray `{` inside the runaway comment offset the orphaned `}`
it created.

**Only the real parser is a syntax check.** In PowerShell that is
`[System.Management.Automation.Language.Parser]::ParseFile`. Run it. Everything
else is a smell test.

### D2. Beware the check that can only pass

A verification that inspects the thing just written, in the shape it was
written, will agree with itself. Two checks once both confirmed an edit had been
applied; neither asked whether the result was valid.

**The honest question to ask of any new check: which defect would make this
fail?** A check that has never been shown to fail on a real defect has not been
tested — so re-inject the bug deliberately and confirm it fires.

### D3. A fix that does not move the symptom means the diagnosis is wrong

Not incomplete — wrong. This appears three separate times in the source
material, each time costing more than the last, because a *real* defect in the
neighbourhood of the traceback absorbed the attention.

**Reproduce by hand and diff the working invocation against the failing one:
arguments first, then environment.** One command in a plain shell has repeatedly
settled in seconds what rounds of reasoning did not.

### D4. Scripted edits apply an assumption uniformly

When the assumption is wrong it is wrong in every file at once — speed and blast
radius are the same property.

- **Locate edits structurally, not lexically.** "The next `})` at column zero"
  is a guess about formatting; walking brace depth from a known opening position
  is a fact about the code.
- **Assert the match count before editing.** `assert count(old) == 1` turns an
  ambiguous target into a failure instead of a silent mis-edit.
- **When a bulk edit is driven by a rule, assert the rule's premise separately
  first.** A rule that fired 12 times with 2 true positives would have destroyed
  a working tool if acted on.
- **Verify the property, not the edit.** Counting braces proves nothing about
  whether a `try` has a `catch`.

### D5. Weak evidence looks like strong evidence

`grep -c 'Common.ps1'` counts *mentions* — including a comment saying the file
deliberately does **not** load it, which is the exact opposite of what the count
implies. A grep that answers a nearby question is more dangerous than one that
answers nothing.

### D6. Output honesty

- **Never write a zero-row CSV.** An empty report reads as "found nothing"
  rather than "the run failed".
- **A partial run must say so** — in the summary *and* in the audit row. A
  report that is incomplete but looks complete is worse than one that failed.
- Silent-wrong beats loud-wrong in cost, so spend effort making wrong answers
  loud.

### D7. Running it remains the only complete test

Three defects in one week passed every static check: a doubled quote inside a
string, a rounding cast producing plausible numbers, and a block comment
containing its own terminator. In each case the check could not see the thing
that was wrong.

<br>

==============================================================================

# ◆ DOCUMENT 3 OF 10 — `02-ARCHITECTURE-SPEC.md`

> Structure, config schemas, conventions, console layout.
>
> *Standalone document. Referred to elsewhere as `02`.*

==============================================================================

# ARCHITECTURE SPEC — structure, schemas, conventions

Target: **PowerShell 5.1**, WinForms, Exchange hybrid, **one on-premises AD
domain**, Exchange Online + SharePoint Online.

5.1 is the target because it is the only version where all three modules work
without qualification. Only Exchange Online is PS7-native; SharePoint Online
needs `-UseWindowsPowerShell` there, and the legacy directory modules do not
work at all.

---

## Layout

```
Toolkit\
  AdminConsole.ps1        launcher: connection state, buttons, layout
  Common.ps1              shared layer, dot-sourced by every tool
  ToolManager.ps1         edits Tools.json / Sections.json / Settings.json
  Test-Toolkit.ps1        the verifier (03)
  Connect-*.ps1           one per service
  <Tool>.ps1              one file per tool
  Config\
    Tenant.json           tenant FACTS       — same on every machine at a tenant
    Settings.json         engineer PREFERENCES — differs per machine
    Tools.json            button registry
    Sections.json         console layout
    Themes.json           palettes — per-deployment, NOT shipped in a handout
    QuickLaunch.json      the operator's own links (Stage G)
```

**The split between `Tenant.json` and `Settings.json` is load-bearing.** Tenant
facts are identical for everyone at that organisation and belong in source
control; paths and browser profiles are personal and must not be. Putting them
in one file guarantees somebody overwrites a colleague's paths while updating a
domain.

---

## `Tenant.json` — the only file allowed to name the organisation

This is the single exemption from the tenant-neutrality rule. If a tenant name,
domain or server appears in a `.ps1` file, that is a bug and `TEN001`/`TEN002`
will fail the build.

```jsonc
{
  "TenantName": "",
  "PrimaryDomain": "",
  "OnMicrosoftDomain": "",

  "Capabilities": {
    "ExchangeOnline":   true,
    "SharePointOnline": false,  // FIRST PASS ONLY. True at the tenant, but no
                                // first-pass tool uses it, and a capability
                                // with nothing behind it is a permanently red
                                // indicator.
                                // FLIP THIS TO true AT STAGE A, when Connect-SPO
                                // arrives. Leave it false and the SPO button is
                                // not greyed -- it is absent (see Tenant Manager
                                // in 09), which reads as a build defect.
    "OnPremAD":         true,
    "OnPremExchange":   true,
    "Pim":              true    // rights are ELIGIBLE, not standing
  },

  // A LIST, even at one entry. A tenant with one domain has one; a tenant with
  // five has five; a cloud-only tenant has none and every AD tool greys out on
  // its own. Never a hardcoded pair, and never a ValidateSet naming them.
  "Directories": [
    {
      "Id": "Corp",                 // identifier: short, stable, written to the audit log
      "Label": "Corporate",         // descriptive, changeable
      "Domain": "",
      "DomainControllers": [],      // EMPTY IS USUALLY RIGHT — see below
      "SharedMailboxOu": "",        // blank means "ask each time"
      "DistributionGroupOu": "",
      "IsDefault": true
    }
  ],

  "OnPremExchange": { "ServerUri": "", "PreferredDomainController": "" },

  // DurationHours MUST match tenant PIM policy -- 9 here. The console header
  // counts down from it. It is not a timeout the toolkit enforces; it is a
  // display of how long Entra granted, and a wrong value is a confidently
  // wrong countdown. No script may hardcode a duration.
  "Pim": { "DurationHours": 9 },

  "LicenseGroups": [],              // { Name, Sku, Directory, Purpose }

  // ADMIN CENTRE URL IS A TENANT FACT, NOT A PREFERENCE, and lives ONLY here.
  // Blank means "derive it from the *.onmicrosoft.com label". See Connect-SPO
  // in 09 -- an earlier version read this from Settings.json and then guessed,
  // so the value typed here was inert.
  "SharePoint":    { "AdminUrl": "", "MySiteHost": "" },

  // Where Report a Bug and Feature Request send to (Stage H). ONE key, shared
  // by both tools -- a second key would be a second thing to get wrong. Both
  // tools fail with "No recipient is configured" until it is set.
  "FeatureRequest": { "Recipient": "" },

  "Theme": ""
}
```

**`_comment*` keys are part of the schema.** The template carries
`_comment`, `_comment_Capabilities`, `_comment_Directories` and others, and they
are the only in-file explanation an engineer editing this by hand will see.
**Tenant Manager must round-trip every key it does not recognise, unchanged.**
An editor that silently drops unknown fields deletes a live setting eventually —
that has already happened once in this toolkit's history.

**`DomainControllers` empty is the better default.** Naming none lets AD locate
a controller, which is what you want. Naming one pins writes to that box and
owns the consequences when it reboots — and see constraint C6: a wrong pin is
worse than no pin.

**`LicenseGroups` needs `Purpose`.** A list that serves two unrelated jobs with
nothing recording which is which produced a real failure: a leaver was removed
from a reporting group they were never in, the step reported success, and the
licence stayed assigned. `Purpose` distinguishes `Termination` (safe to remove)
from `Report` (must not be touched).

**Capability flags drive greying, not deletion.** A false here disables the
tools that need it, with a tooltip saying why. That is how you ship one codebase
to environments of different shapes without conditional branches rotting inside
the tools.

---

## `Common.ps1` — the shared layer

Dot-sourced by every tool. Responsibilities:

**Config** — `Get-TenantConfig` (cached process-wide, **never throws**; a
toolkit that will not start because a config file is missing is worse than one
that starts and reports every capability off), `Get-TenantDirectory` (list, or
one by Id), `Get-DefaultTenantDirectory`, `Get-ToolkitSettings`.

**Connections** — one probe per service, plus `Test-ExoInProcess` used by every
on-premises writer to refuse when a cloud session is live in the same process
(C2).

**Capability gating** — decide which connections a tool needs by **reading its
source and matching cmdlet patterns**, not by a hand-maintained table. It stays
correct as scripts change. Cache the result keyed by path + last-write time, in
`%LOCALAPPDATA%`, not beside the toolkit — the toolkit runs from a share and
several operators would contend over one file. Every failure path must fall
through to a rescan; an optimisation that can fail the program is not one.

**Audit log** — one CSV, append-only, with columns:
`Timestamp, Operator, Machine, Tenant, Script, Action, Target, Result, Detail, RunId, Environment`.

- `Tenant` and `Environment` are different questions. `Environment` is a
  *directory within* a tenant; `Tenant` is which organisation. One console
  shipped to several tenants writes the same format everywhere, and without
  `Tenant` two files side by side cannot answer "who changed what".
- `Target` is **the primary SMTP address whenever the tool resolved a
  mail-enabled object** — never a display name. With several mailboxes in play a
  display name does not identify what changed.
  **Where there is no such object** — a connection script, a config editor, the
  publish chain — `Target` is a **stable literal scope name** (`SharePoint
  Online`, `Toolkit`, `Config`, a report path). Never fabricate an address to
  satisfy the rule; a literal *is* the identifier in those cases. See
  `09-FULL-TOOL-CATALOGUE.md` §4.
- `Result` is `Success` / `Partial` / `Failed`. A partial run must say so.
- Version the schema and refuse to append to a file of a different version;
  `Export-Csv -Append` throws on a header mismatch in 5.1, and silently widening
  someone's audit file is not a logging function's decision.

**Error display** — `Show-HandlerError` converts a .NET exception into message /
script / line / statement and writes a `Failed` row (B1).

**Geometry** — corner placement clamped to `WorkingArea`; console top-left,
tools bottom-left so a tool never covers the console that launched it.

**Theming** — semantic colours tagged and skipped; status colouring applied
after theming so it always wins.

---

## `Tools.json` — the button registry

```jsonc
{
  "Name": "Mailbox Health",
  "Script": ".\\MailboxHealth.ps1",
  "Section": "Admin Tools",
  "Order": 3,
  "ObjectTypes": [],        // empty = never greyed on the LOOKED-UP OBJECT.
                            // Capability gating still applies.
  "Group": "",              // optional. Same string on several tools folds them
                            // behind ONE button with a chooser. See 09 section 3.
  "NewProcess": false
}
```

**Pass a tool only the parameters it declares.** Ask via
`Get-Command -CommandType ExternalScript` rather than passing a fixed set. A
script declaring *some* of them rejects the others with a binding error, while
one declaring **no** `param()` block silently absorbs extras into `$args` — so
the fixed-set approach works right up until someone writes a normal script.

**A section holds `Rows` x `ButtonsPerRow` buttons and no more.** Register a
nineteenth tool in a section with eighteen slots and it has nowhere to draw --
the entry is valid, the file parses, and the button is simply absent. Check the
arithmetic when adding tools, and have `ToolManager` report a section that is
over capacity rather than letting the console silently drop the overflow.

**`NewProcess` is a preference, not a safety mechanism.** A tool that must not
share a process declares it in its own header (C2). Asking an operator to
remember a checkbox for correctness is a design defect.

Tools launched in their own process must **not** be greyed on the console's
connection state — they will use a different process, so gating them here blocks
them for a condition with no bearing on whether they can run.

---

## Sections — how buttons are grouped

`Sections.json` defines the console's rows. The shipped reference has four
sections, but **the first pass ships two**, because a section with no tools in
it renders as an empty box:

```jsonc
{
  "ButtonsPerRow": 5,
  "Sections": [
    { "Name": "Connect",     "Rows": 1 },   // 5 slots
    { "Name": "Admin Tools", "Rows": 2 }    // 10 slots, 8 used
  ]
}
```

**Connect comes first, and Activate comes first within it** -- because that is
the order the operator must follow. Rights are eligible rather than standing, so
a connection made before activating carries a token that does not contain the
access activated afterwards (see the Activate spec). Putting Activate at the
left of the top row makes the sequence the layout rather than something to
remember:

| Section | Buttons, in order |
|---|---|
| **Connect** | Activate · Connect Exchange Online · Connect On-Premises |
| **Admin Tools** | Mailbox Health · Unified Lookup · Group Manager · Mailbox Permissions · Calendar Permissions · Out-of-Office · Inbox Rules · Message Trace |

**Group by what the operator is DOING, not by which service serves it.** Message
Trace sits with the mailbox tools even though it is a different Exchange surface,
because somebody chasing a delivery problem is doing mailbox work. The source
toolkit later added `Daily Scripts` and `Reporting` on the same principle: add a
section when it has three or more tools, not before.

**`Order` within a section is the button order**, left to right then down. Keep
the most-used tool first -- in the first pass that is `Mailbox Health`, the
answer to "something is strange with my email", which is the most common ticket
there is. **From Stage G, `Order: 0` belongs to `Dashboard`**, which is the
start-of-day page and is meant to be the first thing on the console.

**Grow `Rows` as a section fills.** These numbers are the first pass, not the
destination: `09` §2 gives the finished layout, and each stage that adds buttons
must raise `Rows` in the same commit. **An over-capacity section does not
error — the surplus button is simply absent**, which is why `WF011` exists.

---

## Console behaviour

- Ends in `$form.ShowDialog()` — a **modal loop**. Anything that raises an
  external sign-in window must run *after* it returns (B2), then relaunch the
  console with `&` to preserve the session (C17).
- Launch tools with `&`, same process, so they inherit connections.
- Buttons always visible, greyed when their capability is unavailable, with a
  tooltip saying which one. A missing button is a mystery; a greyed one with a
  reason is an answer.
- Width: action tools match the console's width. Viewers — log and report
  browsers — are deliberately wider, because comparing many columns is their
  entire job.

---

## Conventions

- **Reasoning lives in notes, not in the script.** A short guardrail comment at
  the line where somebody would reintroduce a known bug earns its place. The
  essay does not.
- **Prefer absence to a dead branch.** A conditional path for an environment
  that cannot occur is one nobody exercises, so it rots silently and fails on
  the day it matters.
- **Search upward for `Common.ps1`; never count `..`.** A hardcoded climb
  encodes the folder depth into every file at the bottom of the tree, and fails
  in the worst way — a plausible directory that exists, so the error names the
  wrong problem.

```powershell
$root = $PSScriptRoot
while ($root -and -not (Test-Path -LiteralPath (Join-Path $root 'Common.ps1'))) {
    $parent = Split-Path -Parent $root
    if (-not $parent -or $parent -eq $root) { $root = ''; break }   # guard: drive roots loop forever
    $root = $parent
}
if (-not $root) { throw "Common.ps1 not found above '$PSScriptRoot'." }
. (Join-Path $root 'Common.ps1')
```

- **State shared with launched tools must be `$global:`, not `$script:`.** A
  tool invoked with `&` runs in its own script scope; functions from a
  dot-sourced module remain callable, but any `$script:` variable they read
  resolves against the *child's* scope where it was never assigned. The function
  works and its state is `$null`. A `Get-Command` guard proves the function
  exists and nothing about whether its state does.
- Paths derived inside a function come from `$PSScriptRoot`, which resolves to
  the file that *defined* the function regardless of caller.

<br>

==============================================================================

# ◆ DOCUMENT 4 OF 10 — `03-SELF-TEST-SPEC.md`

> The verifier. Build it first.
>
> *Standalone document. Referred to elsewhere as `03`.*

==============================================================================

# SELF-TEST SPEC — `Test-Toolkit.ps1`

**Build this first, before any tool.** It is the gate between every step, not a
phase at the end.

The reason it comes first is specific: most of the defects in
`01-HARD-CONSTRAINTS.md` do not fail visibly. They produce plausible output.
A rule that catches one mechanically is worth more than an instruction to
remember it, because the instruction competes with everything else in context
and the rule does not.

---

## The one check that is actually a syntax check

```powershell
$tokens = $null; $errors = $null
[System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
```

Everything else in this file is a smell test. Brace counting is **not** a syntax
check: comments and string literals contribute, so a file can balance while
being unparseable, and can fail to balance while being perfectly valid. Report
every parse error with its line, and treat any parse error as a build failure.

---

## THE REFERENCE VERIFIER USES DIFFERENT RULE CODES. Do not renumber it.

`reference/Test-Toolkit.ps1` already implements most of what follows, under its
own codes — the audit-target rule is `TGT001`, the WinForms layout rules are
`LAY001`/`LAY002`/`HDL001`, and a few codes (`LOG002`, `INT001`) exist in **both**
with **different meanings.**

**Run the reference verifier and extend it. Do not reimplement it from this
spec, and do not renumber its rules to match these tables.** Where a rule below
already exists in the reference under another code, that is a naming difference,
not a missing check — say so in your report rather than adding a duplicate.

**Add these tables' rules only where the reference has no equivalent check.**
The one that is definitely missing is `PS7002`, because the reference was
written on the target platform and never needed it.

Where you do add a rule, take the next free number in that family rather than
reusing one — `WF004` and `LOG002` are both already taken, in this file and in
the reference respectively, and a code meaning two things is worse than a code
nobody recognises.

---

## Rules

Each rule gets a stable code so findings can be discussed and suppressed
individually. Severity `ERROR` fails the build; `WARN` is reported and
tolerated.

### Encoding and file shape

| Code | Severity | Check |
|---|---|---|
| `ENC001` | ERROR | File is UTF-8 with exactly **one** BOM. Two BOMs is the dangerous case — the second becomes a zero-width character inside line 1 and silently breaks `#Requires`. |
| `ENC002` | ERROR | Line endings are CRLF throughout. LF-only renders as a single line in ISE. |
| `ENC003` | ERROR | No byte above 0x7F unless the file is confirmed UTF-8. Flags adopted CP1252 files before they are read and destroyed. |
| `ENC004` | WARN | `#Requires -Version 5.1` present in the first lines. |

### Syntax-adjacent defects the parser reports badly

| Code | Severity | Check |
|---|---|---|
| `SYN001` | ERROR | Parse errors from `ParseFile`. Report line and message verbatim. |
| `SYN002` | ERROR | A block-comment terminator with prose after it on the same line. A block comment **cannot contain its own terminator** and there is no escape; a real terminator sits alone or at end of line. |
| `SYN003` | ERROR | A string that closes and immediately reopens (`.""` / `.''`). Requires walking real string state — `'..'` with `''`, `".."` with `""` and backtick escapes, `$( )` nesting, here-strings. Also catches an unterminated string at EOF. |
| `VOID001` | ERROR | `[void]<expr>` piped to `Out-Null`. **Compile-time** failure, surfaces as a CLR dialog with no line number. Must be block-comment aware or it flags its own documentation. |

### Language traps

| Code | Severity | Check |
|---|---|---|
| `RET001` | WARN | A function containing both `return @()` and `return $null`. The empty array unrolls to `$null`, so the two are indistinguishable to the caller. |
| `LST001` | WARN | `Generic.List[` combined with `@(` on the same variable. 5.1's array binder throws on these, including when empty. Prefer `ArrayList`. |
| `INT001` | WARN | `[int](` containing `/`. The cast rounds half-to-even; a unit conversion needs `[Math]::Floor`. |
| `PS7001` | ERROR | Use of `$IsWindows` / `$IsMacOS` / `$IsLinux` without a `$PSVersionTable` guard first. These are `$null` in 5.1, so the branch is silently false. |
| `PS7002` | **ERROR** | **Any PowerShell 7-only construct.** See below — this is the most important rule in the file when the build environment is PowerShell 7. |
| `PS7003` | ERROR | `ConvertTo-Json` without an explicit `-Depth`. 5.1 defaults to 2 and silently serialises anything deeper as `System.Object[]`. |
| `SPL001` | WARN | `@(` immediately following a call operator where splatting is intended — an array subexpression passes the whole dictionary as one positional argument. |

### `PS7002` — the rule that exists because the build environment lies

If the toolkit is written on PowerShell 7 (including any Linux agent sandbox)
and runs on Windows PowerShell 5.1, **the build environment's parser is more
permissive than the target's.** Seven-only syntax parses clean where it is
written and fails to load where it runs — so `SYN001` cannot catch it, because
`SYN001` uses the local parser.

This rule is the substitute for a 5.1 parser. Reject, as ERROR:

| Construct | Pattern to detect |
|---|---|
| Null-coalescing | `??` and `??=` outside strings and comments |
| Ternary | `?` … `:` in expression position |
| Pipeline chain operators | `&&` and `\|\|` |
| Null-conditional access | `?.` and `?[` |
| `ForEach-Object -Parallel` | the parameter, with or without `-ThrottleLimit` |
| `Sort-Object -Stable` | |
| `ConvertFrom-Json -AsHashtable` or `-Depth` | `-Depth` on **ConvertFrom** is 6.2+ |
| `Join-String`, `Test-Json`, `Get-Error` | cmdlet names |
| `Split-Path -LeafBase` | |
| `$PSStyle` | |

Detect against comment- and string-stripped source, or `&&` inside a message
string will produce noise and the rule will be switched off — which is the
failure mode described in the tenant-neutrality notes above.

**Demonstrate this rule firing before trusting any clean run.** Write
`$x = $a ?? $b` into a scratch file, confirm the rule reports it, delete it. A
check never shown to fail has not been tested, and this is the one whose silent
failure costs the whole delivery.

### WinForms

| Code | Severity | Check |
|---|---|---|
| `WF001` | ERROR | Every `Add_<Event>({` handler body is wrapped in try/catch. **Verify the property, not the edit**: walk brace depth from the handler's opening position and assert matched try/catch counts, a correct terminator, and that no handler contains another. Do not locate the end by searching for a line-initial `})` — that is a guess about formatting and it has failed before. |
| `WF002` | ERROR | Two controls added to the same parent at identical `Location`. Both draw; one is invisible underneath. |
| `WF003` | WARN | `BackColor` set on a Button without `FlatStyle = 'Flat'` — the colour is ignored. |
| `WF004` | WARN | `Screen.PrimaryScreen.Bounds` used for placement. Should be `WorkingArea`, or bottom-aligned windows go under the taskbar. |
| `WF005` | WARN | A multiline output TextBox without `WordWrap = $true`. A diagnosis off the right edge is a diagnosis nobody reads. |
| `WF006` | WARN | `.GetNewClosure()` on a handler built in a loop. Use the control's `.Tag`. |
| `WF007` | ERROR | A tool form without `Set-FormScreenCorner -Corner 'BottomLeft'`. Only the console may use `TopLeft`; anything else covers it and hides the connection indicators. |
| `WF008` | ERROR | A tool without `Add-ToolHelpButton`, or with the call but no `TOOLKIT-HELP-BEGIN` block in its header. Both halves are needed -- the button reads the header, so one without the other shows an empty box. |
| `WF009` | WARN | `Add-ToolHelpButton` called BEFORE `Set-ConsoleTheme`, or `Add-ToolTargetBar` called AFTER any control it would shift. See B3b -- the order is fixed and both mistakes are invisible until the form is drawn. |
| `WF010` | ERROR | **Console height.** The total height implied by `Sections.json` — rows × button height, plus section headers, lookup controls and the output box — exceeds `WorkingArea` on a 1920×1080 screen. Distinct from `WF004`, which is about *which API* you call; this is about whether the result fits. Irrelevant at fifteen buttons and load-bearing at fifty-five. See `09` §2. |
| `WF011` | ERROR | **Section over-capacity.** A section in `Tools.json` holds more buttons than `Rows × ButtonsPerRow`, **counting a `Group` as one button**. The surplus entry is valid, the file parses, and the button is simply absent. See `09` §3. |

`WF002` must also account for controls that compute their position at run time —
those have no literal `Location` and are exactly the ones most likely to
collide. Feed their computed positions in explicitly rather than reporting a
clean run you cannot support.

### Exchange safety

| Code | Severity | Check |
|---|---|---|
| `EXO001` | ERROR | An enumerating Exchange cmdlet call with no `ResultSize`. The default is 1000 and the short answer is silent. |
| `EXO002` | WARN | A call passing a small literal `ResultSize` without `-WarningAction SilentlyContinue`. It will warn on every run and train people to ignore warnings. |
| `EXO003` | ERROR | A recipient-type test using `-match` or `-like` against `'mailbox'`. `GroupMailbox` contains the substring and is not a mailbox. Match exactly. |
| `EXO004` | ERROR | `Set-Mailbox`, `Set-MailUser` or `Set-MailContact` with `-PrimarySmtpAddress`. On-premises only; on `Set-MailUser` it is accepted and silently inert. |
| `EXO005` | ERROR | `-EmailAddresses` assigned a bare list without evidence the existing collection was read first. A bare assignment replaces everything, losing X500 and breaking replies to historical mail. |
| `EXO006` | WARN | A write cmdlet whose `-Identity` argument came from a previously returned object's `.Identity`. Bind on `Guid`. |
| `EXO007` | ERROR | An on-premises write path with no in-process cloud-session check. See constraint C2. |
| `FLT001` | ERROR | A `-Filter` or `-LDAPFilter` string interpolating a variable that did not come from `ConvertTo-OpathValue` or `ConvertTo-LdapFilterValue`. See C19a -- an unescaped apostrophe yields a query that is wrong rather than one that fails. |
| `GRF001` | **ERROR** | Any use of Microsoft Graph: `Connect-MgGraph`, any `*-Mg*` cmdlet, `Import-Module Microsoft.Graph*`, or a literal `graph.microsoft.com` URL. See constraint C20 — the permissions need admin consent that may not exist, and Graph authenticates successfully without it before refusing every call. Allow the string inside comments, since the constraint is discussed there. |

### Tenant neutrality

| Code | Severity | Check |
|---|---|---|
| `TEN001` | ERROR | Anything shaped like an internal FQDN, a vanity SharePoint host, a `*.mail.onmicrosoft.com` address, or a ticketing host, in a `.ps1` file. |
| `TEN002` | ERROR | Any value named in `Tenant.json` — tenant name, domains, directory Ids, licence group names — appearing in a `.ps1` file. |
| `TEN003` | ERROR | `Tenant.json` parsed but named nothing, meaning `TEN002` checked nothing this run. |

Three implementation notes that were each learned by getting them wrong:

- **Match on word boundaries, not substrings.** A short tenant name inside a
  longer ordinary word produces dozens of findings on correct code, and a rule
  that cries wolf gets switched off.
- **Skip matches that are variable interpolation.** A template that *builds* a
  host from a configured value is the correct answer, not a leak. A rule that
  flags the fix is a rule people disable.
- **Document-example domains are fine.** `contoso.com` names nobody. Flagging it
  teaches people to stop writing examples.

### Audit and logging

| Code | Severity | Check |
|---|---|---|
| `LOG001` | ERROR | A script containing a write cmdlet but no audit-log call. |
| `LOG002` | ERROR | An audit call whose target is a display name rather than a primary SMTP address. With several mailboxes in play, a display name does not identify what changed. **It must NOT flag a literal string** — `-Target 'SharePoint Online'`, `'Toolkit'`, a report path. Those are the non-object targets, where there is no address to prefer and the literal *is* the identifier. The correct call shape is `-Target (Get-LogTarget $x.PrimarySmtpAddress $x.Name)`, which falls back deliberately. |
| `LOG003` | WARN | A `catch` block around a write that does not log a failure row. An error the operator dismissed otherwise leaves no trace. |
| `LOG004` | WARN | An audit call whose `RunId` comes from a fresh `New-Guid` at the call site rather than `Get-ToolRunId`. RunId must be stable across the rows of one operation; a per-row id looks populated and correlates nothing. |

### Structural hygiene

| Code | Severity | Check |
|---|---|---|
| `FUN001` | ERROR | A function redefined locally **in a file that dot-sources the shared module**. A private copy is a fork nobody reviews, and every drift found in practice was a fix silently reverted. |
| `FUN002` | INFO | The same, in a file that does *not* load the shared module. Legitimate — but a second place to maintain that will drift. |
| `REF001` | ERROR | A registered tool whose script file does not exist. |
| `CMD001` | WARN | A call to a function defined nowhere, resolvable neither locally, nor on the allowlist, nor by `Get-Command`. |

`FUN001` must test **dot-sourcing, not the name**. Its first version skipped
that test and produced twelve findings of which two were real — and acting on
the other ten would have destroyed a working standalone tool.

`CMD001` needs two passes, because a string is both hiding place and false
friend: the inside of every `$( )`, which is code even inside a string; and the
line with string literals removed, which is ordinary code. Scanning raw lines
instead flags ordinary prose and gets the rule switched off.

---

## How to know the verifier works

**A check that has never been shown to fail on a real defect has not been
tested.** For each rule, deliberately re-inject the defect it targets, confirm
the rule fires, then revert. Do this before trusting a clean run.

The corollary matters just as much: **beware the check that can only pass.** A
verification that inspects the thing just written, in the shape it was written,
will agree with itself. The honest question for any new rule is *which defect
would make this fail?*

---

## Reporting

- Findings as objects with `Severity`, `File`, `Line`, `Rule`, `Message`.
- Exit non-zero on any `ERROR`.
- **Report what was checked, not only what failed.** A rule that silently
  checked nothing produces a clean run that reads as proof — which is how
  `TEN003` came to exist.
- Say nothing on the normal path beyond the summary. A diagnostic you need a
  memorised baseline to interpret is one you skim past for six months and then
  miss anyway.

<br>

==============================================================================

# ◆ DOCUMENT 5 OF 10 — `08-COMMON-CONTRACT.md`

> Fixed function names every tool calls, plus a size budget.
>
> *Standalone document. Referred to elsewhere as `08`.*

==============================================================================

# COMMON.PS1 CONTRACT — the names every tool depends on

**These names are fixed.** Every tool in `05-TOOL-SPECS.md` and the skeleton in
`06` calls them. If `Common.ps1` exports different names, or the same names with
different shapes, nothing fits together — and because tools are written one at a
time, the mismatch appears late and everywhere at once.

Rule `CMD001` in the verifier enforces this: a call to a function defined
nowhere is an error. Run it after every tool.

The reference `Common.ps1` implements all of these. Take the implementations
from there; this document exists so you know **which** of its 164 functions you
actually need, and in what order.

---

## Size budget — read this before opening the reference

The reference `Common.ps1` is **7,200 lines and 164 functions**, grown over
months against a tenant with more moving parts than this one.

**Do not port all of it.** A shared layer that takes the whole budget and leaves
no tools is the most likely way this build fails while looking productive. The
Phase 1 target is roughly **1,000–1,600 lines** covering the Tier 1 functions
below. Everything else is added when a tool actually needs it.

### THIS LIST IS A FLOOR, NOT A CEILING

**It is not complete and cannot be.** A count of the reference files shows
**33 further `Common.ps1` functions** called by tools that are in the build
scope — licence formatting, filter escaping, console geometry, directory
resolution — none of which are named below. Tier 1d lists the ones worth knowing
about in advance; there will be others.

So the rule is a process, not a list:

1. Port Tier 1 and build.
2. **Run `CMD001` after every file.** It reports any call to a function defined
   nowhere.
3. **Every `CMD001` finding is a function to take from the reference** — not a
   prompt to write your own. The reference implementation encodes a constraint;
   yours will not, and you will not know which one until it fails in production.
4. If the function genuinely is not in the reference either, say so before
   writing it, and record it in `UNVERIFIED.md`.

That loop is what makes an incomplete contract safe. Treating the list as
exhaustive is what makes it dangerous.

This resolves the apparent contradiction in `04-REFERENCE-FILES.md`: *"lift
`Common.ps1` largely intact"* means **do not rewrite these functions from
scratch** — their bodies encode constraints you cannot re-derive. It does not
mean copy all 7,200 lines. Take the ones listed here, verbatim where they fit,
and leave the rest in the reference folder.

---

## Tier 1 — required before any tool (Phase 1)

### Config

```powershell
Get-TenantConfig [-Refresh]        # cached process-wide; NEVER throws.
                                   # Missing/invalid file returns an object with
                                   # every capability false and .Problem set.
Get-ToolkitSettings                # engineer preferences (paths, profiles)
Get-TenantDirectory [-Id <string>] # ALL directories, or one by Id
Get-DefaultTenantDirectory         # IsDefault, else first, else $null
Test-TenantCapability -Name <string>   # 'ExchangeOnline' | 'OnPremAD' | ...
```

> **`Get-Prop` is NOT in the reference `Common.ps1`.** It lives in
> `TenantManager.ps1`, which is not shipped. You need something like it —
> `ConvertFrom-Json` returns a `PSCustomObject` and a missing property throws on
> strict access — so write a three-line helper and put it in `Common.ps1` where
> everything can reach it. Mentioned explicitly because an earlier draft of this
> contract listed it as if it already existed, which is exactly the kind of
> error `CMD001` is for.

`Get-TenantConfig` never throwing is load-bearing, not defensive style. **A
toolkit that will not start because a config file is missing cannot tell you
what is wrong; one that starts and reports every capability off can.**

### Connection state

```powershell
Test-ExoConnectionState    # -> [PSCustomObject] @{ Connected; Detail; Organization }
Test-SpoConnectionState    # same shape
Test-EmsAvailable          # on-premises Exchange cmdlets present AND answering
Test-AdAvailable           # AD module usable
Test-ExoInProcess          # is a CLOUD session live in THIS process?
```

`Test-ExoInProcess` is the one people leave out and it is the important one.
Any tool that **writes on-premises** must check it and **refuse** — not warn —
when a cloud session is also present. Both editions export identically-named
cmdlets and the last import owns the name, so a tool can truthfully report
"connected to on-premises" and send every call to the cloud (constraint C2).

### Audit log

```powershell
Write-CompletedLog -Action <string> -Target <string> `
                   -Result <'Success'|'Partial'|'Failed'> `
                   [-Detail <string>] [-Environment <string>]

Write-FailedLog    -Action <string> -Target <string> `
                   [-Detail <string>] [-ErrorRecord <ErrorRecord>] `
                   [-Environment <string>]
```

`-Target` is **a primary SMTP address whenever the tool resolved a mail-enabled
object**, never a display name. With several mailboxes in play a display name
does not identify what changed.

**Where the tool has no such object** — connection scripts, config editors, the
publish chain, report writers — pass a **stable literal scope name**:
`'SharePoint Online'`, `'Toolkit'`, `'Config'`, the report path.
`reference/Connect-ExO.ps1` is the pattern. **Do not fabricate an address to
satisfy the rule.** The reference verifier already exempts literal targets for
exactly this reason — see `03`'s target rule and `09` §4.

Plus three the log cannot work without:

```powershell
Get-ToolRunId                 # -> a short id, STABLE per tool per process
Get-CompletedLogPath          # the file to append to, or $null if the caller must warn
Get-LogEnvironmentLabel       # fallback only; callers that know pass -Environment
```

**`RunId` groups the rows of ONE operation, and this is the part that gets
implemented wrong.** It is not a per-row identifier. A tool that creates a
resource and then applies settings writes *two* rows; without a shared id there
is nothing tying them together, or tying either to the failure row when a third
step dies. Generate it once per tool per process and cache it — a fresh GUID
per row is worse than an empty column, because it looks populated.

**`Get-CompletedLogPath` returns `$null` rather than guessing.** A v1-era file
is never rewritten or appended to: `Export-Csv -Append` throws on a header
mismatch in 5.1, and silently widening somebody's existing audit file is not a
decision a logging function makes on its own. Newer rows go to a sibling file
**named for the schema version** until the old one is moved aside.

`Write-FailedLog` takes the optional context parameters deliberately: without
them the failure row landed with a blank tenant and no run id while the success
row beside it had both — so the two rows you most need to correlate were the
ones you could not.

**Log failures, not just successes.** A log that records only what worked cannot
be used to reconstruct a bad afternoon, which is the one job an audit trail
has.

### Errors

```powershell
Show-HandlerError  -ErrorRecord <ErrorRecord> [-Context <string>]
Format-ErrorDetail -ErrorRecord <ErrorRecord> [-Context <string>]  # -> string
```

`Show-HandlerError` converts a .NET exception into message / script / line /
statement and writes a `Failed` audit row. Without it the operator gets a CLR
crash dialog with **no script name and no line number** (constraint B1).

### Form furniture

```powershell
Set-ConsoleTheme     -Form <Form>
Set-FormScreenCorner -Form <Form> [-Corner <'TopLeft'|'BottomLeft'>]
Add-ToolHelpButton   -Form <Form> -ScriptPath <string> [-Title <string>]
Test-ConsoleLaunch                  # -> bool; was this started by the console?
```

`Set-FormScreenCorner` must shrink an over-tall form **before** pinning it, or a
form taller than the working area loses its title bar off-screen and cannot be
dragged. It uses `WorkingArea`, never `Bounds` — `Bounds` ignores the taskbar.

`Test-ConsoleLaunch` is a courtesy, **not a security control, and must never be
described as one.** The script is plain text in a writable folder. Its message
says *unsupported*, not *unauthorized* — the latter is a claim about identity,
and this only reads an environment variable.

---

## Tier 2 — add when the first tools need them (Phase 3)

```powershell
Add-ToolTargetBar -Form <Form> ...   # the standard address/lookup bar
Resolve-ToolTarget [-SmtpAddress] [-Environment] [-ToolTitle]
                                     # resolve an object, prompt if not supplied
Get-TargetSmtp -Text <string>        # normalise typed input to an address
Get-CmdletParameterNames -Name <string>   # ask the installed cmdlet what it declares
Invoke-EXOCommand -Name <string> -Parameters <hashtable>
                                     # force resolution to the cloud module (C1)
Get-LogTarget ...                    # pick the best SMTP for an audit Target
```

`Get-CmdletParameterNames` and `Invoke-EXOCommand` are not conveniences. The
first is how you avoid passing a parameter the installed module does not declare
(which fails the whole call, C13); the second is the only reliable defence
against the on-prem/cloud collision (C1).

**A cloud-only object in a hybrid tenant is not an error.** `Resolve-ToolTarget`
must treat "found in the cloud, absent on-premises" as a valid answer.

---

## Tier 1b — PIM, if the tenant uses it

**Required when `Capabilities.Pim` is true, which it is for this tenant.**

```powershell
Get-TenantPimSetting -Name <string>     # from Tenant.json
Get-PimDurationHours                    # tenant policy, NOT a constant
Get-PimStatePath                        # where the activation time is recorded
Get-PimActivation / Set-PimActivated / Clear-PimActivation
Get-PimRemaining                        # -> TimeSpan
Format-PimSpan -Span <TimeSpan>         # -> "3h 42m"
Get-PimStatusLine / Get-PimStatusColour # what the console header shows
```

**DO NOT TRIM THESE OUT.** They are already implemented in the reference
`Common.ps1`, and `AdminConsole.ps1` calls them for the countdown in its header.
Removing them while keeping the console leaves the header calling functions that
do not exist — and the size budget above makes that an easy mistake to make.

Two things the reference implementation gets right and a rewrite would not:

- **`Format-PimSpan` must `[Math]::Floor` the quotient.** PowerShell's `[int]`
  cast **rounds, half to even** (constraint A4), so 3h59m renders as "4h 59m" —
  which reads as a five-hour policy on a four-hour one. Plausible, wrong, and in
  a display nobody cross-checks.
- **`Update-PimLabel` takes its controls as parameters.** It used to read a
  script-scoped `$lblPim`, and when the dashboard defined its own, the console's
  countdown was written into the *dashboard's* label instead. In a tenant
  without PIM that resolved to `$null` and threw; **in a tenant with PIM it did
  not throw, which is worse.**

## Tier 1c — module prerequisites

**Required. The console cannot honestly report connection state without these.**

```powershell
Get-ToolkitModuleDefinition [-Id <string>]  # what the toolkit needs, and from where
Test-ToolkitModule -Id <string>             # installed? importable? which version?
Import-ToolkitModule -Id <string>
Install-ToolkitModule -Id <string>
Request-ToolkitModuleInstall -Id <string>   # ASK the operator before installing
Get-ToolkitModuleStateMap                   # everything at once, for the console
```

The two that matter are **`ExchangeOnlineManagement`** and
**`Microsoft.Online.SharePoint.PowerShell`**. The AD module arrives with RSAT and
is either present or not.

**Why this is Tier 1 and not a nicety:** without it, *"not connected"* and *"the
module is not installed"* look identical to the operator — same red indicator,
same inability to do anything. One is fixed by clicking Connect and the other
never will be, and the console is the only thing positioned to tell them apart.
A first run on a fresh machine is exactly where this bites, and it is also
exactly what you cannot test from a sandbox.

**Never install silently.** `Request-ToolkitModuleInstall` asks first. Installing
a module changes the operator's machine, may require elevation, and may be
against policy. Offer, explain, and let them decline.

### The tool help mechanism

`Add-ToolHelpButton -ScriptPath` does **not** take help text as a parameter. It
reads the script's own header, between two markers:

```powershell
# TOOLKIT-HELP-RUNS: one line on what this tool touches. Optional.
# TOOLKIT-HELP-BEGIN
# Plain language, for an operator. What question does this answer, and when is
# it the WRONG tool? The second half matters more than it looks -- somebody who
# reaches for the wrong tool concludes the data is missing.
# TOOLKIT-HELP-END
```

The text lives in the script rather than a help file for two reasons: a help
file drifts the first time somebody edits a script without editing it, and it
does not travel when the script is copied. If the markers are absent the button
says so explicitly rather than showing an empty box, which is how you notice.

---

## Tier 1d — the ones the in-scope tools actually call

Found by scanning the reference files that correspond to tools in the build
scope. Take each from the reference when the tool that needs it is built.

**Filter escaping — take these before you write any search.**

```powershell
ConvertTo-OpathValue      -Value <string>   # Exchange -Filter strings
ConvertTo-LdapFilterValue -Value <string>   # LDAP filters, RFC 4515
```

They are **not interchangeable**, and using the wrong one produces a query that
is *wrong* rather than one that fails. OPATH cares about the single quote and
escapes it by doubling; LDAP escapes backslash, parentheses, asterisk and NUL as
hex.

Every caller builds a filter around something an operator typed — a partial
name, an office. **One apostrophe in an Irish surname breaks the query.** This
is a correctness fix rather than a security one (these filters read, and the
operator already has rights to read), but a lookup that silently returns nothing
for `O'Brien` is a bug people report as "the tool cannot find anyone".

Take the reference implementations verbatim. The inline version was written
twice and **threw every time it ran**: `[char]0` in an escape table makes the
first argument a char, `String.Replace` has a `(char, char)` overload, and
PowerShell binds to it and then cannot convert a four-character replacement.

**Licence and mailbox state — Mailbox Health needs all three.**

```powershell
Get-UserLicenseState / Format-UserLicenseState   # group-based; NOT Graph (C20)
Get-SimpleMailboxType                            # the type an operator recognises
Get-InboxRulesForMailbox                         # shared with the Inbox Rules tool
```

**Directory resolution.**

```powershell
Get-TenantDirectoryServer          # a server for a directory; may return a DOMAIN
Resolve-DirectoryDomainController  # a real CONTROLLER -- Exchange rejects a domain (C8)
Find-TargetEnvironment / Resolve-Target
Invoke-OnPremCommand
```

`Get-TenantDirectoryServer` deliberately returns a domain when a directory names
several controllers. That is right for the AD module's `-Server` and **fatal for
Exchange**, which rejects a domain with error 0x52. Know which you are calling.

**Console and forms.**

```powershell
Get-ConsoleButtonWidth / Get-ConsolePanelWidth / Get-ConsoleResizeCommand
Get-CapabilityStateMap / Get-ToolAvailability     # what greys out, and why
Get-ConsoleTheme / ConvertTo-ThemeColor / Add-FormLogo
Set-FormResizable                                  # viewers only; see below
Start-UrlInBrowserProfile                          # activate.ps1 and Threat Explorer
```

`Start-UrlInBrowserProfile` is needed in **Phase 2** — `activate.ps1` opens the
PIM portal with it, and `Start-Process` alone would land in whichever browser
profile was used last, which for an admin portal is the wrong identity.

**Config and integrity.**

```powershell
Get-ConfigFilePath / Get-LogPath / Get-MailDomain
Get-ConnectedTenantDomain / Clear-ConnectedTenantDomainCache
Test-TenantDomainMismatch      # amber warning, never blocking
Test-ToolkitIntegrity / Save-ScriptCapabilityCache
Invoke-ConfigMigration
```

`Test-TenantDomainMismatch` is deliberately loose and non-blocking. A tenant can
legitimately have several accepted domains, so this is a question worth asking,
not a fault worth stopping on — but connected to the *wrong tenant*, the label
otherwise reads perfectly correct.

### One that the reference calls and you should NOT port

> **This list was computed against the eight-tool first pass and is not valid
> beyond it.** Two functions that were on it have since moved into scope, and
> the row for each now says when. Recompute the list yourself at the start of
> each stage in `09` — `CMD001` is the mechanism.

Deliberately excluded, and knowing why is worth more than the name:

| Function | Why not |
|---|---|
| ~~`Add-ReportFileControls`~~ **— NOW IN SCOPE** | The shared report-save controls. Not needed for the first pass. **Port it at Stage D**, with the first report; Transport Rules (Stage A) and Cert Check (Stage H) also call it. See `09-FULL-TOOL-CATALOGUE.md`. |
| ~~`Resolve-AdObjectType`~~ **— IN SCOPE, FIRST PASS** | **Port it.** An earlier draft of this document called it dual-directory machinery; that was wrong. It classifies **one** AD object's type from its attributes, and both `Unified-Lookup.ps1` (a first-pass tool) and `AD-Lookup.ps1` call it. Omitting it produces a `CMD001` on tool 3 of 8. |
| `Start-ToolInOnPremProcess` | **The only genuine exclusion.** Launches a tool in its own process **with directory routing attached**. The routing belongs with a second directory, which this tenant does not have. |

**Do not confuse that with own-process isolation, which IS required.** Isolation
is the console's own `Test-ToolDeclaresOwnProcess` + `Start-ToolScriptInNewProcess`
pair, driven by each script's `# TOOLKIT-REQUIRES: OwnProcess` header, and it
defends against the cmdlet collision (C1) rather than against a second
directory. **Three tools declare it: `Group-Manager-Editor.ps1` in the first
pass (`05` §8), and `ManageMailObject.ps1` and `Bulk-DLMembership.ps1` at Stage
F (`09`).**

If you find yourself wanting `Start-ToolInOnPremProcess` specifically, stop —
something has been assumed about the tenant shape that this build does not have.

---

## Tier 3 — later, and only if scoped in

Report file helpers, theme *editing* (the ThemeManager UI — the palette itself
is Tier 1, taken as-is from `Themes.TEMPLATE.json`), capability caching. All
present in the reference; **none needed for the first pass, all needed by the
end of Stage H** — see `09-FULL-TOOL-CATALOGUE.md`. Report file helpers arrive
at Stage D, theme editing at Stage H.

> **`Start-UrlInBrowserProfile` is NOT Tier 3.** It is needed in **Phase 2** —
> see the Tier 1b note above. Quick Links (Stage G) and M365 Reports (Stage D)
> also use it, but it must exist long before either.

---

## Rules for anything you add

- **One definition, in `Common.ps1`.** A private copy in a tool is a fork nobody
  reviews. In the source toolkit, five tools carried 13 local copies of shared
  functions and **seven had drifted — every drift a fix silently reverted.** The
  console was running the pre-fix version of a function it dot-sourced, and
  nothing said so.
- **A tool that deliberately runs standalone may carry its own copy** — that is
  legitimate, and the verifier distinguishes the two by whether the file
  dot-sources `Common.ps1` at all, not by the name.
- **State shared with launched tools must be `$global:`, never `$script:`.** A
  tool invoked with `&` runs in its own script scope; a dot-sourced function
  remains callable but any `$script:` variable it reads resolves against the
  *child's* scope, where it was never assigned. The function works and its state
  is `$null` — and a `Get-Command` guard proves the function exists while
  proving nothing about whether its state does.
- **Paths derived inside a function come from `$PSScriptRoot`**, which resolves
  to the file that *defined* the function regardless of caller.

<br>

==============================================================================

# ◆ DOCUMENT 6 OF 10 — `04-REFERENCE-FILES.md`

> How to use the reference files.
>
> *Standalone document. Referred to elsewhere as `04`.*

==============================================================================

# REFERENCE FILES — what ships, and how to use them

**Nineteen** working scripts and five config files from the original toolkit
ship in `reference/`, delivered across two concatenated files and extracted in
Step 0. They are **reference implementations, not a starting codebase.** That
distinction is the most important thing in this document and is spelled out
below.

- `DEVIN-REFERENCE-CODE.md` — 11 scripts + 5 config. Infrastructure and core
  tools. Described in the two tables immediately below.
- `DEVIN-REFERENCE-CODE-2.md` — 8 scripts. **An exemplar for every remaining
  tool in the build scope**, so that no tool has to be written from prose alone.
  Listed under "Part 2" further down.

---

## The core five

Take these first. If only five files were available, these are the five.

| File | ~lines | The reference for |
|---|---|---|
| `Common.ps1` | 7,200 | **The spine.** Config readers, connection probes, capability gating, audit log, error display, theming, geometry. Every convention in `02` is realised here. |
| `Test-Toolkit.ps1` | 1,950 | The verifier from `03`, already built. Run and extend it — do not reimplement it from the spec. |
| `AdminConsole.ps1` | 2,230 | Launcher: connection state, capability-driven greying, JSON-driven layout, the modal-loop deferral (B2). |
| `Unified-Lookup.ps1` | 2,760 | The most complex read tool. Two directories searched in one pass, results classified, routing to the right editor, a capped result set filtered. |
| `Group-Manager-Editor.ps1` | 1,250 | **The write reference.** Directory resolution before a write, identity binding, the on-prem/cloud refusal guard, audit rows on success and failure. |

## The recommended six

Added in value order. Cut from the bottom if the set has to shrink.

| # | File | ~lines | Why it earns a slot |
|---|---|---|---|
| 6 | `MailboxHealth.ps1` | 373 | **The canonical small tool.** All five above are large; none shows the whole shape at a glance. This one does, and it is tool #1 in the build order — an exact target rather than a description. |
| 7 | `Connect-ExO.ps1` | 251 | Encodes **B2** (a modal loop blocks any window a tool raises) and **C17** (connections are per-process). Two of the least intuitive constraints in the pack, in one of the smallest files. |
| 8 | `Connect-ExSE.ps1` | 203 | The on-premises half: session setup, `-ViewEntireForest` asserted rather than assumed (C6), and why it avoids `Get-Credential` entirely. |
| 9 | `ToolManager.ps1` | 948 | Config round-trip. Without it you will invent config editing, and **silently dropping fields the editor does not understand is a documented past bug** — it once deleted a live setting on first save. |
| 10 | `Protocols.ps1` | 404 | The tool that taught **C15** (slice a large query; never retry an oversized one) and **C13** (ask the cmdlet what it declares). The best single artefact on bulk queries. |
| 11 | `MessageTrace.ps1` | 688 | A clean mid-size tool: grid, filter-without-re-running, export, and an explicit statement of what the tool is *not* for. |

## Part 2 — one exemplar per remaining tool

These arrive in `DEVIN-REFERENCE-CODE-2.md`. They exist because after part 1,
five of the eight tools in the build scope had no working example, while three
did — and that asymmetry is exactly where conventions drift.

| File | ~lines | The reference for |
|---|---|---|
| `activate.ps1` | 326 | **Tool 0, built with the console in Phase 2.** Rights here are eligible, not standing, so nothing else works until it has been used. Read its ORDER MATTERS header. |
| `MailboxPermissions.ps1` | 347 | Tool 4. Full Access and Send As are separate grants and constantly confused. |
| `CalendarPermissions.ps1` | 246 | Tool 5. Folder permissions; the calendar folder name is localised. |
| `OOFEditor.ps1` | 240 | Tool 6. Two messages, internal and external, usually only one wrong. |
| `InboxRules.ps1` | 320 | Tool 7. Delete is irreversible; disable is not. |
| `Group-Manager.ps1` | 732 | Tool 8's **front door** — classifies, then routes to the editor in part 1. |
| `MailAttributes.ps1` | 739 | **Constraint C11 in real code**: the proxy-address rewrite that destroys X500 if done naively. The most damaging mistake available here. |
| `ExternalForwarding.ps1` | 405 | A complete CSV report: shared report controls, partial-run labelling, never a zero-row file. |

**Every tool in the first pass now has a working example.** If you want another
reference beyond these, ask first — the answer is usually that the pattern already
exists in one of these.

---

Also in `reference/Config/`: `Tenant.TEMPLATE.json`, `Settings.json`,
`Sections.json`, `Tools.json`, `Themes.TEMPLATE.json`. The config split is half the architecture and is
hard to infer from prose.

> `Tools.json` is a reference for **schema first, contents second**. It
> registers tools from the source toolkit, most of which are **deferred out of
> the first pass rather than out of scope** — they are specified in
> `09-FULL-TOOL-CATALOGUE.md` and are built in Stages A–H. For the first pass,
> copy the shape and write your own eight entries. From Stage A onward, `09` §2
> gives the destination sections and `09` §3 the `Group` field this file's copy
> does not exercise.

---

## How to use them

**These are the reference for CONVENTIONS, not a codebase to extend.**

The toolkit they come from served a different tenant with a different shape.
Copying a tool wholesale imports assumptions that stay invisible until they are
wrong — which is the failure mode this entire pack exists to prevent.

### Do

- **Match the conventions exactly.** Naming, parameter shapes, how a tool
  dot-sources `Common.ps1`, how it logs, how it reports partial results, how it
  lays out a form. Consistency across the toolkit is worth more than any
  individual improvement.
- **Take `Common.ps1` functions verbatim — but only the ones you need.** Their
  bodies encode constraints you cannot re-derive, so do not rewrite them from
  scratch. But the file is **7,200 lines and 164 functions**, and porting all of
  it would consume the build.
- **`08-COMMON-CONTRACT.md` is a FLOOR, not a list of everything you need.** It
  names the minimum. There are at least 33 further `Common.ps1` functions called
  by in-scope tools that it does not name. The discovery mechanism is `CMD001`:
  build, run it, and **every unresolved call is a function to take from the
  reference** rather than one to write yourself.
- **Run `Test-Toolkit.ps1` unchanged, first**, then extend it — starting with
  the `PS7002` rule from `03`, which the original does not have because it was
  written on the target platform.
- **Read the comments.** They are dense and load-bearing. Where a comment
  explains why something is done a surprising way, that explanation usually cost
  somebody a day. **Treat a comment that contradicts your instinct as
  information, not noise.**

### Do not

- **Do not copy a tool and adjust it.** Write each new tool from its spec in
  `05`, consulting the reference for *how* rather than *what*.
- **Do not carry over multi-directory machinery.** The source tenant had two AD
  domains and much of the directory-resolution code exists only to disambiguate
  them. This tenant has **one**. Keep the config-driven shape — `Directories`
  stays a list — but do not port cross-directory pinning, per-directory
  controller selection, or side-routing. Dead branches here.
- **Do not port anything you cannot explain.** If a block's purpose is not clear
  from the code and its comment, leave it out and note the omission. Unexplained
  code that is carried forward becomes permanent.
- **Do not assume the reference is bug-free.** It is working production code,
  not scripture. If something looks wrong, say so — several of the constraints
  in `01` exist because a previous version of these very files had the bug.

### Expect to find

Comments describing bugs that were fixed, sometimes at length, sometimes dated.
These are deliberate: they exist to stop the bug being reintroduced by the next
person to edit that line. Preserve the practice — short guardrail at the line,
longer reasoning in notes.

You will also find comments admitting a known limitation or an open question.
Those are honest, not stale. Do not "fix" them by deleting them.

---

## The one thing to do first

Run `Test-Toolkit.ps1` against `reference/` and read its output. In about a
minute that tells you which rules exist, what they are called, what a clean run
looks like, and what the report format is.

It is also the fastest way to learn what this toolkit considers a defect, which
is most of what `01` is trying to convey in prose.

<br>

==============================================================================

# ◆ DOCUMENT 7 OF 10 — `05-TOOL-SPECS.md`

> Per-tool specs for the first pass, in build order, each with an exemplar and a done-when.
>
> *Standalone document. Referred to elsewhere as `05`.*

==============================================================================

# TOOL SPECS — build order, with the traps that apply to each

Build in this order. Read-only tools first, so every write tool stands on a read
path already proven against the live tenant.

Codes in the **Traps** rows refer to `01-HARD-CONSTRAINTS.md`. They are the
specific ones that have actually bitten these tools — not a general reminder to
be careful.

**Every tool below has a working exemplar in `reference/`**, named at the top of
its section. Read the exemplar first, then write the tool from the spec. The
exemplar answers *how* — layout, guards, logging, error handling; the spec
answers *what*. Do not edit a copy of the exemplar: the source tenant had two AD
domains where this one has one, and copying carries assumptions that stay
invisible until they are wrong.

Every tool, without exception:

- dot-sources `Common.ps1` by upward search (never a counted `..`)
- wraps every handler (**B1**) and turns word wrap on in output boxes (**B6**)
- writes an audit row with the **primary SMTP address** as Target, never a
  display name
- opens **bottom-left** so it never covers the console (only the console uses
  top-left)
- carries a **Help button** and a filled-in `TOOLKIT-HELP` block -- no exceptions
- follows the fixed end-of-form order: target bar, theme, corner, help button
- passes `Test-Toolkit.ps1` clean before it is considered done

---

## Phase 2 — with the console

### 0. Activate (PIM)

> **Exemplar: `reference/activate.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Opens Privileged Identity Management so the operator can turn their admin
access on.*

Built with the console rather than among the tools, because **nothing else works
until it has been used.** Admin rights at this tenant are *eligible, not
standing*: until an operator activates, they are signed in with no admin access
and every other tool fails in ways that look like faults rather than like a
missing activation.

| | |
|---|---|
| **Input** | None |
| **Does** | Opens the PIM portal in the admin browser profile; records the activation time locally |
| **Writes** | Nothing in the tenant. A local state file only |
| **Traps** | **A4** — `[int]` rounds half-to-even, so the remaining-time display must `[Math]::Floor` or a 8h59m remainder renders as "9h 59m". This exact bug shipped once. **B2** — the portal opens in a browser, i.e. a separate process, so it is unaffected by the console's modal loop. |

**ORDER MATTERS, and it is the part people get wrong:**

> 1. **Activate FIRST**
> 2. **Connect SECOND** (Exchange Online, SharePoint, on-premises)

A connection made *before* activating keeps the token it was issued with, and
**that token does not contain access activated afterwards.** The symptom is a
tool that authenticates perfectly and then refuses every operation, which reads
as a permissions problem and is not one. If the operator has already connected,
they must activate and then reconnect. Say this in the tool's help text.

**The console header countdown** reads `Pim.DurationHours` from `Tenant.json` —
**9 at this tenant** — and shows remaining time via `Get-PimStatusLine` and
`Get-PimStatusColour` (contract Tier 1b). It is not a timeout the toolkit
enforces; it is a display of how long Entra granted. A wrong duration produces a
confidently wrong countdown, which is worse than none: an operator who trusts it
starts a long job believing they have hours in hand.

**Done when:** the portal opens in the admin browser profile (not whichever
profile was last used), the activation time is recorded locally, and the console
header counts down from **9 hours**. Check the countdown against a clock after a
few minutes -- a rounding error here reads as a longer policy than you have.

**Out of scope: programmatic self-activation.** PIM for Groups is a Graph-only
API needing admin-consented delegated permissions
(`PrivilegedEligibilitySchedule.Read.AzureADGroup` and
`PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup`). Unless that consent
exists, open the portal. Do not attempt a Graph implementation without
confirming consent first — and if you do confirm it, ask before building.

---

## Phase 3, then Phase 4a — read-only

> **Tool 1 alone is Phase 3, and it ends at a hard stop** — see
> `00-BUILD-PROMPT.md` §6. Tools 2 and 3 below are the start of **Phase 4** and
> must not be written until the operator has run tool 1 on Windows.

### 1. Mailbox Health

> **Exemplar: `reference/MailboxHealth.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*The first place to look when someone says "my email is doing something
strange".*

One mailbox, one screen, read-only. Answers in a single view what would
otherwise be six commands: size and quota, mailbox type, forwarding, delegates
and permissions, auto-reply state, litigation hold, archive status, and whether
the account is licensed.

| | |
|---|---|
| **Input** | One SMTP address or alias |
| **Reads** | `Get-Mailbox`, `Get-MailboxStatistics`, `Get-CASMailbox`, `Get-MailboxPermission`, `Get-MailboxAutoReplyConfiguration` |
| **Writes** | Nothing |
| **Traps** | **C9** — archive state comes from `ArchiveGuid`, never `ArchiveStatus`, which lags replication and reports `None` on an archive that exists. **C14** — `EwsEnabled` is often `$null` meaning *inherit*; resolve against `Get-OrganizationConfig` and show the effective value. **C12** — do not classify recipient type by substring. |

Build this first even though it is not the most useful tool, because it
exercises the whole read stack — resolution, connection gating, form layout,
audit logging of a read — with nothing at risk.

**Done when:** every field is either populated or explicitly says why it is not.
A blank field that means "not applicable" and a blank field that means "the
lookup failed" must not look the same.

---

### 2. Message Trace

> **Exemplar: `reference/MessageTrace.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Follow a message through Exchange Online — where it went, whether it was
delivered, what happened if it was not.*

| | |
|---|---|
| **Input** | Sender, recipient, subject, date range, status filter |
| **Reads** | `Get-MessageTrace`, `Get-MessageTraceDetail`, quarantine cmdlets |
| **Writes** | Quarantine release only |
| **Traps** | **C3** — the trace enumerates; pass `ResultSize = 'Unlimited'`. **A4** — any latency arithmetic must `[Math]::Floor`, or 119 seconds renders as "2m 59s". **D6** — never write a zero-row export. |

Two design points that matter more than they look:

- **Filter the results without re-running the trace.** A search box over the
  returned rows, not a new query. Re-running is slow and the operator has
  usually already got what they need on screen.
- **Say what the tool does not do.** Message trace answers "what happened to
  this message". It is not threat hunting — no URL verdicts, no campaign views,
  no bulk remediation. Put that in the help text and link out to the right
  place, because the alternative is somebody concluding the data is missing.

Trace data is retained for roughly ten days in the detailed view. State that in
the UI rather than returning an empty grid for an older date.

**Done when:** a trace returns rows for a message you can verify independently,
the search box narrows them **without re-running the query**, and a date outside
the retention window says so rather than returning an empty grid.

---

### 3. Unified Lookup

> **Exemplar: `reference/Unified-Lookup.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Find any object without knowing where it lives.*

The most complex read tool and the one worth the most. One search box; searches
Exchange Online **and** the on-premises directory in the same pass, reports what
each side found, classifies each result, and opens the right editor for that
combination.

| | |
|---|---|
| **Input** | A name fragment, alias, or SMTP address |
| **Reads** | `Get-Recipient`, `Get-Mailbox`, `Get-DistributionGroup`, `Get-UnifiedGroup` in the cloud; AD in the directory |
| **Writes** | Nothing directly — it routes to editors |
| **Traps** | **C12** — `GroupMailbox` contains "mailbox" and is a Microsoft 365 Group. **C3** — every one of these enumerates. **A1** — the result filter returns a possibly-empty selection; do not also use `$null` for cancelled. **C4** — carry the GUID from the search into whatever tool is opened; never re-resolve by the display Identity. |

Two behaviours worth specifying explicitly:

- **Cap the result set, and say so.** A common word matches thousands. Cap at a
  few hundred, and when the cap is hit **say the results are incomplete** —
  silently truncating leaves the operator filtering a set that was already
  wrong, which looks like a tidy complete answer.
- **A name existing in two places is not an ambiguity to resolve.** It is the
  answer to the ticket. Show both.

**Done when:** an object that exists only in the cloud, only on-premises, and in
both, each produce a correct and visibly different result.

---

## Phase 4b — low-risk writes

From here on, every tool writes. Each needs: a confirmation naming exactly what
will change, an audit row on success **and** on failure, and the guard from
**C2** if it touches on-premises.

### 4. Mailbox Permissions

> **Exemplar: `reference/MailboxPermissions.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Grant, view or remove access to somebody else's mailbox.*

Two things that are constantly confused and are granted separately:

- **Full Access** — open and read the mailbox
- **Send As** — send mail appearing to come from it

*"I can see the shared mailbox but can't send from it"* is Full Access without
Send As, and it is the most common ticket in this area. Handle both in one tool
and label them unmistakably.

| | |
|---|---|
| **Reads** | `Get-MailboxPermission`, `Get-RecipientPermission` |
| **Writes** | `Add-/Remove-MailboxPermission`, `Add-/Remove-RecipientPermission` |
| **Traps** | **C12** — a Microsoft 365 Group is not a mailbox; access to one comes from *membership* and these cmdlets do not apply. Detect and refuse with an explanation rather than failing obscurely. **C4** — bind on GUID. |

Note on AutoMapping: it can be set when granting Full Access but there is **no
documented way to read it back** in Exchange Online. Do not display a value you
cannot verify — offer it as a write-time choice and say that current state
cannot be shown.

Access granted here does not expire. Say so in the UI.

**Done when:** Full Access and Send As can each be granted, listed and removed
**independently** -- granting one must not grant the other -- and pointing the
tool at a Microsoft 365 Group produces a clear refusal explaining that access
there comes from membership, not from these cmdlets.

---

### 5. Calendar Permissions

> **Exemplar: `reference/CalendarPermissions.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*See and change who can view or edit someone's calendar.*

Answers "my assistant can't see my calendar" and "someone can see more than they
should".

| | |
|---|---|
| **Reads** | `Get-MailboxFolderPermission` on the calendar folder |
| **Writes** | `Add-/Set-/Remove-MailboxFolderPermission` |
| **Traps** | The calendar folder path is localised — resolve it rather than hardcoding `:\Calendar`. `Default` and `Anonymous` are pseudo-users that always exist and cannot be removed, only set to `None`. |

Show current permissions for everyone who has one, and offer the standard access
levels rather than raw rights strings.

**Done when:** the current list matches what Outlook shows for the same mailbox,
`Default` and `Anonymous` appear and can be set to `None` but not removed, and a
non-English calendar folder still resolves.

---

### 6. Out-of-Office Editor

> **Exemplar: `reference/OOFEditor.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Read and set someone's auto-reply without needing their mailbox open.*

For the classic ticket: somebody left, or is off sick, and their auto-reply
needs setting or clearing on their behalf.

| | |
|---|---|
| **Reads** | `Get-MailboxAutoReplyConfiguration` |
| **Writes** | `Set-MailboxAutoReplyConfiguration` |
| **Traps** | There are **two** messages — internal and external — and they are usually different, with only one of them wrong. Show both, always. The scheduled window matters: `Scheduled` with a past end date reads as "on" in some views and sends nothing. |

**Done when:** both messages are visible without scrolling or switching tabs, a
scheduled window shows its start and end, and auto-reply can be turned off on a
mailbox you do not own.

---

### 7. Inbox Rules

> **Exemplar: `reference/InboxRules.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*View, disable, enable and delete rules on a mailbox.*

The usual reasons to look: mail vanishing into a folder nobody created, mail
forwarded somewhere it should not be, or a rule left behind by a compromise.

| | |
|---|---|
| **Reads** | `Get-InboxRule` |
| **Writes** | `Enable-/Disable-/Remove-InboxRule` |
| **Traps** | **Deleting cannot be undone; disabling is reversible.** Default to disable and make delete the deliberate choice. **C4** — bind on the rule's identity carefully; rule names are not unique and are frequently blank. |

A rule forwarding externally is a security finding, not just a nuisance —
highlight it.

**Done when:** rules list with their conditions readable, **disable is the
default action and delete requires a separate deliberate confirmation**, a rule
with a blank name is still addressable, and external forwarding is visibly
flagged.

---

## Phase 4c — groups

### 8. Group Manager, and Group Manager Editor

> **Exemplar: `reference/Group-Manager.ps1` (front door) and `reference/Group-Manager-Editor.ps1` (editor).** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Start here for any group; it works out what the group is and opens the right
editor.*

Two tools. The front door classifies; the editor edits.

**The front door** looks the group up and determines two things — **what** it is
and **where** it lives:

| What | Where | Opens |
|---|---|---|
| Distribution list / mail-enabled security group | synced from on-premises | Editor, against on-premises |
| Distribution list / mail-enabled security group | cloud-native | Editor, against Exchange Online |
| Microsoft 365 Group | (always cloud) | Group editor — members and owners |

**A Microsoft 365 Group is a different animal** and that is why it gets its own
destination. Access comes from **membership**, not mailbox permissions, and it
is managed by an entirely separate cmdlet family (`*-UnifiedGroup`) that the
distribution-list editor cannot speak.

| | |
|---|---|
| **Reads** | `Get-Recipient` to classify, then `Get-DistributionGroup` / `Get-UnifiedGroup` |
| **Writes** | Membership, owners, delegation, addresses |
| **Traps** | **C12** — the classification hinges on not matching `GroupMailbox` as a mailbox. **C18** — `IsDirSynced` decides which side to write; say so *before* the attempt, since the cloud refuses most changes to a synced object. **C4** — resolve once, bind on GUID; a group's `Identity` renders as the AD CN and re-binding it is a fresh forest-wide lookup. **C10/C11** — address edits rewrite the whole proxy collection; carry X500 across or replies to historical mail break. **C1/C2** — this is the tool family most exposed to the cmdlet collision. |

> **`Group-Manager-Editor.ps1` MUST declare `# TOOLKIT-REQUIRES: OwnProcess` in
> its own header, in the first 40 lines.** Its on-premises path uses
> `Add-`/`Remove-ADPermission`, and those names exist in both Exchange editions
> — in one process the last import owns them (C1/C2). This is the only tool in
> the first pass that needs the marker, and it is the most exposed file in it.
> Two more declare it at Stage F; see `09`.
>
> **The front door has a console button; the editor does not.** The editor
> refuses to open standalone — it is reached only by routing from Group Manager,
> whose window title reads **DL Manager**, the name operators use for it.

Why the routing matters operationally: during a migration, groups move
on-premises → cloud a batch at a time, and **nothing about the group's name or
address changes when it does.** There is no way to tell by looking, which is
precisely why the tool should determine it rather than the operator guessing.

**Done when:** a synced distribution list, a cloud-native distribution list and a
Microsoft 365 Group each route to the correct editor, **and the front door says
which it found and why** before opening anything. A silent correct route and a
silent wrong one look identical.

---

## After these

**The remaining forty-eight tools are specified in
`09-FULL-TOOL-CATALOGUE.md`**, in eight lettered stages. Do not open it as a
build instruction until this first pass has been run on Windows and accepted per
`07`.

They are variations on patterns established above: attribute editors (apply
**C10** and **C11** exactly — this is where the X500 loss happens), reports
(**C3** and **C15** — slice large queries, never retry an oversized one), room
administration, and on-premises provisioning.

**A leaver / termination workflow is NOT in the catalogue** and is not
authorised work. If it is ever asked for, it is the highest-risk tool in any
toolkit of this kind: multi-step, partially reversible, and the step that fails
is usually not the step that reports. It would need a retry queue — a record of
what could not be finished, so a half-completed run is visible rather than lost.
The console's Dashboard already surfaces a retry-queue panel for exactly that
shape of tool.

<br>

==============================================================================

# ◆ DOCUMENT 8 OF 10 — `09-FULL-TOOL-CATALOGUE.md`

> The destination: every remaining tool, the finished console layout, and the stages that get there.
>
> *Standalone document. Referred to elsewhere as `09`.*

==============================================================================

# FULL TOOL CATALOGUE — every remaining tool in the finished toolkit

`05-TOOL-SPECS.md` specifies the first pass in depth — Activate plus eight
tools, ten scripts — and stops there. **This document specifies everything else:
forty-eight further scripts in eight stages, plus additions to three files
already built** (`AdminConsole.ps1`, `ToolManager.ps1`, `Test-Toolkit.ps1` — see
§5). Together they describe the whole toolkit, and the whole toolkit is the
destination.

**A note on names.** `Group-Manager.ps1` is called **Group Manager** in `00`,
`02`, `04` and `05`. Its window title, and what operators call it, is **DL
Manager**. They are the same tool. This document uses *Group Manager* for
consistency with the rest of the pack and notes the window title where it
matters.

---

## 0. READ THIS FIRST — how this document relates to the rest of the pack

**`00-BUILD-PROMPT.md` §5 defines the FIRST PASS. This document defines the
DESTINATION. They are not in conflict, and neither one cancels the other.**

Several documents in this pack say things are "out of scope" — reports,
SharePoint tools, provisioning, theme editing. **Read every one of those as "not
in the first pass."** They are all in this catalogue and all get built. The
sequencing is deliberate: the first pass proves the conventions on eight tools,
and everything here is built on top of conventions that have already been
confirmed on a real tenant.

Specifically, these read as *deferred* rather than *excluded*:

| Where | Deferred thing | Where it lands |
|---|---|---|
| `00` §5 | provisioning | Stage F |
| `00` §5 | reporting suite | Stage D |
| `00` §5 | SharePoint tools | Stage A (Connect-SPO) and Stage D |
| `00` §5 | theming beyond one palette | Stage H |
| `04` | "`Tools.json` … most of which are deferred out of the first pass" | Sections and `Group` field: §2 and §3 below |
| `08` | `Add-ReportFileControls` | Stage D, with the first report |
| `08` Tier 3 | theme editing | Stage H |

**Genuinely excluded, and NOT re-authorised by this document:** everything in §6
below, and **a leaver / termination workflow**, which is not in this catalogue,
is named as unauthorised in `05`, and would be the highest-risk tool in the
toolkit. If it is ever wanted, it is new work with its own conversation.

**`00` §7 tells you to stop and ask when you are about to build something not in
§5. That instruction still stands for anything not in THIS document.** Work
listed here is authorised; work listed nowhere is not.

**Do not begin any stage below until the first pass has been run on Windows and
accepted** per `07-ACCEPTANCE-CHECKLIST.md`. Everything here compounds the
conventions established there — including the mistakes.

### Stage numbering

This document uses **lettered stages (A–H)** deliberately, because
`00-BUILD-PROMPT.md` §6 already numbers Phases 0–5 and its **Phase 5 is
handover**. Stages A–H sit between `00`'s Phase 4 and its Phase 5. When you
report progress, say "Stage C" or "Phase 4" — never a bare number.

---

## 1. What "done" means

The finished toolkit is **55 console buttons over 65 `.ps1` files.** Every entry
exists because somebody needed it on a real ticket.

That reconciles as:

```
55 console buttons
  – 1  the "Rooms" button, which fronts 5 scripts (see §3)
  + 5  those five scripts
= 59  scripts reachable from a console button
  + 1  Group-Manager-Editor.ps1   opened BY Group Manager, no button of its own
  + 2  Common.ps1, AdminConsole.ps1        infrastructure
  + 2  Deploy.ps1, Publish-Toolkit.ps1     run from a shell, never a button
  + 1  UnBlock.ps1                shipped inside handouts; not a tool
= 65  .ps1 files
```

**Build in stage order.** Do not build breadth-first across all of them. Each
stage ends where a human can run what exists and find out whether the
conventions are holding, and a convention that has drifted across forty files is
far more expensive to correct than one that has drifted across five.

**Piecemeal delivery is expected and fine. A partial tool is not.** A tool is
delivered or it is not started — no stubs, no `TODO`, no half-wired buttons. A
button that opens something broken is worse than a button that is absent,
because the operator has already stopped looking for another way to do the job.

---

## 2. The finished console

```
Connect        1 row  ×  5   =  5 slots     4 buttons
Admin Tools    5 rows ×  5   = 25 slots    24 buttons
Daily Scripts  2 rows ×  5   = 10 slots     5 buttons
Reporting      3 rows ×  5   = 15 slots    12 buttons
Toolkit        2 rows ×  5   = 10 slots    10 buttons
                               ---------   ----------
                               65 slots    55 buttons
```

Admin Tools' 24 = **8 from `05`** (Mailbox Health, Message Trace, Unified
Lookup, Mailbox Permissions, Calendar Permissions, Out of Office, Inbox Rules,
Group Manager) **+ 16 from this document**.

`Sections.json` for the finished state:

```jsonc
{
  "ButtonsPerRow": 5,
  "Sections": [
    { "Name": "Connect",       "Rows": 1 },
    { "Name": "Admin Tools",   "Rows": 5 },
    { "Name": "Daily Scripts", "Rows": 2 },
    { "Name": "Reporting",     "Rows": 3 },
    { "Name": "Toolkit",       "Rows": 2 }
  ]
}
```

**Ship sections as they fill.** A section with no tools renders as an empty box,
so add each one when it has three or more tools, not before. The table above is
the destination, not the starting `Sections.json`.

> **Two deliberate moves from the source toolkit.** `Cert Check`
> (`CertificateExpiry.ps1`) and `Tester` (`Test-Toolkit.ps1`) sit in **Admin
> Tools** in the source `Tools.json`; here they move to **Toolkit**, which is a
> new section. Neither is an operator tool for a ticket. This move is what keeps
> Admin Tools at 24 of 25 slots — if you put them back, Admin Tools overflows
> and the console silently drops two buttons.

### The height check is now load-bearing

Thirteen rows of buttons plus five section headers plus the lookup controls plus
the output box is a tall window. On a 1080p screen with a taskbar it is close to
the limit.

**At console startup, compute the total required height from `Sections.json` and
compare it against `WorkingArea` — not `Bounds`, which includes the taskbar
(B4; verifier rule `WF004`).** If it does not fit, say so plainly in the output
box and name the section that overflowed. Do not silently draw buttons below the
bottom edge, where they are unreachable and invisible, and do not silently
shrink them.

This did not matter with fifteen buttons. It matters at fifty-five.

---

## 3. Grouping — one button, several scripts

**A `Group` is several tools behind one console button.** Clicking it opens a
small chooser listing the members by name; picking one runs it. This is how
fifty-nine button-backed scripts fit into fifty-five buttons without a wall of
near-identical controls.

`Tools.json` gains one optional field:

```jsonc
{
  "Name": "Room Regional Settings",
  "Script": ".\\RoomRegional.ps1",
  "Section": "Admin Tools",
  "Group": "Rooms",          // <-- same string on every member
  "Order": 22,
  "ObjectTypes": [],
  "NewProcess": false
}
```

Rules, all of which the reference console already implements:

- **Same `Group` string on every member.** Blank means the tool stands alone,
  which is the default and what most tools should stay as.
- **The group button takes the position of the member with the lowest `Order`**,
  so the group lands where you expect rather than wherever the alphabet puts it.
- **Members are still created as buttons — just hidden.** Two reasons, both
  real: the console's search filter can then reveal a matching member directly,
  so a grouped tool stays reachable by name; and the grey-out pass can compute
  the group button's state from members it already has in front of it.
- **The chooser repeats each member's availability.** A member that cannot run
  is disabled *in the chooser too*, with the same tooltip reason the console
  would have shown. Otherwise the group button becomes the place where greying
  goes to die, and the operator learns the reason one dialog too late.
- **Capacity counts buttons, not scripts.** A group of five occupies **one**
  slot. Get this wrong and §2's arithmetic is wrong.

> **Fully implemented in `reference/AdminConsole.ps1`** — `Get-ToolGroupName`,
> `Show-ToolGroupChooser`, and the button-build loop — and editable via the
> `Group` column in `reference/ToolManager.ps1`, which carries its own
> explanation. Take both verbatim rather than reimplementing.

**Groups in the finished toolkit: `Rooms` only** (five scripts, Admin Tools).
Others may be introduced if a section overflows, but grouping tools that are not
genuinely siblings makes them harder to find, not easier.

---

## 4. How to read an entry

Each entry heads with:

```
`Section` · ObjectTypes · capability · variant · Stage · ~lines
```

- **`ObjectTypes: []`** means the button is **never greyed on the looked-up
  object** — the tool is not about one object. It does **not** mean the button
  is never greyed at all: capability gating still applies, per
  `02-ARCHITECTURE-SPEC.md`. A tool marked `ObjectTypes: []` · *needs EXO* is
  greyed when Exchange Online is unavailable and enabled regardless of what was
  looked up.
- **capability** names what must be available for the button to be live —
  *needs EXO*, *needs EMS* (on-premises Exchange), *needs AD*, *needs SPO*, or
  *none*. Absent capability ⇒ greyed, with a tooltip saying which one.
- **variant** is the `# TOOLKIT-VARIANT:` marker the script carries.

**Where an entry names a constraint (C4, C11, …), that constraint has bitten
this specific tool before.** It is not a general reminder.

**Where an entry names a cmdlet, verify it against Microsoft Learn before use**,
per `00` §4A, and record the citation at the call site. This matters more here
than in `05`: these tools reach into corners of Exchange — `Set-Place`,
`Set-CalendarProcessing`, `New-RemoteMailbox`, blocked-sender release — where
parameter names and accepted values are easy to get plausibly wrong and the
failure is silent.

**Where you cannot tell what a tool should do from its entry, ask.** Do not
invent behaviour, and do not build a thinner version and call it done. Every
entry summarises something that already exists and is in daily use, so a missing
detail is knowable rather than a matter of taste.

### What every tool carries

The standard header block; the upward search for `Common.ps1`; the connection
guard; wrapped handlers (B1); bottom-left placement (B3a); the `?` help button
added **last** (B3b); UTF-8 with a single BOM and CRLF (A8).

### Audit `Target` — use `Get-LogTarget`, and stop thinking about it

`02`, `08` and `03`'s `LOG002` all require the **primary SMTP address** as
`Target`. That is correct for every tool acting on a mail-enabled object, and it
is most of them.

**About half the tools in this catalogue have no mailbox in play** — every
report (whose target is the report path), the three viewers, Header Analyzer,
Transport Rules, Connect-SPO, Dashboard and the whole Toolkit section.

**`Common.ps1` already solves this. Use it:**

```powershell
Write-CompletedLog -Action 'Grant' -Result 'Success' `
    -Target (Get-LogTarget $mbx.PrimarySmtpAddress $mbx.Name)
```

`Get-LogTarget` takes candidates in preference order and returns the first
non-empty one. **The fallback is deliberate, not lazy** — a plain security
group, a room list before Exchange assigns an address, a failed create: some
things genuinely have no address, and a name is a worse identifier than an
address but a far better one than nothing.

**Where there is no object at all, pass a literal scope name:** `'SharePoint
Online'`, `'Toolkit'`, `'Config'`, the report path. `reference/Connect-ExO.ps1`
is the pattern.

**The reference verifier already exempts literal targets** — it flags
`.Name` / `.DisplayName` / `.SamAccountName` but deliberately skips
`-Target '...'`, because those are the non-object cases where the literal *is*
the identifier. **Do not "fix" that exemption, and do not fabricate an address
to satisfy a rule.** A made-up address in an audit log is worse than a plain
label, because it looks searchable.

**A tool that reads nothing and changes nothing need not log at all.** Quick
Links and M365 Reports open a browser; there is no event worth a row.

---

## 5. Stages

Each stage ends with a **stop**: report what was built, what the verifier says,
and what went into `UNVERIFIED.md`. Do not roll into the next stage unprompted.

| Stage | What | Scripts | Why here |
|---|---|---|---|
| **A** | Read-only tools, plus Connect-SPO | 6 | Lowest risk in the set. Establishes the query-into-a-grid pattern that C–D reuse. SPO must exist before the SharePoint reports in D. |
| **B** | Cloud group tools | 2 | Depends on the classification logic built for Group Manager in Phase 4. |
| **C** | Cloud writes on a single object | 4 | First writes beyond the first pass. C10/C11 live here. |
| **D** | The reporting suite | 12 | Highly patterned and near-identical to one another. The volume is here and the risk is low — build one properly, then the rest as a batch. |
| **E** | Rooms | 5 | Self-contained, cloud-only, and the first `Group`. |
| **F** | On-premises and hybrid writes | 6 | **Highest risk in the toolkit.** C1/C2/C7/C8 apply to every one. Do not start before Stage C's audit rows have been confirmed on a real tenant. |
| **G** | Console furniture | 3 | Dashboard reports on everything above, so it goes after them. |
| **H** | Toolkit maintenance and the publish chain | 10 | Needed to keep and hand on the toolkit, not to use it. |

48 scripts. **Stage F is the one to slow down on.** Every tool in it can write
to the wrong directory silently, and "no exception was thrown" is not evidence
that it wrote where you intended (C1, C10).

### Work that is NOT a new script, and when it happens

Three pieces of this catalogue land in files built during the first pass. They
have no stage of their own and are easy to miss:

| What | Where | When |
|---|---|---|
| **The `Group` mechanism** — `Get-ToolGroupName`, `Show-ToolGroupChooser`, the grouped button-build loop (§3) | `AdminConsole.ps1` | **Before Stage E.** Take verbatim from the reference |
| **`Group` column + over-capacity warning** (§2, §3) | `ToolManager.ps1` | **Group column before Stage E; the capacity warning before Stage A**, which is where a section first overflows |
| **`WF010` / `WF011` + the `LOG002` literal exemption** | `Test-Toolkit.ps1` | **Before Stage A** |

**`Sections.json` `Rows` must grow with each stage, in the same commit.** The
first pass ships `Admin Tools: Rows 2` — ten slots, eight tools. Stage A adds
six Admin Tools buttons. **Fourteen buttons in ten slots does not error; four
buttons are simply absent.** Raise `Rows` to 4 at Stage A and to 5 by Stage E,
and add the `Toolkit` section at Stage H. §2 has the destination.

### Own-process isolation, and what it is not

Three tools here declare `# TOOLKIT-REQUIRES: OwnProcess`. That mechanism is
**the console's**, not a `Common.ps1` function: `AdminConsole.ps1` reads each
script's own declaration via `Test-ToolDeclaresOwnProcess` and launches it with
`Start-ToolScriptInNewProcess`. Both are in `reference/AdminConsole.ps1` and are
in scope.

`08-COMMON-CONTRACT.md` tells you **not** to port `Start-ToolInOnPremProcess`.
That is a different thing — it is multi-directory *routing*, which this
single-domain tenant does not need. **That prohibition stands. Own-process
isolation is still required**, because it defends against the cmdlet collision
(C1), not against a second directory.

---

# THE CATALOGUE

---

## Connect

### SPO — `Connect-SPO.ps1`
`Connect` · `ObjectTypes: []` · none · **Stage A** · ~167 lines

Connects the console to SharePoint Online. **Same pattern as `Connect-ExO.ps1`:**
the console closes, the sign-in happens, and the console reopens with the
connection live — because a sign-in window cannot accept typing while the
console's own window is in front of it (B2).

Needed by the OneDrive and SharePoint reports at Stage D, and by Mailbox /
OneDrive Perms at Stage F. **The Exchange tools do not use it and must not be
greyed on its state.**

**Flip `Capabilities.SharePointOnline` to `true` in `Tenant.json` as part of
this stage.** `02` ships it `false` because nothing in the first pass uses it,
and an unticked capability makes its tools **absent rather than greyed** — which
reads as a build defect.

> The exemplar is `reference/Connect-ExO.ps1`. **There is no
> `reference/Connect-SPO.ps1`** — only the ExO and ExSE connect scripts ship in
> the reference set. Follow ExO's shape.

**The admin URL is a TENANT FACT and lives in `Tenant.json`, at
`SharePoint.AdminUrl` — never in `Settings.json`.** This is a corrected bug, not
a preference: the source read `Settings.json` and then guessed, never looking at
`Tenant.json` at all, so the value Tenant Manager wrote and displayed was inert.
An engineer could type the correct admin centre into the tool whose entire job is
tenant facts, save it, and watch Connect SPO use a derived host that had never
existed. `Settings.json` no longer carries the key at all.

**Derivation is the fallback, not the source.** When `SharePoint.AdminUrl` is
empty, derive `https://<label>-admin.sharepoint.com` from the **session
organisation** — the SharePoint tenant label follows the `*.onmicrosoft.com`
prefix, **not** the vanity mail domain. They match at most tenants; they do not
always, and a tenant that renamed its mail domain would otherwise get a URL that
never existed.

Logs with `Target` = `SharePoint Online`.

---

## Admin Tools

### Dashboard — `Dashboard.ps1`
`Admin Tools` · `ObjectTypes: []` · none · **Stage G** · ~919 lines

The start-of-day page, and `Order: 0` — the first button on the console.

**Six panels**, each answering a question worth asking before touching anything:

| Panel | Question |
|---|---|
| **ENVIRONMENT** | Am I connected to what I need, and is PIM still active? |
| **TOOLKIT HEALTH** | Is my copy complete and current? |
| **TODAY** | What has run today, and did any of it fail? |
| **TOP TOOLS** | What actually gets used. |
| **FAILURE HOTSPOTS** | What keeps going wrong, so it can be fixed. |

> **The source has a sixth panel, RETRY QUEUE, and you should NOT build it.** It
> reports on a leaver/termination workflow, which is not in this catalogue and is
> not authorised work (`05`). Nothing in scope writes a retry queue, so the panel
> would be permanently empty and would read as a defect rather than as good news.
> Build five panels. If a multi-step workflow is ever added, the panel comes with
> it.

**TOOLKIT HEALTH reads `Version.json`, which only `Publish-Toolkit.ps1` writes
(Stage H).** Until then it must say "not stamped" rather than "incomplete" — an
unstamped working copy is the normal state, not a fault.

**READ-ONLY.** Nothing here changes a mailbox, a group or a file. Its buttons
open other tools; they do not act.

**The environment and toolkit-health panels must call the SAME functions as the
console header and footer**, not work it out again. Otherwise the two can
disagree, and **a second opinion about your own state is worse than one.** This
is a hard requirement, not a tidiness preference.

**Never fold an unknown result into `Success`.** Any audit row without a
`Result` counts as `Unknown` and is displayed as such. A dashboard that
overstates what worked is worse than one with a visible gap.

**It is built at Stage G** because five of its six panels report on tools built
in Stages A–F. Building it early produces a page reporting on nothing.

> Dashboard can launch tools that have no console button. `ToolStatus.ps1`
> recognises that as a legitimate state — see its entry and §7.

### AD Lookup — `AD-Lookup.ps1`
`Admin Tools` · `ObjectTypes: []` · **needs EMS** (degrades to AD-only) · **variant: OnPremAD** · **Stage F** · ~1,383 lines

Type part of a name, get back every matching on-premises object with **what it
is** and **where it lives**.

**It sits in Stage F, not with the read-only tools, because its attribute editor
writes to on-premises AD with `Set-ADObject`.** That puts it behind the same
gate as the other on-premises writers.

**How it reads, which is the thing most easily got wrong:** the search calls
`Get-Recipient` **unqualified against on-premises Exchange**, routed through
`Invoke-OnPremCommand`. That routing is not optional — an Exchange Online
session in the same process would answer for the cloud and **the OU column would
be empty on every row** (C1).

**When there is no on-premises Exchange in the process at all**, it falls back
to Active Directory and derives the type from attributes via
`Resolve-AdObjectType` in `Common.ps1`. That is a **degraded path**, not the
primary one.

> **Do not build this on `Get-ADObject` alone.** `msExchRecipientTypeDetails`
> comes back from AD as an integer, and the source explicitly considered and
> rejected decoding that bitmask. Guessing it produces a Type column that is
> confidently wrong, which is worse than an empty one.

**The Type column is the point.** These look alike in a name list and behave
nothing alike:

| Type | What delegation means |
|---|---|
| `UserMailbox` / `SharedMailbox` | Full Access and Send As both work |
| `MailUniversalDistributionGroup` | Send As works; Full Access is meaningless |
| `MailUniversalSecurityGroup` | A distribution list that also grants rights |
| `MailContact` | No mailbox anywhere; nothing to delegate |
| `RemoteSharedMailbox` | The mailbox is in the cloud; this is the on-premises stub that syncs it |

Getting this wrong is not a typo — it is twenty minutes and a confused reply.

**The OU column** answers the other half: which business unit owns this, and
which directory it came from.

**Double-click puts the address on the clipboard**, ready to paste into whichever
tool does the actual work. That is the primary workflow.

**It works on GROUPS**, which the older attribute tools did not — they used
`Get-ADUser` and simply failed on anything that was not a user. **Half of what
this finds is a distribution list.**

**"Also search Active Directory"** is a checkbox that widens the Exchange search
to include objects with **no mailbox at all** — the plain security group sitting
next to the shared mailbox with almost the same name. It cannot be delegated to,
but seeing it explains why the name looked ambiguous. **Off by default, because
it is slower.**

#### The attribute editor

Opens the selected object. Everything populated is shown; **46 attributes are
editable** — names, job and organisation, every phone field, the full address,
and all fifteen `extensionAttribute` fields.

**Five classes are READ-ONLY, and the list is short on purpose. All five must be
enforced:**

| Class | Fields | Why |
|---|---|---|
| **addresses** | `mail`, `proxyAddresses` | They belong to Exchange. Editing them behind its back **is how an object gets two primaries** (and see C11). Send the operator to Mail Attributes |
| **identity** | `sAMAccountName`, `userPrincipalName`, `cn`, `name` | Changing one of these is a migration, not an edit |
| **relationships** | `member`, `memberOf`, `managedBy`, `manager` | The value is a distinguished name, and a text box invites somebody to type a person's name into it |
| **state** | `userAccountControl`, the timestamps | Enabling an account by hand-editing a bitmask is a way to be wrong quietly |
| **multi-valued** | anything shown as `multi (n)` | **Writing one value back would silently discard the rest** |

Reading is safe. **The attribute editor is the only thing that writes**; it
confirms every change and records it in the audit log.

Escape every operator-typed value before it reaches an LDAP filter — RFC 4515
hex escaping, **not** the OPath doubling used for Exchange filters (C19a). The
two are not interchangeable, and using the wrong one is a silent wrong-results
bug rather than an error.

### Header Analyzer — `HeaderAnalyzer.ps1`
`Admin Tools` · `ObjectTypes: []` · none · **Stage A** · ~824 lines

Paste raw internet headers, get a readable account of what happened to the
message:

- the delivery path, **oldest hop first**, with the time spent at each hop
- SPF, DKIM and DMARC results, plus Microsoft's composite verdict
- the Exchange Online filtering result — spam confidence, bulk level, and *why*
  it was filtered, in words rather than numeric codes
- anything worth a second look, such as a `Reply-To` pointing somewhere other
  than the `From` address

Answers "why did this go to junk", "why was this blocked", and "is this actually
from who it says it is".

**Everything happens on this machine and nothing is uploaded.** That is the main
reason the tool exists — headers carry internal server names, addresses and
subject lines that must not be pasted into a public analyser website. Say so in
the help text.

Connects to nothing; parses text. No connection guard. Nothing worth auditing.

### DL Memberships — `DLMemberships.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **variant: Cloud** · **Stage A** · ~328 lines

Every distribution list a person belongs to. Answers "why is this person getting
this mail" and "what will they stop receiving when they leave" — the second is
worth checking *before* a termination rather than after.

**One filtered query, not an enumeration.** It should return in a second or two
however many groups the tenant has. Fetching all groups and filtering
client-side is the wrong shape (C15).

**Two limitations, both stated in the window and not only in the help:**

1. **DIRECT MEMBERSHIPS ONLY.** If someone is in "All Staff" because "Accounts"
   is in "All Staff", this lists Accounts and not All Staff. That is the reason
   a person can still receive mail from a list not shown here.
2. **Microsoft 365 groups and dynamic distribution groups are not included.**
   Dynamic groups have **no stored membership to query** — they are evaluated
   when mail is sent. Do not attempt to enumerate them.

Read-only. Point the operator at Group Manager or Bulk DL Membership to change
anything.

### Group Editor — `GroupEditor.ps1`
`Admin Tools` · `ObjectTypes: ["UnifiedGroup","OnpremGroup"]` · needs EXO · **variant: Cloud** · **Stage B** · ~601 lines

View and change **owners and members** of Microsoft 365 groups and of
on-premises groups synced to the cloud.

**Owners matter more than they look.** A group with no owner can be managed by
nobody except an admin, and nobody is notified when the last one leaves. Surface
the owner count prominently and warn before removing the last one.

**C18 applies and must be handled before the attempt, not after it.** `OnpremGroup`
means `IsDirSynced` is true — the object is mastered on-premises and **the cloud
will refuse most changes to it.** Check `IsDirSynced` and say plainly that the
change must be made on-premises, rather than issuing a write that appears to
succeed. Reporting success on a change that did not happen is the exact failure
mode `00` §1 names as the reason this pack exists.

Changes that do go through take effect immediately and are written to the audit
log.

Reached from Group Manager's routing as well as from its own button.

### UnifiedGroup Editor — `UnifiedGroup.ps1`
`Admin Tools` · `ObjectTypes: ["UnifiedGroup"]` · needs EXO · **Stage B** · ~344 lines

Membership of a Microsoft 365 group specifically.

**A Microsoft 365 Group is not a distribution list**, even though it receives
mail like one. It also carries a Teams team, a SharePoint site and a shared
mailbox — so adding somebody here gives them considerably more than an email
feed, **and the confirmation must say that.**

Owners and members are separate lists and both matter; ownership changes belong
in Group Editor.

Its recipient type is `GroupMailbox`, which **contains the substring "mailbox"
and is not one** (C12). Any code deciding "is this a mailbox" by substring
classifies this wrongly.

### Mail Attributes — `MailAttributes.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **variant: Cloud** · **Stage C** · ~739 lines

Addresses, display name, alias, and how the person appears in the address book.
**Mostly used for renames** — someone marries, changes name, or was created with
a typo.

**This is the most damaging tool in the toolkit if written naively, and it is
where C11 lives.** Writing a bare `EmailAddresses` list **REPLACES the whole
proxy-address collection.** The X500 addresses that keep old Outlook replies and
calendar items resolving are in that collection, they are invisible in most UIs,
and losing them is not obvious until weeks later when replies start bouncing.

**Add and remove individual entries. Never assign the whole collection.**

**Adding the new address while keeping the old one is the safe move**, because
mail to the old address keeps arriving. Changing the *primary* changes what
recipients see and what appears in the directory; the old address should normally
be kept as a secondary rather than deleted, and the tool should default to that.

**C10 applies with force.** `Set-Mailbox -PrimarySmtpAddress` is on-premises
only, and `Set-MailUser -PrimarySmtpAddress` is **accepted and silently ignored**
in Exchange Online. A tool that calls it, sees no exception and reports success
is lying. **Read the value back and compare before reporting anything.**

### All Attributes — `AllAttributes.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **variant: Cloud** · **Stage C** · ~449 lines

A wider view of the directory attributes on an account, beyond the mail ones.
For when something is set that should not be, or missing when it should be
there, and the operator needs the **raw value** rather than what a UI reports.

**Editing here writes directly to the attribute with no validation beyond what
the directory itself enforces.** Say that in the confirmation dialog, not only in
the help text.

Point the operator at Mail Attributes first if the field is a mail one — that
tool understands addresses properly and this one does not.

### Audit UserMB — `AuditUser.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **Stage A** · read-only · ~441 lines

Reports the **mailbox audit settings** on a user mailbox — which actions are
recorded, separately for the owner, for delegates, and for administrators.

Used when someone asks what can be proven about access to a mailbox, or before
an investigation, to establish whether the events of interest were being logged
**at the time**. That is the tool's whole value: it answers a question about the
past, and the answer is often no.

### Audit SMB — `AuditSMB.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **Stage A** · read-only · ~497 lines

The same report for a **shared** mailbox, and worth being a separate tool.

Shared mailboxes usually have several delegates, so **delegate actions are the
events that matter most** — and they are the ones most often not being recorded.
Order and emphasise the output accordingly rather than mirroring the
user-mailbox layout.

### Blocked Senders — `BlockedSenders.ps1`
`Admin Tools` · `ObjectTypes: []` · needs EXO · **variant: Cloud** · **Stage C** · ~481 lines

Every account the service has blocked from sending mail, and the ability to
release one.

An account lands here after sending mail the service classified as spam. In
practice that is usually a deliberate bulk send that went out the wrong way —
**but it is also exactly what a compromised mailbox looks like, and the two are
indistinguishable from this list alone. Find out which before releasing.** That
sentence belongs next to the release button, not in a footnote.

Columns shown before any action:

| Column | Why |
|---|---|
| **Reason** | why the service blocked them, in its own words |
| **Blocked** | when it happened |
| **Released here** | how many times *this toolkit* has released this address, and when it last did |

The last column is computed from the toolkit's own audit log. A second release is
a coincidence; a fourth is a pattern — usually a team that keeps forgetting the
bulk-send tool. Releasing again fixes today and not next week.

**Microsoft limits how many times an account can be released.** Exceed it and
the release fails and only a support case clears it. **The count is a budget, not
just a habit tracker.**

**The ticket reference the operator types is recorded with the release** — in the
audit log and in the service — so "who released this and why" is answerable
later.

Releasing does not fix the cause. **If the account was compromised, releasing it
hands the mailbox back to whoever took it** — put that in the confirmation.

### Transport Rules — `TransportRules.ps1`
`Admin Tools` · `ObjectTypes: []` · needs EXO · **variant: Cloud** · **Stage A** · ~427 lines

Every mail-flow (transport) rule in the organisation, in one list.

**It does not change, enable or reorder any rule.** The only thing it writes is a
snapshot file (below).

**Why not edit here, and the reasoning belongs in the help text.** A transport
rule applies to every message the organisation sends or receives, and
`Set-TransportRule` carries well over a hundred conditions and actions. A form
covering all of them would be the admin centre rebuilt worse; a form covering
some of them could only edit the rules it understood and **would quietly drop the
parts it did not.** The admin centre's condition builder is the one thing it
genuinely does well. Author there. If asked to add editing, say why not.

**Four capabilities, and the last two are why the tool earns a button:**

| | |
|---|---|
| **Search** | one box, matched against the **whole rule** — name, description, every condition and every action. "Which rule mentions this domain" is one keystroke here and a long afternoon in the admin centre |
| **Priority** | the running order in a single list. **Which rule wins is decided by this, and it is the thing the admin centre shows least clearly** |
| **Detail** | every property the rule actually **sets**, with the couple of hundred empty ones left out |
| **Snapshot** | the whole collection to a dated XML file. **Mail flow has no change history, so a snapshot before anyone edits is the only way to answer "what did this look like last week"** |

**"Take a snapshot BEFORE changing rules, not after. It is the only version of
yesterday you will ever have."** Put that sentence in the tool.

The snapshot is the one write, and it is to a file. Audit it with the snapshot
path as `Target`.

### Change Organizer — `ChangeMeetingOrganizer.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **variant: Cloud** · **Stage C** · ~382 lines

Moves a meeting or series from one organizer to another. The usual case is
someone leaving who owns a recurring meeting nobody else can edit.

Flow: look up the **current** organizer → type the subject → name the new
organizer → **Preview** → **Transfer**.

**Four warnings. All four must appear BEFORE the transfer, not after.** An
operator who learns any of these afterwards has made things worse than before
they started.

1. **Teams.** Changing the organizer does **not** change the organizer of the
   Teams online meeting attached to it. The new organizer will not be able to
   manage or start that Teams meeting and must replace the online meeting details
   manually afterwards.
2. **Who notices.** Attendees *in this tenant* see nothing — silently updated, no
   re-RSVP. **External and on-premises attendees get a cancellation from the old
   organizer and a new invitation from the new one.** They must re-RSVP, and any
   Reminder, Category or Show As changes they made are lost.
3. **One way, once.** After a transfer you cannot transfer the part of the series
   left behind on the original calendar again.
4. **Scope.** **User mailboxes only, default calendar only.** Shared and room
   mailboxes are **not supported** — give the new person access to the mailbox
   instead.

Point 4 is a functional refusal, not advice. `ObjectTypes: ["Mailbox"]` admits
shared and room mailboxes, so **the tool must check the recipient type itself and
refuse** with that explanation.

### Rooms — group of five
`Admin Tools` · `Group: "Rooms"` · **Stage E** · one button, five scripts

Room administration splits into five genuinely different things that get
confused with one another constantly. The chooser naming them is most of the
value.

All five: `ObjectTypes: []` · needs EXO · **variant: Cloud**.

#### Room Lists — `RoomList.ps1` · ~429 lines

*Why a room does not appear in Outlook's Room Finder.*

**The Room Finder does not list rooms. It lists ROOM LISTS, and then the rooms
inside them.** A room mailbox belonging to no room list is invisible there no
matter how correct its mailbox and place record are — and the symptom people
report is "the room does not exist", which points nowhere near the cause.

A room list is an ordinary distribution group flagged as a room list. **It cannot
be flagged after the fact through this tool**: a normal group converted to a room
list keeps behaving oddly in the Room Finder, so **New is the only route
offered.** That is a deliberate omission, not a missing feature.

Membership changes take effect immediately, **but Outlook and Teams cache the
Room Finder aggressively — allow up to 24 hours** before concluding it did not
work. Verify in this tool rather than in Outlook.

#### Room Places — `RoomPlace.ps1` · ~493 lines

Capacity, floor, building, devices, accessibility, tags — the room metadata
behind the Room Finder.

**This is not the mailbox.** A room has a mailbox (name, address, permissions)
*and* a "place" record (what the room is like), and **only the place record
drives room suggestions.** A room that never appears in a search usually has an
empty place record rather than a broken mailbox.

Writes with `Set-Place` immediately. **No staging and no undo** — say so.

**In a hybrid organisation some fields cannot be set here at all.** If the room
mailbox originated on-premises and syncs up, Microsoft blocks `Set-Place` from
changing `City`, `CountryOrRegion`, `GeoCoordinates`, `Phone`, `PostalCode`,
`State` and `Street`. Those come from the on-premises directory and must be
changed there with `Set-User` or `Set-Mailbox`.

**Mark those rows in the UI rather than letting the operator discover it from a
failure.** This tenant is hybrid, so this is the normal case, not the edge case.

#### Room Regional — `RoomRegional.ps1` · ~522 lines

The timezone a room thinks it is in, and the hours it considers bookable.

A room created without a timezone inherits one, and that is how a room accepts a
9am booking that shows as 3am to the people in it. **Nothing errors; the meeting
is simply at the wrong time.**

**Two timezones live here and they are not the same setting:**

- **Regional TimeZone** — how the mailbox stores and displays times
- **Working hours TimeZone** — the zone the working hours are stated in

They should almost always match. When they do not, "only during working hours"
means something nobody intended, **and this tool must say so** rather than
displaying both without comment.

`ScheduleOnlyDuringWorkHours` in Calendar Processing depends entirely on the
working hours set here. Turning that on without checking these is how a room
starts declining bookings for no visible reason — cross-reference the two tools
in both directions.

#### Calendar Processing — `CalendarProcessing.ps1` · ~498 lines

How a room decides what to do with a meeting request: accept, decline, hold for
approval, and what the reply says. This is where "the room accepts everything",
"the room declines recurring meetings" and "nobody ever approves the bookings"
are actually decided.

**`AutomateProcessing` is the master switch** and must be understood before
anything else on the form is touched:

| Value | Behaviour |
|---|---|
| `AutoAccept` | the room books itself by the rules below |
| `AutoUpdate` | requests are marked tentative; a delegate decides |
| `None` | nothing is automatic; the mailbox behaves like a person |

**Most of the other settings only take effect under `AutoAccept`.** Lay the form
out so that switch reads as governing the rest, because it does, and grey or mark
the settings that are inert under the current value.

**The same "list alone does nothing" trap as Booking Policy applies here.**
`BookInPolicy`, `RequestInPolicy` and `RequestOutOfPolicy` hold lists of people
and interact with the `AllBookInPolicy` / `AllRequestInPolicy` switches.
**Setting a list while the matching `All-` switch is still true does nothing
visible.** Set them together and say so.

Writes with `Set-CalendarProcessing` immediately. **No undo.**

#### Booking Policy — `Set-BookingPolicy.ps1` · ~548 lines

**Who** is allowed to book a room. Everyone not on the list has their request
declined automatically.

Add by typing an address or alias and pressing Enter. **Bare aliases work** — if
`jsmith` does not resolve on its own, retry against the tenant's own primary
domain before giving up. Remove by selecting a row and pressing Delete, or via
the Remove button.

**THE LIST ALONE DOES NOTHING.** A room also carries a master switch,
`AllBookInPolicy`, and while it is `True` the room accepts bookings from everyone
and **the list is ignored entirely.** Restricting a room means setting both. This
tool sets them together and says so before it writes.

**An empty list with restriction on means nobody can book the room at all.** A
valid thing to want and a terrible thing to do by accident — confirm it
separately, with different wording from the normal confirmation.

### Resource Duplicator — `ResourceDuplicator.ps1`
`Admin Tools` · `ObjectTypes: []` · **needs EMS *and* EXO** · **Stage F** · ~819 lines

Copies a room or equipment mailbox: reads every setting off an existing resource,
creates a new one, and applies the same settings — booking rules, delegates,
capacity, auto-accept behaviour. More reliable than setting a dozen options by
hand and hoping you remembered them all.

**This is a HYBRID, THREE-STEP tool and the hardest thing in this catalogue to
get right. It is not an ordinary Admin Tools button.** The form is laid out as
three labelled groups, in this order:

```
Step 1  --  source resource
Step 2  --  new resource  (created on-premises)
Step 3  --  apply settings  (after directory sync; works in a later session)
```

**THE CAPTURE FILE IS THE MECHANISM, and it is the thing you would not guess.**
Step 1 reads every setting off the source and **writes them to a capture file on
disk**. Step 3 reads that file — **not the live session**. That is what makes
"come back after the next sync cycle" workable: the operator can close the
window, go home, and finish tomorrow. The tool says so out loud:

> *"You can close this window: Step 3 reads the capture file, not this session."*

Show the available captures in Step 3, with `(no captures yet -- run Step 1)`
when there are none. **Without the capture file the operator must leave a
WinForms window open across a directory-sync cycle, which is not a workflow.**

**Step 2 — create, on-premises.** Calls `New-RemoteMailbox` with `-Room` or
`-Equipment`. That cmdlet exists **only in the on-premises Exchange Management
Shell**; check for it and fail with that explanation rather than a
`CommandNotFoundException`:

```powershell
if (-not (Get-Command -Name 'New-RemoteMailbox' -ErrorAction SilentlyContinue)) {
    throw 'New-RemoteMailbox is not available. This step needs the on-premises Exchange Management Shell.'
}
```

**Ask the cmdlet what it declares before using a switch** (C13) — `-Room` and
`-Equipment` are not present on every build:

```powershell
$declared = Get-CmdletParameterNames -Name 'New-RemoteMailbox'
if ($declared -notcontains 'Room') { throw 'New-RemoteMailbox does not declare -Room on this build.' }
```

**Step 3 — configure, in the cloud, AFTER directory sync.** The new mailbox does
not exist in Exchange Online until sync publishes it. The tool must **detect
that and stop cleanly**, not fail obscurely:

```powershell
throw "'$Identity' is not in Exchange Online yet. Directory sync has not published it -- wait for the next cycle and click Apply again."
```

**The operator clicks Apply again after the next sync cycle, possibly from a
fresh launch of the toolkit.** A single-pass implementation fails on every run.
This and the capture file are the two most important things to get right here,
and neither is visible from the tool's own help text.

**Every cloud call goes through `Invoke-EXOCommand`** — `Get-Mailbox`,
`Get-CalendarProcessing`, `Set-CalendarProcessing`, `Set-Mailbox`,
`Get-MailboxFolderPermission`, `Get-MailboxRegionalConfiguration`,
`Get-DistributionGroup`, `Get-Place` / `Set-Place`. Several of those names exist
in **both** editions (C1); the wrapper is what stops the on-premises session
answering a call meant for the cloud in a process that has both loaded.

**Its mistakes get copied too.** Show the source configuration for review before
creating anything, and make the operator confirm against what they can see rather
than against a room name.

Stands outside the `Rooms` group deliberately: the other five configure an
existing resource, this one creates one.

### Console — `console.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **Stage G** · ~158 lines

A PowerShell prompt that **shares the console's live connections.** For the times
a tool does not cover what is needed and the operator wants to run cmdlets by
hand without connecting to everything again in a new window.

Whatever the console is connected to, this is connected to — which necessarily
includes the ability to change things, **with none of the previews or
confirmations the tools provide.** State that in the help text.

**Must run in the console's own process**, since sharing the connections is the
entire point (C17). It is the one tool for which `OwnProcess` would defeat the
purpose.

---

## Daily Scripts

### New User — `NewUser.ps1`
`Daily Scripts` · `ObjectTypes: ["Mailbox"]` · needs EMS · **Stage F** · ~584 lines

Creates the mailbox for a new starter in the hybrid setup. Enables a **remote
mailbox** on-premises — the on-premises object that tells the directory this
person's mailbox lives in the cloud — along with their archive.

**The mailbox appears in the cloud after the next directory sync, not
immediately.** That gap is normal and is the usual reason a brand-new account
looks broken for a while. **The completion message must say this**, or the tool
generates its own support tickets.

**C7 applies and is the trap here.** This tool sets Exchange's AD session scope.
`Set-ADServerSettings` is **process-wide and outlives the tool that set it** — a
tool that sets it and does not restore it poisons every subsequent tool in the
console's process. Either restore the previous value in a `finally`, or pass
scope per call and never set it globally. Note which you chose in
`UNVERIFIED.md`.

**C8 also applies:** Exchange's `-DomainController` takes a **server**, not a
domain. Passing a domain name returns error `0x52`, which reads like a typo and
is not.

**Derive the set of directories from what was actually resolved; never hardcode
a list.** The source looped over two literal names, so an account resolving to
anything else matched no iteration and was **silently never written** — no error,
no audit row, no mailbox.

> **This tenant has ONE directory, and `04-REFERENCE-FILES.md` and `00` §8 both
> tell you not to port multi-directory machinery. Both still apply.** Do not
> build cross-directory routing, per-directory controller selection, or
> side-routing. The point here is narrower: where the code iterates over
> resolved objects, iterate over **what it resolved** rather than a literal —
> which with one directory is a one-element set and costs nothing.

### Mailbox / OneDrive Perms — `MailboxOneDriveAccess.ps1`
`Daily Scripts` · `ObjectTypes: []` · **needs EXO *and* SPO** · **Stage F** · ~808 lines

Grants or removes access to somebody's mailbox **and** their OneDrive in one
place. The usual case is someone leaving or off unexpectedly when their manager
needs both — **and doing them separately is how one of the two gets forgotten.**

- **Mailbox** — Full Access and Send As
- **OneDrive** — site collection administrator on their personal site

Gate on both connections and **say which one is missing**, not "not connected".

**OneDrive access is visible to the person if they return.** Appropriate for a
departure, less so for a temporary absence, and either way it should be removed
when no longer needed. Put that in the confirmation.

**Audit each half separately.** A run that granted mailbox access and failed on
OneDrive is a **Partial** — not a Success and not a Failure. If the audit row
cannot distinguish them, the log cannot reconstruct what actually happened, which
is the one job it has.

### Manage Mail Object — `ManageMailObject.ps1`
`Daily Scripts` · `ObjectTypes: []` · needs EMS · **variant: OnPrem** · **`OwnProcess`** · **Stage F** · ~645 lines

Creates and removes shared mailboxes and distribution groups on-premises.

Flow: type the address (with or without domain) → pick object type → pick Add or
Remove → **Preview** → **Apply**.

**Declares `# TOOLKIT-REQUIRES: OwnProcess` in its own header, within the first
40 lines.** `New-DistributionGroup` and `Remove-DistributionGroup` exist in
**both** Exchange editions, and in one process the last import owns the name
(C1/C2). Getting this wrong creates or deletes the object in the wrong place with
no error at all.

**On Add** it needs a directory, because that decides which OU the object lands
in. Directories with an OU configured in Tenant Manager use it without asking;
those without prompt for it.

**Shared mailboxes are created as REMOTE mailboxes** — the AD account is created
on-premises, the mailbox lives in the cloud. That is why they appear after a sync
cycle rather than immediately.

**On Remove it looks the object up first, shows what it found, and refuses if the
object is not the type selected.** Typing a person's address into "Remove Shared
Mailbox" would otherwise delete their mailbox. **This guard is the single most
important line of code in the tool.**

- **Remove** deletes the AD account. The cloud mailbox is soft-deleted and
  recoverable for a limited window.
- **Disable** keeps the AD account and strips its mail attributes. **Disable is
  the reversible one**, and is offered as a checkbox on Remove.

**Owner is optional.** On a shared mailbox it grants FullAccess and SendAs; on a
distribution group it sets `ManagedBy`, which also lets that person edit
membership from Outlook.

Distribution groups reject external senders unless the box is ticked.

### Bulk DL Membership — `Bulk-DLMembership.ps1`
`Daily Scripts` · `ObjectTypes: []` · needs EXO or EMS · **`OwnProcess`** · **Stage F** · ~692 lines

Adds or removes a set of people across a set of distribution lists in one go.
Distribution lists in one box, users in the other.

**The operation is every user against every list.** Ten users across four lists
is forty changes, **and the preview says that number before anything runs.**

- **Check** — resolves everything and shows exactly what will happen
- **Run** — does it, after a confirmation

**Anything that will not resolve is named individually and skipped**, rather than
failing the whole batch. A batch that aborts on item three of forty is worse than
one that completes thirty-seven and names the three it could not.

**One audit row per person**, not one per batch. There is no undo: a removal is
recoverable only by adding people back, and **the audit log is the only record of
who was in what.**

**Which environment it writes to depends on how it was opened.** Opened from **Group Manager**, it follows the group — a synced group is edited on-premises in its own
window, a cloud-managed one in Exchange Online in the console's session. **The
banner at the top of the window always names the environment it is about to write
to.** Read it before pressing Run.

**It refuses to run in a process where both environments are loaded** (C1). Not a
warning — a refusal.

### Quick Links — `QuickLaunch.ps1`
`Daily Scripts` · `ObjectTypes: []` · none · **Stage G** · ~354 lines

The operator's own set of links, as buttons. Each opens in the Edge profile it
was configured with, so admin portals open with the elevated account and everyday
links do not.

Links marked as admin render **dark blue**, so which is which is visible before
clicking. A Find box narrows a long list.

Reads `Config\QuickLaunch.json`. Edited by Quick Launch Tool (Stage H).

Opens links in a browser. **Connects to nothing and must not be greyed on any
connection state.** Nothing worth auditing.

`Start-Process -ArgumentList` does not quote for you — build the profile argument
carefully (A11).

---

## Reporting

Twelve buttons. **Nine of them produce a CSV; three are viewers.** The rules
below are stated separately for each, because applying the CSV rules to the
viewers produces requirements that cannot be met.

### The nine CSV reports

`Protocols`, `DisabledLicensed`, `StaleAccounts`, `DLInventory`,
`ExternalForwarding`, `PermissionsAudit`, `OrphanedOneDrive`, `ExternalSharing`,
`OwnerlessGroups`. All `ObjectTypes: []`, all read-only, all **Stage D**.

Shape: run a query, write a timestamped CSV to the reports folder, show results
in a grid, offer to open the file.

- **Use the shared report-file controls** — `Add-ReportFileControls` from `08`
  Tier 2 — rather than each tool inventing its own save path and naming.
  **`08` says to add it with the first report. This is that point.**
- **Slice large queries; never retry an oversized one** (C15). A query that
  failed for size fails again the same way.
- **`ResultSize` explicitly on every enumerating *Exchange* call** (C3). The
  default is 1000 and the short answer arrives silently. **`DLInventory` queries
  Active Directory, where `ResultSize` does not exist** — use `-ResultPageSize`
  / paging there instead, and do not bolt an Exchange parameter onto an AD
  cmdlet.
- **Never write a zero-row file.** An empty CSV and a failed run look identical
  in a folder listing a week later.
- **Label a partial run as partial**, in the file and in the audit row. A report
  covering 80% of mailboxes that does not say so is worse than no report, because
  it will be used as though it were complete (D6).
- `Target` in the audit row is the report name or output path.

### The three viewers

`ReportBrowser`, `M365Reports`, `LogViewer`. These run no Exchange query, write
no CSV, and the four rules above do not apply to them. Specified individually
below.

---

### Protocols — `Protocols.ps1`
needs EXO · **Stage D** · ~404 lines

Which client protocols are enabled per mailbox — POP, IMAP, ActiveSync, Outlook
on the web and the rest.

Security review material: **legacy protocols are the ones most used in
password-spray attacks**, and this shows where they are still on.

This is the tool that taught **C15** (slice the query) and **C13** (ask the
cmdlet what it declares rather than assuming the property set). It is in
`reference/` — **treat it as the reference implementation for every bulk query
in this section.**

### Disabled/Licensed — `DisabledLicensed.ps1`
needs EXO + AD · **Stage D** · ~345 lines

Accounts disabled in Active Directory but still in a licence group — **still
costing money.** Straightforward cleanup list, and usually longer than expected.

**The licence group name comes from config, never a literal.** Removing the
licence is a separate, deliberate step and is not offered here.

### MB w No Activity — `StaleAccounts.ps1`
needs EXO · **Stage D** · ~348 lines

Accounts enabled and licensed that have not signed in for a while. Two reasons to
care: a licence paid for and not used, and an account nobody is watching that
still works.

**Not every result is dead** — shared and service accounts appear here
legitimately. Present it as a list to **review**, not a list to disable, and say
so in the tool rather than only in the help.

### DL Inventory — `DLInventory.ps1`
needs AD · **variant: OnPremAD** · **Stage D** · ~404 lines

Inventories the mail-enabled groups in **on-premises Active Directory** for
migration planning: what exists, how big it is, who manages it, in one sortable
file.

> **The source carries `# TOOLKIT-VARIANT: Cloud` while reading on-premises AD.
> That marker is wrong in the source.** Write `OnPremAD` and note the correction
> in `UNVERIFIED.md` rather than propagating it.

### External Forwarding — `ExternalForwarding.ps1`
needs EXO · **Stage D** · ~405 lines

Mail leaving the organisation, from **both** places it can be configured:

1. **mailbox forwarding**, set by an admin or by the user
2. **inbox rules**, which users create themselves — and which are the ones
   nobody thinks to check

**The classic sign of a compromised mailbox is a quiet rule forwarding everything
to an outside address. This is the report that finds it.** Checking only the
first source is the common mistake, and it produces a clean report on a
compromised tenant.

Also the reference for the shared report controls and partial-run labelling. It
ships in `DEVIN-REFERENCE-CODE-2.md`; after Step 0 it is at
`reference/ExternalForwarding.ps1`.

### Permissions Audit — `PermissionsAudit.ps1`
needs EXO · **Stage D** · ~393 lines

Every mailbox delegation across the organisation — Full Access, Send As, Send on
Behalf.

Answers "who can get into what", **a question that only ever gets asked under
time pressure** — which is the argument for it being a saved report rather than
something assembled live. Access granted for a good reason years ago is still
access.

It reports; it removes nothing.

### Orphaned OneDrive — `OrphanedOneDrive.ps1`
needs SPO + AD · **Stage D** · ~361 lines

OneDrive sites whose owner is disabled or no longer in Active Directory.

When someone leaves, their OneDrive does not leave with them. It sits there
holding whatever they had, consuming storage, with nobody responsible for it.

**Read-only on purpose** — deciding what to do with the content is a
conversation, not a script.

### External Sharing — `ExternalSharing.ps1`
needs SPO · **Stage D** · ~436 lines

How each SharePoint site is configured for external sharing, **and** which guest
accounts actually have access.

**Both halves matter and neither is enough alone:** a site permitting external
sharing with no guests on it is a **setting**, while a guest with access to a
site is a **person who can read your files.** Report both and keep them visually
distinct.

### Ownerless Groups — `OwnerlessGroups.ps1`
needs EXO · **Stage D** · ~349 lines

Microsoft 365 groups with no effective owner.

An ownerless group cannot be managed by anyone but an admin — nobody can add
members, change settings or delete it — **and it happens quietly whenever the
last owner leaves. Nobody is notified.**

### Report Browser — `ReportBrowser.ps1`
`ObjectTypes: []` · none · **Stage D** · ~279 lines

Opens any CSV the reporting tools have written, without going to find it in
Explorer. Lists everything in the reports folder, **newest first, across every
report type**, so last month's can be re-read without re-running it.

**Depends on every report using the shared save path.** A report that invents its
own location is invisible here — which is the practical argument for
`Add-ReportFileControls`.

Reads the filesystem. Connects to nothing.

### M365 Reports — `M365Reports.ps1`
`ObjectTypes: []` · none · **Stage D** · ~304 lines

Buttons that open Microsoft 365 admin portal reports — usage, activity,
licensing, security — in the browser.

**These are reports Microsoft already computes from data a script cannot get at
directly. Rather than reproduce them badly, this puts them one click away.**

Opens each in the elevated Edge profile, so the operator lands signed in as the
right account.

**Deep links need no Graph consent** (C19), which is what makes this viable in a
tenant that will not consent to Graph (C20). **Reads nothing, changes nothing,
opens links.** Nothing to audit.

### Log Viewer — `LogViewer.ps1`
`ObjectTypes: []` · none · **Stage D** · ~518 lines

A read-only view of the audit log — every change the toolkit's tools have made,
who made it, when, and to whom. **The answer to "when did that happen and who did
it".**

Search and filter across everything.

**Never fold a row with no `Result` into `Success`.** Count it as `Unknown` and
show it as such. A log that overstates what worked is worse than one with a
visible gap.

> The source handles a legacy two-column log schema. **That schema has never
> existed at this tenant** — see §6 — so do not build conversion or dual-schema
> reading. The `Unknown` rule above still stands as defensive handling.

Nothing here changes anything. The log is the record; this reads it.

---

## Toolkit

Maintenance and configuration. Needed to *keep* and *hand on* the toolkit, not to
use it. All **Stage H** unless noted.

### Tenant Manager — `TenantManager.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~536 lines

Edits `Tenant.json` — the facts about this tenant: name, domains, which services
it has, its directories. Every other tool reads it, and it is **the only file in
the toolkit allowed to contain a domain or server name.**

**Why this tool exists:** capabilities decide which buttons appear at all, and an
absent `Tenant.json` means every capability reads as false. A new tenant opened a
console with the on-premises tools simply missing and nothing saying why. Every
other config file had an editor; **the one that decides what the console IS had
none.**

**Two modes, chosen automatically:**

| Condition | Mode |
|---|---|
| No `Tenant.json` | first-run setup, started from the shipped template |
| `Tenant.json` present | editor, showing what is there now |

**`Detect`** fills in what it can from the connections the window already has, so
most of the work is confirming rather than typing.

**CAPABILITY IS A CLAIM, NOT A PROBE.** Ticking Exchange Online says this tenant
*has* it, not that you are connected right now. Unticking something the tenant
genuinely does not have makes its tools **disappear** rather than sit greyed.

**Preserve keys it does not understand.** The source dropped `_comment_*` keys on
save — it carried the root `_comment` and lost `_comment_Capabilities`,
`_comment_Directories` and the rest. An editor that silently discards fields it
was not taught about will eventually delete a live setting. **Round-trip
unknown keys unchanged.**

**Also owns two keys other tools depend on and nothing else surfaces:**
`SharePoint.AdminUrl` (see Connect-SPO) and `FeatureRequest.Recipient`, which
Report a Bug and Feature Request share. Both are in `02`'s schema; both fail
quietly if this editor does not expose them.

### Tool Manager — `ToolManager.ps1`
`Toolkit` · `ObjectTypes: []` · none · **built in Phase 2 — additions only, no stage** · ~948 lines

Already required infrastructure in `00` §5. Listed here for **two additions the
finished toolkit needs:**

1. **The `Group` column** from §3, with its explanation.
2. **An over-capacity warning.** Report a section that has more buttons than
   `Rows × ButtonsPerRow` rather than letting the console silently drop the
   overflow — **and count buttons, not scripts.**

### Theme Manager — `ThemeManager.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~604 lines

The console's colours, and which theme each tenant gets.

**COLOUR IS NOT DECORATION HERE.** One console folder per tenant, all looking
identical, is how a change lands in the wrong tenant. A distinct theme per tenant
means the operator can tell at a glance which one is in front of them, **before
clicking anything.**

**Add shipped themes** brings in the bundled palettes. Pick one, Save, relaunch
the console.

**Tenant binding** ties a theme to a tenant so the right colours load
automatically. The console reads `Theme` from `Tenant.json`; this is where that
value comes from.

**CONTRAST MATTERS MORE THAN TASTE.** A palette where disabled text is nearly
the same as enabled text **hides the grey-out that tells you a tool cannot run.**
**The shipped themes are checked for that; hand-rolled ones are not** — say
exactly that in the help text. This is a stated limitation, not a feature: the
source tool does no contrast checking and you should not invent one.

The logo is cosmetic and optional. It is the one thing here that is.

**`Themes.json` is per-deployment, like the rest of `Config\`, and is NOT
shipped.** Note the tension with Deploy's rule that "everything else in `Config`
is tenant fact and does travel" — `Themes.json` is the documented exception, and
the source carries both statements. Resolve it in favour of not shipping it, and
record the decision in `UNVERIFIED.md`.

### Tool Status — `ToolStatus.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~725 lines

Lists every `.ps1` beside the console and says **what loads it:**

| State | Colour | Meaning |
|---|---|---|
| **AdminConsole** | green | a console button points at it |
| **Dashboard** | green | the dashboard can launch it |
| **Not loaded** | yellow | nothing opens it — safe to park |
| **Parked** | grey | already moved to `DisabledTools` |

**Yellow does not mean useless.** A script may be deliberately shelved, or called
by another script rather than opened by a button. **Read it before parking it** —
the `?` column, or a double-click, shows that tool's own help text without
leaving the window.

**`Common.ps1`, `AdminConsole.ps1`, `Publish-Toolkit.ps1` and this tool are not
listed.** They are infrastructure and must never be parked.

**Parking MOVES a file to `DisabledTools`.** It deletes nothing, and Restore puts
it back. **This deployment keeps only what it uses; the master copy keeps
everything.**

**Last run comes from the audit log**, so it is when the tool last did something
worth recording — **not when it was last opened.** Label the column so that is
obvious.

**After parking anything, re-run Publish-Toolkit.** The version manifest lists
files by name, so a parked script reads as *missing* until the toolkit is
re-stamped.

### Quick Launch Tool — `QuickLaunchTool.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~371 lines

Edits the links Quick Links shows.

- **Section** — any name; links group under it, and sections appear in the order
  they first appear in the list
- **Move Up / Move Down** — controls both section order and order within a
  section
- **Edge profile** — picked from the profiles that actually exist on the machine.
  Blank means Edge's Default profile
- **Admin** — makes the button dark blue, for links needing an elevated account
- **Test Selected** — opens the highlighted row exactly as Quick Links would, so
  a wrong profile is found here rather than in front of a customer

**Nothing is saved until Save is pressed.** Say so.

### Config Backup — `ConfigBackup.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~307 lines

Copies `Config\*.json` somewhere safe, and puts them back. These are the files
that are **the operator's** rather than the toolkit's: `Tools.json`,
`Settings.json`, `Sections.json`, `QuickLaunch.json`, `Themes.json`,
`Tenant.json`.

**Copying a new console folder over an old one replaces every one of them, and
nothing warns you.** The loss shows up later as blank paths and missing buttons,
long after the cause.

**RUN A BACKUP BEFORE COPYING A NEW CONSOLE OVER THIS ONE.** That is the whole
reason the tool exists, and it belongs at the top of its help text.

**Backups are written OUTSIDE the console folder, beside it**, so that replacing
or renaming the console cannot take the backup with it. **Each backup is a
timestamped folder** holding the files exactly as they were.

**Restore copies a backup back over `Config\` — and takes a fresh backup of
what is there first**, so restoring the wrong snapshot is not final either.

### Cert Check — `CertificateExpiry.ps1`
`Toolkit` · `ObjectTypes: []` · needs EMS · read-only · **Stage H** · ~459 lines

Certificates on every on-premises Exchange server and how long each has left.
Worth running monthly.

**An expired certificate breaks hybrid quietly.** The first symptom is usually
somebody reporting that free/busy has stopped working between on-premises and the
cloud, **which does not look like a certificate problem** — hence the tool.

**Separate certificates in use from ones merely present**, since only the first
kind breaks something.

**Takes `-Environment` as a parameter and must USE it.** The source declared the
parameter and then passed a hardcoded value at the call site, so the parameter did
nothing. If you accept a parameter, use it or do not declare it.

### Tester — `Test-Toolkit.ps1`
`Toolkit` · `ObjectTypes: []` · none · **built in Phase 0 — additions only, no stage** · ~1,950 lines

The verifier, given a console button so it can be run without a shell.

By the finished state it carries **every rule in `03-SELF-TEST-SPEC.md`** —
including `PS7002`, `WF002` (control overlap) and `LOG001` (a write with no audit
call), all of which are already specified there. Do not treat those as extra
rules to invent.

**Three additions belong to this catalogue, and all three are needed BEFORE
Stage A:**

| Rule | What |
|---|---|
| `WF010` | Console height exceeds `WorkingArea` — §2. **Not `WF004`, which is already taken** and asks a different question (which API you call, not whether the result fits). |
| `WF011` | A section holds more buttons than `Rows × ButtonsPerRow`, a group counting as one — §3. |
| `LOG002` exemption | The rule must not flag a literal `-Target 'Toolkit'` — §4. The reference verifier already behaves this way; **check before changing anything.** |

> **The reference verifier uses its own rule codes** and several collide with
> `03`'s. See the note at the top of `03`: run and extend the reference, do not
> renumber it, and take the next free number in a family rather than reusing one.

### Report a Bug — `BugFix.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~449 lines · *sends one email; changes nothing else*

Reports something broken.

**One send. The operator is the `To:`, and the real destination is `Bcc:`.** That
is how the reporter gets their own copy without a second message, and it is why
there is no direct-reply path — **requests arrive through these buttons, not as
ad-hoc mail.**

**THE FIELD THAT MATTERS MOST IS "Steps to reproduce".** A bug that can be
reproduced is usually fixed the same day; one that cannot may never be. Even "it
happens on shared mailboxes but not user ones" halves the search. **Weight the
form accordingly** — that field must not look like one of several equal boxes.

Also ask for:

- **The exact error text, pasted rather than described.** "It said something
  about a domain controller" and the actual `0x52` message are different amounts
  of information.
- **What you EXPECTED.** Half of all bug reports turn out to be the tool doing
  what it was built to do, and knowing what the reporter expected instead is what
  turns that into a real answer rather than "working as designed".
- **Whether it happens every time.** An intermittent fault and a constant one are
  chased completely differently.

**"Blocking me right now" is a checkbox**, and it is the only thing on the form
that jumps the queue.

"You do not need to know what caused it. Report it anyway."

### Feature Request — `FeatureRequest.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~363 lines · *sends one email; changes nothing else*

Asks for something the console does not do yet. Three boxes, Send, and the same
`To:` operator / `Bcc:` destination pattern.

**The frequency question decides what gets built first.** "Twice a day" and "once
a quarter" are both fine answers — a small annoyance that happens constantly is
usually worth more than a large one that almost never does, and there is no way
to know which is which unless the requester says.

> **Both tools share ONE recipient key: `FeatureRequest.Recipient` in
> `Config\Tenant.json`.** Do not add a separate `BugFix.Recipient` — a second
> thing to configure is a second thing to get wrong, and the source deliberately
> collapsed them.
>
> **Scrub that address out of any error surfaced to the operator.** A failed send
> should not print the destination into a window the operator may screenshot.

---

## Infrastructure — no console button

### Deploy — `Deploy.ps1`
**Stage H** · ~633 lines · run from a shell

Builds a handout copy of the toolkit and zips it. **One build; everyone gets the
same one.**

Creates a staging folder beside the current one —
`adminconsole3.1_2026-08-13_0930`, version and timestamp — so what shipped and
when is visible at a glance, and an older build is still there to go back to.
Zips it and drops a `README.txt` inside for the recipient.

**No name prompt.** An earlier version asked who the copy was for and baked the
name into the folder. Wrong twice over: the build is identical whoever receives
it, and a name baked into a folder is stale the moment it lands on their machine.
**The timestamp is a BUILD label, not an install path.**

**Excluded from the handout, and why:**

| Excluded | Reason |
|---|---|
| `Logs`, `Reports`, `RetryQueue` | your history, not theirs — they start empty |
| `DisabledTools` | tools deliberately out of service *here* |
| `Config\Settings.json` | per-engineer paths and email |
| `Config\Themes.json` | **per-deployment, like `Settings.json`.** Everything else in `Config` — `Tenant.json`, `Tools.json`, `Sections.json`, `QuickLaunch.json` — is tenant or layout fact and does travel |
| `Publish-Toolkit.ps1` | stamping is the master copy's job |
| **`Deploy.ps1`** | **copies do not make copies** |
| `*Cache*.json` | stale scan data is worse than none — theirs rebuilds on first use |

**`UnBlock.ps1` IS included** — it is the first thing the recipient needs,
because files out of a downloaded zip arrive blocked by Windows.

**NOTHING IS EVER OVERWRITTEN.** Every run makes a new timestamped folder, so a
build can be reissued, compared, or gone back to. Old ones are deleted by hand.

**The manifest (`Version.json`) is REBUILT for the copy, not copied** — otherwise
the recipient's self-test reports the excluded files as missing **forever.**

### Publish Toolkit — `Publish-Toolkit.ps1`
**Stage H** · ~605 lines · run from a shell · **the last thing before handing the toolkit on**

**RUN THIS ON THE MASTER COPY ONLY.** On someone else's copy it re-stamps their
edits as official, which is exactly backwards. **This is the only thing that
writes `Version.json`.**

In order:

1. **UNBLOCK every file.** Anything downloaded or copied from another machine
   carries a mark-of-the-web tag that makes PowerShell refuse to run it. The
   files actually blocked are listed — that list should match what was just
   downloaded, and anything surprising on it is worth a look.
2. **SHOW what changed** since the last stamp.
3. **CONFIRM WHICH TENANT** this build is for, and **refuse to continue unless
   `Config\Tenant.json` already says the same thing.**

   > This is the one thing nothing else can catch. `Config\` is excluded from the
   > version manifest, and `Test-Toolkit` reads the tenant names *out of*
   > `Tenant.json` — so a toolkit built against one tenant and published for
   > another would ship the first one's domains and servers **past every other
   > guard, and work perfectly, at the wrong company.**

4. **LIST tools this tenant can never run**, so they can be removed before
   shipping rather than greyed out on someone else's screen.
5. **ASK minor or major** and derive the number from the last one — 1.03 becomes
   1.04, or 2.0. Nothing to type, so nothing to mistype.
6. **STAMP `Version.json`** with the version, the tenant, and a SHA-256 per
   script.

   **Unblocking happens BEFORE hashing, on purpose.** Whatever clearing the tag
   does or does not do to a file, the manifest then describes the bytes exactly
   as they will be handed over.

---

## 6. Deliberately excluded

Do not build these. If you think one is needed, ask first.

| Script | Why not |
|---|---|
| `PimActivation.ps1` | **Programmatic PIM self-activation is impossible in this tenant.** PIM for Groups is a Graph-only API, and the tenant will not consent to `PrivilegedEligibilitySchedule.Read.AzureADGroup` / `PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup` (both admin-consent delegated permissions). The retired AzureADPreview cmdlets covered directory *roles*, not groups. **`activate.ps1`, which opens the PIM portal, is the supported route and is already in scope** (C19/C20). |
| `Convert-LogHistory.ps1` | A one-time migration from a two-column audit-log schema **that has never existed at this tenant.** Nothing to convert. |
| `UnifiedGroupMembership.ps1` | Superseded by `UnifiedGroup.ps1`, which covers the same ground. Building both produces two buttons that disagree. |
| `Settings.ps1` | A 52-line shim. Settings are edited through Tool Manager. |
| Anything Graph-based | C20. Consent is the blocker, not capability. Deep links (C19) are the workaround and need no consent. |

**`UnBlock.ps1`** (11 lines — note the capital B in the filename) is a helper
shipped *inside* handouts, not a tool. Include it in Deploy's output; it needs no
button, no help block and no spec.

---

## 7. Definition of done — the whole toolkit

Beyond the per-script definition in `00-BUILD-PROMPT.md` §10:

- [ ] Every button in `Tools.json` points at a script that exists
- [ ] Every `.ps1` beside the console is one of: **registered in `Tools.json`**;
      **launchable from Dashboard**; **opened by another tool**
      (`Group-Manager-Editor.ps1` is the one — Group Manager routes to it and it
      refuses to open standalone); **infrastructure** (`Common.ps1`,
      `AdminConsole.ps1`, `Publish-Toolkit.ps1`, `Deploy.ps1`, `UnBlock.ps1`,
      `Test-Toolkit.ps1`); or **listed in `UNVERIFIED.md`** with a reason
- [ ] **Tool Status shows no unexplained yellow.** It reports "nothing loads
      this — safe to park" for anything not on a button, which is correct for
      `Group-Manager-Editor.ps1`, `Deploy.ps1` and `UnBlock.ps1` and **wrong to
      act on**. Parking the group editor silently breaks Group Manager
- [ ] `Sections.json` capacity ≥ button count **per section**, counting a group
      as **one** button
- [ ] Total console height fits `WorkingArea` on a 1920×1080 screen with a
      taskbar
- [ ] The verifier passes, **and can still be made to fail** on a planted defect
- [ ] Every tool that writes has an audit row on success **and** on failure
- [ ] Every audit `Target` is a primary SMTP address, or a documented scope name
      for the tools with no mailbox in play (§4)
- [ ] `UNVERIFIED.md` lists every cmdlet used that could not be confirmed against
      Microsoft Learn, and every assumption made about tenant shape
- [ ] No tenant name, domain or server in any `.ps1` — `TEN001`/`TEN002` clean

**A toolkit where fifty buttons work and five are honestly listed as unbuilt is a
success. A toolkit where fifty-five buttons exist and five of them quietly do the
wrong thing is not, and is considerably harder to recover from.**

<br>

==============================================================================

# ◆ DOCUMENT 9 OF 10 — `06-TOOL-SKELETON.ps1`

> A complete annotated tool in the required shape.
>
> *Standalone document. Referred to elsewhere as `06`.*

==============================================================================

*This document is a PowerShell file. Reproduce it as `06-TOOL-SKELETON.ps1`, not as markdown.*

```powershell
﻿#Requires -Version 5.1

# TOOLKIT-HELP-RUNS: reads one mailbox. Changes nothing.
# TOOLKIT-HELP-BEGIN
# One or two sentences an OPERATOR would find useful, in plain language.
#
# Say what question this answers and when it is the wrong tool. The second half
# matters more than it looks: somebody who reaches for the wrong tool concludes
# the data is missing rather than that they are in the wrong place.
#
# Read-only. Nothing here changes anything.
# TOOLKIT-HELP-END

<#
    <ToolName>.ps1
    One line on what this is.

    Reasoning: ScriptNotes/<ToolName>.ps1.notes.md
        -- OPTIONAL and out of scope for the first pass. The convention is that
        the long reasoning lives beside the toolkit in a notes file while the
        script keeps only short guardrail comments. Do not create the folder
        unless you have something to put in it; an empty convention is worse
        than none.

    THE SHAPE OF THIS FILE IS THE POINT. Every tool in the toolkit follows it,
    so an operator who has read one can read any of them, and a reviewer knows
    where to look for the parts that matter -- the connection guard, the
    handler wrapping, and the audit call.
#>

# -----------------------------------------------------------------------------
# Parameters
# -----------------------------------------------------------------------------
# DECLARE ONLY WHAT THIS TOOL USES. The console asks each script what it
# declares and passes only that. A script declaring SOME of the common
# parameters rejects the others with a binding error, while one declaring no
# param() block at all silently absorbs extras into $args -- which is why a
# fixed-set launcher works right up until somebody writes a normal script.
param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -----------------------------------------------------------------------------
# Shared layer
# -----------------------------------------------------------------------------
# SEARCH UPWARD, NEVER COUNT '..'. A hardcoded climb encodes the depth of the
# folder tree into every file at the bottom of it, and fails in the worst
# available way: the wrong number of levels usually lands on a directory that
# EXISTS, so Join-Path succeeds and the error names a plausible folder instead
# of saying the file moved.
$ToolkitRoot = $PSScriptRoot
while ($ToolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $ToolkitRoot 'Common.ps1'))) {
    $parent = Split-Path -Parent $ToolkitRoot
    # Without this guard a drive root or certain UNC forms loop forever.
    if (-not $parent -or $parent -eq $ToolkitRoot) { $ToolkitRoot = ''; break }
    $ToolkitRoot = $parent
}
if (-not $ToolkitRoot) { throw "Common.ps1 not found above '$PSScriptRoot'." }
. (Join-Path $ToolkitRoot 'Common.ps1')

# -----------------------------------------------------------------------------
# Connection guard
# -----------------------------------------------------------------------------
# A TOOL THAT WRITES ON-PREMISES MUST REFUSE, NOT WARN, when a cloud session is
# live in the same process. Both editions export identically-named cmdlets and
# the last import owns the name, so a tool can report "connected to
# on-premises" truthfully and still send every call to the cloud. See C1/C2.
#
# Delete this block in a read-only tool. Keep it in anything that writes.
#
# A tool that cannot share a process declares it in its own header, in the
# first 40 lines, and the console obeys that rather than a checkbox somebody
# has to remember.
# Started by the console, or double-clicked in Explorer? A tool run on its own
# has no looked-up object, no directory detection and no connection guard, so it
# fails somewhere confusing instead of at the front door. Courtesy, NOT a
# security control -- see the contract note in 08.
if (-not (Test-ConsoleLaunch)) { return }

$exo = Test-ExoConnectionState
if (-not $exo.Connected) {
    [System.Windows.Forms.MessageBox]::Show(
        "Not connected to Exchange Online. Use the Connect button on the console first.`r`n`r`n$($exo.Detail)",
        'Not connected', 'OK', 'Warning') | Out-Null
    return
}

# AND IF THIS TOOL WROTE ON-PREMISES, it would also have to REFUSE -- not warn
# -- when a cloud session is live in this same process, because both editions
# export the same cmdlet names and the last import owns them (C1/C2):
#
#     if ((Test-EmsAvailable) -and (Test-ExoInProcess)) { ...refuse... }
#
# This tool is read-only against the cloud, so the guard is shown and not used.

# -----------------------------------------------------------------------------
# Work
# -----------------------------------------------------------------------------
function Get-ToolData {
    <#
        Keep the work in functions and the UI below. It makes the tool
        readable and the logic testable without a form.
    #>
    param([Parameter(Mandatory)][string]$Identity)

    # RESOLVE ONCE, THEN BIND ON A KEY THAT CANNOT BE RE-INTERPRETED (C4).
    # A returned object's .Identity renders as the AD CN -- no domain, no
    # address -- and passing it back is a fresh forest-wide name lookup that
    # can resolve to a DIFFERENT object, not a bind.
    $mbx = Get-Mailbox -Identity $Identity -ErrorAction Stop
    $key = $mbx.Guid.ToString()

    # ENUMERATING CALLS NEED ResultSize EXPLICITLY (C3). The default is 1000
    # and the short answer is silent. A call that fetches ONE object by
    # identity does not need it.
    $perms = @(Get-MailboxPermission -Identity $key -ResultSize Unlimited -ErrorAction Stop |
               Where-Object { -not $_.IsInherited })

    return [PSCustomObject]@{
        Guid        = $key
        Primary     = "$($mbx.PrimarySmtpAddress)"
        DisplayName = "$($mbx.DisplayName)"
        Permissions = $perms
    }
}

# -----------------------------------------------------------------------------
# UI
# -----------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = '<Tool Name>'
$form.ClientSize = New-Object System.Drawing.Size(840, 560)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

$txtId = New-Object System.Windows.Forms.TextBox
$txtId.Location = New-Object System.Drawing.Point(12, 12)
$txtId.Size = New-Object System.Drawing.Size(300, 24)
$txtId.Text = "$SmtpAddress"
$form.Controls.Add($txtId)

$btnGo = New-Object System.Windows.Forms.Button
$btnGo.Text = 'Look up'
$btnGo.Location = New-Object System.Drawing.Point(320, 11)
$btnGo.Size = New-Object System.Drawing.Size(100, 26)
$form.Controls.Add($btnGo)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(12, 48)
$txtOut.Size = New-Object System.Drawing.Size(816, 500)
$txtOut.Multiline = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.ReadOnly = $true
# WORD WRAP ON, ALWAYS (B6). A real diagnosis once sat entirely in the second
# clause of an error message, off the right edge of a window with wrapping
# disabled, and three rounds of fixes went to the wrong layer. The cost of a
# horizontal scrollbar is that nobody scrolls.
$txtOut.WordWrap = $true
$form.Controls.Add($txtOut)

# -----------------------------------------------------------------------------
# Handlers
# -----------------------------------------------------------------------------
# EVERY HANDLER IS WRAPPED (B1). An exception inside a WinForms handler is
# caught by .NET rather than PowerShell, so the operator gets a CLR crash
# dialog -- a hundred lines of internals, no script name, no line number.
# Show-HandlerError converts it back and writes a Failed audit row, because an
# error the operator dismissed is otherwise the one failure that leaves no
# trace at all.
$btnGo.Add_Click({
    try {
        $id = "$($txtId.Text)".Trim()
        if (-not $id) { return }

        $data = Get-ToolData -Identity $id

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("$($data.DisplayName)  <$($data.Primary)>")
        [void]$sb.AppendLine('')
        foreach ($p in $data.Permissions) {
            [void]$sb.AppendLine("  $($p.User)  --  $($p.AccessRights -join ', ')")
        }
        $txtOut.Text = $sb.ToString()

        # TARGET IS THE PRIMARY SMTP ADDRESS, NEVER A DISPLAY NAME. With
        # several mailboxes in play a display name does not identify what was
        # looked at, let alone what changed.
        Write-CompletedLog -Action 'Read' -Target $data.Primary -Result 'Success' `
            -Detail "$(@($data.Permissions).Count) explicit permission(s)"
    }
    catch {
        # Never an empty catch. Surface it AND record it.
        Show-HandlerError -ErrorRecord $_ -Context 'Mailbox lookup'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox lookup'
        Write-FailedLog -Action 'Read' -Target "$($txtId.Text)".Trim() -ErrorRecord $_
    }
})

# -----------------------------------------------------------------------------
# Show
# -----------------------------------------------------------------------------
Set-ConsoleTheme -Form $form
# BOTTOM-LEFT. The console owns the top-left corner and stays there; a tool
# that also opens top-left lands on top of it and hides the connection
# indicators you may need to read while the tool is running.
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title '<Tool Name>' | Out-Null

[void]$form.ShowDialog()
$form.Dispose()
```

<br>

==============================================================================

# ◆ DOCUMENT 10 OF 10 — `07-ACCEPTANCE-CHECKLIST.md`

> For the OPERATOR, not the agent. Windows validation.
>
> *Standalone document. Referred to elsewhere as `07`.*

==============================================================================

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

<br>

==============================================================================

# END OF BUILD PACK

10 documents. Reference code is in `DEVIN-REFERENCE-CODE.md` and `DEVIN-REFERENCE-CODE-2.md`.
