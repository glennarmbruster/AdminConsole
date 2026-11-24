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
