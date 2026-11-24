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
