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
