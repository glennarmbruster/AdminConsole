# Test-Toolkit.ps1

Static self-test for a toolkit folder. Nothing is executed, nothing is
connected to. Run it before every ship.

Built 2026-08-03. Point it at any of the three toolkit folders.

---

## Why this exists, priced honestly

One day's work on this toolkit produced twelve defects. Nine of them were
mechanically detectable:

| Defect | Rule |
|---|---|
| `ResourceDuplicator.ps1` unterminated `<#` — dead file, undetected for days | `SYN001` |
| `ManageOwnerGroups.ps1` was CP1252 with no BOM | `ENC001` |
| Its `param()` did not match what the console passes — would fail on click | `CFG004` |
| `Disable-User.ps1` `try` with no `catch` | `SYN001` |
| `$lblDCHint.Visible` set before the control existed | `REF001` |
| `Import-ScriptCapabilityCache` calling itself | (dropped — see below) |
| `` `u{1} `` — PowerShell 6+ syntax in a 5.1 script | `VER001` |
| `Get-AcceptedDomain` unbounded, silently truncating | `EXO001` |
| `-ResultSize 1` warning noise | `EXO002` |

The three it cannot catch are logic errors: connect buttons gated on the
connection they create, a splat written as `@(...)`, and a probe whose warning
was correct all along.

## The check that matters most

```powershell
[System.Management.Automation.Language.Parser]::ParseFile($File, [ref]$tokens, [ref]$errors)
```

PowerShell's **own parser**, without executing anything. Real errors, real line
numbers, no false positives.

Every brace-counting and regex approximation used before this — including the
ones used to verify this toolkit all through the day it was written — is a poor
substitute. Those kept producing imbalances in `Common.ps1` and `LogViewer.ps1`
that had to be hand-waved away as artifacts of a naive string stripper. The
parser has no such problem, because it is the thing that decides.

## Calibration cost two rounds, and that is the point

**First run: 6 errors, 74 warnings. Every one of the 6 was a false positive.**

`REF001` flagged `$lblConn.Text` and `$txtOut.Text` being set before their
controls were created — but all six were **inside function bodies**, which run
at call time, long after the form is built. The rule did not model scope.

Scoped to top-level statements only, it drops to zero false positives and still
catches the real one (`$lblDCHint`, which genuinely was at top level).

`HDL001` fired 67 times on "handler with no try/catch", most of them one-liners
like `$btnClose.Add_Click({ $form.Close() })`. Sixty-seven warnings nobody reads
is the same as none. Narrowed to handlers that call Exchange or AD **directly**:
down to a handful worth looking at.

`EXO001` still cannot see through splats. `OwnerlessGroups.ps1` sets
`$params['ResultSize']` conditionally and calls `Get-UnifiedGroup @params`;
`NewUser.ps1` puts `Identity` in a hashtable. Both were flagged, both are
correct code. **The rule is advisory, never a gate** — an AST-based version that
resolves splat variables would fix it and is the obvious next improvement.

`FUN001` self-recursion was dropped entirely: doc-comment examples that name
their own function (`Get-EdgeProfiles`, `New-ProgressReporter`) made it
unreliable, and the AST would be needed to do it properly.

**A linter's first job is to earn trust.** One false positive teaches people to
ignore it, and then the true positives go unread too. Two rounds of removing
rules was the work, not overhead.

## What it found on the real toolkits

After calibration: **0 errors** across all three, and these warnings worth a
look:

- `Set-MailboxAccess.ps1` and `onedriveaccess.ps1` contain literal **U+FFFD
  replacement characters** on five lines each. These files have already lost
  characters to a bad encoding round-trip — the originals are gone, not merely
  misread. Both are superseded and unregistered, so it is cosmetic, but it is
  the same failure `ManageOwnerGroups.ps1` had.
- `MessageTrace.ps1` uses `ResultSize 5000` with no `-WarningAction`, so it will
  print the "more results available" warning on every run — exactly the noise
  that cost a day to track down.
- `Common.ps1` line 45 contains an en-dash, inside a comment that is *about* Word
  inserting en-dashes. Correctly flagged, deliberately kept.

## Rules that need judgement rather than obedience

Printed at the bottom of every run, because a rule nobody understands gets
disabled rather than fixed:

- `ENC005` — non-ASCII is sometimes deliberate.
- `HDL001` — a handler delegating to a function that catches its own errors is
  fine.
- `EXO002` — a probe that deliberately asks for one row is correct. Silence the
  warning; do not raise the limit.

## Running it

**As a button.** Name `Toolkit Self-Test`, Script `.\Test-Toolkit.ps1`, Section
`Reporting`. Browse to any toolkit folder. Writes a timestamped CSV and an audit
row like every other report.

**Headless, for the ship pipeline:**

```powershell
.\Test-Toolkit.ps1 -Path ..\CloudOnly -Quiet
if ($LASTEXITCODE -ne 0) { 'do not ship' }
```

`-FailOn Error` (default), `Warn`, or `Never`. Exit code 1 means stop.

## Still open

- **`EXO001` should use the AST** to resolve splatted hashtables. Until then it
  is advisory and will occasionally cry wolf on correct code.
- **No Tier 4 smoke test.** Dot-sourcing `Common.ps1` into a fresh runspace and
  asserting the exported function count would catch a load-time error that
  static analysis misses — the `Import-ScriptCapabilityCache` recursion, for
  instance. Left out because a runspace that dot-sources a file has side
  effects, and a test that can change state is not a test you run casually.
- **No version stamp.** The intended companion is writing a build number on a
  clean run so a tenant can say which build they have and whether it was green.
- **`FUN001`** (a tool redefining a `Common.ps1` function) is implemented in
  `Test-CommonContract` and is the highest-value contract rule here — it is
  what would have caught the `ManageOwnerGroups` logging bug months earlier.

## TEN000 / TEN003 -- proving the leakage rule actually ran

A predicted six TEN002 findings did not appear in a Test-Toolkit run. The six
leaks were real (`@ExpressPros.com` in three `Projects\DLtoSharedMbx\`
scripts) and the regex, Tenant.json and file discovery all checked out, so the
clean result was not evidence of clean code.

Two defects came out of it, both about a check that can fail invisibly:

1. `$names` was populated inside an empty `catch {}`. Any failure reading
   Tenant.json left the list empty, TEN002 could then never fire on any file,
   and the run reported a pass. Now: a read failure or an empty list raises
   **TEN003 (ERROR)** and returns, and a successful load emits **TEN000 (INFO)**
   naming every value it is matching on. "No TEN002 findings" and "TEN002 never
   ran" are no longer the same output.

2. `Test-TenantLeakage` was handed `FileInfo` objects while its sibling
   `Test-CommonContract` was handed `.FullName` strings -- an inconsistency that
   already bit once. The loop now normalises to a full path up front and uses it
   for both `Get-Content` and the findings.

`Projects\` is now exempt, and that is a decision rather than an accident. The
folder is tenant-scoped by construction and is pulled before the toolkit ships
elsewhere, which is the same declaration an `EX-` prefix makes. The difference
was that `EX-` was chosen and this was not -- those files simply never produced
a finding. TEN000 now reports the skipped count so the exemption is visible and
arguable, with the standing caveat: it is only safe while `Projects\` really is
pulled before shipping.

Method note: the cause of the original miss was never proven. Brace-counting and
a Python re-implementation are not `[Parser]::ParseFile`, and this session had
already been burned once treating an approximation as the parser. The fix makes
the failure mode impossible to hide rather than claiming to have diagnosed it.

## INFO findings were never rendered

The TEN000 note added to prove the leakage rule had run did not appear in the
next report. The rule had run; `Format-TestSummary` simply never printed it.

`Add-Finding` has accepted `INFO` since it was written, and the report loop read
`@('ERROR', 'WARN')`. Every INFO ever raised was recorded and discarded at
the last step -- including **FUN002**, the rule that flags a standalone tool
carrying its own copy of a Common function. Those have been silently invisible
for as long as the rule has existed.

Worse, the clean path returned at `Clean. Safe to ship.` before reaching the
loop at all, so the diagnostics vanished precisely in the run where knowing
which checks ran is the whole point of the sentence.

Now: `INFO` renders as its own section, the clean banner no longer returns
early, and the Result line carries a note count. Exit codes are unchanged --
`INFO` must not gate shipping.

This is the third instance of one fault in this file: a check that passes
because it cannot see, cannot run, or cannot report. The tool that exists to
catch that fault had it in its own reporting layer.

## VER001 no longer flags itself

The finding message spelled the escape sequence literally, and
`Remove-PowerShellComment` strips comments but not string literals -- so the
text inside a finding is still code to the scanner. The message is now
assembled from `[char]0x60`.

The cost was one permanent ERROR, which matters more than one: `-FailOn Error`
gates shipping, so once the remaining encoding errors are cleared the toolkit
could never have reported a clean run. A check that is always red is a check
people learn to scroll past.

## FUN003 -- a function that was never written

`Get-MailDomain` was called in 23 places across 12 files and defined in none.
The tenant refactor replaced hardcoded domains with calls to a helper that was
never added. AdminConsole threw on launch.

Every call site is inside a string -- `"...@$(Get-MailDomain)"`. A subexpression
is not evaluated until the string is built, so the file PARSES, the encoding is
clean and the manifest hashes fine. Nothing in the tester could see it.

It was quiet in a second place too. `Get-SpoAdminUrl` uses it as a fallback,
inside a `try`, so the SPO settings hint read "nothing could be derived" -- which
looks like an unset option, not a missing function. The screenshot that prompted
the tenant-agnostic complaint was showing this bug.

A sweep of every Verb-Noun token in the toolkit found `Get-MailDomain` was the
ONLY undefined call. Three others looked undefined and were not: two appear only
in comments discussing them, and `Get-ADServerParam` is guarded by `Get-Command`
before use.

### Why the rule is two passes

FUN003 resolves in the order defined-here, allowlist, `Get-Command`. Using
`Get-Command` is what makes it affordable -- it knows `New-Object` without a
list. The allowlist covers server cmdlets absent until a session connects;
flagging `Get-Mailbox` because Exchange is not loaded would make the tester
useless offline.

Scanning matters more than resolving:

- **Pass A** reads inside every `$( )`, which is code even in a string. Without
  it the rule misses the exact bug it exists for.
- **Pass B** reads the line with string literals removed.
- **Neither pass reads a here-string body** for pass B. First attempt flagged
  seven lines of help text -- "copy the sender or Message-ID", "the SMB-Access
  group" -- because `Remove-PowerShellComment` strips comments and a here-string
  is a string, so prose arrives looking like code. Pass A still reads them,
  since `$( )` in a double-quoted here-string is evaluated.

Measured before shipping: raw scanning gives 5,166 findings, anchoring without
`Get-Command` gives 5,166, and the final form gives 0. A rule that fires on
prose is a rule that gets switched off.


## LOG001 / LOG002 -- audit coverage (2026-08-09)

`Write-CompletedLog` and `Write-FailedLog` had been in Common.ps1 for months,
and the comment above them states the intent plainly: "a log that only records
successes cannot be used to reconstruct a bad afternoon, which is the one job an
audit trail has."

The code said that. The tools did not do it. `Write-FailedLog` was called by
exactly ONE script out of sixty, while six writers held sixty-six catch blocks
between them. Six more scripts under `Projects\DLtoSharedMbx\` create
mailboxes, create and delete distribution groups, and grant permissions without
writing a single row.

This is the house failure in its purest form: nothing broken, nothing thrown,
and the gap invisible because the only place it shows is a log nobody reads
until they need it. Whether an action was logged cannot be noticed by USING the
console. It can only be noticed by a check.

### Severity was chosen, not defaulted

LOG001 is WARN and LOG002 is INFO, not ERROR. `-FailOn Error` gates shipping,
and a rule that lands red on a script with a defensible reason not to log -- a
preview-only tool, a report whose filter happens to name a write cmdlet -- is a
rule that gets switched off, after which it catches nothing. VER001 taught that
once already.

### Strings and here-strings are removed before matching

Tool help text is full of cmdlet names, and the ? button help lives in
here-strings, so stripping comments alone is not enough. Counting prose as code
is how the "23 of 24 tools unguarded" miscount happened.

### Its own try

`Test-AuditCoverage` is wrapped separately from `Test-Policy` (INT002, not
INT001). A here-string tracker is the kind of code that throws on one odd file,
and losing every audit finding in a folder because of it would reproduce exactly
the silent gap the check was added to find.

### Verified against a hand scan

A hand count found six scripts that write and never log. The rule as implemented
returns those six and no others.
