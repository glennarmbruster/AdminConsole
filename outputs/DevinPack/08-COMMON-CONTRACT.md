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
