

---

## 45. 2026-09-12 -- brought back from the second tenant's console

Everything below was built in the second tenant's console first and carried
here the same day. None of it has been run here either.

### 45a. Config additions go through a MERGE, not a file

Eight new tools and five new themes had to get into `Tools.json` and
`Themes.json`, and 44a says no update may ship those files. So the update
ships `Config\Additions-2026-09-12.json` and `Add-ConsoleAdditions.ps1`,
which reads the live files, adds only what is missing, and writes them back
after a dated backup under `Config\Backup\`. Without `-Apply` it only prints
the plan. Run twice, the second run does nothing. `Themes.TEMPLATE.json` got
the five themes directly, because a template is not a live file.

The tools land in `Daily Scripts` / `Admin Tools` / `Reporting` as a first
guess and a section's `Rows` is grown if it can no longer hold its tools;
never shrunk. Accents are applied only to sections that have none.

### 45b. What arrived

- **Tools:** `QuarantineManager`, `MigrationStatus`, `MailFlowConnectors`,
  `DynamicDLPreview`, `SpamPolicyViewer`, `HybridAttributeCheck`,
  `DomainMailAuthCheck`, `FreeBusyViewer`. Learn-checked, not run. Two
  facts from that check: the quarantine type value is `HighConfPhish`, and
  `Set-MoveRequest -BadItemLimit` is retired in Exchange Online.
- **Common.ps1:** `Set-FormWindowChrome` (dark caption bar when the theme is
  dark, DWM rounded corners on Windows 11, no-op on 10; hooked from
  `Set-ConsoleTheme` on `HandleCreated`) and `Show-InputDialog` (one-line
  prompt, used by two of the new tools).
- **AdminConsole.ps1:** an eighth view, *What I can run right now*, grouped
  by what each tool needs; optional per-section `Accent` in `Sections.json`
  (heading colour, panel tint); "signed in as" beside the title; Ctrl+F to
  the Find box and Esc to clear it from anywhere; click a section heading
  to fold it (`[+]`), searching unfolds; window position and enlarged size
  remembered at close. Fold state and window bounds live in the per-operator
  `ConsoleState.<operator>.json`, so nothing shipped changes.
- **ToolManager.ps1:** an Accent column in the Sections dialog, because that
  dialog rebuilds `Sections.json` and would otherwise delete the key on Save.
- **Themes:** Navy, Mist, Powder, Sapphire Light, Sapphire Bright.

### 45c. The one thing the second console has that this one does not

Its `Common.ps1` carries a `#region Tenant additions` below everything above,
with 36 functions the ported tools call. None of the eight tools brought here
use any of them (checked by name against this file); `Show-InputDialog` was
the only one, and it is now above the line in both.

### 45d. Look Up IS Unified Lookup

The console's Look Up button opens Unified Lookup with the text AS TYPED and
the search already run. "glenn" finds every Glenn, exactly as typing it in
Unified Lookup does. The console's own resolution -- the pass that greys the
grid and hands every other tool an address -- runs afterwards and ONLY when
the text has an `@`. The first cut ran it first on everything, and on a
fragment it appended the mail domain, hunted for glenn@..., and reported a
miss while Unified Lookup had already found six of him. Order matters here.

`Unified-Lookup.ps1` gained a `param` block (`SmtpAddress`, `Environment`,
`DomainController`). It had none, so `Get-ToolParameterSet` passed nothing
and the console's context was swallowed into `$args`. A tool that wants the
console's address MUST declare it.

The 140px output box under the input row is hidden -- `$txtOut` still
exists because eleven places write to it, and its `TextChanged` raises the
text as a dialog (quietly, to the status bar, while Look Up is running) so
nothing can vanish into an invisible control. `SectionStartY` went from 308
to 164.

### 45e. Three silent rules that got a voice

- **Arrange refused drops into a full section, silently.** Rows x columns
  was the cap and nothing on screen said so. A drop into a full section now
  grows it by one row; Save Arrangement writes the new `Rows` to
  `Sections.json` (read and written whole, so `ButtonsPerRow` and `Accent`
  survive). Never shrunk.
- **Arrange only worked on tools that were available.** A disabled WinForms
  control raises no mouse events at all, so every greyed-out button was
  nailed to the grid and the layout you could reach depended on what you
  happened to be connected to. Arrange mode now enables every tool button
  for the duration -- arranging is not running -- and `Update-ToolButtonState`
  stands down while it is on, so a lookup or a background capability refresh
  cannot re-grey a button under the operator's mouse mid-drag. The real
  state is recomputed when Arrange goes off rather than parked and restored:
  a parked answer goes stale the moment anything else happens.
- **Tool Manager forbade sorting, for a real reason** -- Order was renumbered
  from row position and Save wrote it. Every row now carries its saved place
  on `.Tag`; Order, Save and the JSON read `.Tag`, never the row index, so
  any column can be sorted as a view (Section sorts in Sections.json order).
  Move Up/Down are enabled only in saved-order view.
- **The A-/A+ stepper in every view changed the font only after the next
  click** because its handler referenced the locals of the function that
  built the buttons, which are gone by the time the message loop fires it.
  State now rides on the button and is read through `$this`. Same class of
  fault as 14c; the tell is "works on the second action, not the first".

### 45f. Small ones

- `Get-PimDurationHours` reads `Tenant.json` `Pim.DurationHours` first, then
  `Settings.json`, then 4. It only read Settings, so a tenant file saying 9
  counted down from 4.
- `Start-ToolScript` refreshes the PIM header label when a tool returns, so
  a tool that starts or clears the countdown shows it at once.
- `TenantManager.ps1` Save carries through every top-level key it has no
  field for, and every `OnPremExchange` key it does not edit. It carried
  `Pim` and `FeatureRequest` by special case; the rule is general.
- Navy's `AccentBack` was 4.21:1 against white and failed the readability
  gate; now `#2A62B8` (5.94:1).

### 45g. The self-test was reporting 180 errors and none of them were real

A full `Test-Toolkit` run on 2026-09-12 produced ~180 ERRORs. One was a real
bug. The rest were the tester being wrong about itself, which is the worst
kind of finding: it buries the one that matters.

**The real one.** `Member_BulkADD_CSV.ps1` had a type cast in front of a
COMMAND -- legal in front of a method, which is what the line was before the
MessageBox sweep rewrote it. The file had not parsed since that sweep and
nothing in it could run. SYN001 found it; nobody had opened the tool.

**PRM001 (~180 of them).** The rule that checks `-Parameter` names against
Common.ps1's `param()` blocks read those blocks with a regex, and the regex
started its paren-matching walk one character short -- on the `(` itself. The
depth counter opened at 1, sailed past the real closing paren and read to the
end of the FUNCTION BODY, so every local variable looked like a parameter and
every parameter written `[string]$Text)` -- closing paren on the same line --
was missed entirely. `Get-TargetSmtp` was reported as declaring `-t` when it
declares `-Text`; `Get-ConfigFilePath` as declaring nothing. It now reads the
param blocks from the AST, which this same file was already using twenty
lines away in `Get-TopLevelParameterName`. A parser knows what a parameter
is; a regex is guessing.

**FUN003.** `Get-ComplianceSearch` and family, the in/outbound connectors,
`Get-QuarantineMessageHeader` -- real cmdlets that exist only once a session
is connected, which is precisely what the allowlist is for. Added. And
`Send-As` / `Send-To` were prose: permission names, in captions and audit
text, with `Send` being an approved verb. The rule's own docstring predicted
this failure; there is now a short list of hyphenated terms that are not
commands.

**EXO002.** `-WarningAction SilentlyContinue` beside a deliberately bounded
`-ResultSize`, at the sites that lacked it. Applied per hashtable, not per
line -- several already had the key on the following line, and a second one
in the same literal is a "duplicate keys" parse error, which is how a tidy-up
becomes an outage.

**Left alone, deliberately.** The six FUN001 redefinitions (a local
`ConvertTo-OuPath` in AD-Lookup and Unified-Lookup, `Get-Prop` in
TenantManager, and so on) are real -- the local copy wins and a fix in Common
never reaches the tool -- but they are not equivalent copies. TenantManager's
`Get-Prop` defaults to `''` where Common's defaults to `$null` and treats a
present-but-null value as absent. Deleting the local one changes behaviour in
a config editor, and none of this can be tested today. They stay, with the
finding standing as the reminder. The two HDL001 handler warnings are the
judgement calls the rule says they are.

The lesson is the one Test-CommonContract's own docstring already states:
twelve findings of which two are real is how a rule gets ignored. A checker
that is wrong at scale is worse than no checker, and it has to be held to the
standard it holds the toolkit to.

---

## 46. ONE SCRIPT, TWO CONSOLES -- and the script is the whole delivery

As of 2026-09-12 this toolkit and the second tenant's console are the same
code. The differences are countable: `Common.ps1` there carries a
`#region Tenant additions` with 36 functions below the marker, its Connect
section has Set-PimTimer where this one has Activate, it has no Tenant
Editor, and the four `Config\*.json` files are its own. Everything else --
every shared function, every view, every convention in this document -- is
byte-identical.

That makes a new tool portable by default, and it is worth keeping it that
way deliberately rather than by luck.

### 46a. THE DELIVERY IS THE .ps1 FILE. NOTHING ELSE.

**Never ship `Tools.json`, `Sections.json`, `Settings.json` or `Tenant.json`
to either console** -- 44a, and it matters more here than anywhere. Both
installs have been arranged by hand: buttons dragged, sections grown a row,
accents chosen, paths set. A new tool arriving with a Tools.json would
replace all of that to add one button.

So a new tool ships as ONE FILE, and the person who receives it adds the
button themselves in their own Tool Manager. The covering note needs exactly
two facts: **which section** the button belongs in, and **what the tool needs
connected**. It cannot be added for them from here and must not be.

The same is true in reverse: a tool built for the second tenant arrives here
as a file, and gets a button here the same way.

### 46b. What makes a tool portable

- **Build against the reference `Common.ps1`** -- everything above the
  tenant marker. Those functions exist in both consoles at the same
  version. A tool that calls one of the 36 functions BELOW that marker runs
  in the second console only, and will not say so: it fails at the click,
  in a handler, looking like a connection fault.
- **Keep organisation vocabulary out from the start** (44a-ii). A tool
  written here with this org's names in its comments cannot cross without a
  scrub, and a scrub is a second edit nobody remembers to make. Write
  `contoso.com`, `ABCD`, `WXYZ` the first time and it travels.
- **Declare the console's context if the tool wants it** --
  `param([string]$SmtpAddress, [string]$Environment, [string]$DomainController)`.
  A script with no param block gets nothing and does not know it (45d).
- **Do not assume a section name.** The two consoles have different sections.
  An unknown section in Tools.json does not error -- `Get-ToolSection` drops
  the tool into the first section, which looks like the console losing it.
  Name the section in the covering note; let the operator pick the real one.

### 46c. What is NOT portable, and should be said out loud when it applies

A tool that is genuinely tenant-specific -- one that needs a function from
the tenant region, or a directory only one tenant has -- is fine, but it has
to be labelled as such when it is handed over. The failure mode for an
unlabelled one is silent: it is dropped into the wrong console, it appears on
the grid like any other button, and it breaks the first time somebody in a
hurry presses it.

### 46d. Server names were the last thing hiding in code

The second tenant's build had four Exchange hostnames written as literals in
five separate scripts -- Check Queues, Clear IIS Logs, Exchange Health Check,
onprem mailtest, and one loop in Unified Trace. Five copies of one fact, none
of them in `Tenant.json`, which is the exact failure 44a-ii describes: the
toolkit looked portable and was not, and the symptom on another tenant would
have been an admin-share timeout rather than anything naming the cause.

`Get-TenantExchangeServerList` in `Common.ps1` now answers it once:
`OnPremExchange.Servers` from `Tenant.json` if present, else `Get-ExchangeServer`
from an open on-premises session, else nothing -- with `ExcludedServers`
subtracted either way and matching that ignores case and domain suffix. An
empty list is reported to the operator, because a scan of zero servers that
reports success is worse than an error.

The key is absent-means-derive, so this toolkit needs no `Tenant.json` change
to benefit: connect OnPrem and the four tools work against whatever this
directory actually has.

The second tenant's copy also carries `markdown\SERVERNAMES.md`, a record of
the names that were removed. It is a deliberate, tenant-local exception to the
vocabulary rule -- a record of what was taken out of the code, which would be
useless with the names taken out of it. It does not exist here and should not.

### 46e. A third button style: Enhanced (3D)

`Settings.ConsoleButtonStyle` now takes `Enhanced` as well as `Raised` and
blank. Raised was reported as "VERY subtle" on 2026-09-12, and it is -- a 20%
lift off the face colour is most of nothing on the darker palettes.

Enhanced is not a second drawing routine. `Get-RaisedButtonDepth` returns a
multiplier (1.0 for Raised, 2.1 for Enhanced) and `Add-RaisedButtonChrome`
applies it to each of the four numbers that make the bevel -- the gradient's
two ends and the two edge tones -- clamped below a blend of 1.0, because a
pure white top edge and a pure black bottom one stop reading as a lit surface
and start reading as a fault. Past about 1.5x the colours are near that clamp
and stop adding anything, so depth also widens the bevel to two pixels and
draws an inner highlight line just inside the top edge, which is the part that
actually says "lit from above". One routine, so a fix to the raised look is a
fix to the enhanced one.

`Test-RaisedButtonStyle` answers true for both, which is what kept the change
to one file: every other place that asks "is this button raised" -- the corner
radius, the glyph handler standing down, the Danger outline redraw -- needed
no edit.

**The depth is resolved at ATTACH time and parked on the button's `Raised3D`
state, beside the corner radius.** The first version asked for it inside the
paint handler, which stats `Settings.json` on every hover of every button --
exactly what the comment above `Radius` already warns against. Writing that
warning down did not stop it being broken on the same afternoon; reading the
state object is the only version that is cheap.

### 46f. The footer, twice over

Two small things, one cause each.

**The operator name was on the version line AND on the status strip.** It was
put on the version line as backlog #187, before the bottom strip existed; the
strip now ends with `DOMAIN\user` permanently. Removed from the version line,
which is also the line that was running into its own ellipsis.

**A band of dead space under the version line.** The initial fit gives back
any height the reflow freed -- except it was skipping windows restored from
saved bounds, on the grounds that the size was the operator's choice. That
reasoning is right about width and position and wrong about height: the saved
height came from a session with a different number of rows in it, so a
restored window kept a gap nobody could remove except by resizing it by hand
at every launch. Height now follows the content always; width and position
still stay where they were put.

### 46g. The last of the tester's own findings

With PRM001 reading the AST and the allowlist filled in, a full run came back
with ten findings. Eight were still the tester being wrong, and both causes
are now fixed.

**PRM001 x2 -- real code reported as broken.** `AuditUser.ps1` and
`UserDefinedDelegates.ps1` each carried their OWN copy of the function they
were calling. The local `Invoke-MailboxAuditSearch` declares
`-ProgressCallback`; Common's declares `-OnProgress`. The local
`Resolve-Target` declares `-Identity`; Common's declares `-Smtp`. The calls
bind to the local copies and always have. The rule was checking them against
a function they never reach, and now skips any name the file defines itself.

**HDL001 x2 -- not directory calls at all.** MobileDevices' loop is already
inside a try/catch. Unified-Lookup's `Update-HandOffButton` calls
`Get-CloudEditPlan` and `Get-MailboxSettingsPlan`, which touch nothing: they
string-match a row and return a plan object. Left alone; the rule's own
JUDGEMENT clause is what it is for.

**FUN001 x6 -- real, and now fixed by renaming rather than deleting.** Every
local copy genuinely diverges from Common's: different parameters, different
return shapes, a third to a fifth of the size in two cases. They are older,
differently-shaped functions that happened to share a name. Deleting one
would have bound its calls to Common's incompatible version and failed at the
click -- which is exactly the trap the shadowing set for whoever tidied up
next. So they are renamed, bodies untouched:

| File | was | now |
|---|---|---|
| `AD-Lookup.ps1` | `ConvertTo-OuPath` | `ConvertTo-AdOuPath` |
| `Unified-Lookup.ps1` | `ConvertTo-OuPath` | `ConvertTo-AdOuPath` |
| `AuditUser.ps1` | `Invoke-MailboxAuditSearch` | `Invoke-AuditUserSearch` |
| `AuditUser.ps1` | `ConvertTo-MailboxAuditRecord` | `ConvertTo-AuditUserRecord` |
| `TenantManager.ps1` | `Get-Prop` | `Get-TenantProp` |
| `UserDefinedDelegates.ps1` | `Resolve-Target` | `Resolve-DelegateTarget` |

Each carries a banner saying what it used to be called and why it is not
called that any more. Converging them onto Common's versions is still the
right end state, but that is a refactor with a test for each tool, not a
rename -- and the rename removes the trap in the meantime.

**A clean run now means something.** That is the whole point of the day's
work on this file: 180 findings of which none were real, then ten of which
six were, then zero.

### 46h. ButtonsPerRow is a LIMIT now, not a suggestion

The console used to divide the available width by the cell size and lay out
however many columns came out. `ButtonsPerRow` only set the DESIGN width, so
the number in `Sections.json` was a starting point and the real column count
was whatever the window, the theme and a scrollbar happened to allow. Set it
to 10 and the console drew 9, or 12 after a resize. Reported 2026-09-12.

Two changes make it exact:

- **`Get-ConsoleZoomFit` takes the column count as given.** It now finds the
  largest factor at which `PerRow` cells AND the whole stack of rows both fit,
  instead of deriving the columns from the width. Widening the window scales
  the buttons and their text up; it does not add a column. That is what "10
  columns" should mean, and it is the answer to "why is my grid a different
  shape on every machine".
- **`Set-SectionPanelGrid` forces the wrap point.** A FlowLayoutPanel wraps
  when the next control will not fit, so the usable width inside the panel is
  set to exactly `PerRow` cells -- the PerRow+1'th cannot fit, the PerRow'th
  always can -- and whatever the panel has beyond that becomes padding, half
  each side, so the grid stays centred. It is called at build, on every width
  change, and at the end of each zoom pass, always AFTER the buttons have
  been resized, because the padding is computed from the new cell size.

**Width and height are now two separate scales, and that is the point.** A
single zoom factor could not satisfy both: maximised, it could not grow past
what the HEIGHT allowed, so the grid sat in the middle of the window with
dead space either side; narrowed, it bottomed out at 1.0 and the panel stayed
wider than the window, so the flow dropped to seven columns -- the exact
behaviour the hard limit was meant to remove.

So `Get-ConsoleZoomFit` now answers one question, how tall may a row be, and
its floor is 0.6 rather than 1.0 -- a window narrower than the design has to
shrink its buttons, because the alternative is dropping a column. The button
WIDTH is not searched for at all: it is whatever divides the available width
into exactly `PerRow` cells, so the grid fills the row edge to edge at any
size. Type size takes the larger of the two factors, capped a fifth above the
vertical fit so the caption cannot outgrow the row it sits in.

Derivation order matters: padding first (from the height factor), then the
button width from what is left, because a cell is the button plus its two
margins.

**Section panels no longer AutoScroll.** A vertical scrollbar takes its width
out of the client area, which silently removed a column from the fullest
section -- the one place an operator would notice. There is nothing to scroll
to: the row count comes from the tool count, and a drag that overfills a
section grows a row (46e) rather than hiding the overflow behind a bar.

### 46i. The Catalogue says how it is sorted, and offers two more ways

It was always in console order -- section by section, then the order the
buttons sit in -- which is a reasonable default and was invisible. Nothing on
screen said so, and a hundred tools in an unannounced order reads as no order
at all. Reported 2026-09-12 as "I have no idea how it's sorted".

There is a Sort box on the bar now: **Console order** (unchanged default),
**Name**, **Modified**. The title line names the current one, so even the
default is now stated rather than assumed.

Modified is the file's `LastWriteTime`, stamped onto the model ONCE when the
view opens -- one stat per tool is nothing, one per keystroke in the filter
box across a hundred files is a stutter the operator would feel and never
connect to a sort they are not using. It sorts newest first, because the
question behind it is always "what changed", and the date appears in the list
rows only when it is what you sorted by; a date on every row the rest of the
time is a column of noise. It is in the detail pane always.

Two things fixed on the way:

- **The shell sized every ComboBox to five ems**, which fits the drop arrow
  and about four characters. Any dropdown with a real word in it opened
  showing an ellipsis. It measures its widest item now.
- **The right-hand cluster places right to left**, so the first control added
  sits furthest right -- the box has to be added before its label or the bar
  reads "[Console order] Sort:".

### 46j. Publish-Toolkit offers new tools a button

`Add-ConsoleAdditions.ps1` was written for one delivery and is now the general
answer to 44a's awkward corner: no update may ship `Tools.json`, and a new
script therefore arrives with no button. It reads what is there, adds only
what is missing, backs up first, and is a no-op on a second run.

It now finds work two ways. **`Config\Additions-*.json`** is a curated list
that names each tool and the section it belongs in -- what a considered
delivery comes with. **A scan** finds `.ps1` files in the folder that no
button points at, which is the safety net for a script dropped in and
forgotten. A scanned tool has no section, so it lands in the first one and
says so; the caption is the file name tidied up, which is a starting point
rather than a decision.

Infrastructure is never offered -- the console, `Common.ps1`, the publisher,
the tester, this script, the dot-sourced `*Query.ps1` libraries. Decline
anything else and it is written to `Config\Additions-Ignore.json` so it is
not raised again; delete that file to be asked afresh.

**Publish-Toolkit runs it as step 5**, after the stamp, and asks. That is the
right moment: the new scripts are in the folder because you just stamped
them, and the publishing copy is the one whose `Config\` is real. After
rather than before because a failed publish must not have edited
`Tools.json` -- a step that half-happens is worse than one that does not.
`Config\` is excluded from the manifest, so adding a button cannot
invalidate the hashes either way.

**One thing this does not solve.** Publish runs on the master copy only, and
an engineer receiving the build is told to copy the `.ps1` files and leave
`Config\` alone -- so they get the scripts and not the buttons. The script
ships with the build and they can run it themselves; that is the intended
route, and it is worth saying out loud in the covering note rather than
leaving them to find a tool that is present and unreachable.

### 46k. WHAT NEVER GETS A BUTTON, AND WHY THE SCAN CANNOT WORK IT OUT

Eleven scripts are already reachable from the console itself and must never be
offered a place on the grid. Ten open from the **Dashboard**:

| Script | |
|---|---|
| `Settings.ps1` | paths, title, button style, icons, dialogs |
| `TenantManager.ps1` | tenant facts -- capabilities, directories, licence groups |
| `ThemeManager.ps1` | palettes |
| `ToolManager.ps1` | buttons, sections, columns |
| `ToolStatus.ps1` | which scripts the console uses, and parking the rest |
| `QuickLaunchTool.ps1` | the header shortcuts |
| `ConfigBackup.ps1` | copy `Config\*.json` somewhere safe |
| `UnBlock.ps1` | clear the mark-of-the-web on this folder |
| `BugFix.ps1` | report a fault |
| `FeatureRequest.ps1` | ask for something |

These are **configuration and housekeeping, not tools you run against a
mailbox**. A second copy on the grid gives two routes to one window, invites
the question of whether they differ, and spends a button on something nobody
reaches for from the grid anyway.

The eleventh is **`Unified-Lookup.ps1`**, and it is on the list for the same
reason by a different route: since 45d the console's own **Look Up** button
opens it with whatever was typed and the search already run. It is the front
door. A grid button would open it EMPTY, sitting a few inches below a box that
opens it filled in -- two controls that look identical and are not, which is
the worse half of the pair winning by being where the eye lands. The file
stays where it is; the console launches it by name.

Note it is NOT the same case as the Dashboard ten. Those are reached from a
window the operator opens on purpose; this one is reached from a control on
the main screen that is used a hundred times a day. Both end at the same rule:
one route is enough.

**The scan cannot deduce this.** They are ordinary scripts with ordinary
`TOOLKIT-HELP` blocks and they look exactly like tools -- there is no property
of the file that says "the Dashboard already owns me". So the decision is
written down, once, in `$script:NeverOffer` in `Add-ConsoleAdditions.ps1`,
rather than being re-declined at every publish.

Two lists, and they are not the same thing:

- **`$script:NeverOffer`** is what the TOOLKIT has decided. It travels with
  the code, applies to every install, and is where a script goes when it
  earns a permanent home on the Dashboard. Expect it to grow.
- **`Config\Additions-Ignore.json`** is what ONE OPERATOR declined on THEIR
  machine. It is written when you answer no, and deleting it asks again.

Put a Dashboard tool in the operator file and the next console starts offering
it again. Put a one-off local preference in the code and you have made a
decision for five other tenants.

---

## 47. 2026-09-13 -- the reference console's own Tools.json, reorganised

### 47a. The one-time exception to the no-shipping-JSON rule

The standing rule is section 34's: **never ship `Settings.json`, `Tools.json`,
`Sections.json`, `Tenant.json`, `QuickLaunch.json` or `Themes.json` over a
working install.** New keys arrive absent-means-off, or through Tool Manager.
The rule exists because those files are the operator's, they drift per machine,
and an overwrite silently throws away work nobody can see was lost.

On 2026-09-13 that rule was suspended ONCE, explicitly and in writing, because
the preconditions that make it dangerous were all absent for one afternoon:

- the working copy and the reference copy had just been reconciled and were
  confirmed identical, so there was no per-machine drift to destroy;
- the whole `AdminConsole` folder was going to be copied over wholesale, so the
  edited files WERE the install, not an overlay onto it;
- the reorganisation being made -- splitting one 56-button section into five --
  is not expressible as a merge. Tool Manager can move a button; it cannot
  re-cut the sections and renumber every entry without a person doing 97 drags.

**This is not a precedent.** The next config change goes through Tool Manager or
through the merge path in 45a. If it ever needs suspending again, it needs the
same three conditions and the same explicit say-so, recorded the same way.

The pre-change files are kept at `ScriptNotes/Config-Backup-2026-09-13/`.

### 47b. What changed in Tools.json

**Sections, replacing the old Admin Tools / Daily Scripts / Reporting /
Projects.** Same six-section shape and the same accent colours as the second
tenant's console, so the two consoles look like siblings:

| Section | Accent | Buttons |
|---|---|---|
| Connect | `#3D9970` | 7 |
| Daily Scripts | `#2E6BD6` | 16 |
| Mailbox and Access | `#7B61C4` | 26 |
| Groups and DLs | `#C4772E` | 14 |
| Reporting | `#1F9BA8` | 26 |
| Projects | `#C24E4E` | 8 |

**Activate is the first button, Dashboard is the second.** Asked for directly.
Connect then runs Activate, Dashboard, ExO, OnPrem, SPO, MSGraph, Console --
which is close to the order they get pressed in.

**Seven duplicate buttons removed.** Each was the same script twice under two
captions, differing only in path case and slash style -- `./ownerlessgroups.ps1`
against `.\OwnerlessGroups.ps1`, and six more like it. They came from the
additions tool being run against a folder listing on a case-insensitive
filesystem: to the scanner `ResourceDuplicator.ps1` and the already-buttoned
`resourceduplicator.ps1` are different strings, so it offered a tool that was
already there. The readable caption survived in every case.

**`Unified-Lookup` lost its button** and is now reached from Look Up on the
console header, which is what that button has done since section 46's header
work. It is on the permanent `$script:NeverOffer` list so the additions tool
will not put it back.

**Two broken Projects paths fixed.** `Edit DL-Input.csv` and
`DL Inventory - All` were both missing their `Projects\` prefix and had never
worked from a button. Both now point into the live tree rather than the older
`Projects\ABCD\` snapshot -- `Projects\DLtoSharedMbx\dl-input.ps1` and
`Projects\Migration\Inventory-AllGroups.ps1`.

**Every `Script` value normalised to `.\ProperCase.ps1`,** matching the real
filename on disk. This is what stops the duplicate-button problem recurring:
the additions scanner compares strings, so the strings have to be the ones the
filesystem uses.

`Rows` in `Sections.json` is recomputed from the button count and
`ButtonsPerRow` (7), so no section opens with an empty row under it.

### 47c. Parking, and one file that must not be restored

Nineteen scripts were parked into `DisabledTools\`. Fifteen of them were
already parked -- a folder copy had put the root copies back, byte-identical --
and four were genuinely retired: `Connect-SCC.ps1`, `MailObject.ps1`,
`Object-Lookup.ps1`, `SearchAndPurge.ps1`, each verified as referenced only
from comments. Root is now 106 `.ps1`, `DisabledTools\` 20.

`DisabledTools\UnifiedGroupMembership.ps1` was **not** a parked copy of the
live tool -- it was an OLDER, different version (2026-08-13, 330 lines; the
live one is 2026-09-11, 354 lines). Restoring it from the park UI would have
quietly overwritten the newer file with a month-old one and the only symptom
would have been features going missing. It is renamed
`UnifiedGroupMembership.ps1.OLD-2026-08-13.bak` so nothing can restore it by
accident.

**The general hazard:** the park folder is a flat namespace with no version
check. Anything restored from it overwrites root by name. Before restoring,
compare dates and sizes -- the park does not.

## 48. The console header is two labelled rows

Reported 2026-09-13: the header had a full-width lookup box on one line and a
`Find tool:` box floating in the middle of the next, and it was not obvious
which box did which job. They do completely different things -- one finds a
person, the other filters the buttons -- and they looked like one control
block.

The header is now two parallel rows:

```
Unified Lookup:  [_____________________]          [ Clear ][ Look Up ]
Find tool:       [_____________________]  [Show all] [n of m] [View] [Arrange]
```

Each row starts with its own label at x=12. Both boxes start at the same x --
measured from the WIDER of the two labels -- and share one width, which is the
SMALLER of what the two rows can spare. Both rows end flush at the same right
margin as the button grid.

### 48a. One function owns both rows

`Set-ConsoleFindRowLayout` places everything, right to left, from `.Width` read
at the moment of placement. It is called after the theme pass, at the end of
every zoom, from `Update-SectionPanelWidth`, and at Shown.

There used to be a second copy of this arithmetic near the top of the script,
computing X positions from literal control widths. It was always stale:
`Set-ConsoleTheme` runs after it and changes button widths, so every position
derived from `74`, `90`, `50` and `62` was wrong by however much the buttons
grew. That block is gone. **If a control is added to either row, it goes in the
list in `Set-ConsoleFindRowLayout` and nowhere else.**

The "not enough room, give the find row its own line" fallback is gone with it,
along with the `SectionStartY` shift that went with it. It existed because Look
Up / Clear / Close occupied the first 318px of the find row; the find row no
longer shares a line with them and starts at x=12, which is roughly 250px more
room than the fallback was written to rescue.

`txtInput` also came out of the "controls that span the grid" list in
`Update-ConsoleZoom` and `Update-SectionPanelWidth`. It no longer spans the
console, and leaving it there made every zoom stretch it full width and then
immediately shrink it back.

### 48b. Close is gone, Clear is not

**Close was removed.** The window's own X does the same job, and the row is
worth more than a second way to shut the console. `$form.CancelButton` went
with it, which means Escape now has no default behaviour to override -- the
Escape-clears-the-filter handler stays anyway, because clearing the filter is
the useful thing for Escape to do and the console should not depend on which
buttons happen to exist.

**Clear stayed, because it is not decoration.** Look Up on a full address sets
`$script:LastSmtp` / `LastType` and un-greys the tool buttons for that object;
Clear is the only way to put them back to "nothing selected" without typing a
different address. Its tooltip now says so.

**Enter in the lookup box runs the lookup.** Bound to `$txtInput` directly and
not to `$form.AcceptButton`, because AcceptButton would fire Look Up from
anywhere on the console -- including from the find box, where Enter should do
nothing.

### 48c. Look Up sits above Show All; Clear is gone

```
Unified Lookup:  [_____________________]  [ Look Up  ]
Find tool:       [_____________________]  [ Show all ] [n of m] [View] [Arrange]
```

Right-aligning row 1 put Look Up over Arrange -- the control it has least to do
with. It now takes Show All's LEFT EDGE AND WIDTH, copied from the button after
row 2 has been placed rather than computed a second time. Each row is a labelled
box followed by that row's own action, in one column.

That copy also collapses the box width to one number. Both rows end at the same
x by construction, so what row 1 can spare and what row 2 can spare are the same
value -- the `[Math]::Min` of the two is gone.

**Clear was removed, but its behaviour was not.** Emptying the box is cosmetic;
what mattered was that Look Up on a full address sets `LastSmtp` / `LastType`
and un-greys the tool buttons for that person, and Clear was the only way back
to "nothing selected". Leaving the buttons pointed at somebody no longer on
screen is the state most likely to run a tool against the wrong mailbox. So it
is now `Clear-ConsoleTarget`, on **Escape in the lookup box**.

ESCAPE PRECEDENCE IS DELIBERATE. The form-level `KeyPreview` handler claims
Escape first, but only while the find box has text. Filtering the buttons,
Escape clears the filter; not filtering, Escape in the lookup box clears the
lookup and the targeting. The filter wins because abandoning a search is far
more frequent, and is the case where losing the other state would surprise.

Enter in the lookup box runs the lookup -- bound to the control, not to
`$form.AcceptButton`, which would fire Look Up from the find box too.

## 49. UnBlock.ps1 works on the folder it is in

It ran `Get-ChildItem "E:\Scripts\M365\AdminConsole"` -- one machine's path,
baked into a file that gets copied to other machines. On the machine it was
written for it worked, which is why it survived; anywhere else it unblocked
somebody else's toolkit, or nothing at all, and said so either way.

`$PSScriptRoot` is the fix: the script unblocks the folder it is sitting in, so
a second copy of the toolkit cannot unblock the first. It also gained the
`TOOLKIT-HELP` block every other tool has, so the console's `?` has something to
show.

**Still outstanding, and NOT fixed here:** seven scripts under
`Projects\DLtoSharedMbx\` carry `E:\scripts\m365\dlconversion` as a literal --
`New-SalesMB`, `NewCloudMB`, `6-NewCityState`, `CLOUD-managedby`,
`Set-Managers\Manager-set` and two more. They are project one-offs rather than
console tools, and none of them is on a button, so they were left alone rather
than changed without being run.

## 50. Searching must not resize the console

Reported 2026-09-13: *"when I search for a tool, the font changes and everything
gets larger; when I clear it or select Show all, the screen never goes back."*

Two faults, one root, and the second is the nastier of the pair.

### 50a. Control.Visible does not return what you assigned

`Update-ConsoleSearch` counted matches per section like this:

```powershell
$ctrl.Visible = -not $grouped
if ($ctrl.Visible) { $secVisible++ }     # WRONG
```

**A child control's `Visible` getter returns EFFECTIVE visibility** -- its own
state AND every parent's. A section emptied by the previous filter has a hidden
panel at this point, so every button in it read back `$false` however many times
it was set to `$true`. `$secVisible` stayed 0, the section was judged empty, and
the panel was hidden again. **Permanently** -- there was no state from which it
could recover, because every subsequent clear repeated the same reasoning.

The fix is to count the value being ASSIGNED, not the property read back:

```powershell
$want = -not $grouped
$ctrl.Visible = $want
if ($want) { $secVisible++ }
```

The filtering branch was already correct -- it counted `$match`, a local. Only
the unfiltered branch, the one that restores, had the bug. That is the shape of
it: the code that hides worked, the code that un-hides did not.

**The rule: never read a WinForms `Visible` back to learn what you just set it
to.** If the answer matters, keep it in a variable.

### 50b. The zoom fit was measuring the filtered grid

`Get-ConsoleZoomFit` sizes the console to the buttons that are visible. Filter a
hundred down to four and the content nearly vanishes, so the search finds room
for a factor of 3.0 and fonts, buttons and headings all balloon.

Filtering is not a resize. The operator did not ask for bigger buttons; they
asked which buttons match. So `$script:FilterActive` is set at the top of
`Update-ConsoleSearch` -- **before** anything there can call `Update-ConsoleZoom`
-- and while it is set the fit returns the factor already applied. Clearing the
filter unsets it and the refit happens against the full grid.

50a made this permanent rather than merely annoying: with whole sections reading
as empty, even the post-clear refit saw a nearly bare console and kept the huge
factor. Fixing either alone would have left half the complaint.

**Known and accepted:** resize the window WHILE a filter is on and the buttons
keep their old size until it is cleared. The alternative -- fitting against a
count of what WOULD be visible unfiltered -- means a second copy of the group
and fold rules in a function that has no other reason to know them.

### 50c. The logo and the '?' now follow the right edge

`Add-FormLogo` and `Add-ToolHelpButton` place their controls once, against
whatever `ClientSize` the form has at the time -- the DESIGN width, before saved
window bounds are restored and before any zoom. On a wider console they stayed
put and ended up mid-header beside Look Up instead of in the corner.

`Set-ConsoleFindRowLayout` now re-places both from the right margin on every
pass, using the same 12px Arrange uses, so the corner is one column. They are
found by `.Name` (`pbToolkitLogo`, `btnToolHelp`) rather than by variable,
because neither is a variable in this script's scope.

---

## 51. 2026-09-13 -- ServerHealth.ps1

Disk, patch and update state across a list of servers. Modelled on
Exchange-HealthCheck.ps1 -- same tabs, same one-round-trip-per-server
collection, same "a server that did not answer is a ROW, not a gap" rule -- but
asking operating-system questions rather than Exchange ones.

### 51a. A second server list, and why it is not the existing one

`OnPremExchange.Servers` means "servers the Exchange tools may run cmdlets
against", and three tools already read it with that meaning. Server Health asks
about disks and patches on a WIDER set -- domain controllers, file servers, a
jump box, things with no Exchange on them. Overloading one list would either
drag non-Exchange servers into Exchange tooling or keep them out of the health
check, so this is `ServerHealth.Servers`, read by `Get-TenantHealthServerList`.

```json
"ServerHealth": {
  "Servers": [
    { "Name": "SRV01", "Exchange": true,  "Notes": "hub" },
    { "Name": "SRV02", "Exchange": false }
  ]
}
```

A bare string is accepted too and reads as `Exchange: false`. Absent means off.

**`Exchange` is a HINT, not a gate.** The collector looks for the Exchange
install on every server it reaches and reports what it finds; the flag only
decides whether ABSENCE is a finding. A server marked true with no Exchange is
flagged; a server marked false WITH Exchange is also flagged, because the flag
must not be authoritative over the machine.

Tenant Manager edits it as one text box -- `SRV01*; SRV02` -- where a trailing
`*` means "expect Exchange". The `*` is parsed off on save and a real JSON
boolean is written, so nothing downstream has to know about a convention
invented for one field. Notes on existing entries are carried through by name.

### 51b. "Missing updates" means the server's own Windows Update agent

The check runs `IUpdateSearcher::Search("IsInstalled=0 and Type='Software' and
IsHidden=0")` inside the `Invoke-Command` on each server. So the answer is what
that server would install if you clicked Check for updates -- on a WSUS estate,
approved updates rather than everything Microsoft ships.

**The synchronous `Search` is supported for remote callers with administrator
rights; `BeginSearch`/`EndSearch` are not** (learn.microsoft.com, "Using WUA
From a Remote Computer"). That is why this uses the blocking call inside the
session rather than the async pair.

`IsHidden=0` leaves out updates an administrator declined on that server. They
are hidden because somebody decided so, and listing them as missing every month
is how a report trains people to ignore it.

**A failed search fails LOUDLY into `WuStatus` with its HRESULT**, and the
Overview shows `unknown` rather than `0`. "No updates pending" and "could not
ask" must not look the same -- and the code matters: `0x8024402C` is a WSUS it
cannot reach, `0x80072EE2` is a timeout, `0x80240044` is the call being
refused, and each sends you somewhere different.

The search is also most of the runtime, so a tickbox turns it off for a fast
disk-and-patch-date pass.

### 51c. Exchange build is REPORTED; currency is not claimed

The build comes from `ExSetup.exe`'s file version -- the method Microsoft's own
build-numbers page gives as option 2, and the only one that reflects an SU.
`Get-ExchangeServer`'s `AdminDisplayVersion` shows the CU and NOT the SU on top
of it, so a fully patched server and an unpatched one report the same string.

`Resolve-ExchangeRelease` names it from a table in the script, and the table's
own as-of date is printed in the summary. **It does not compute "behind by N".**
The only way to be right about that is to know what Microsoft has shipped since
the table was written, and the failure mode of guessing is a report that says
CURRENT next to a server three CUs behind. Refresh the table by re-reading the
Learn page; an unrecognised revision degrades to naming the CU plus the raw
build rather than to blank.

### 51d. Two thresholds per drive, not one

`Get-DriveState` trips on percentage **or** an absolute floor. Percentage alone
is wrong on very large volumes (10% of 4TB is 400GB and fine); an absolute floor
alone is wrong on small ones (20GB on a 60GB drive is a third of it). Requiring
both would miss the case each exists for, so it is an OR.

`Set-RowColourByState` then colours from the STATE COLUMN rather than
re-deriving the rule, which is where Exchange Health Check has the threshold
written twice. One rule, and it is visible in the report rather than only in the
colour.

### 51e. OS build needs UBR

`Win32_OperatingSystem.Version` stops at `10.0.17763` for every Server 2019 ever
built. The revision that says which patch level lives in the registry as `UBR`
under `CurrentVersion`, and without it two servers a year apart in patching
report an identical OS version.

---

## 52. 2026-09-13 -- one toolkit folder, several tenants (phase 1)

The JSON files were the only thing separating one tenant's console from
another's, so a second tenant meant a second copy of a hundred scripts -- and
every fix then had to be made twice, which is how two copies drift. This is the
first half of removing that: one folder, one set of scripts, a tenant chosen at
launch.

```
AdminConsole.ps1                     the default tenant
AdminConsole.ps1 -Tenant wxyzpros    by file name
AdminConsole.ps1 -Tenant wxyz        by nickname
```

### 52a. ABSENT MEANS OFF, AND THAT IS THE ENTIRE REVERT PLAN

Multi-tenant mode is on when `Config\Tenants` exists AND holds a `.json`.
Everything -- the switch, the per-tenant folders, the tenant in the title, the
write confirm -- asks `Test-MultiTenantMode` and does nothing when the answer is
no. `Get-TenantConfigPath` returns `Config\Tenant.json` exactly as it always
did.

So reverting is renaming one folder, and `REVERT-MultiTenant.ps1` does that --
renaming rather than deleting, and refusing to run if there is no
`Config\Tenant.json` to fall back to, because reverting into a console with no
tenant file at all is worse than the thing being reverted.

**An empty `Tenants` folder is NOT multi-tenant mode.** A folder somebody
created and left empty is not a configuration, and treating it as one breaks a
working console with nothing to fall back to.

### 52b. Why a global, and why it is set before anything reads config

`$global:ActiveTenantKey`, not `$script:`. The console launches every tool with
`&` into its own process -- there are no `NewProcess` entries in Tools.json --
so a global set once at startup is inherited by every tool without any of them
declaring a `-Tenant` parameter or remembering to pass it on. A `$script:`
variable in Common.ps1 would not be: a tool dot-sources its own copy and gets
its own script scope.

`Set-ActiveTenant` runs in AdminConsole.ps1 **immediately after Common.ps1 is
dot-sourced and before anything touches Tenant.json**, and it drops
`$global:TenantConfig`. Without that drop the console loads the default tenant,
caches it, and then changes the key underneath a cache nobody re-reads -- the
wrong-tenant bug in its purest form, and the same shape as the stale-cache fault
that made Server Health report an empty server list.

### 52c. A name that matches nothing STOPS the console

`Set-ActiveTenant` throws and lists the tenants it knows. It does not fall back
to the default, because a typo quietly opening a different organisation is
precisely the failure this feature has to avoid. `Resolve-TenantKey` matches the
file name, the nickname, or either with the domain suffix dropped -- `srgpros`,
`srgpros.com` and a nickname all find one file -- and returns `''` rather than
guessing.

### 52d. The files are the list; the index only decorates

`Config\Tenants\Tenants.json` holds nicknames and marks the default. It is
**not** the list of tenants: `Get-TenantRegistry` builds that from the `.json`
files actually present. So dropping a file in makes a tenant appear, deleting
one makes it vanish, and a stale index cannot hide a tenant that exists or
invent one that does not. With nothing marked default the first is used, because
an install with tenants but no default must still open.

### 52e. One tenant per folder, and a collision check anyway

`Resolve-ToolkitFolder` appends the tenant key to every folder it resolves --
Reports, Logs, RetryQueue -- including when the tenant set an absolute path,
since two tenants pointed at one share is the likeliest way to end up with mixed
reports. A CSV named for its report and its timestamp says nothing about whose
data is in it.

`Test-TenantFolderCollision` still runs, for the case the segment cannot catch:
a configured path that already ends in ANOTHER tenant's key.

Single-tenant installs get no segment and the same paths as yesterday. That is
what makes the folder rename a complete revert.

### 52f. The confirm is a setting, not a hardcoded dialog

`Confirm-TenantForWrite` names the tenant before a tool changes anything, and is
silent on single-tenant installs -- a prompt that always says the same thing is
one people learn to click through, which would make it useless on the day it
finally had something to say.

```json
"ConfirmTenantOnWrite": false
```

turns it off everywhere. Absent or true means on. It is meant to come off once
switching tenants is second nature, and taking it off should be one key rather
than a sweep through five scripts undoing a dialog somebody hardcoded.

### 52h. SWITCHING TENANTS MEANS A NEW POWERSHELL PROCESS

Raised 2026-09-14, and it is the most important thing in this section.

Swapping the config file does NOT swap:

- the Exchange Online session and its access token
- the on-premises PSSession and the cmdlets `Import-PSSession` proxied in
- `Connect-SPOService`
- `Connect-MgGraph`
- whatever PIM activated, for the roles you hold in THAT tenant

Those belong to the **powershell.exe**, not to the console, and closing the
console window does not end them. So running

```
.\AdminConsole.ps1 -Tenant wxyz
```

from a prompt that earlier ran the ABCD console gives you ABCD's live
Exchange Online session with WXYZ's configuration on top -- a console correctly
labelled one tenant while every cmdlet it runs reaches the other. That is worse
than either tenant alone, and nothing on screen would say so.

**`Set-ActiveTenant` now refuses it.** Changing tenant within one process, while
`Get-ToolkitLiveConnection` reports anything live, throws and prints the command
to run in a new window. The FIRST tenant in a process is always allowed --
whatever is connected then was connected deliberately, and this is not the place
to second-guess it. Only a CHANGE is blocked.

**Disconnecting instead was considered and rejected.** Tokens cache, modules
keep state, and `Import-PSSession` leaves proxy functions that still resolve
after the session they proxied is gone. "Probably clean" is not a basis for
pointing a disable tool at a directory. A fresh process has none of that by
construction.

`Start-TenantConsole` launches one, and Tenant Manager's **Open in new window**
button is a click for it -- naming, in the confirm, whatever this window is
signed in to, so the rule reads as a reason rather than ceremony. `-NoExit` is
deliberate: a startup failure on a shell that closes itself takes its own
explanation with it.

## 53. 2026-09-14 -- park in place (phase 2)

`ParkInPlace.ps1`, on the Dashboard. A tool is hidden from ONE TENANT's console
while the file stays exactly where it is.

### 53a. Why this is not DisabledTools\

`DisabledTools\` MOVES the file, for the whole install, for every tenant. That
is right when a tool is retired. It cannot say "the on-premises Exchange
tools are real tools worth keeping, they are simply meaningless in a cloud-only
tenant" -- because one folder now serves several tenants, the file is shared,
and the answer is not.

So parking is a LIST, per tenant, of scripts the console behaves as though it
had never been given. Un-park by removing the entry; nothing was moved, so
nothing has to be found and put back.

`Config\Tenants\<key>.parked.json` in multi-tenant mode -- beside the tenant
file, so the two travel together and deleting a tenant takes its exclusions with
it -- or `Config\Parked.json` on a single-tenant install.

### 53b. Matching is on the BARE FILE NAME

Tools.json says `.\Foo.ps1`, the scanner says `Foo.ps1`, a hand-edited file
says `./foo.ps1`, and all three mean one file. `Get-ParkedScriptIndex` keys on
the lowercased leaf. Getting this wrong shows up as a park that silently does
nothing, which is the failure mode hardest to notice.

### 53c. Filtered at the source, not at the buttons

`$script:AllTools` is filtered immediately after Tools.json is read. Filtering
where the buttons are built would leave every other consumer -- the Catalogue
view, the find box's "n of m", Arrange, the section row maths -- counting tools
that are not on screen. One list, filtered once, is the only version where those
agree.

`Add-ConsoleAdditions` skips parked scripts too, which is half the point: a
parked tool that gets re-offered on every publish is not parked. Note the
distinction from `$script:NeverOffer` -- that is what the TOOLKIT has decided
for every install, permanently; this is what THIS TENANT decided, and it changes
when the tenant does.

### 53d. Park by capability, because parking one Connect button is a trap

Parking `Connect-ExSE.ps1` alone leaves twenty tools on the grid that cannot
work without it, each greying out with its own explanation at the moment
somebody presses it. So the tool offers the CAPABILITY: pick "On-premises
Exchange" and every script declaring `TOOLKIT-VARIANT: OnPrem` is ticked.

The variant is read from each script's own header rather than from a table here,
so a tool added next month is grouped correctly without this file knowing it
exists.

It TICKS rather than saves -- the grid then shows exactly what is about to
happen, and Save is still a separate press.

### 53e. No select-all, and a floor

Asked for explicitly, and right. Parking every tool produces an empty console
and is never something somebody meant to do in one click. So:

- no select-all box, and the tick column is `SortMode = NotSortable` (sorting by
  the tick invites "sort, then drag down the column", which is select-all with
  extra steps)
- Save REFUSES a list that would park everything, and says so

`$script:NeverPark` hides the console, Common.ps1, this tool itself and the
config editors from the list entirely. A parking screen that can lock you out of
parking is not a safety feature. Hidden rather than shown-and-refused, because a
row you cannot use is a row somebody spends time trying to use.

### 53f. The model is the truth, the grid is a view

Ticks live in `$script:Model`, not in the DataGridView. If they lived in the
grid, typing in the filter box would rebuild the rows and throw away every tick
made so far -- silently, which is the worst way to lose work. The sort is
preserved across redraws for the same reason.

### 53g. Still to do

- Wiring `Confirm-TenantForWrite` into the write tools. The function exists and
  is tested; no tool calls it yet.
- Tenant colour on the console for non-default tenants.
### 53h. What Test-Toolkit checks, and the claim I got wrong

An earlier note here said Test-Toolkit would report a parked tool as an orphan.
**That was wrong** -- Test-Toolkit has no orphan check at all. Finding scripts
with no button is `Add-ConsoleAdditions`, and that already skips parked ones.

The check actually worth having is the opposite direction, and is now CFG006: a
parked entry naming a script that is NOT THERE. Parking matches on the file
name, so a renamed or deleted script leaves a dead entry that parks nothing --
and the tool it was meant to hide is quietly back on that tenant's console.
Nothing else can report that, because the console cannot tell "parked" from
"never mentioned", by design.

It reads EVERY tenant's `*.parked.json`, not the active one: the tester runs
from a prompt with no console, so there is no active tenant, and asking for one
would check a single file and call the estate clean.

**A parked script still listed in Tools.json is NOT a finding.** That is how
parking works -- one Tools.json for every tenant, filtered per tenant at load.


### 53i. Sections park independently of scripts (2026-09-14)

The question was "if I remove a section will it remove it only from one
tenant? Tools.json should be tenant specific -- please verify it is."

**It is not, and that is deliberate.** Tools.json is one file for every tenant
(53c). A Tools.json per tenant means adding a tool in N places and N chances to
forget one. What IS per tenant is the parked list, and it now carries two
independent lists:

```
{ "Parked":   [ { "Script": ".\\Projects\\X\\Foo.ps1", "Reason": "", "When": "" } ],
  "Sections": [ "Projects" ] }
```

- **A parked section** drops its heading and every tool whose `Section`
  matches, including tools added to it next month. The console removes the
  name from `$script:SectionNames` and the tools from `$script:AllTools` at
  load, so Catalogue, find-box counts and Arrange all agree.
- **A parked script** is one tool. Project scripts under `Projects\<name>\`
  are now listed and parkable -- the first version scanned the root only.
- The two do not imply each other. Park the section and leave its scripts
  unticked; park half its scripts and leave the section up; or both. Un-park
  the section and the individually parked scripts stay parked.

`Save-ParkedScript -Entries -Sections` writes both every time, because it
replaces the file. A caller that sent only scripts would silently un-park
every section; Park In Place therefore always sends both.

### 53i-fix. The parked section left a blank band and a scrollbar

First test: Projects parked, heading gone, but the window was the same height
and now scrolled. The filter had removed the name from `$script:SectionNames`
while the panel loop, the row maths and the window height all read
`$script:SectionCfg.Sections` directly -- so the panel was still built, empty,
and its rows still counted. The parked names are now removed from
`SectionCfg.Sections` itself, before anything is derived from it. Save
Arrangement re-reads Sections.json from disk, so the in-memory cut cannot leak
into the file. The last remaining section is never parked.

### 53j. 53b corrected: the key is the path relative to the root, not the leaf

53b said matching was on the bare file name. That was wrong the moment
project scripts became parkable: `1-ONPREM-Export.ps1` exists in more than one
project folder, and a leaf-only match would have parked all of them when one
was ticked. `Get-ParkedKey` now normalises to the lowercased path relative to
the toolkit root, no leading `.\`, backslashes -- `foo.ps1` for a root tool,
`projects\x\foo.ps1` for a project tool. Root keys are identical to the old
leaf keys, so an existing parked list reads unchanged. Test-Toolkit's CFG006
resolves the same relative path.

### 53k. NeverPark is three names, not eleven

The first list also kept the publisher, tester, managers and
`REVERT-MultiTenant.ps1` off the grid, and the first question was why the
revert script could not be parked. There was no good answer: parking moves
nothing, the Dashboard launches those by path rather than by button, and the
only real cost of parking them is that Add-ConsoleAdditions stops offering
them -- which is the point. NeverPark is now `AdminConsole.ps1`, `Common.ps1`
and `ParkInPlace.ps1`: the three where parking would lock you out of
un-parking. `REVERT-MultiTenant.ps1` and `Diagnose-Theme.ps1` were added to
Add-ConsoleAdditions' NeverOffer, which is why the revert script was being
offered in the first place.

### 52i. The tenant crosses the process boundary in an environment variable

52b said a `$global:` reaches every tool because the console launches them all
in-process. **That is true of the tool grid and false of the Dashboard.** The
Dashboard opens Theme Manager, Tenant Manager, Park In Place, Tool Manager and
the tester through `Start-ToolDetached` -- a new `powershell.exe`. In those
processes the global is empty, `Get-TenantConfigPath` falls back to
`Config\Tenant.json`, and every one of them was silently editing a file the
console does not read.

Found 2026-09-14 as "I change the theme and nothing happens": Theme Manager had
written the theme, correctly, into the wrong tenant file.

`Set-ActiveTenant` now also sets `$env:ADMINCONSOLE_TENANT`, and
`Get-ActiveTenantKey` reads it when the global is empty. Environment variables
are inherited by child processes however they are started, so nothing in
`Start-ToolDetached` had to change.

Two guards: the variable is always WRITTEN, never only-when-unset, so a stale
value from an earlier console in the same window is overwritten rather than
trusted; and it is only HONOURED when a matching tenant file exists, so a
leftover from a deleted tenant reads as "no tenant" rather than as a path to
nothing.

## 54. 2026-09-14 -- the tenant band, and one file per tenant

### 54a. Every window says which tenant it is on

With two consoles open side by side, a Disable-User form for one tenant looked
exactly like the other's: the tenant was in the console title and nowhere
else. `Add-TenantBand` (Common.ps1) puts a 6px strip of colour along the top
of every form and appends `[label]` to the title bar when the label is not
already there. It runs at the end of `Set-ConsoleTheme`, which every form
already calls, so no tool changed. Single-tenant installs get nothing.

Colour: the tenant file's `Band` key (hex, set in Tenant Manager), else one of
eight palette colours picked by hashing the tenant key, so two tenants that
never chose still differ. The band is tagged NoTheme so the theme walk leaves
it alone, and it is added AFTER the walk for the same reason.

This also closes the environment-variable worry from 52i: whatever a process
resolved as its tenant is now visible on every window it opens.

### 54b. Everything about a tenant is in its one file

The parked list was a fourth place to look (`<key>.parked.json`). It is now a
`Parked` block inside the tenant file:

```
"Parked": { "Scripts": [ { "Script": ".\\Foo.ps1", "Reason": "", "When": "" } ],
            "Sections": [ "Projects" ] }
```

`Get-ParkedStore` reads the block; if the tenant file has none, it reads the
old side file, so nothing already parked is lost. The first save writes the
block and renames the side file to `.migrated`. Test-Toolkit's CFG006 now
walks every tenant file (and the old side files, for tenants that have not
saved since).

`Set-TenantFileKey -Name -Value` is the one writer for any per-tenant key set
by a tool other than Tenant Manager: read the file, set the key, write it
back, clear the cache. Theme Manager now uses it for `Theme`; Park In Place
for `Parked`. Tenant Manager already round-trips keys it has no field for.

Single-tenant installs use `Config\Tenant.json` for the block too; the standing
rule that Tenant.json is never shipped over a working install is unchanged.

### 54c. Not done, deliberately

Themes.json still carries `FollowTenant` / `TenantThemes`. It is a second way
to say what the tenant file's `Theme` now says. Left for a separate pass
because removing it touches Theme Manager's UI, and the per-tenant Theme key
wins over it already.

## 55. 2026-09-15 -- the cloud module is the one behind the LIVE connection

Change Meeting Organizer failed in the console with "a server side error has
occurred" and worked at a prompt. `Get-ConnectionInformation` in the console's
process showed four Exchange Online connections, three expired. The module
keeps a `tmpEXO_*` module for every connection ever made; `Get-ExchangeModuleMap`
took the FIRST cloud-looking module it found, and cached it for the life of
the process. Calls through an expired module's proxies come back as the
generic server-side error, not "token expired" -- and `Get-Mailbox` happened
to resolve well, which made it look like a calendar problem for two rounds.

Three changes: `Get-ExchangeModuleMap` asks `Get-ConnectionInformation` for
the Active connection's `ModuleName` and uses that module, falling back to
the export scan only when no connection reports one; the cached map is
re-checked against `Get-Module` so a reconnect cannot leave it naming a module
that is gone; `Get-EXOCommand`'s fallback refreshes the map instead of taking
the first `tmpEXO_*`. `Connect-ExO.ps1` clears expired connections before
signing in (all of them, when none is live; with `-Force` when one is) and
refreshes the map after.

Also learned from the live cmdlet: `Invoke-ChangeMeetingOrganizer -Subject`
is a SUBSTRING match, and its multiple-match output offers a `-MeetingId`
(hex GOID) that Microsoft Learn does not document alongside `-EventId`. The
tool accepts either and sends `-MeetingId` only if the loaded cmdlet has it.

### 55a. The actual cause: two calendar cmdlets reject -ErrorAction

After the session was proven clean (one connection, right account, right
module), a same-process comparison found it: `Invoke-ChangeMeetingOrganizer
... -WhatIf` works; the identical call with `-ErrorAction Stop` -- typed or
splatted -- returns "a server side error has occurred". `-ErrorVariable` is
accepted. `Get-CalendarDiagnosticObjects` behaves the same. `Get-Mailbox`
and everything else in the toolkit tolerate `-ErrorAction`, which is why the
toolkit's habit of `ErrorAction = 'Stop'` in every splat had never bitten.

`Invoke-EXOCommandNoEA` (Common.ps1) strips `ErrorAction` from the splat,
runs the cmdlet with `-ErrorVariable`, and throws the first collected error,
so callers keep their try/catch. Use it for any REST cmdlet that fails
generically when the session report shows nothing wrong. The session
hardening in 55 stays: it was real, just not this.

## 56. 2026-09-15 -- one toolkit, a view of it per person (phase 1)

Two operators on one folder: the second one to drag a button silently
rearranged the first one's console, because Arrange wrote straight back to
the shared `Tools.json` and `Sections.json`. The fix is a per-person overlay.

### 56a. What is shared and what is not

`Tools.json` and `Sections.json` still describe what the toolkit HAS, and
they are still shared -- adding a tool is one edit in one place, and a tool
renamed or retired there is renamed or retired for everyone. What is per
person is how that toolkit is ARRANGED: sections and their rows and accents,
columns, which tool sits where, which tools hide behind a group button, which
are hidden, and the theme.

### 56b. Config\Users\<windows-account>.json

Keyed on the Windows account, not the admin UPN: the UPN needs a live
connection, is not known at startup, and changes with which admin signed in,
while the layout belongs to the person at the keyboard.

A `Default` block and a block per tenant key. The tenant block wins KEY BY
KEY, not merged within a key -- if it names Sections, its Sections are used
whole. Half-merged layouts cannot be predicted or explained.

**Absent means unchanged.** No file, no overlay, and the console behaves
exactly as it did before any of this. That is also what Reset restores.

### 56c. Composition order at load

Shared `Tools.json` and `Sections.json`, then the tenant's parked block, then
the person's profile. Two rules fall out:

- A tool the TENANT parked stays gone whatever the profile says. Tenant
  parking is a fact about the tenant; personal parking is taste
  (`Get-EffectiveParked` unions the two and nothing subtracts).
- A tool the profile does not mention keeps its `Tools.json` section, so a
  tool added next month appears for everybody without anyone editing a
  layout. Absence means 'not moved', never 'hidden'.

### 56d. Arrange asks once

With no profile, the first Save Arrangement asks: just me, or everyone?
'Just me' creates the profile and every later drag goes there silently;
'everyone' writes the shared files as before. Asked once, because a dialog on
every drag is a dialog nobody reads.

The profile stores PLACEMENT ONLY -- section, order, group per tool -- never a
copy of Tools.json. `Theme` and `Parked` are carried forward on every write,
because other windows own those and rebuilding the block without them would
quietly undo whatever they set.

### 56e. Customizations.ps1: save and load

On the Dashboard. Save writes the current arrangement to a file the operator
keeps; Load puts one back as their profile, for one tenant or all.

**It works with no profile too, and that is the point.** It captures whatever
the shared files currently give you -- which is the migration path for a
console somebody arranged the old way, by editing the shared files, that is
about to take a build replacing them. Save before the upgrade, load after.

The file declares `Kind = AdminConsoleCustomizations` and is checked before it
is trusted: it came off a desktop and could be any JSON at all.

### 56f. Still to do

- Tool Manager's section editor and Theme Manager still write the shared
  files. Until they are redirected, section edits and theme picks made THERE
  are still everyone's; Arrange and Park In Place are the per-person paths.
- `Add-ConsoleAdditions` still reads the tenant parked list only, so a tool
  one person hid personally is still offered to them on the next publish.

## 57. 2026-09-15 -- three display bugs found on one FH machine

An engineer at FH had a console nobody else could reproduce: headings buried,
captions cut mid-word, the version line sliced through. Devin worked it with
him and the record they kept (AdminConsole-LookAndFeel-Changes.txt) is worth
reading. Three separate bugs, all pre-existing, all in this file. Ported here
2026-09-15 and mirrored.

### 57a. Labels and Buttons do not size themselves; TextBoxes do

The root cause, and the reason it looked like three other things first. Every
Label and Button in the header carried a HEIGHT WRITTEN AS A 96 dpi LITERAL --
18, 24, 34 -- while the theme sets the font in POINTS. At 192 dpi a point
renders about twice the pixels, so the text needed 47px in a 22px box. The two
input boxes were always fine because a single-line TextBox derives its own
height from its font. That is the whole tell.

The zoom scaled box and text by the same factor, so it neither caused nor
fixed it: this was never a factor problem. `Set-ConsoleFindRowLayout` now
measures every label and button from the font it is actually wearing,
grow-only, exactly as the text boxes already did.

**The machine reported "96 dpi / 100%" to everything that asked**, because the
process asking was DPI-unaware and Windows answered in virtualised units.
Measure inside the process or do not bother.

### 57b. Section headings: a 20px box and a hardcoded 22px offset

Same class of bug, different control. `$lbl.Size = ...20)` with the panel
placed at `$y + 22`, and nothing ever sized the box from the text. The theme
re-faces the label, the accent pass makes it bold, and `Update-ConsoleZoom`
scales the heading FONT while never touching its Height -- so the heading
clipped, which reads exactly like the panel below has been drawn over it.

`Set-SectionHeadBox` measures 'Xg' (cap-height to descender, so every heading
gets the same box whether or not its name has a descender) IN THE FACE THE
THEME WILL APPLY, which is what lets this be a single pass instead of a
re-stack. `Get-SectionHeadOffset` replaces the literal 22 at all four stacking
sites: build loop, Update-ConsoleSearch, Update-ConsoleZoom, Get-ConsoleZoomFit.
`Expand-LabelToFont` does the same grow-only job for the version footer.

Re-stacking after the theme was tried at FH and broke the console --
`Update-ConsoleSearch` also decides Visible for every button, and running it
early makes those decisions against half-initialised state. Recorded so it is
not retried.

### 57c. Fonts stop shrinking; bounds do not

A second operator, same spec as a colleague, same 96 dpi: header crushed, grid
perfect. The only difference was session size, 1716x844 against 2268x1131. At
844px the stack does not fit, so the factor search engaged for him and not for
her -- and at f = 0.70, 8.25pt scaled to 7.13pt still renders at 13px while an
18px label row becomes 13px. Below roughly 8pt the pixel metrics round to the
same value, so the text keeps 100% of its height while its box loses 30%. The
button grid survived on headroom the label rows did not have.

`$fTop = [Math]::Max($f, 1.0)` and `$fFontTop` likewise: the header furniture
scales UP with the zoom and never down.

### 57d. -ZoomDiag

New switch. Dumps the zoom factor, the screen working area, and for every
header control its box, its font, the pixels its text needs and the percentage,
to the startup log. This is what found 57a after three wrong theories argued
from screenshots. Use it before theorising about anyone's display again.

### 57e. What was NOT taken

The Gui3D theme retrofit from the same piece of work (a shared PS_Scripts
library, applied through the single `Set-ConsoleTheme` choke point) lives in a
separate AdminConsole-Themed folder and is not in this toolkit. Their own
record flags a latent bug in that library worth knowing if it is ever adopted:
`Get-Gui3DPalette -AsHex` calls `.Clone()` on an ordered hashtable, which
5.1 does not support; the error is non-terminating, so execution falls through
and returns Color objects where the caller asked for hex strings, silently.

Their header rework also renamed `$lblInput` to `$lblUnified` and added its own
column-width and box-share logic. Not taken: it would undo the two-row header
agreed on 2026-09-13. Only the measurement was ported.

## 58. 2026-09-15 -- multi-user, phase 2: every writer now asks whose it is

Phase 1 (56) built the profile and pointed Save Arrangement at it. Three other
windows still wrote the shared files, which meant the console was only
half per-person: arrange your buttons and nobody noticed, change a section or
a theme and everybody did.

### 58a. The same question, asked the same way, in four places

Save Arrangement, Theme Manager, Tool Manager's section editor and Park In
Place all now ask **just me, or everyone** -- once. After the first answer the
profile exists (or does not) and that is the answer from then on, because a
dialog on every colour tweak is a dialog nobody reads. Save Arrangement, Theme
Manager and Tool Manager infer it from `Test-UserProfile`; Park In Place is the
exception below.

Every one of them carries the REST of the layout across when it writes.
`Save-UserLayout` replaces the tenant's block whole, so a theme change that
rebuilt the block without Placement would silently reset somebody's
arrangement. Theme writes Theme, sections write Sections and ButtonsPerRow,
parking writes Parked, and each leaves the others as it found them.

### 58b. Park In Place has two scopes, and they say different things

Not a preference dressed as one:

- **Everyone** -- this tenant does not HAVE these tools. A fact about the
  tenant, true for whoever opens the console. Tenant file.
- **Just me** -- I do not want to look at these. This person's profile.

Tenant parking wins. A person can hide more; nobody can un-hide what the tenant
parked. The grid gained a **Hidden by** column saying which list is doing it,
the radio defaults to whichever list is already in use so the common case
cannot move one into the other, and switching scope rebuilds the grid because
the ticks mean a different list either side of it.

**The untick is guarded.** In 'just me' scope, unticking a tenant-parked row
would drop it from the personal list, change nothing visible, and read as the
tool being broken. The tick is put back with a sentence naming the scope that
owns it.

### 58c. The scanner and the tester

`Add-ConsoleAdditions` reads `Get-EffectiveParked` now, so a tool somebody hid
from themselves stops being offered back on the next publish -- they said no
once already. `Test-Toolkit` walks every `Config\Users\*.json`, Default block
and each tenant block, and reports a placement or a parked entry naming a
script that is not there, or a placement into a section no list declares. Same
CFG006 as the parked lists, same reason: both fail silently.

### 58d. Verified

Two profiles plus one account with none, resolved against the same folder and
the same tenant parked list: independent themes, columns and sections; both
inherit the tenant's hidden script and section; the account with no profile
gets the shared layout unchanged. A profile carrying an EMPTY parked list
still inherits the tenant's -- which is the rule in 58b, checked rather than
asserted.

## 59. 2026-09-16 -- pre-ship review of the FH console

Four questions: mistakes, view fixes, performance, parity. Findings and what
was done about each.

**Parity.** Every shared script diffed reference against FH. All differences
are deliberate scrub lines except three where FH is the BETTER copy:
Connect-ExSE, Bulk-DLMembership and CertificateExpiry derive the audit log's
Environment from `Get-DefaultTenantDirectory` where the reference still
writes a literal directory name. Worth bringing back to the reference; not a
ship concern. Exchange-HealthCheck calls Devin's `Invoke-EmsCommand` in FH and
`Invoke-OnPremCommand` here, by design, and both exist where they are called.
UnBlock's body was already identical; only its header note was stale.

**Mistakes.** `Show-InputDialog` was defined twice in FH's Common.ps1 -- once
in the shared part, once in Devin's region, byte-identical -- a graft
artefact. The region copy was removed. No top-level forward calls in
AdminConsole.ps1 (the load-order class that bit twice this week). Every
Verb-Noun the toolkit calls resolves to a definition or a real cmdlet.

**Performance -- one real regression, fixed.** On 2026-09-15 a staleness
check put `Get-Module` BEFORE the cache hit in `Get-ExchangeModuleMap`, which
runs on every cloud cmdlet through `Get-EXOCommand`: one enumeration of every
loaded module per Get-Mailbox, two thousand times in a two-thousand-mailbox
loop. The cache hit now returns first. Staleness is caught where it is free:
Connect-ExO refreshes after sign-in, and a module-qualified lookup against an
unloaded module misses, which is the one moment the cache can be wrong.

Second, smaller: `Get-ParkedStore` re-read and re-parsed the tenant file on
every call, four times per console start, when `Get-TenantConfig` already
holds it parsed and every writer clears that cache. Now reads the cache.

**View -- one real gap, fixed.** All three 57 fixes are present. But Park In
Place grew to 854px and never called `Set-FormScreenCorner`, the only clamp
there was -- so on the 844px session in 57c its Save button would sit below
the screen with nothing to scroll. `Limit-FormToScreen` now runs from
`Set-ConsoleTheme`, which every form passes through: taller or wider than the
working area is clamped and given a scrollbar, anything that fits is left
alone. One clamp in the shared path instead of one per tool.

**Ship folder gate.** 128 scripts, all balanced (the checker now understands
here-strings; new-OOF was a false positive), UTF-8 BOM / CRLF / ASCII, no
banned words, 106 Tools.json entries all resolving, no duplicate functions, no
stray files, Config\ excluded by the standing rule.

## 60. 2026-09-16 -- Refresh, and why a drag session left the console wrong

After moving buttons between sections: a heading displaced by one slot, one
heading hidden under it, a blank row under two sections. A restart fixed it,
which says the persisted layout was fine and the on-screen one had drifted.
Two causes, both fixed, plus a Refresh button as the safety net.

### 60a. Rows grew on hover and never shrank

The DragOver handler grows a section by a row the moment a dragged button
CROSSES into it while it is full -- before any drop, and it re-parents the
button on the way through. Drag across three full sections and back and all
three keep a blank row. It said so in its own comment: "never shrunk here:
rows the operator declared are theirs". Right principle, wrong variable: it
was protecting the transient working count, not the declared one.

`$script:SectionRowsDeclared` now holds what Sections.json or the profile
said, and every re-stack settles `SectionRows` to max(declared, needed).
The operator's declared rows are still the floor; only growth nothing ended
up using is given back.

### 60b. A save did not re-stack

`Save-ToolArrangement` only called `Update-ConsoleSearch` when a section had
grown, and the personal-profile path (56d) returned before even that. A
plain move between sections changes how many buttons two panels hold, and
panel heights and heading positions are recomputed nowhere else. So the
console showed the pre-drop geometry until something else happened to
re-stack it, and whatever ran next inherited stale rows. Both save paths
now end with a full re-stack.

### 60c. Refresh

Header button, left of View. `Invoke-ConsoleRedraw` runs the same passes the
console runs at startup, in order, from what is on screen: re-stack (rows,
headings, panels, zoom if in force), header row from captions, panel widths,
button state from the current connection. Reads no file, safe to press as
often as you like. Arrange became 'Arrange Off' / 'Arrange On' and grew to
110px; it was clipping at 74.

### 60d. Small screens, found on the way

`Update-ConsoleZoom` started the sections at `StartY0 * f` while the header
above them is placed at `fTop = max(f, 1)`. Below f = 1 -- the 844px session
in 57c -- the sections started above where the header ends and slid under
it. Now `max(StartY0, StartY0 * f)`, and the zoom fit reserves the same.

### 60e. Refresh must end where startup ends

First Refresh grew scrollbars and blank rows where a clean start had neither.
The two ran different LAST passes: startup ends with `Set-ConsoleInitialFit`,
whose zoom sizes every panel from the buttons it holds and widens the window
until nothing scrolls; Refresh ended with `Update-ConsoleSearch`, which sizes
panels from their DECLARED rows and, with no zoom in force, sets the window
from that -- taller than needed, clamped to the screen, scrollbars. Refresh
now runs the startup passes in the startup order and toggles AutoScroll so
a scrollbar already drawn goes away.

### 60f. The result line is gone; the status strip has the words

"Not looked up yet." and everything else written to `$lblResult` now goes to
the status strip along the bottom through `Set-ConsoleResult`. The strip was
already receiving the lookup result (`Set-ToolStatus -Done` mirrored it), so
the line was the same sentence twice at the cost of a row. The control stays,
hidden, because a dozen places read its Text back. Sections start at 136.

Found on the way and closed: 57a measured the header controls but never
moved what sits BELOW them. On a 192 dpi display the Find row's buttons are
59px, the row ends past the literal the sections start at, and the first
heading is drawn across it. `Set-ConsoleFindRowLayout` now takes the header's
real bottom edge, converts it to design units, and raises `SectionStartY` and
the zoom's `StartY0` when the header is taller than designed. Grow-only.

### 60g. The phantom scrollbar

A vertical bar with nothing behind it, on a console whose status strip was
fully visible, and a freeze when it was dragged. `$form.AutoScroll` was set
once, at build, and left on; WinForms draws a bar whenever the lowest control
lands within a pixel or two of the docked strip, and the initial fit --
trimming the window to exactly what was used -- arranged that on purpose.
Dragging the bar re-ran the zoom under the pointer; that is the freeze.

`Set-ConsoleScrollState` measures (from the origin, same arithmetic as the
fit) and switches AutoScroll on only when something is genuinely below the
edge. Called wherever the layout settles: end of the fit, end of every zoom
pass (which restores the previous state first, so a throw mid-zoom still
leaves a scrollable window, then decides), and Refresh. The fit's trim also
leaves 12px rather than 0.

Also 60f's first cut: the call-site rewrite that routed `$lblResult.Text`
writes through `Set-ConsoleResult` rewrote the one inside the helper too,
which then called itself until the stack ran out. Fixed; recorded because
"the grep for leftovers returned one line" is not the same as checking the
helper still had its own.

### 60h. Refresh opens every section, and the header offset stops ratcheting

Refresh with two sections folded fitted the window to the folded stack and
left a 110px hole under the header. Two causes.

The hole was 60f's header-bottom measurement reading `$script:Zoom.Applied`
from inside the zoom pass -- a field that pass writes at its END. It got the
previous factor, divided the already-scaled header bottom by too small a
number, over-estimated the design offset, and, being grow-only, ratcheted
the sections down the window one pass at a time. `Set-ConsoleFindRowLayout`
now takes `-HeaderFactor` from the zoom pass, which knows it, and settles to
max(design literal, measured) rather than max(current, measured) so an
over-estimate can come back down.

The folded sections: Refresh now clears every fold, in memory and in the
operator's state file, before it re-stacks. A redraw means 'everything as it
really is', and with per-person profiles about to group more, the folded
state is the one thing a refresh should not preserve.

## 61. One codebase, two installs (2026-09-16)

`AdminConsole` and `AdminConsoleFH` are now the same program. Every `.ps1` is
byte-for-byte identical; only `Config\*.json` differs. That is the rule, and
`Compare-Toolkit.ps1` enforces it:

    .\Compare-Toolkit.ps1 -Other ..\AdminConsoleFH

It checks three things — shared scripts identical, no undeclared extra files,
no organisation names in the other install's code — and exits 1 if any fail, so
it can gate a packaging step. Run it before every ship.

**Why.** The two folders had drifted into eight differing scripts plus a
`#region Tenant additions` block that made `Common.ps1` a different file in each
install. The cost showed up as a bug that could not be reproduced: per-user
themes were gated on `Test-MultiTenantMode`, so on the single-tenant install the
feature silently did not exist and two operators overwrote each other's theme
all day. A code path only one install exercises is a code path nobody tests.

**What changed.** The second install's 35 extra functions were folded into
`Common.ps1` (11 that nothing called were dropped). The reference was made
tenant-neutral so the mirror carries no scrub lines — including one regex that
named two directories outright and now reads a bracketed label from the end of
the sentence instead.

**Site-only tools are not drift.** Each install has tools the other does not:
this one provisions and terminates staff in a directory the other does not have;
the other has tools for its own suppliers, tracing and bulk jobs. Neither set
belongs in the other folder, and neither list is expected to shrink. Seven tools
here — `Disable-User`, `NewUser`, `NotDeadYet`, `FranUser_Office`,
`ManageOwnerGroups`, `Rename-OfficeGroups`, `EX-ServiceNowQueue` — and
twenty-three there are named in
`Compare-Toolkit.ps1`'s `-SiteOnly` list so the check can tell 'deliberately not
shipped' apart from 'somebody edited one copy'. The rule is about code, not
inventory: a script present in both places must be the same script.

## 62. How a change reaches both installs (2026-09-16)

Standing workflow, set after a run of changes went to one folder and had to be
chased into the other by hand.

**Every change is made in BOTH folders.** `AdminConsole` and `AdminConsoleFH`
are the same program (section 61); a fix that lands in one and not the other is
the drift that section exists to stop. There is no "do it here and port it
later" step — porting later is what produced eight differing files and a bug
nobody could reproduce.

**Every new or changed file is also staged in TWO ship folders (2026-09-17;
this replaced the single `ships\` folder):**

- `outputs\shipFH\` — for the other organisation's install. The file is copied with
  `.txt` appended: `AdminConsole.ps1` arrives as `AdminConsole.ps1.txt`. The
  `.txt` is what gets a `.ps1` past a mail filter, and comes off at the far end
  with `rename.ps1`.
- `outputs\shipEX\` — for this install's own reference copy. The file is copied **with its
  name unchanged** (`.ps1`, `.json`, `.md`, whatever it is). That install is
  updated directly, so nothing needs to be renamed on arrival.

Subfolders are part of the name in both: `Views\View_List.ps1` is staged as
`shipFH\Views\View_List.ps1.txt` and `shipEX\Views\View_List.ps1`.

Each folder is the ship list for its install: if something is in it, it needs
to travel; if the folder is empty, that install is current. Each is cleared
once its files have landed, so the next change starts from an empty list
rather than an accumulating pile nobody trusts.

**Before any file is changed or deleted, the old version goes to `outputs\undo\`**
(2026-09-17). Path: `undo\YYYY-MM-DD_HHMM_<what>\<same relative path as in the
install>` -- `undo\2026-09-17_1420_theme-fix\AdminConsole\Common.ps1`, say.
Only the files touched, not the folder: three files change, three files are
saved. Rolling back is copying that entry's contents back over the install.
Nothing reads from `undo\` and nothing ships from it; entries can be deleted
once the change they protect has been in use long enough to trust.

**What never goes through either ship folder:**

- `Config\*.json` — the other install's `Tenant.json`, `Tools.json`,
  `Sections.json`, `Themes.json` and `QuickLaunch.json` are its own. Copying
  this install's over them replaces that organisation's settings with these.
  This is the oldest rule in the toolkit and NEITHER ship folder bends it.
  Corrected 2026-09-18: an earlier version of this section carved out
  `shipEX\` on the reasoning that it goes to the reference install's own
  config. That was wrong in the way that matters. The reference copy in
  `outputs\` is not the live install -- the live one has been edited by
  Tool Manager, has per-user layouts and parked scripts in it, and copying
  a staged `Tools.json` over it would silently discard all of that. The
  rule has no exceptions now, which is also one fewer thing to remember
  at the moment somebody is copying files.

  So a NEW TOOL reaches an install as a `.ps1` and is added there with
  Tool Manager or Add-ConsoleAdditions; a NEW CONFIG KEY is typed in with
  Tenant Manager. Both were built for this and neither destroys what is
  already in the file.
- Anything under `Logs\`, `Reports\`, `DisabledTools\`.
- `Compare-Toolkit.ps1` — a tool for comparing the two installs, which belongs
  with the reference only.

**THE BANNED-WORD SCAN IS NOT OPTIONAL, AND IT IS NOT GLENN'S JOB.** Every
file staged for `shipFH\` -- and every file in `AdminConsole` that could ever
be staged there, which is all of them -- must be free of the names of this
organisation, its directories and its other tenants. THE WORD LIST ITSELF IS
NOT REPEATED HERE, because this document ships too and a rule that spells out
the names it bans defeats itself. The list lives in the `-Banned` default of
`Compare-Toolkit.ps1`, which is reference-only and never staged, and that is
the single copy. What the list covers: the parent company name, both on-prem
directory nicknames, the operating-company name in its three spellings, and
the two sibling tenants (`contoso.com` / `ABCD` / `WXYZ` are the
placeholders). A name that reaches the other organisation's code is
not a cosmetic slip: it tells them who else this toolkit serves. So the scan
is run by WHOEVER STAGES THE FILE, every time, before it goes into a ship
folder -- the same regex `Compare-Toolkit.ps1` uses, over every file being
staged. That is the gate. It does not depend on anyone remembering to run a
tool afterwards.

`Compare-Toolkit.ps1 -Other ..\AdminConsoleFH` is the belt to that braces:
it re-scans BOTH folders in full and checks that no script has been copied
into the site-only folder. Run it occasionally -- after a big batch, before a
packaged release -- not as a per-ship ritual. It exits non-zero on any
finding, so it can gate a packaging step if one ever exists.

---

## 63. The views moved out of the console (2026-09-16)

**What changed.** The eight alternative views — Tool list, Recent and pinned,
Command palette, Status board, Catalogue, 3D buttons, What I can run — no longer
live in `AdminConsole.ps1`. Each is its own file in `Views\`:

```
Views\View_List.ps1        Views\View_Catalogue.ps1
Views\View_Recent.ps1      Views\View_3DButtons.ps1
Views\View_Palette.ps1     Views\View_Ready.ps1
Views\View_Status.ps1
```

`AdminConsole.ps1` went from 6,849 lines to 5,865.

**Why.** They were 1,090 lines sitting between the console's own code and its
startup — so every experiment with a view was an edit to the file that draws the
button grid, and a typo in the catalogue could stop the console opening. For
code the console calls once, on a menu click, that is a bad trade. Shipping a
new view is now one file.

**How a view declares itself.** Five comment lines at the top, above the block
comment:

```powershell
#Requires -Version 5.1
# VIEW-ID: Catalogue
# VIEW-NAME: Catalogue - what each tool does
# VIEW-NOTE: Every tool with the help text from its own header.
# VIEW-FUNCTION: Show-ConsoleCatalogueView
# VIEW-ORDER: 50
```

`Read-ConsoleViewManifest` reads the head of each file, stopping at the first
`function` line so a tag quoted inside a comment further down is not mistaken
for a declaration. **There is no list to edit.** Drop a file in the folder and
it is on the View menu; take it out and it is gone.

**Dot-sourced, not run.** The loader is at script level in `AdminConsole.ps1`,
deliberately not inside a function — a dot-source inside a function defines into
that function's scope and the definitions vanish when it returns. Because the
files are read into the console's own scope, a view sees `$form`,
`$script:AllTools` and every helper exactly as it did before the move. **Nothing
in the view bodies changed.**

**What stayed in the console.** The shell all eight are built on:
`Get-ConsoleToolModel`, `New-ConsoleViewShell`, `Update-ConsoleViewShell`,
`Add-ConsoleViewFontStepper`, `Get-ConsoleViewColors`, `Invoke-ConsoleViewTool`,
`New-ConsoleToolTile`. Those are shared and a view file may call any of them.

**Failure is on the menu, not in a log.** A missing `Views\`, an empty one, a
file that will not parse, a duplicate `VIEW-ID`, or a `VIEW-FUNCTION` the file
does not define — each costs that one view and adds a disabled *"Some views did
not load"* line to the View menu, with the detail in its tooltip. A console that
quietly offers one view because a folder was not copied is a support call that
starts with "the views are gone".

**`Show` holds the function NAME, not a scriptblock.** `& $view.Show` runs
either. A name survives being written to a user profile and read back, which is
what the per-user default view needs next.

**Shipping.** `Views\` is shared code under the mirror rule (section 61) and
`Compare-Toolkit.ps1` checks it. That same pass added `Projects\`, `probes\` and
`RetryQueue\` to the comparer's skip list: they are per-install by nature and
were producing forty-odd permanent false differences, which is how the one that
matters gets missed.

**Still to come.** A per-user default view stored in
`Config\Users\<account>.json`, so the console opens straight into the view that
operator chose — without changing what anyone else sees.

---

## 64. Which view the console opens into, per person (2026-09-16)

**The problem.** Two engineers, one folder. One of them does not get on with the
button grid and wanted to change the console to suit himself, which would have
changed it for the other. The views existed already; what was missing was the
ability to *keep* one.

**The answer.** `Config\Users\<account>.json` gains two keys beside `Theme` and
`Groups`:

```json
"Default": {
    "Theme":       "Slate",
    "DefaultView": "List",
    "Views":       { "List": { "Font": 14 } }
}
```

`DefaultView` is a `VIEW-ID`; empty or `Buttons` means the grid. `Views.<id>` is
that view's own remembered settings — today just the A+/A- font size, which is
saved on close and only if it changed.

**New in Common.ps1:** `Get-UserDefaultView`, `Set-UserDefaultView`,
`Get-UserViewSetting`, `Set-UserViewSetting`, `Remove-UserLayoutKeyFromTenants`.

**New tool: `ViewManager.ps1`** (Dashboard → View Manager). Lists Button Layout
plus every `View_*.ps1`, marks which one this account opens into, and sets or
clears it. It cannot *open* a view — a view is drawn by the console and needs
the console's window — and it says so rather than offering a button that fails.

**"Open straight into it."** `Open-ConsoleDefaultView` runs from `Add_Shown`,
after the initial fit, on a 450ms timer. The view opens **over** the console,
not instead of it: the lookup box, connect buttons, PIM clock and status strip
live on the console, and replacing it would mean rebuilding those in every view
file. Closing the view lands on a complete console. The way back is an Esc.

### What the review caught, because it is the interesting part

Two hostile passes found eleven defects. Three are worth recording because they
are a *family*, and the family will recur:

**1. A grown block, rewritten from a list of names.** Save Arrangement built its
layout block from scratch and carried forward `Theme` and `Parked` by name.
Correct when written; silently wrong the moment a fifth key existed. Moving one
button would have wiped the operator's chosen view and every font size with it.
Now it starts from `Get-UserLayoutOwn` like the other three writers. **Rule: a
writer that does not own the whole block must start from the block.**

**2. Read scope and write scope must match.** `Set-UserDefaultView -AllTenants`
read the *tenant* block (`Get-UserLayoutOwn` had no switch) and wrote it over
`Default`. With no tenant block it read `$null` and replaced `Default` with a
single key — theme, groups, sections and placement gone, and `Save-UserLayout`
keeps no `.bak`. `Get-UserLayoutOwn` now takes `-AllTenants` and every caller
passes the same switch to both halves.

**3. "For everyone" has to mean everyone.** The layout merges Default with the
tenant block *key by key, tenant wins*. So setting a preference "in this tenant"
and later "in every tenant" left the narrow copy winning, with no window able to
clear it. `Remove-UserLayoutKeyFromTenants` now strips the key from every tenant
block on an every-tenant save.

**And one aliasing bug my own fix introduced:** `Get-UserLayoutOwn` returned the
live object inside `$global:UserProfile`, so editing it edited the cache — Save
Arrangement then read the cache back to recover parked tools' placements and
found its own edit, carrying nothing. Worse, a save that *threw* left the cache
holding an arrangement that never reached disk, ready for the next writer in the
process to persist. It now returns a shallow copy. All four existing callers
benefit; the comment in Save Arrangement names the dependency so nobody
optimises the copy away.

**Two smaller ones worth the line:**

- `$mine = Test-UserProfile` asked "does a profile file exist" as a proxy for
  "has this person arranged anything". Those were the same question only while
  arranging was the only thing that created a profile — picking a default view
  creates one now, and so does nudging A+. Save Arrangement, Tool Manager and
  Theme Manager now each test the key *they* write.
- A `Views\` folder that cannot be **listed** (an ACL, an offline share) is not
  an empty folder. `Get-ConsoleViewCatalogue` reports that case as one entry
  with no `File` and the reason in `Error`, so the console stops saying "Views\
  is empty" about a folder with seven files in it.

---

## 65. 'Button Layout' is now a file; the console is called 'Console' (2026-09-16)

**Two things, and the second only makes sense because of the first.**

**`Views\View_ButtonLayout.ps1`** is the tool grid as a view: the same tools,
the same sections, the same order, drawn as ordinary console buttons, honouring
this operator's `ButtonsPerRow`, in ~200 lines that do nothing else. It is a
copyable starting point for anyone who wants the grid to look different.

**The built-in entry is renamed from 'Button Layout' to 'Console'.** Its `Id`
is still `Buttons` — that is what a per-user default view is stored as, and
renaming it would silently reset everyone who had chosen the grid.

**Why the rename.** The console IS a button grid, so the old name was accurate
and still wrong: it made the window and the grid sound like the same thing. They
are not — the window is also the lookup box, the connect buttons, the PIM clock,
the status strip, drag-to-arrange, collapsible sections and the zoom walk. Once
the grid existed as a file, two menu entries reading "Button Layout" would have
been unexplainable.

**Why the grid could not simply be extracted from the console.** It is not a
separable function in there: `Build-ToolButtons` is wound through the zoom
baseline, `Set-ConsoleInitialFit`, the collapse state, the find row and the drag
handlers. Handing an engineer 5,800 lines and saying "make the buttons the way
you like" is not a small ask. The view is a *sibling* built on the same model
(`Get-ConsoleToolModel`), not a copy of the console's rendering — so the two
cannot disagree about which tools exist or which can run, but they are free to
look different.

### The thing to be clear about with anyone customising one

`Views\View_ButtonLayout.ps1` is **shared code** under the mirror rule (section
61). Editing it changes the Button Layout view for everyone who uses the folder.

What is personal is **which view you open into**. So the workflow for an
engineer who wants their own grid is:

1. View Manager -> select Button Layout -> **Copy to a new view...** -> name it.
2. Edit `Views\View_<theirs>.ps1` to taste.
3. View Manager -> select it -> **Open into this one**.

Their file is on everyone's View menu — it is a file in a shared folder, there
is no way around that — but nobody else opens into it, and the original Button
Layout is untouched. That is the same split as everywhere else in this toolkit:
**code is shared, choices are personal.**

---

## 66. On the Dashboard means only on the Dashboard (2026-09-17)

**The rule.** A script that has a Dashboard button is reached from the Dashboard
and from nowhere else. `Add-ConsoleAdditions` must never offer it as a tool-grid
button.

**Why.** Two buttons for one window gives the operator two routes to the same
place and invites the question of whether they differ — a question with no good
answer, asked every time. It also spends a grid button on configuration and
housekeeping, which is not what the grid is for: the grid is tools you run
against a mailbox or an account.

**This is about the SCAN, not about permission.** Tool Manager will still add any
of these by hand. Somebody who genuinely wants Theme Manager on their grid can
put it there. The rule is that the toolkit does not *suggest* it — a suggestion
accepted by reflex is how a second button appears without anyone deciding.

**How it is enforced.** `Add-ConsoleAdditions.ps1` excludes:

1. `$script:NeverOffer` — a hand-written list of infrastructure (`AdminConsole`,
   `Common`, `Publish-Toolkit`, `Deploy`, `Test-Toolkit`, `Diagnose-Theme`, this
   file) plus the Dashboard tools and `Unified-Lookup`.
2. **Whatever `Dashboard.ps1` currently declares**, read at run time by
   `Get-DashboardScript` from the `Script = '...'` entries in `$configTools`.

**Two lists because the first one drifts.** A hand-written list of Dashboard
tools is correct on the day it is written and wrong the first time somebody adds
a Dashboard tool and forgets to update it. That is not hypothetical: it is
exactly what happened to **ViewManager.ps1** on 2026-09-16 — it shipped with a
Dashboard button and the next publish offered it as a grid tool anyway. Reading
`Dashboard.ps1` makes the rule enforce itself, so the *next* Dashboard tool is
covered on the day it is added.

The scrape can only ever ADD exclusions. If the parse finds nothing — a reshaped
Dashboard, a renamed variable — the hand-written list stands exactly as before,
so the failure mode is "back to where we were", not "everything is offered".

**When you add a Dashboard tool**, you now have to do nothing. Adding it to
`$script:NeverOffer` as well is belt and braces and still worth doing for
anything that is *not* on the Dashboard but should never be offered — a
dot-sourced library, a build-time script, a diagnostic run from a prompt.

**Not the same as `Config\Additions-Ignore.json`.** That is one operator
declining one script on one install. This is what the toolkit has decided, for
every install, permanently.

## 67. The own-process marker is `TOOLKIT-REQUIRES: OwnProcess` and nothing else (2026-09-17)

`ServerReboot.ps1` shipped with `# TOOLKIT-OWN-PROCESS: ...` in its header. Nothing
reads that line. `Test-ToolDeclaresOwnProcess` (AdminConsole.ps1) looks for
`TOOLKIT-REQUIRES:\s*OwnProcess` in the first 40 lines, and finding nothing it
loaded the tool **in-process** -- into the console's PowerShell, which has Exchange
Online connected and no on-premises session. The worker runspaces were created in
that process, so every remote call failed with "WinRM cannot complete the
operation". The same script run by hand in the child window that
`Connect-ExSE.ps1` had prepared worked, because that window IS the process an
own-process launch produces. The symptom looked like a connection fault in the
script; it was a spelling fault in the header.

The same header also carried `TOOLKIT-VARIANT: OnPremExchange` and
`TOOLKIT-REQUIRES: OnPremExchange`, neither of which is a value anything
recognises. The variant is `OnPrem`. A marker is only a marker if the reader
agrees on the spelling, so when adding one, grep for the regex that consumes it
first -- `Test-ToolDeclaresOwnProcess`, `Get-ScriptVariant` -- rather than
inferring a name from the others.

## 68. One folder of shared code; the other install keeps only its own scripts (2026-09-17)

Section 61 made the two folders the same program. Once they were, the second
copy was doing nothing but demanding a mirror step on every change -- and
section 62's two ship folders already route each change to each install. So
the mirror is gone.

**`AdminConsole`** holds ALL shared code: the console, `Common.ps1`, `Views\`,
every shared tool, the templates. One copy. A change is made once, here, and
staged for both installs (`shipEX\` as-is, `shipFH\` with `.txt`).

**`AdminConsoleFH`** now holds ONLY what belongs to that install alone: its
twenty-three site-only scripts (AvePoint, Proofpoint, SendGrid, the
Unified-Trace family, Get-Queues, Search_Purge, RenameUser, and so on) and its
`markdown\` docs. No `Common.ps1`, no console, no `Config\` -- nothing that
could drift, because nothing there exists anywhere else. A change to one of
those scripts is staged to `shipFH\` only.

**`Compare-Toolkit.ps1` was reworked to match.** It no longer compares hashes
-- there is nothing to compare. It checks that no script exists in BOTH folders
(a copy in the site-only folder is the drift returning), that every name in
`-SiteOnly` still matches a file (the list is the record of which of THIS
folder's scripts never travel), and that no organisation name appears in
shipped code in either folder. Same command as before:

    .\Compare-Toolkit.ps1 -Other ..\AdminConsoleFH

The full mirror it replaced is under `undo\2026-09-17_AdminConsoleFH-full-mirror\`
along with its `Config\`, in case anything in it turns out to be wanted.

## 69. Microsoft Graph: one app, delegated, scopes declared by the tool (2026-09-17)

The tenant registered a single Graph application -- delegated permissions only,
no secret, no application permissions -- and the console signs in THROUGH it as
the operator. Five tools landed on it at once: Sign-in Activity, Risky Users,
Auth Methods, Licence Assignment and Service Health.

**Config, not code.** The client id and tenant id are tenant facts and live in
`Tenant.json` under `Graph`, with `Capabilities.MicrosoftGraph` as the on/off
switch; Tenant Manager edits both and validates that they are GUIDs. A ticked
capability with a blank client id reads as "not configured" and greys the Graph
tools with a tooltip, rather than letting them fail at sign-in. One install only
for now -- the other has the capability off, so the tools are absent there.

**Each tool declares its own scopes**, in its header:

    # TOOLKIT-GRAPH-SCOPES: AuditLog.Read.All, User.Read.All

`Connect-MSGraph.ps1` scans every script in the folder and signs in ONCE with
the union, so one prompt covers every Graph tool and the toolkit asks for
nothing it does not use. The alternative -- a list of scopes maintained inside
the connect script -- is correct on the day it is written and wrong the first
time a tool is added, which is the same drift `Get-DashboardScript` exists to
stop (section 66). The connect script PRINTS the list and which script asked for
each one, so when consent is refused the request to the app owner is already
written out.

**Tools never sign in themselves.** A tool runs inside the console's process
while the console holds the thread in a modal loop, and an interactive sign-in
raised there draws and never takes input -- the fault that moved the Exchange
Online connect out of its button handler. Every Graph tool calls
`Assert-ToolkitGraphScope` first instead: it reports which scopes are missing
on its own status line, disables its buttons, and points at the connect button.

**REST, not the generated cmdlets.** Only `Microsoft.Graph.Authentication` is
required; everything goes through `Invoke-ToolkitGraph`, which prefixes the v1.0
base, follows `@odata.nextLink` on `-All`, stops at `-MaxRows`, and turns a
failure into a sentence -- including the 403 hint, because a 403 means EITHER
the app lacks the scope OR the operator lacks the Entra role, and the two are
indistinguishable from the call site.

**403 is a two-part answer, and both parts are in every tool's help.** Scopes
are granted to the app; roles are held by the person. Sign-in Activity needs
Reports Reader or better; Risky Users needs a security reader role and Entra P2;
Auth Methods needs Authentication Administrator. A tool whose help named only
the scope would send people to the app owner for a problem the app owner cannot
fix.

**PIM stays on the default Graph app.** `PimActivation.ps1` needs two
`Privileged*.AzureADGroup` scopes this app does not carry, so it was left alone
-- it is the one Graph caller not on MSGraph-M365. If those scopes are ever
added to the app, moving it over is a one-line change.

**Threat Explorer was NOT built.** The delegated route to that data is Advanced
Hunting (`POST /security/runHuntingQuery` over `EmailEvents`), which needs
`ThreatHunting.Read.All` -- not on the app. The direct `analyzedEmails` API is
application-permission and beta only, so it is not an option at all from a
delegated console. Decided 2026-09-17: build it when the scope is granted, not
before, because a tool that cannot run is worse than no tool.

## 70. The meeting-handover sweep, and the window that nearly broke it (2026-09-18)

`ChangeMeetingOrganizer.ps1` gained **Everything they organise...**: name the
leaver and their replacement, and it lists every recurring series that mailbox
organises for a tick-and-transfer. It exists because the old flow needed
somebody to already know the subject of every meeting a departing person owned,
which nobody does.

**It must run before the mailbox is converted to shared.** Microsoft: "You
can't transfer meetings in Group or Shared mailboxes." Disable-User converts to
shared as one of its steps, so after a termination the handover is no longer
possible without converting back and re-licensing. The tool already refuses a
shared mailbox by name; the help now says when to run it. The two scripts are
NOT coupled -- by decision, so a termination does not wait on a calendar read.

**The trap, written down because it is invisible.** The date range on
`Get-CalendarDiagnosticObjects` filters `OriginalLastModifiedTime` -- when a
meeting was last EDITED, not when it runs. A 90-day window therefore misses the
weekly series set up two years ago that nobody has touched, which is precisely
the meeting that needs a new owner. The sweep uses two years for that reason.

**And why it is sliced.** `ResultSize` caps at 1000 and a row is one CHANGE, not
one meeting, so one busy series can be dozens of rows and a single wide query
truncates silently. The window is cut into quarters, each with its own 1000-row
budget, and a quarter that fills up sets `Truncated` -- reported in red in the
picker rather than hidden. A list that might be short is useful; a list that
pretends to be complete is not.

**Transfers go by SUBJECT, not by id, and that is forced.** The calendar log
identifies a meeting by GOID; `Invoke-ChangeMeetingOrganizer` takes `EventId`
(the store id) or `Subject`, and nothing converts between them -- the documented
parameter sets are ByEventId and BySubject. Subject is a substring match, so the
cmdlet is what saves this: given an ambiguous subject it transfers nothing and
returns candidates with EventIds. `Invoke-SeriesTransfer` then picks the
candidate whose subject matches EXACTLY and re-runs by id; if two share the
exact subject it reports AMBIGUOUS and refuses rather than guessing.

`Get-MeetingFacts`'s row-reading was extracted to `ConvertTo-MeetingFactList` so
Find and the sweep decide "is this a series" with the same code. Two copies of
that judgement inside one tool would be one of them wrong the first time a
property name was corrected in the other.

## 71. Mail link-rewriting corrupts shipped code (2026-09-18)

`Common.ps1` came back from the other install with three URLs rewritten by
Proofpoint URL Defense -- the email security product that replaces links in
mail so they can be checked at click time. It does not know the difference
between a link in a message and a string literal in a script, so this:

    return "http://$configured/PowerShell/"

had become a `urldefense.com/v3/__...` wrapper, and the same happened to the
two SharePoint URL builders. A fourth string gained a trailing space.

**What that would have done.** `Get-ExchangeServerUri` builds the remote
PowerShell endpoint for on-premises Exchange. Rewritten, it returns an
unreachable address, so Connect-ExSE fails on a line that looks correct in
every editor unless you scroll right. The SPO admin URL is the same shape of
fault. These are not cosmetic.

**The `.txt` rename does not protect against this.** That convention exists to
get a `.ps1` past attachment filters; it does nothing about link rewriting,
which happens to the CONTENT. Anything mailed in either direction is exposed.

**So the gate before staging now includes a URL check**, alongside the
banned-word scan: grep for `urldefense`, and for a URL string ending in a
space. Both are one line and both would have caught this. If a file arrives
from the other install, it is checked BEFORE being treated as authoritative --
on this occasion the reference copy was clean and simply replaced it, but a
file with real changes in it would have needed the rewrite undone by hand.

**Better still, do not mail code.** A shared folder, a zip with the files
stored rather than linked, or any transport that is not an email body avoids
the problem entirely. That is a decision for whoever owns the process, not
something the toolkit can enforce.

## 72. The Mailbox Editor: one window instead of fourteen tools (2026-09-18)

`MailboxEditor.ps1` is a tabbed window over one mailbox. Load an address once
and every tab acts on it: **Overview, Addresses, Access, Folder Permissions,
Mail Flow, Settings, Retention, Automatic Replies, Recovery**. It absorbs what
fourteen separate tools did -- permissions, permission copy, OneDrive access,
calendar permissions, mailbox type conversion, archive, hold and retention,
automapping repair, delegates, retention policy, forwarding, hide/unhide,
health and statistics.

**Why.** The tools were each right and the set was wrong. Answering "what is
going on with this mailbox" meant opening five windows, pasting the address
into each, and holding the answers in your head, because no window could see
what another had found. Looking at a mailbox and editing it is one job, and it
now has one window.

### The hybrid write-target map is the reason this is not just a merge

On a synced mailbox some fields are mastered in Active Directory. Exchange
Online **accepts** a write to one of them and directory sync reverts it within
about half an hour, with no error anywhere -- the operator sees success, the
user sees nothing change, and the ticket comes back next week. Hide-Unhide.ps1
learned this the hard way.

So `Common.ps1` carries `$script:MailboxFieldOrigin`, a three-valued map:

| Origin | Meaning | In the UI |
|---|---|---|
| `Cloud` | mastered in Exchange Online | always editable |
| `Directory` | mastered on-premises when the mailbox is synced | disabled, with the reason, on a synced mailbox |
| `Unverified` | not confirmed either way | editable, with a caution |

`Get-MailboxWriteTarget` returns `Ok` / `Directory` / `Warn` plus a Reason, and
every Apply calls `Test-EditorFieldWritable` before it writes. `Unverified` is
deliberately a value rather than a guess: claiming certainty we do not have is
how the original bug got written.

### Rules the window keeps

- **Per-section Apply, never a single Save.** One Apply writes one group of
  fields. A window-wide Save would write everything the operator did not touch,
  and on a mailbox that is an audit line claiming six changes for one edit.
- **Tabs load on first view, and only cache a read that worked.** A tab whose
  read failed is left uncached (`$script:TabReadFailed`), so it retries rather
  than showing empty for the rest of the session.
- **Every write is audited against the address**, including failures -- a
  failure logged against a button name cannot be traced to a mailbox later.
- **Lossy renderings are never written back.** Delivery restrictions come back
  as recipient objects and render as names; the box remembers what was drawn
  and is only sent if the operator actually edited it. `ForwardingAddress` is
  shown in its label rather than in the edit box for the same reason -- a value
  the tab could not accept back should not sit in a field with an Apply button.
- **The address list is rebuilt from a fresh read.** Making an alias primary
  replaces the whole `EmailAddresses` collection, so it is built from the
  mailbox as it is now, not from the snapshot taken at Load. An X500 added
  since -- by another admin, or by a mailbox move -- would otherwise be
  deleted, which breaks replies to every message already sent.
- **Automapping repair closes the gap it opens.** The repair removes Full
  Access and re-adds it; if the re-add fails it retries, and a final failure
  says in plain words that the trustee currently has no access.

### Unified Lookup: two buttons became one

'Open permissions' and 'Mailbox settings' both pointed at a slice of what this
window does, so they are now one **Mailbox Editor** hand-off. `Set-CloudPrimaryAddress`
moved from Unified-Lookup.ps1 into Common.ps1 unchanged, so the primary-address
rebuild has one implementation behind both front doors.

### The fourteen absorbed tools are still installed, on purpose

They are **not** retired in the same change that replaces them. A consolidation
that removes the old path on day one has no way back when the new window turns
out to be wrong about something on a mailbox shape nobody tested. They move to
`DisabledTools\` after a fortnight of live use, and the tool catalogue is
regenerated then -- not before.

## 73. The Mailbox Editor edits on-premises too, on a switch (2026-09-18)

Section 72 above described a window that **refused** the directory-mastered
fields on a synced mailbox and said where they really lived. That was true and
useless: the operator was left holding a correct explanation and no way to act
on it. `MailboxEditor.ps1` is now `TOOLKIT-VARIANT: Dual` and carries an
**Editing** drop-down in the top right -- Cloud (Exchange Online) or
On-premises (AD / Exchange).

**A drop-down, not a toggle button.** A button captioned "Cloud" cannot be
read: is that the mode you are in, or the mode you will get if you press it?
A list shows the current mode and the alternative at once.

### Three on-premises shapes, three cmdlet families

`Resolve-OnPremMailboxTarget` probes in the order that answers fastest for the
common case:

| Kind | What it is | Written with |
|---|---|---|
| `RemoteMailbox` | the usual hybrid stub; mailbox has moved to the cloud | `Set-RemoteMailbox` |
| `Mailbox` | never migrated; still an on-premises mailbox | on-premises `Set-Mailbox` |
| `User` | a directory user with no mailbox | `Set-User` only |

Directory fields (title, department, office, manager) come from `Get-User` for
every shape, because they are user properties and are not on a RemoteMailbox
object at all.

### The rules this mode adds

- **It degrades, it does not refuse.** No on-premises Exchange in the tenant,
  or none connected in this window, greys the switch out with the reason. The
  cloud side is untouched. Nobody loses the tool because a server is down.
- **Only `Directory`-origin fields are writable on-premises.** The first cut
  allowed `Unverified` through as well; that was wrong and dangerous, because
  every Unverified field -- forwarding, delivery restrictions, send-on-behalf
  -- is handled by a cloud handler that calls `Invoke-EXOCommand` directly.
  The gate said yes, the write went to Exchange Online, and the audit line
  stamped it OnPrem. *Unverified means we do not know, and the answer to not
  knowing is not to write.*
- **Six tabs are disabled, not merely skipped.** Blocking a tab's *load* left
  every button on it live over a grid still holding rows read for the cloud
  mailbox -- "Remove selected" would have removed real permissions under a
  banner that said on-premises. The TabPages are disabled and the grids
  cleared, and all sixteen handlers on them also refuse in code.
- **One router, `Invoke-EditorSetMailbox`.** Handlers name no cmdlet and
  supply no Identity; both come from whichever object the mode loaded. The
  alternative -- an if/else inside each handler -- gives four handlers four
  chances to disagree about which cmdlet on-premises wants.
- **`Environment` in the audit log is the mode**, not the constant `Cloud`. An
  address change made on-premises and one made in the cloud are different acts
  with different blast radii.
- **The remote routing address cannot be made primary.** On a remote mailbox
  the `*.mail.onmicrosoft.com` entry is what `targetAddress` points at -- the
  pipe that carries on-premises mail to the cloud mailbox, not a reply address.
  Promoting it breaks mail flow for that person. The cloud path cannot produce
  this mistake; this one can, so it is refused by name.
- **An enabled e-mail address policy is reported, not hit.** While
  `EmailAddressPolicyEnabled` is set, Exchange rebuilds the proxy list from the
  policy, so an address written here is rejected or reverted. The cmdlet's own
  error names the policy but reads like a syntax fault, which sends people
  looking in the wrong place.

### Sync is stated, never triggered

Every on-premises write ends with one sentence, from one place
(`$script:OnPremSyncNotice`): the change is made on-premises and Exchange
Online will not show it until the next directory sync, so check back in about
30 minutes. The window does **not** offer to start a sync cycle -- that is not
this toolkit's to run -- and it does not poll, which would park the operator
in front of a spinner for half an hour.

### Known limit

A mailbox that **never migrated** cannot be opened at all: `Invoke-EditorLoad`
resolves in Exchange Online first, and an on-premises-only mailbox is not
there. The failure now says which it is rather than reporting "not found",
so nobody concludes the mailbox is gone. Lifting the limit means the whole
window running off an on-premises object with no cloud counterpart, and that
is a separate change.

## 74. The Mailbox Editor stops asking people to type things (2026-09-18)

Three changes, one principle: **if the toolkit knows the valid answers, it
should not ask for one to be typed.**

### Access and Folder Permissions

Both tabs gained a **tick-box column**, and Change and Remove act on every
ticked row rather than on the highlighted one. Row selection was the wrong
mechanism: it is invisible once the grid loses focus, it is lost when the grid
is re-read, and it cannot say "these four, spread down a list of thirty".

- **Grant** takes a list of MAILBOXES, a list of PEOPLE and tick boxes for the
  kinds of access. The request that actually arrives is "give these three
  people access to these two shared mailboxes", and the old dialog answered it
  one grant at a time -- six passes, with the words `FullAccess`, `SendAs` or
  `SendOnBehalf` typed each time and rejected on a typo.
- **Change** takes the ticked rows down to the distinct people behind them and
  offers two columns of tick boxes: GIVE them, TAKE AWAY. Two explicit lists
  rather than one tri-state column, because a checkbox with three meanings is
  a checkbox nobody trusts.
- **Folder roles are a list you click, not tick boxes**, and that is Exchange's
  doing rather than a preference: a folder permission stores ONE role per user
  per folder. Owner and Reviewer together is not a state the server can hold,
  so offering both ticks would be offering something that cannot be written.

**Anything ticked in both columns is a restamp, and goes through
`Repair-MailboxAutoMapping`** rather than a bare remove-then-add. That function
exists precisely because the re-add commonly fails on throttling; a bare pair
leaves the trustee with the permission removed and nothing put back.

**A bulk grant audits each mailbox against its own address.** A log whose lines
all point at the mailbox that happened to be loaded cannot answer "who gave
them access to that one", which is the question an audit log is for. The
dialogs report per target too -- four mailboxes where one failed is not a
success and not a failure, and a status line saying "Applied" would be a lie
about the one that did not.

**Send on Behalf is gated per mailbox in the list, not on the loaded one.** It
is directory-mastered on a synced mailbox, and checking only the banner meant a
cloud-only mailbox in the box suppressed the warning for the synced ones beside
it -- whose write directory sync would quietly revert.

### Automatic Replies

The tab held two plain multiline text boxes. Exchange stores auto-reply text as
HTML, so the box showed markup as text and wrote back whatever was in it:
opening the tab on a message somebody had formatted in Outlook and pressing
Apply flattened it. Two editors for one field, and the simpler one silently
destroyed the other's work.

`OOFEditor.ps1`'s editor -- a WebBrowser with `contenteditable`, a real
toolbar, templates and a preview -- **moved into Common.ps1**, and both front
doors now open the same one. OOFEditor.ps1 is a window around the shared
function; the tab shows state, audience, schedule and the message as plain text
and hands over. It keeps being its own tool because setting one person's
out-of-office is a job of its own and does not need eight other tabs in front
of it.

### Health, Statistics and the licence

- **Licence is on Overview**, read-only. It decides whether mail keeps
  arriving and it is the first thing asked about a mailbox that has stopped
  working, so it belongs where you land. Group-based, so it names the GROUP
  that drives the assignment rather than the SKU -- the group is the thing
  anybody can act on.
- **Health** carries only what no other tab shows: protocol access (OWA,
  ActiveSync, POP, IMAP, MAPI, EWS), whether the directory account is disabled,
  the hybrid ExchangeGuid/ArchiveGuid comparison, and inbox rules that forward,
  redirect or delete. Its verdict box does cover the other tabs, but it states
  the FINDING and names the tab -- "Forwarding is set to an external address --
  see Mail Flow" -- rather than becoming a second copy of Mail Flow.
- **Statistics loads on a button, never on tab selection.**
  `Get-MailboxFolderStatistics` enumerates every folder in the mailbox; on a
  large one that is tens of thousands of rows. No other tab should ever wait
  behind it, and opening Statistics by accident should cost nothing.
- **The folder table sorts on the byte count, parsed from the text.**
  `ByteQuantifiedSize` has a `ToBytes()` method on a live object, but over a
  remote Exchange Online session the property arrives deserialised as
  `1.9 GB (2,040,109,465 bytes)` with no methods on it. Relying on the method
  scored every row zero and the table came back in Exchange's own order looking
  sorted.

## 75. Enable mailbox, from Unified Lookup (2026-09-18)

The one thing you could see in that window and not act on: an account whose AD
object exists and whose mailbox does not. The button runs the same sequence
New User runs -- `Enable-RemoteMailbox`, then the archive -- from
`Enable-ToolkitRemoteMailbox` in Common.ps1, with a confirmation and no second
window, because the alias is already on the row.

Four things it gets right that the first cut did not:

- **`'AD only'` is an on-premises row.** A user with no mailbox has no proxy
  addresses, which is exactly what makes the search classify the row that way
  -- so the first version rejected the only rows it existed for. `OnPremType`
  was also being filed under `CloudType` for those rows, saying the opposite of
  the truth.
- **The identity is the distinguished name, not the alias.** Alias is
  `mailNickname` falling back to `sAMAccountName`; for an account with no
  Exchange recipient object it is neither guaranteed unique nor guaranteed
  resolvable, and an ambiguous identity here mail-enables somebody else.
- **The on-premises session is checked before the object.**
  `Enable-RemoteMailbox` lives in the on-premises Exchange module; without it
  loaded the button would confirm, act, and fail on the cmdlet lookup -- after
  the operator had said yes to creating a mailbox.
- **The routing domain is derived in one place**,
  `Get-TenantRemoteRoutingDomain`. The naive replace is not idempotent: a
  Tenant.json already holding the `mail.` form passes Tenant Manager's suffix
  check and produced `<tenant>.mail.mail.onmicrosoft.com` -- an address that
  looks plausible, is accepted by `Enable-RemoteMailbox`, and silently does not
  deliver.

## 76. The search was hiding half of a hybrid mismatch (2026-09-18)

`Get-OnPremAdObject` dropped **every** Active Directory object that had
proxyAddresses, on the reasoning that on-premises Exchange had already
returned it with better information:

    if ($mailed -and $ExchangeAnswered) { continue }

That is true for an ordinary mail-enabled object and **false for the one shape
it matters most for**: an AD user carrying mail attributes but no Exchange
recipient object at all. On-premises Exchange does not return that user --
there is nothing there to return -- and this line threw the AD row away. The
person vanished from the on-premises half of the search entirely and appeared
as cloud-only, in a window whose whole purpose is to show both halves at once.

That shape is not exotic. **It is what a mailbox created directly in the cloud
leaves behind on-premises**, which is how the tenant's existing GUID mismatches
got there in the first place.

**The fix is to de-duplicate against what Exchange actually returned**, not
against an assumption about it. The Exchange pass now collects the
distinguished names and primary addresses of the rows it produced, and the AD
pass skips exactly those. `-ExchangeAnswered` (a switch) became
`-AlreadyReturned` (the keys).

### What Enable mailbox decides on

The button's eligibility is now keyed on the **on-premises** object and nothing
else. `Enable-RemoteMailbox` acts on an Active Directory user, so the only
question that matters is whether that user already has a mailbox on-premises --
an on-premises mailbox or a remote mailbox stub.

What Exchange Online holds is **information, not a veto**. Keying the refusal
on the cloud type meant the button refused the exact case it was asked for: a
cloud mailbox with no on-premises stub. A `MailUser` counts as "no mailbox" for
this purpose -- it is an AD account with mail attributes and no mailbox of any
kind, which is precisely what a cloud-first mailbox leaves behind.

**It runs what New User runs, and nothing more.** `Enable-RemoteMailbox`, then
the archive. When a cloud mailbox already exists the confirmation says so and
says the new stub's ExchangeGuid will not match it -- stated rather than
blocked, and not silently repaired, because stamping GUIDs is a different piece
of work from enabling a mailbox and the tenant has a backlog of those to
consider as one job.

### A greyed button has to say why somewhere you can read it

WinForms does **not** show a ToolTip for a disabled control. So the explanation
attached to the Enable button existed only in the one state where it could
never be read: grey button, no tooltip, nothing on screen. That is
indistinguishable from broken, and it was reported as broken. The reason now
goes on the status line whenever a row is selected that the button cannot act
on. Any other greyed button in the toolkit that carries its reason only in a
tooltip has the same fault.

## 77. The Room Editor and the Group Editor (2026-09-18)

Two more consolidations on the Mailbox Editor's pattern: one target, tabs
across it, per-section Apply, no global Save, every write audited against the
object.

### RoomEditor.ps1 -- nine tools into eight tabs

Overview, Booking, Who Can Book, Replies, Place, Regional, Room Lists,
Calendar. **The tab order is the order of the questions**, not alphabetical
and not the order the old tools happened to be written in: "why will it not
accept" is Booking, "why will it not accept for HER" is Who Can Book, "why can
nobody find it" is Room Lists and Place.

The Overview verdict is the reason the window exists. *The room is not
accepting bookings* has about six causes living on five tabs, and reading them
one tab at a time is exactly what the old seven tools made you do.

**What each old tool contributed, and what must not be lost:**

- **CalendarProcessing.ps1** -- the writable parameter set is discovered from
  the cmdlet at run time rather than hand-listed, so a setting Microsoft adds
  next year appears without this file changing. Its **INERT note column** is
  carried over whole: most booking settings do nothing unless
  AutomateProcessing is AutoAccept, Exchange accepts the write either way, and
  the ticket that follows says the tool does not work.
- **Set-BookingPolicy.ps1** -- the bug it was written to fix is now stated on
  the Overview verdict *and* prevented by construction: `BookInPolicy` and
  `AllBookInPolicy` are on the same tab and written in **one call**. Setting
  the list without the switch leaves a room accepting everything while holding
  a list of four names that nothing reads.
- **RoomList.ps1** -- membership is asked per list, because a room carries no
  usable "lists I belong to" property. A list that will not read shows `?` and
  refuses both buttons; saying "not a member" about a list we could not read
  would invite adding the room twice. **Creating** a room list stays in
  RoomList.ps1: it makes a tenant-wide object and belongs to no room.
- **RoomPlace.ps1** -- the hybrid-blocked address block is refused *before* the
  call, and **every write is read back and compared**, because Set-Place
  reports success either way. Both safeguards, not one: the old tools had one
  each.
- **RoomRegional.ps1** -- the two timezones (regional, and working-hours) are
  on one tab with a mismatch warning, because ScheduleOnlyDuringWorkHours
  depends on them entirely and they live in different cmdlets. Timezones are a
  list built from the machine: Exchange wants `Central Standard Time`, not
  `CST` and not an IANA name.

**Deliberately left out:** Protocols.ps1 (tenant-wide report, no per-room
dimension), ResourceDuplicator.ps1 (creates a directory object across two
sessions via a capture file -- a provisioning wizard, not an editor),
CalendarReset.ps1 (a destructive bulk flow whose Preview-then-Apply gate a tab
would dilute).

### Group-Editor.ps1 -- the router and both editors in one window

The three things that looked like group editors were really a **router plus
two half-editors**: `Group-Manager.ps1` decided which to open,
`Group-Manager-Editor.ps1` did distribution lists properly on both sides, and
`GroupEditor.ps1` existed because a Microsoft 365 group answers
`*-DistributionGroup` with *"The current operation is not supported on
GroupMailbox"*.

The routing did not go away -- the cmdlet families really are different. It
moved inside, where it decides which tabs are live rather than which program
to start.

| Kind | Owners and members | Notes |
|---|---|---|
| Distribution list, mail-enabled security | `ManagedBy`, `*-DistributionGroupMember` | on-premises or cloud by where it is mastered |
| Microsoft 365 group | `*-UnifiedGroupLinks` | no on-premises half, ever |
| Dynamic distribution group | none -- membership is a filter | shown, not editable |

**The rules it keeps:**

- **`-BypassSecurityGroupManagerCheck` on every distribution-group write.**
  Exchange checks whether *you* are in the group's ManagedBy regardless of
  your RBAC role. It is **not** sent to `Set-DynamicDistributionGroup`, which
  does not declare it -- that would fail at parameter binding before reaching
  Exchange.
- **`ManagedBy` and `AcceptMessagesOnlyFromSendersOrMembers` are replaced
  whole.** Adding one owner means writing the other four back, and getting it
  wrong does not error -- it silently removes everybody else.
- **Send As is genuinely two different things.** On-premises it is an AD
  permission taking the group's **distinguished name** and the trustee's
  **sAMAccountName** with `'Send-As'`; in the cloud it is a recipient
  permission taking two SMTP addresses with `'SendAs'`. The two spellings are
  not interchangeable and neither errors usefully when given the other's.
  Removing a **Deny** needs `-Deny` or the cmdlet removes the Allow
  underneath, reports success, and changes nothing.
- **Everything is resolved and read in the directory that will be written.**
  Exchange Online's `Get-Recipient` does not populate SamAccountName, and an
  on-premises write is invisible in the cloud until the next sync -- so a
  cloud re-read after an on-premises Apply repaints the tab from the
  pre-change copy and contradicts what was just done.
- **`-ResultSize Unlimited` everywhere.** The default is 1000 and it truncates
  silently; a partial member list is how people get removed from groups they
  were never seen in.
- **Adds before removals**, so a run that fails halfway leaves the group with
  too many members rather than too few.

**Deliberately left out:** creation and deletion (an editor that can delete
its own subject is a different kind of tool), the bulk many-to-many membership
tools, the tenant-wide reports, and `DLtoSharedMB.ps1`, which is a multi-day
workflow that deletes its own subject halfway through.

### One more thing both of these did to Common.ps1

`New-TabLabel`, `New-TabValue`, `New-TabButton`, `New-TabGrid`,
`Add-GridPickColumn`, `Get-GridPicked`, `Set-GridPicked`, `New-DialogCheck`,
`New-DialogList`, `ConvertFrom-DialogList` and `Show-ResultListDialog` moved
out of MailboxEditor.ps1 and into Common.ps1 the moment a second editor needed
them. Three private copies of a tick-box column is three places for "the
selected rows" to quietly come to mean three different things.

## 78. Three editors, one shape (2026-09-18)

The Mailbox Editor is the reference and the other two follow it. Where they
had drifted, the drift was closed in whichever direction was already right.

### The Editing switch is now built once, in Common.ps1

`New-EditorModeSwitch` puts the identical Cloud / On-premises drop-down in the
same place with the same wording in all three windows, and
`Set-EditorModeAvailability` and `Sync-EditorModeCombo` go with it. Three
hand-typed copies of **the control that decides which directory a write lands
in** is three chances for one of them to be labelled the other way round.

`Set-EditorModeAvailability` deliberately does **not** write its reason to the
status line by default. Both new editors load a target immediately after
probing, and the load's own status message overwrote the explanation before
anybody could read it -- on top of WinForms refusing to show a tooltip on a
disabled control. The reason goes on a banner that survives.

### Hide-from-the-address-book moved to the first tab in all three

It was on Settings in the Mailbox Editor, on Overview in the Room Editor, and
missing entirely from the Group Editor. *Can people find this* is a first-tab
question, so the Room Editor's placement won. The old locations keep a
**read-only echo** pointing at Overview, because somebody who has used the
tool for a fortnight will look where it used to be and an absence reads as a
bug.

### The Group Editor was blind to half the directory

`Resolve-GroupTarget` asked Exchange Online and nothing else. A group in a
child domain, a group that had not synchronised, a group simply not
mail-enabled in the cloud -- all came back as *"not found in Exchange
Online"*, and **every on-premises group in a directory the cloud does not
carry showed nothing at all**. That was reported as a whole category of group
being invisible.

Two fixes, and neither of them needs a per-organisation build:

- **Both directories are asked**, in the order the caller implies: on-premises
  first when on-premises was requested, cloud first otherwise. Which one
  answered is recorded in `FoundIn`, and a group the cloud has never heard of
  is treated as on-premises whatever `IsDirSynced` says -- that property is
  meaningless on an object the cloud did not return.
- **`Set-ADSessionScope` runs before the on-premises read.** A child-domain
  object is invisible to a session scoped to its own domain, and *invisible
  looks exactly like absent*. The helper was already there for other tools.

It also refuses to claim the on-premises side when there is no on-premises
session: doing so sent the detail read to a module that is not loaded and made
the group look unreadable when the cloud copy was sitting right there. The
banner now says which copy is on screen and why.

### What a refused mode switch must not do

`Resolve-GroupTarget` assigns `$script:Grp` internally before its detail read,
so that the read can be routed to the right directory. That means a failed
resolve leaves the variable pointing at the broken record, and *not assigning
the result* is not enough to undo it. `Set-GrpMode` now saves the previous
object and mode and puts both back, including the drop-down.

And the object decides the mode, not the box: a Microsoft 365 group forced to
on-premises comes back cloud-side whatever was asked for, and a drop-down left
saying otherwise is a drop-down that lies about where the next write lands.

## 79. An empty list is not an empty group (2026-09-18)

A distribution list in a **child domain** loaded its identity in the Group
Editor and then showed **no members at all**. The group had several.

Two causes, and the second is the one that made it invisible.

**The session was not forest-wide.** An Exchange session sees its own domain
unless told otherwise, and a group whose members live in another domain comes
back with **no members rather than an error**. `Set-ADSessionScope` was being
called only while probing the on-premises side during the lookup -- so a
group the *cloud* answered for first (which every synced group does) never got
forest scope at all, and every later on-premises read ran narrow.

**The controller was the wrong one.** The old code pinned the first configured
directory that resolved to a server. For an object in a child domain that is
usually the *parent's* controller, and a controller that cannot resolve the
members drops them quietly. The directory is now read off the group's own
distinguished name -- the `DC=` parts reassembled into a domain and matched
against `Directories` in Tenant.json, **longest match wins**, because a child
domain's name contains its parent's and only the child is right. When nothing
matches, **nothing is pinned**: forest-wide with no pin is correct and safe,
and a wrong pin is not.

The scope is also set inside `Resolve-GroupTarget` before its own detail read,
which is the first on-premises call for a group and used to run with whatever
the *previous* group had pinned.

### The part that made this hard to notice

Nothing errored. A blank grid under a correct-looking identity banner is
exactly what a group with no members looks like, so the tool was confidently
reporting the wrong answer. That is the worst failure mode a read-only view
has, and the fix is not only to read correctly but to **refuse to present an
unverified empty result**:

> When the on-premises read returns zero members for a synced group, the cloud
> copy is counted as a cross-check. If the cloud has members and on-premises
> does not, the tab says so in red, names the scope and controller in use, and
> disables the membership buttons -- rather than showing an empty grid.

The same shape of check is worth having anywhere a count of zero is both a
legitimate answer and a plausible symptom.

## 80. Status colours follow the theme (2026-09-18)

A warning in the Group Editor was reported as unreadable: dark red text on a
dark console. It was, and the fault was wider than one message.

**Eleven of the twelve shipped themes have a dark form background.** The
editors used hand-picked colours -- `180,0,0` for failures, `150,80,0` for
warnings, `DimGray` for supporting text -- chosen against a white form that
almost nobody runs. On a near-black background the most important text in the
window was the least visible.

`MailboxHealth.ps1` had already solved this, for itself, and even said why:

> A READABLE RED ON BOTH A DARK AND A LIGHT THEME. Pure `Color::Red` is
> unreadable on the dark palettes and shouts on the light ones.

That fix sat in one file while every other tool kept the unreadable version.
It is `Get-ToolStatusColour` in Common.ps1 now, with four kinds -- Bad, Warn,
Good, Muted.

**The background decides, not the theme name.** A theme is dark or light by
the luminance of its `FormBack`, so a palette added next year gets the right
answer without this function knowing it exists. The theme's own `DangerColor`
is preferred for Bad, but only when it is genuinely separated from that
theme's background -- two shipped themes carry a blank or the literal string
`None` there, and a red that works on paper and not on screen is the same
fault in a different colour.

`New-TabLabel -Muted` goes through it too, so every supporting label in every
tabbed editor follows the theme without each tool deciding for itself.

**105 hardcoded colours** came out of the three editors and 41 more out of
Unified Lookup. Nineteen other tools still use `Firebrick` and `ForestGreen`
on status labels and have the same problem; they are listed in the backlog
rather than swept blind, because a few of those uses are on white NoTheme
boxes where a dark red is correct and a mechanical replace would break them.

**The general lesson, which is not about colour.** The readable-red fix
existed, was correct, was commented, and stayed private to the file that
needed it first. A fix that solves a problem the whole toolkit has belongs in
Common.ps1 the moment the second caller appears -- and the second caller is
usually already there, silently getting it wrong.

---

## 81. Five more tabs, one Settings tab gone, and a window that stayed a window (2026-09-18)

Glenn asked six things of the Mailbox Editor in one message. Five are done.
The sixth is honestly reported as not done, because doing it would have meant
shipping a change that cannot be tested here and that fails in a way an
operator would not recognise as a bug.

### The Settings tab is gone entirely

Hide-from-address-lists moved to Overview in the previous batch, which left
Settings holding two things: change mailbox type, and enable archive. Both of
them belong on Overview too -- they are facts about what the mailbox IS, and
that is what Overview is for. So both moved and the tab was deleted rather
than left standing with nothing in it.

Neither needs the explanation it used to carry. Enable Archive in particular:
the button is simply greyed with **Archive already enabled.** beside it when
there is one. A disabled button plus a sentence says everything a paragraph
said, and cannot be out of date.

`$cmbOvType` carries the four types the toolkit converts between. A mailbox
whose `RecipientTypeDetails` is not one of them -- a scheduling mailbox, a
team mailbox -- now clears the selection rather than leaving the PREVIOUS
mailbox's choice sitting in the box while the note beside it says something
else. An empty picker is a true statement; a stale one is not.

### Inbox Rules is a tab now

The Health tab already counted rules, which is the point at which somebody
wants to see them. `Update-RulesTab` lists them; `Show-RuleDetail` opens one;
`Invoke-RuleVerb` enables, disables or removes.

`Invoke-RuleVerb` builds the cmdlet name from a **literal** per-verb string,
not by concatenating the verb onto `-InboxRule`. That is deliberate: the
capability scanner reads source for cmdlet names, and a name assembled at
runtime is invisible to it. Three literals cost three lines and keep the tool
honest about what it calls.

### Recover Items, and the parameter that excluded the whole point

The Recovery tab restores a deleted MAILBOX. This is the other thing the EAC
does: restore an ITEM a user deleted, including after Deleted Items was
emptied.

The first cut passed `-SourceFolder DeletedItems` to `Get-RecoverableItems`.
That is not a filter meaning "things that were deleted" -- it is the name of
one specific folder, and passing it excluded Recoverable Items and Purged
Items, which is exactly the case the tab exists for. Omitting the parameter
searches all three. It is also omitted on the `Restore-RecoverableItems` call,
because naming a source at restore time that the search did not use is a good
way to restore nothing.

The grid gained a **Was in** column and lost a **From** column.
`Get-RecoverableItems` does not return a sender. A column that is always empty
is worse than no column, because it reads as "this message had no sender"
rather than "this tool cannot tell you".

### Quarantine, and a cap that now says so

`Get-QuarantineMessage` pages. The tab pulls 1,000 per page up to twenty pages
and, **if it hits that ceiling, says so in red**: the note reads that this is
not all of them and to narrow the window or the type. A truncated list
presented as a complete one is the single worst thing a search tool can do --
the operator concludes the message is not in quarantine and goes looking
somewhere else.

Phish and malware release asks a second, differently-worded question. Releasing
a phish to a user's inbox is not the same act as releasing a false-positive
newsletter, and the confirmation should not pretend it is.

### Automatic Replies: it CAN be a tab, and it is still a window

Glenn asked whether the out-of-office editor has to open in an external
screen. The answer is no -- and it still does, for now.

**A plain PowerShell scriptblock is not a closure.** It does not capture the
variables around it; it resolves them when it FIRES, against whatever scope is
live at that moment. `Show-OutOfOfficeForm` gets away with this today only
because `ShowDialog()` blocks inside the function, so the function's own scope
is still on the stack every time a handler runs.

Build it into a tab and the function returns immediately. Its scope is gone by
the time anyone clicks anything, and all fifteen handlers then resolve against
the CALLER -- MailboxEditor.ps1. There, `$wbInt`, `$wbExt`, `$btnSave` and
`$doLoad` are `$null`, which is survivable. The unsurvivable part is that
`$form`, `$tabs` and `$lblStatus` all EXIST there and are the wrong objects:
the editor's internal/external switch would drive the Mailbox Editor's own
TabControl, its status line would write into the main window, and its Close
button -- `{ $form.Close() }` -- would close the entire Mailbox Editor.

The fix is `.GetNewClosure()` on every handler in that function. That is real
work with real testing, not a parameter added at the end of a long batch, so
the `-Container` parameter is present in `Show-OutOfOfficeForm` and **throws**
if anyone passes it. Failing loudly at the call site beats failing strangely
three clicks later.

What the tab does instead: it shows the state, the audience, the schedule and
the internal message with its HTML stripped, so the whole picture is readable
without leaving the editor, plus a button that opens the real editor. The
stripped text is display-only and is never written back -- the editor owns the
value, so losing formatting in the preview costs nothing.

### The near-miss worth recording

The Automatic Replies replacement was done as a string splice: everything
before the Replies region, plus the new text, plus everything from the Inbox
Rules region onward. But Inbox Rules had been inserted AFTER Statistics
earlier in the same batch, so "everything from Inbox Rules onward" started 551
lines further down than intended and the splice **deleted the Recovery, Health
and Statistics regions**.

It would not have been subtle in use -- `Invoke-EditorLoad` calls
`$gridStats.Rows.Clear()` and would have thrown on the first Load of any
mailbox -- but nothing in the edit itself complained. Bracket balance was
fine. Encoding was fine. A 551-line hole passes every check that counts
characters.

What caught it was reading the file afterwards and asking which functions
should be there. What fixed it was `shipEX\`, which had the previous version.
The lesson is the reason the staging folders exist at all, and the reason an
edit that splices by index should assert on what it is about to remove.

