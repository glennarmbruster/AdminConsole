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
