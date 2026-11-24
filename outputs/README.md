# Exchange / Microsoft 365 admin tooling

Four independent things live here. Each folder is **self-contained** — every
script resolves its neighbours through `$PSScriptRoot`, so a folder can be copied
anywhere and still works.

| Folder | What it is | Status |
|---|---|---|
| **`AdminConsole/`** | The live console. Hardcoded to expresspros.com on purpose. Hybrid: on-prem AD, Exchange Online, SharePoint. | In daily use |
| **`FH-Agnostic/`** | The same toolkit with nothing tenant-specific in it. Single forest, hybrid or cloud, every location configurable. | Ready; not yet run against a live tenant |
| **`CloudOnly/`** | `FH-Agnostic` with every Active Directory dependency removed rather than made optional. For tenants with no on-premises directory at all. | Built 2026-08-03 |
| **`obsolete/`** | Archived: ANZ, SRG, and the paused browser-hosted suite. Nothing here is referenced by anything that runs. Has its own README. | Reference only |

**`FH-Agnostic` and `CloudOnly` are named for different things on purpose.** One
names the environment shape it was built against; the other names the hard
constraint it enforces. Both are tenant-agnostic in the sense that matters —
nothing in either is hardcoded to an organisation, a domain or a server.

Supporting material, not deployable:

- **`ScriptNotes/`** — the reasoning behind the code, kept out of the scripts so
  they can be shared without shipping the thinking. `ARCHITECTURE.md` is the
  cross-cutting one; the rest mirror script names, extension included.

Everything archived now lives under `obsolete/`, including `obsolete/_archive/`
— the pre-comment-stripping originals and the extraction tooling that produced
`ScriptNotes/`. Three folders at the top level are live; the fourth is history.

## Where to start

**`AdminConsole/AdminConsole.ps1`.** Run it from a PowerShell window that already
has your connections — Exchange Online and SharePoint sessions belong to the
process that made them, so connecting in a different window does not count. The
console will say so if you do.

`AdminConsole/ToolManager.ps1` edits what appears on it: tool buttons, sections,
theme, and the paths and Edge profile names in `Settings.json`. Run it directly —
it is deliberately not a button on the console.

## `Config\` — the ownership line

Every JSON lives in `Config\`: `Tools.json`, `Sections.json`, `Themes.json`,
`Settings.json`, plus `Images\` for tenant logos. Nothing else does.

The split exists so that shipping an update is **"replace the files, leave
`Config\` alone"** rather than *"copy these forty files but not those four"* —
an exception list is a rule that gets broken, and the shipped toolkits are
installed by hand, by someone else, roughly monthly. An install predating this
layout is migrated automatically on first launch, and the console reports what
it moved rather than doing it silently.

Output is a third category and stays out of `Config\`: `reports\` (with
`reports\captures\` beneath it) and `Logs\`, both relocatable through
`Settings.json`.

## Two facts that explain most of the design

**On-prem and cloud cmdlets collide.** `Get-Mailbox`, `Set-Mailbox`,
`Get-DistributionGroup` and friends exist in *both* the Exchange Management Shell
and the Exchange Online module, and which one answers depends on module load order
rather than on anything visible at the call site. Hence the warnings in the
distribution-list tools, and hence the console greying out tools whose
prerequisites are missing instead of letting them run and guess.

**Tools launch with `&`, inside the console's own process.** They inherit its
connections — which is what makes the grey-out accurate — but *not* its `$script:`
scope. A `$script:` variable read inside a function resolves against the launched
script's scope, where it was never assigned, and comes back `$null`. Paths come
from `$PSScriptRoot`; shared state goes in a global.

## Encoding

Every `.ps1` is UTF-8 with a **single BOM and CRLF** line endings. Not cosmetic:
an LF-only file renders as one continuous line in the ISE, and a double BOM leaves
an invisible zero-width character inside the first line that breaks the `#Requires`
directive. Preserve both if you edit outside the ISE.
