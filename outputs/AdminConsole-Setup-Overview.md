# Admin Console — the config files, and setting up for the first time

*September 2026*

The console itself holds no settings. Everything it knows comes from six JSON files in
`Config\`, and everything an individual operator changes about how it looks comes from a
seventh file named after their Windows account. Understanding which file owns what is most
of understanding the console.

---

## The files in `Config\`

**`Tenant.json` — what this organisation HAS.** The most important file. It names the tenant,
its domains, and a `Capabilities` block of true/false flags: Exchange Online, SharePoint
Online, on-premises AD, on-premises Exchange, PIM. A capability set to false greys out every
tool that needs it — that is the intended way to hide functionality, rather than deleting
tools. It also carries the on-premises Exchange server list, licence groups, the SharePoint
admin URL, the tenant's theme and colour band, and the parked list (below). In multi-tenant
mode there is one of these per tenant, in `Config\Tenants\<key>.json`.

**`Tools.json` — what tools exist.** One entry per button: `Name`, `Script`, optional
`ObjectTypes` (which kinds of object the tool accepts — a tool declaring `Mailbox` greys out
until a mailbox is looked up), `Section`, and optional `Group`. Give several tools the same
`Group` word and they collapse behind a single console button, outlined in white so it is
obvious at a glance that more sits behind it. Edit through **Tool Manager**, not by hand.

**`Sections.json` — the shape of the console.** The section names in order, how many rows
each gets, an optional accent colour, and `ButtonsPerRow` for the whole window. Sections are
the bands the buttons sit in — Connect, Daily Scripts, Mailbox and Access, and so on.

**`Themes.json` — the palettes.** A named set of colours each, plus `ActiveTheme`. A theme
with `Enabled: false` is ignored, which is the quiet one: the save succeeds and the console
opens in the old colours with nothing saying why. Edit through **Theme Manager**.

**`QuickLaunch.json` — shortcuts.** Links and commands for the Quick Links button.

**`Settings.json` — how the toolkit behaves.** Paths (reports, logs, retry queue), the
operator email, the console title, and the look: button style, corners, icons, dialog style.
Note that the preferred domain controller is *not* here — it belongs to the organisation, so
it lives in `Tenant.json` under `OnPremExchange`.

Two of these ship as `.TEMPLATE.json` — copy to the real name and fill in.

> **The standing rule when upgrading:** never copy `Config\` over a working install. The
> `Config` in a package is a neutral scaffold. Copy the `.ps1` files and leave `Config\`
> alone, or you will overwrite the tenant's own settings with an empty skeleton.

---

## First-time setup

1. **Unblock the files.** Anything mailed or downloaded carries a block flag that stops
   PowerShell running it: `Get-ChildItem -Recurse -File | Unblock-File`.
2. **Fill in `Tenant.json`.** Copy `Tenant.TEMPLATE.json` if there is no real one yet. Name,
   primary domain, the `.onmicrosoft.com` domain, and then the capability flags — be honest
   here, since a flag set true for something the tenant does not have produces tools that
   fail rather than tools that are greyed out. Use **Tenant Manager** rather than an editor;
   it validates as you go.
3. **Check `Settings.json` paths.** Reports and logs default to folders under the toolkit.
   Point them at a share if the team needs to read each other's output.
4. **Run `.\AdminConsole.ps1`.** Connect buttons are at the top; the status strip along the
   bottom shows which services this PowerShell window is signed in to.
5. **Run Test-Toolkit** from the Dashboard. It checks the config files against the scripts
   actually present and reports what is missing, mis-referenced, or pointing at nothing.

If the console opens unthemed or in unexpected colours, run `Diagnose-Theme.ps1` from a
prompt. It prints every step of theme resolution and names the winner, which is faster than
guessing.

---

## Multi-tenant

Multi-tenant mode is **absent-means-off**. One `Config\Tenant.json` and nothing else: the
console behaves exactly as it always did. Create `Config\Tenants\` with one file per tenant
and the console becomes tenant-aware.

Start it with `.\AdminConsole.ps1 -Tenant <key>`, where the key is the file name without
`.json` (a nickname also works). With no `-Tenant`, the tenant marked default is used. A name
that matches nothing is fatal on purpose — opening the wrong organisation because of a typo
is the failure the whole feature exists to prevent.

Every window carries a **colour band** in the tenant's colour, and the title bar names the
tenant. This matters more than it sounds: two consoles open side by side look identical
otherwise, and the tools they launch write to real mailboxes.

**One tenant per PowerShell window.** Switching tenant swaps a config file; it does *not* swap
the Exchange Online session, the on-premises session, SharePoint, Graph, or whatever PIM
activated — those belong to the window and outlive the console. Opening a second tenant in a
connected window is refused, with instructions to open a new window. Take that refusal at
face value.

Each tenant file carries its own theme, band colour, and parked list, so hiding a tool for one
tenant leaves the others alone.

---

## Multi-user

Several people can share one toolkit folder and see different consoles. Personal settings live
in `Config\Users\<windows-account>.json`, created the first time that person saves something —
before that the file does not exist and the shared config is simply used as-is.

A profile holds the theme, the number of columns, the sections and their row counts, where
each tool sits (`Placement`), and a personal parked list. It has a `Default` block that applies
everywhere and an optional block per tenant that wins over it, key by key — so someone can have
one arrangement generally and a different one for a particular tenant.

Four windows ask **"just me, or everyone?"** when saving, and the answer routes the write:

| Window | "Just me" writes to | "Everyone" writes to |
|---|---|---|
| Theme Manager | your profile | the tenant file |
| Tool Manager (Sections) | your profile | `Sections.json` |
| Park In Place | your profile | the tenant file |
| Save Arrangement (drag-and-drop) | your profile | `Sections.json` + `Tools.json` |

**Tenant parking always wins over personal.** A tool the tenant has parked stays hidden no
matter what an individual profile says; a personal park hides a tool for that person only.
Park In Place shows both, with a `HiddenBy` column saying which list is doing it.

**Save / Load Layout** on the Dashboard exports a profile to a file and puts one back — useful
for carrying an arrangement to another machine, or for handing a new starter a sensible
starting point. Reset returns to the shared layout; resetting the whole profile renames the
file rather than deleting it, so an afternoon of arranging is recoverable.

Nothing personal ever changes what a tool *does*. A profile says where a button sits and
whether it is visible. What the tool is, where it points, and what it is allowed to do stay in
`Tools.json` and `Tenant.json`, shared by everyone.

---

## Two things that catch people out

**A theme saved as "just me" does not appear until the console re-reads the profile.** The
console notices the file changing and re-reads it, but a console opened *before* the save
keeps what it had. Close it and open it again.

**The console does not reload config while running.** Tool Manager, Tenant Manager and Theme
Manager all say "reopen the Admin Console to see it", and they mean it.
