# Exchange & Microsoft 365 Admin Toolkit

A WinForms console that launches focused admin and reporting tools for a hybrid
Exchange environment: one Active Directory forest, one Microsoft 365 tenant.

Nothing in this toolkit is hardcoded to an organisation, a domain or a server.
It runs where it is dropped.

## Requirements

- Windows PowerShell 5.1
- `ActiveDirectory` module, for the on-premises tools
- `ExchangeOnlineManagement`, for the cloud tools — `Connect-ExchangeOnline`
- `Microsoft.Online.SharePoint.PowerShell`, for the OneDrive and sharing tools —
  `Connect-SPOService`
- On-premises Exchange Management Shell, for the distribution-list tools

Connections are **per-process**. Connect in the same window that runs
`AdminConsole.ps1` — connecting in another window does not count, and the console
will correctly tell you it is not connected.

## Getting started

1. Put the whole folder anywhere. `E:\Scripts\Admin`, a share, a USB stick — all
   equivalent. No path is baked in.
2. Connect to whichever services you need, in one PowerShell window.
3. Run `.\AdminConsole.ps1` from that same window.

Tools whose prerequisites are missing grey out, and hovering one explains what is
absent. That is derived by reading each script, so it stays correct as the
scripts change.

## Configuration

Two settings, both blank by default and both resolved at run time. Set them in
**ToolManager → Settings…**, which writes `Settings.json`.

| Setting | Blank means |
|---|---|
| Mail domain | Work out the domain at run time — see **Addresses** below. Set it and a bare alias completes immediately, with no connection needed. |
| Domain controller | Let the `ActiveDirectory` module bind to the nearest DC. Right answer in a single forest, and survives a DC being down. |
| Reports folder | A `reports\` folder beside the toolkit. Point it at a share to give the team one set of reports. Resource captures live in `reports\captures\` and follow it. |
| Logs folder | A `Logs\` folder beside the toolkit, holding `CompletedTasks.csv`. |
| Edge profiles | The placeholder names `Admin` and `Standard`. **Set these** — see below. |

**No path anywhere in this toolkit is hardcoded.** Everything resolves from
`$PSScriptRoot` or from a setting, so nothing assumes a particular drive letter
exists. A relative value in a folder setting is taken as relative to the toolkit
folder; an unreachable share falls back beside the toolkit with a warning rather
than failing to write and leaving no trace of why.

### Edge profiles

Two tools open a browser: the M365 portal launcher wants the elevated account,
everyday links want the ordinary one. The profile names ship as `Admin` and
`Standard`, which are **placeholders and almost certainly wrong on your machine**.

Getting them wrong fails quietly: a profile that does not exist falls back to the
last-used one, which usually works, so the symptom is a portal opening as the
wrong account rather than an error. The Settings dialog lists the profiles Edge
actually has on the machine, so pick from that list.

These are per-machine — two people running this folder from a share will not have
the same profile names.

`Settings.json` is a separate file rather than values written into `Common.ps1`
on purpose: `Common.ps1` gets replaced when the toolkit is updated, and a setting
stored there would vanish on every update — silently, with a symptom (aliases
stop completing) that looks nothing like the cause.

`Common.ps1` still has `$script:MailDomain` and `$script:DomainController` as a
code-level fallback if you would rather not keep a `Settings.json`.
**`Settings.json` takes precedence over them**, so a value set in the UI always
takes effect.

### Addresses

A bare alias is completed to a full address when the domain can be established
with confidence, in this order:

1. `Settings.json` — **ToolManager → Settings…**
2. `$script:MailDomain` in `Common.ps1`, the code-level fallback.
3. The connected tenant's primary accepted domain.
4. The mail domain on your own directory account.

If none of those answer, the alias is passed through untouched and the lookup
reports it as not found — so **full email addresses are required until a domain
is set or Exchange Online is connected.**

The AD DNS domain is deliberately **not** used as a fallback. A forest called
`corp.example.net` routinely sends mail as `example.com`, so guessing from it
would build a plausible wrong address and then blame your typing for the failure.

Set it once in Settings… and you can type `jsmith` instead of
`jsmith@your-domain` everywhere, before connecting to anything. The dialog offers
the tenant's accepted domains as a dropdown when Exchange Online is connected,
and shows you what the console will actually use as you type.

Everything else is configured through the UI:

- **`Config\Tools.json`** — which tools appear, in which section, in what order. Edit
  through `ToolManager.ps1` rather than by hand.
- **`Config\Sections.json`** — the console's sections and how many button rows each
  gets. Edit through **Sections…** in ToolManager.
- **`Config\Themes.json`** — colour palettes. Edit through **Theme…** in ToolManager,
  which opens `ThemeManager.ps1`.
- **`Config\Settings.json`** — mail domain, domain controller, folders and Edge
  profiles. Edit through **Settings…** in ToolManager.

Drop a `Config\Images\logo.png` in and the console picks it up. A
`logo.<your-domain>.png` takes precedence if you run more than one copy.

## What's here

### Admin tools

Mailbox Health · Message Trace · Inbox Rules · Calendar Permissions ·
Out of Office · Mailbox Permissions · Mail Attributes · All Attributes ·
Group Editor · Microsoft 365 Group Editor · DL Memberships · DL Permissions ·
Mailbox Audit · Shared Mailbox Audit · Mailbox / OneDrive Permissions ·
PowerShell Console

### Reports

All reports write a timestamped CSV to a `reports\` subfolder, show a progress
bar, and never write an empty file — a zero-row CSV reads as "found nothing"
rather than "the run failed". An incomplete run is labelled `PARTIAL` in both the
summary and the log.

| Report | Finds |
|---|---|
| Protocols | OWA / ActiveSync / POP / IMAP / MAPI / EWS state per mailbox |
| DL Inventory | Mail-enabled groups, with nesting, owners and staleness |
| External Forwarding | Mailboxes and inbox rules forwarding outside the org |
| Permissions Audit | Full Access, Send As and Send on Behalf delegations |
| Ownerless Groups | Microsoft 365 groups with no owner, or only disabled owners |
| Orphaned OneDrive | OneDrive sites whose owner is disabled or gone |
| External Sharing | Per-site sharing capability and the guests who have access |
| M365 Portal Reports | Launcher for the portal reports Microsoft already computes |
| Report Browser | Every CSV in `reports\`, with row counts. Needs no connection. |

## Two things worth knowing before you edit

**On-premises and cloud cmdlets collide.** `Get-Mailbox`, `Set-Mailbox`,
`Get-DistributionGroup` and friends exist in *both* the Exchange Management Shell
and the Exchange Online module. When both are loaded, which one runs depends on
module load order — not on anything visible at the call site. The distribution-list
tools warn about this in their own windows; don't run them in a window with a
cloud session open.

**Scripts are launched with `&`, in the same process.** They share the console's
connections, which is what makes the grey-out accurate. They do **not** share its
`$script:` scope: a `$script:` variable read inside a function resolves against
the launched script's scope, where it was never assigned, and comes back `$null`.
Use `$PSScriptRoot` for paths and a global for shared state.


## Updating the toolkit

**Replace the `.ps1` files. Leave `Config\` alone.**

That is the whole rule, and it is deliberately not a list of exceptions. Every
file you have customised — tool buttons, sections, colours, paths, your logo —
lives under `Config\`. Everything outside it is code and can be overwritten
wholesale without losing anything.

If you are upgrading from a version that kept the JSONs loose in the folder, you
do not need to move anything by hand: the console detects the old layout on
first run, moves the files into `Config\`, and prints what it moved. Your
settings survive the upgrade.

`reports\` and `Logs\` are output, not configuration, and are untouched either
way.

## File encoding

Every `.ps1` is **UTF-8 with a single BOM and CRLF line endings**. This is not
cosmetic — an LF-only file renders as one continuous line in the ISE and Notepad,
and a double BOM leaves an invisible zero-width character inside the first line
that breaks the `#Requires` directive. Keep it that way if you edit outside the
ISE.
