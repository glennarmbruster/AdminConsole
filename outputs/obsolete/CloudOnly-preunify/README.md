# Exchange & Microsoft 365 Admin Toolkit

**Cloud-only build.** A WinForms console that launches focused admin and
reporting tools for a Microsoft 365 tenant with no on-premises directory.

There is no Active Directory dependency anywhere in this folder — no
`ActiveDirectory` module, no domain controller setting, no on-premises Exchange
Management Shell. Exchange Online *is* the directory. If your environment is
hybrid, use the `FH-Agnostic` build instead; this one will report cloud-only
answers for hybrid objects and will not tell you what the on-premises side
thinks.

Nothing in this toolkit is hardcoded to an organisation, a domain or a server.
It runs where it is dropped.

## Requirements

- Windows PowerShell 5.1
- `ExchangeOnlineManagement` — `Connect-ExchangeOnline`
- `Microsoft.Online.SharePoint.PowerShell`, for the OneDrive and sharing tools —
  `Connect-SPOService`

Those two are the whole list. The console probes for exactly those two
capabilities, so a tool can never be greyed out for a reason you cannot fix from
this machine.

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

All blank by default and resolved at run time. Set them in
**ToolManager → Settings…**, which writes `Settings.json`.

| Setting | Blank means |
|---|---|
| Mail domain | Work out the domain at run time — see **Addresses** below. Set it and a bare alias completes immediately, with no connection needed. |
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

`Common.ps1` still has `$script:MailDomain` as a code-level fallback if you would
rather not keep a `Settings.json`. **`Settings.json` takes precedence over it**,
so a value set in the UI always takes effect.

### Addresses

A bare alias is completed to a full address when the domain can be established
with confidence, in this order:

1. `Settings.json` — **ToolManager → Settings…**
2. `$script:MailDomain` in `Common.ps1`, the code-level fallback.
3. The connected tenant's primary accepted domain.

If neither answers, the alias is passed through untouched and the lookup reports
it as not found — so **full email addresses are required until a domain is set or
Exchange Online is connected.**

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
- **`Config\Settings.json`** — mail domain, folders and Edge profiles. Edit
  through **Settings…** in ToolManager.

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
| DL Inventory | Distribution, mail-enabled security and Microsoft 365 groups — nesting, owners, staleness and dir-sync state |
| External Forwarding | Mailboxes and inbox rules forwarding outside the org |
| Permissions Audit | Full Access, Send As and Send on Behalf delegations |
| Ownerless Groups | Microsoft 365 groups with no owner, or only disabled owners |
| Orphaned OneDrive | OneDrive sites whose owner is disabled or gone |
| External Sharing | Per-site sharing capability and the guests who have access |
| M365 Portal Reports | Launcher for the portal reports Microsoft already computes |
| Report Browser | Every CSV in `reports\`, with row counts. Needs no connection. |

## What this build cannot tell you

Cloud-only means cloud-only, and a few answers are genuinely unavailable rather
than just relocated:

- **`targetAddress`** has no cloud equivalent. In a cloud-only tenant the mailbox
  *is* the destination, so there is nothing for it to point at. Mail Attributes
  omits it rather than showing a permanently blank row.
- **Renaming the directory object.** `Set-User -Name` renames the mailbox, which
  is a different operation from an on-premises RDN change. It is shown read-only
  rather than quietly repurposing the old edit path.
- **`mail` / `WindowsEmailAddress` editing.** `Set-Mailbox -WindowsEmailAddress`
  does nothing while the mailbox is subject to an email address policy, which is
  the normal state. An edit button that silently no-ops is worse than none.
- **`Description`, `Division`, `EmployeeID`, `EmployeeNumber`, `Organization`.**
  No `Set-User` parameter exists, so All Attributes shows them read-only.
- **Reverse group membership.** Exchange Online has no "which groups is this
  recipient in" cmdlet, so DL Memberships enumerates every group and tests each
  one. It is a scan, not a lookup, and says so — with a Cancel button that stops
  between groups so partial results are always true.
- **A directory-synced object.** Anything with `IsDirSynced` True is mastered
  somewhere else and Exchange Online will refuse most changes to it. Group Editor
  and DL Inventory both call this out; in a genuinely cloud-only tenant the count
  should be zero.

## One thing worth knowing before you edit

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
