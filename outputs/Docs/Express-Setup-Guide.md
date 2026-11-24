# Express Admin Toolkit v3.x
## Express Tenant Setup Guide

*Applies to every 3.x release. Setup does not change within a major version.*

You have been given a folder. This gets you from that to a working toolkit.

Follow it in order. Steps 1 to 4 are required; the console will not work
without them. Steps 5 onward make it pleasant to use.

If anything does not match what you see on screen, stop and ask rather than
guessing -- several of these tools write to production.

---

## First-time setup

Everything below is **one-time configuration**. You do it once on this machine
and then it stays done:

- Unblock the files
- Tenant Manager
- Settings
- Theme selection
- Config backup

After that, none of it comes back. Day to day you connect and use the tools.

The steps you **will** repeat are the short ones: connecting at the start of a
session, and taking a config backup before you copy over a newer version.

---

## Before you start

**You need:**

| | |
|---|---|
| Windows PowerShell | 5.1 (the blue one, not PowerShell 7) |
| ExchangeOnlineManagement module | `Install-Module ExchangeOnlineManagement` |
| Microsoft.Online.SharePoint.PowerShell | for the SharePoint tools |
| RSAT: Active Directory module | for the on-premises tools |
| An admin account | with Exchange and AD rights for this tenant |

**Check PowerShell first.** Open Windows PowerShell and run:

    $PSVersionTable.PSVersion

You want major version 5. If it says 7, you have the wrong shell open -- search
for "Windows PowerShell", not "PowerShell 7".

---

## Step 1 -- Put the folder somewhere sensible

Somewhere local, with no spaces in the path if you can help it, that you can
write to. For example:

    E:\Scripts\M365\AdminConsole

Do **not** run it from a network share or from a Downloads folder.

---

## Step 2 -- Unblock the files

**This one catches everybody.** Windows marks files that came from another
machine -- email, a share, a zip -- and PowerShell refuses to run them. The
error looks like a security warning about running scripts, and it appears at
the worst possible moment.

`UnBlock.ps1` ships with the toolkit and clears the whole folder. It is itself
blocked, so unblock that one file by hand first:

1. Right-click **`UnBlock.ps1`** -> **Properties**
2. Tick **Unblock** at the bottom -> **OK**
3. Open Windows PowerShell in the console folder and run:

       .\UnBlock.ps1

That clears every other file in one go. Do this **before** anything else --
until it is done, nothing in the folder will run.

---

## Step 3 -- First launch

    .\AdminConsole.ps1

**If you get a message saying there is no `Config\Tenant.json`**, that is
expected on a brand-new copy. Say **Yes** to opening Tenant Manager and go to
Step 4.

**If the console opens**, the tenant file was shipped with your copy. Go to
Step 4 anyway and *verify* the values -- you are checking, not typing.

> **Why it matters.** `Tenant.json` tells the console what this tenant has.
> Until it exists, every capability reads as OFF and the tools that need one
> are hidden -- so the console looks half empty and nothing explains why.

---

## Step 4 -- Tenant Manager (required)

This is the file that defines the tenant. Open **Tenant Manager** and check
every value below.

### Identity

| Field | Express value | What it does |
|---|---|---|
| Tenant name | `Express` | Shown in the console header. Your "am I in the right tenant" check. |
| Primary domain | `expresspros.com` | The default mail domain. Every "@domain optional" prompt uses it, and tools build addresses from it. |
| onmicrosoft domain | `expresspros.onmicrosoft.com` | Used to work out routing addresses for new mailboxes. |
| Theme | `Express` | Optional. Colour scheme from Themes.json. |

### Capabilities -- what this tenant HAS

For Express, **tick all six**:

    [x] ExchangeOnline      [x] SharePointOnline    [x] OnPremAD
    [x] OnPremExchange      [x] Pim                 [x] PimTimer

> **This is a claim, not a connection.** Ticking Exchange Online says the
> tenant *has* Exchange Online -- not that you are connected right now.
>
> Unticking something makes its tools **disappear** rather than grey out. That
> is deliberate: a greyed button invites you to wonder what is wrong with it,
> and a tool for a service the tenant does not own is just clutter.
>
> Express is hybrid, so it has everything. A cloud-only tenant would untick
> OnPremAD and OnPremExchange, and the on-premises half of the console would
> simply not appear.

### Directories -- one row per AD directory

Express has two:

| Id | Label | Domain | DCs (semicolon separated) | Default |
|---|---|---|---|---|
| `HQ` | `HQ` | `expresspersonnel.com` | `expresspersonnel.com` | |
| `Franchise` | `Franchise` | `franchises.expresspersonnel.com` | `fran-vm-dc01.franchises.expresspersonnel.com; fran-vm-dc02...` through `dc08` | [x] |

Two things worth understanding:

**Why HQ lists a domain and Franchise lists servers.** Where a directory has
one entry and it is the domain name, AD picks a controller for you -- which is
what you want. Franchise names its eight controllers explicitly because that
environment needs a specific one for some operations.

**Why "Default" is on Franchise.** Most day-to-day work at Express is franchise
work, so tools that need a directory and were not told which one start there.

**Ids are not cosmetic.** `HQ` and `Franchise` are referenced by the licence
groups below. Rename one here and you must rename it there.

### Servers

| Field | Express value | Notes |
|---|---|---|
| Exchange server URI | `exch-vm-hybrid1.expresspersonnel.com` | On-premises only. Connect ExSE uses this. |
| SPO admin URL | `https://expresspros-admin.sharepoint.com` | Blank = worked out from the tenant. Filled in here because it is known. |
| MySite host | `https://expresspros-my.sharepoint.com` | Where OneDrive personal sites live. The OneDrive tools build URLs from it. |

### Licence groups

Referenced by the disable tools. **The name must match the AD group exactly** --
a typo means the tool silently reports "not a member" and the person keeps
their licence.

| Group name | Sku | Directory |
|---|---|---|
| `E1 License HQ` | E1 | HQ |
| `E1 License Fran` | E1 | Franchise |
| `E5 License HQ - Full` | E5 | HQ |
| `E5 License Fran - Full` | E5 | Franchise |

### Finish

Click **Validate** -- it reports anything inconsistent, and says what the
consequence would be rather than just naming a rule.

Then **Save**. Windows already open keep their old values until restarted.

> **Detect** fills in what it can from the connections the window already has.
> It only fills fields that are **empty** and never overwrites something you
> typed.

---

## Step 5 -- Tool Manager: Settings (required)

Open **Tool Manager**, then the **Settings** button. These are *your* settings
on *this machine* -- not the tenant's. Everyone who uses this console fills
these in for themselves.

### Folders

| Field | What to put | Why |
|---|---|---|
| **Reports** | `E:\Scripts\M365\AdminConsole\reports` | Where report CSVs are written. Leave blank to use a folder beside the toolkit. |
| **Logs** | `E:\Scripts\M365\AdminConsole\Logs` | Where `CompletedTasks.csv` goes -- the audit trail of what the tools did. Each row records the operator, the tenant, the object, and whether it worked, so logs from several toolkits can be put side by side. Read it with **Log Viewer**. |
| **Retry queue** | *(leave blank)* | Work a tool could not finish and will retry. Blank uses `RetryQueue` beside the toolkit and is created when needed. A path you type here is **never** created automatically, because an empty folder would look like an empty queue. |

> Never hardcode a drive letter anywhere else. `E:\` exists on your machine and
> may not on the next one -- these settings are the reason the scripts do not
> need to know.

### Your email

**Put your own work address here** -- for Express, your `@expresspros.com`
address.

When the disable tool emails a manager about a delegated mailbox, it copies
you. Blank falls back to the mail address on your AD account; if that is empty
too, no copy is sent.

### SPO admin URL / Exchange server

Both can stay blank here -- they come from `Tenant.json`. The boxes exist so
one machine can override the tenant value without changing it for everyone.
The hint under each tells you what will be used if you leave it blank.

### Edge profiles

Names of your Edge profiles, used by tools that open a browser so that links
open signed in as the right account.

Use the profile **name** as shown in Edge, not "Profile 2". If your Edge has
profiles called `Admin`, `GWA` and `ServiceNow`, type those.

---

## Step 6 -- Tool Manager: the tool list

The grid is what the console shows as buttons.

| Column | Meaning |
|---|---|
| **Name** | Button caption |
| **Script** | The .ps1 it runs |
| **ObjectTypes** | Which looked-up object the tool applies to: `Mailbox`, `UnifiedGroup`, `OnpremGroup`. Comma-separated. |
| **Section** | Which panel it appears in |
| **Group** | Optional. Same word on several rows collapses them behind one button with a chooser. See below. |
| **Tier** | Free text for your own notes. The console ignores it. |
| **Own window** | Runs the tool in its own PowerShell process. **Required for on-premises Exchange tools.** See below. |

**Sections...** edits the panels themselves -- their names, their order, how
many rows each one gets, and how many buttons wide every panel is.

Button *width* is not a setting. At launch the console measures every tool
caption and makes all buttons the narrowest width that still fits the longest
one, wrapping to two lines where it needs to. That is usually well under the
old fixed width, and the space it reclaims is what pays for an extra column.
The dialog previews the resulting window width as you change the number, and
warns you if it would be wider than your screen.

### Own window -- required for on-premises Exchange tools

**Any tool that talks to on-premises Exchange must have this ticked.** This is
not a preference. Getting it wrong does not produce an error -- it produces a
change in the wrong directory.

**Why.** On-premises Exchange and Exchange Online export the *same cmdlet
names*: `Get-Mailbox`, `Set-Mailbox`, `Get-DistributionGroup` and dozens more.
Load both into one PowerShell process and **whichever was imported last wins
the bare name**. The console is normally connected to Exchange Online, so an
on-premises tool running inside it can silently resolve to the cloud instead --
or the reverse.

What that looks like in practice:

- You run a DL inventory expecting on-premises groups and get the cloud copies.
- A change you make on-premises appears to succeed and lands in Exchange Online.
- Nothing errors. Nothing warns you. The output looks entirely normal.

**Own window prevents it** by giving the tool its own PowerShell process, which
connects to on-premises Exchange itself and never sees the console's cloud
session. One process, one Exchange, no ambiguity.

**The trade-offs, so nothing surprises you:**

- **Slower to start.** That window has to import the Exchange session -- expect
  a wait before the tool appears. It has not hung.
- **Never greyed out.** The console cannot know what that separate window is
  connected to, so it never disables those buttons. A green button here means
  "this will open", not "this will work".
- **Separate connections.** Signing in there does not connect the console, and
  vice versa.

**How to tell if a tool needs it.** If it uses on-premises Exchange cmdlets --
anything under Connect ExSE, the `AD-` tools that also touch Exchange, the DL
decommission scripts -- tick it. If in doubt, tick it: the cost is a slower
start, and the cost of being wrong is a write to the wrong directory.

Some scripts declare this requirement in their own header, and that declaration
wins over the checkbox -- so a tool built to need its own process gets one even
if the box is unticked.

### Group -- several tools behind one button

Type the same word in **Group** on several rows and they collapse into one
button that opens a chooser naming them. `Rooms` on the room tools is the
example: five tools, one slot on the front page.

Use it for tools that answer the same question different ways -- the Exchange
Online and Active Directory versions of a tool, for instance. The chooser names
the actual scripts, so you pick the one you meant rather than answering a
question about an abstraction.

> **Removed 2026-08-13: the AD twin column.** A tool used to be able to name an
> on-premises twin and become one button that asked which directory to act on.
> It worked, but the pairing was invisible -- nothing on the console said a
> button had two scripts behind it -- and Move Up / Move Down did not carry the
> twin with its row, so reordering silently reattached twins to the wrong
> tools. Grouping does the same job out loud. See ARCHITECTURE section 30.

---

## Step 7 -- Theme Manager (optional but recommended)

Pick a colour scheme. This sounds cosmetic and is not.

You will have several consoles open against different tenants. **Colour tells
you which one you are looking at faster than reading does**, and the tools in
this console write to production.

Click **Add shipped themes** to pull in the supplied palettes -- it only adds
names you do not already have and never touches your own. Pick one, tick
**Enabled**, **Save**.

Express uses the `Express` theme. If you manage more than one tenant, bind a
distinct colour to each so they are never confused.

---

## Step 8 -- Back up your config

Open **Config Backup** and click **Back up now**.

Do this **now**, and again **before you ever copy a newer console folder over
this one**.

`Config\` holds everything you just did. A folder copy replaces all of it and
warns about nothing -- the loss shows up later as blank paths and missing
buttons. Backups are written *outside* the console folder precisely so that
replacing the folder cannot take them with it.

---

## Step 9 -- Connect, and check it worked

Order matters.

**1. Activate first, if your rights are on-demand.** Click **Activate** in the
Connect panel.

> A session made *before* activation holds the token it was issued with, and
> that token does not contain rights granted afterwards. Everything then fails
> as though the activation never happened. Activate, **then** connect.

**2. Connect what you need.** ExO, SPO, OnPrem. The console closes while a
sign-in window is open and reopens on its own -- that is deliberate, because a
sign-in prompt cannot take typing while the console is in front of it.

**3. Check the header.**

    ExO: Connected    SPO: Connected    Tenant: expresspros.com
    EMS: Connected    AD: Connected     PIM: Groups 3h 53m

If the tenant name is not what you expect, stop. A `(?)` beside it means the
connected domain does not match `Tenant.json` -- not necessarily wrong, since a
tenant can have several accepted domains, but answer it before you write
anything.

**4. Look somebody up.** Type a mailbox address and click **Look Up**. You
should get `Found: Mailbox (HQ)` or `(Franchise)`.

That single action proves the tenant config, the directory setup and the
Exchange connection all work.

---

## Validation checklist

Work down this list before using the toolkit for real. Every line is something
you can see on screen -- no judgement required.

    [ ]  Header shows  Tenant: expresspros.com
    [ ]  ExO: Connected
    [ ]  SPO: Connected
    [ ]  EMS: Connected            (only if you will use the on-premises tools)
    [ ]  AD:  Connected            (only if you will use the on-premises tools)
    [ ]  Dashboard -> Toolkit Health -> Files verified   shows a count, not an error
    [ ]  Dashboard -> Toolkit Health -> Differences      shows none
    [ ]  Version line at the bottom of the console is not amber
    [ ]  A mailbox lookup returns  Found: Mailbox (HQ)  or  (Franchise)
    [ ]  Tester reports no errors you have not already been told about

**All ticked -- the toolkit is ready for production use.**

If any line fails, fix it before running a tool that writes. Each one maps to a
step above:

| Failed line | Go back to |
|---|---|
| Tenant name wrong or blank | Step 4 -- Tenant Manager |
| A service shows Not connected | Step 9 -- Connect |
| Files verified / Differences reports a problem | Run **Tester**; the toolkit differs from its stamped version |
| Version line is amber | Same -- files changed since the last stamp |
| Mailbox lookup fails | Step 9, and check you are connected to the right tenant |

> **Two of these are easy to skim past.** *Files verified* and *Differences*
> tell you whether the code you are running is the code that was shipped. A
> difference is not necessarily wrong -- somebody may have fixed something --
> but you should know about it rather than discover it later.

---

## Using it day to day

**Look up first, then click a tool.** Most tools act on whatever you last
looked up.

**Greyed buttons have three different meanings.** Hover to see which:

- *Wrong object type* -- look up something else and it returns.
- *Not connected* -- connect that service. If your rights are on-demand,
  activate first.
- *Not available in this tenant* -- connecting will not help.

**Connections belong to the window, not to you.** A tool opened in its own
window has its own connections. Closing the console drops them.

---

## When something looks wrong

| Symptom | Cause |
|---|---|
| Scripts refuse to run, security warning | Step 2 -- unblock the files |
| Console opens but most buttons are missing | `Tenant.json` missing or capabilities unticked |
| "Not found in the on-premises directory" for an address you know exists | Not connected to on-prem, or looking in the wrong directory |
| Everything greyed after a lookup | Wrong object type -- clear the box and press Look Up with it empty to release the grey-out |
| Version line at the bottom is amber | Files differ from the stamped version. Run **Tester**. |
| A tool worked yesterday and not today | Your session expired. Reconnect. |

**Run Tester** when in doubt. It checks the whole folder and reports what is
wrong in specific terms.

---

## Do not do these

**Do not copy a new console folder over this one without backing up `Config\`
first.** It replaces `Tools.json`, `Settings.json`, `Sections.json`,
`QuickLaunch.json`, `Themes.json` and `Tenant.json`, silently.

**Do not edit the .json files by hand.** Use the manager tools. JSON needs
backslashes doubled -- `".\\Script.ps1"` -- and one wrong character makes the
file unreadable, which takes out **every** button rather than one.

**Do not hardcode a path, address or server name into a script.** Everything
tenant-specific belongs in `Tenant.json`; everything machine-specific belongs
in Settings. This is what lets one toolkit serve every tenant.

---

## Giving a copy to another admin

`Deploy.ps1` on the master builds a handout copy and zips it. No console
button -- run it from the folder. One build; everyone gets the same one.

    adminconsole3.1_2026-08-13_1103.zip

Their logs, reports, retry queue and **your** `Settings.json` are not in it.
Everything else in `Config\` is tenant fact and does travel, so they are not
retyping the tenant setup. `Version.json` is rebuilt for the copy, so their
self-test reads clean rather than reporting the deliberately-excluded files as
missing forever.

A `README.txt` inside the zip walks them through it. Two paths:

**First copy.** Extract somewhere permanent, rename the dated folder to
whatever they like, run `Unblock.ps1` FIRST, then open the console and fill in
Dashboard -> Tool Manager -> Settings.

**Upgrading an existing copy.** Back up config in the current console with
`ConfigBackup.ps1`, rename the old folder aside, extract the new build and
**rename it to exactly the old folder's name**, unblock, then restore the
backup.

> The rename is not cosmetic. Backups are filed under the console folder's
> NAME, outside the folder itself -- so replacing the folder cannot destroy
> them, but a console in a differently-named folder looks somewhere else and
> shows no backups at all.

Deploy is excluded from the copies it makes, as is `Publish-Toolkit.ps1`. A
deployed console cannot re-stamp itself or deploy onward, so every folder
traces back to one master. The trade: you cannot hand-patch a deployed copy
without leaving its self-test permanently orange -- send a fresh build instead.

Full detail in ARCHITECTURE section 31.

---

## Quick reference -- Express

    Tenant name         Express
    Primary domain      expresspros.com
    onmicrosoft         expresspros.onmicrosoft.com
    Capabilities        all six ticked
    Directories         HQ (expresspersonnel.com)
                        Franchise (franchises.expresspersonnel.com) -- default
    Exchange server     exch-vm-hybrid1.expresspersonnel.com
    SPO admin           https://expresspros-admin.sharepoint.com
    MySite host         https://expresspros-my.sharepoint.com
    Licence groups      E1/E5 x HQ/Franchise

**Every tool has a `?` button** explaining what it does and what it touches.
Use it before running anything unfamiliar -- several of these write to
production.
