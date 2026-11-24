# Express Personnel Admin Tooling

PowerShell tooling for Exchange Hybrid and Microsoft 365 administration.

Author: Glenn Armbruster

---

## Before you run anything

**Use Windows PowerShell 5.1.** Not PowerShell 7. The SharePoint and AzureAD
modules do not work the same way under 7, and several tools will fail there.

**Connect first, in the same window.** Exchange Online connections belong to
the PowerShell process that made them. Connecting in one window does not
carry over to another. Use the Connect buttons at the top of the Admin
Console rather than connecting separately.

```powershell
Connect-ExchangeOnline
Connect-SPOService -Url https://<tenant>-admin.sharepoint.com
```

The console shows **ExO** and **SPO** at the top in green when connected, red
when not, alongside the tenant you are connected to. Check that line before
you act — particularly the tenant name.

**On-prem tools need the Exchange Management Shell**, not a plain PowerShell
window, plus the ActiveDirectory module.

---

## Starting up

| I want to… | Run |
|---|---|
| The main on-prem console | `AdminConsole.ps1` |
| The browser-based M365 console | `M365_Host_Web.ps1` |
| ANZ tenant tools | `ANZAdminConsole.ps1` |
| SRG tenant tools | `SRGAdminConsole.ps1` |

The console launches everything else — you shouldn't need to run individual
tool scripts directly. It opens top-left; tools open bottom-left so they don't
cover it.

For the web host: run it, leave the window open, and use the browser tab it
opens. Closing the PowerShell window ends the session. `Ctrl+C` stops it.

---

## Which suite for which tenant

| Suite | Use for |
|---|---|
| `AdminConsole.ps1` and unprefixed scripts | HQ and Franchise (on-prem / hybrid) |
| `ANZ*` | expresspros.com.au |
| `SRG*` | SRG tenant |
| `M365*` | any tenant — whichever one you're connected to |

The `M365*` tools act on **whatever tenant you are currently connected to**.
Confirm the tenant shown at the top before running anything that changes state.

---

## What the tools do

**Look-up / read-only** — safe to run freely.

- `MailboxHealth` — mailbox configuration snapshot
- `MailboxAudit`, `SharedMailboxAudit` — audit settings
- `MessageTrace` — message delivery tracking, with status and timezone filters
- `DLMemberships`, `UnifiedGroupMembership` — group membership lookup
- `MailAttributes`, `AllAttributes` — AD attribute viewer/editor
- `Mailbox-Group-Inspector` — combined mailbox and group view

**Make changes** — these prompt for confirmation. Read the prompt.

- `MailboxPermissions` — Full Access and Send As
- `MailboxOneDriveAccess` — mailbox and OneDrive access in one step
- `CalendarPermissions` — calendar sharing
- `OOFEditor` — out-of-office messages
- `InboxRules` — view, enable, disable, remove inbox rules
- `DL_SendToAs`, `SendToDL`, `SendAsDL` — distribution list delegation
- `GroupEditor`, `DLManager` — group and DL management
- `NewUser` — enables remote mailboxes for accounts that already exist
- `Disable` (ANZ/SRG/M365) — termination workflow

**Utility**

- `ToolManager` — add, remove, and organise console buttons
- `PowerShellConsole` — a PowerShell prompt inside the tool, using its live
  connections
- `ServiceNowQueue` — opens the ticket queue

---

## Things to know

**`NewUser.ps1` does not create accounts.** It mail-enables accounts that
already exist in AD. If an account hasn't been created yet, it will tell you
so and do nothing. Create the account first.

It also finds accounts automatically — you don't need to know whether someone
is HQ or Franchise, and a single batch can mix both.

**Archive shows "provisioning" right after enabling.** That's expected. The
archive is created; Exchange Online just hasn't reported back yet. Nothing to
re-run.

**Don't run the DL tools with an Exchange Online session open.** Several
commands exist in both on-prem and cloud Exchange, and with both loaded the
wrong one can be used. Use a clean Exchange Management Shell window for DL
work.

**Comments marked `GUARD:` are load-bearing.** They sit above confirmation
prompts and scope settings. Don't remove them or the code beneath them.

---

## Logging

Successful changes are appended to:

```
E:\Scripts\M365\Logs\Completed\CompletedTasks.csv
```

Two columns, `Task` and `DateTime`. All tools share this file. Only verified
successes are recorded — skipped and failed operations are shown on screen but
not logged.

---

## When something goes wrong

**"No command named Get-Mailbox"** — not connected in *this* window, or a
previous session is still holding on. Close the window, open a new one, and
connect again.

**Prompted to log in again** — same cause. Connections don't cross processes.

**Web host won't start** — an earlier instance may still hold the port. Close
any stray PowerShell windows and retry.

**A tool is missing from the web console** — its view file isn't in the same
folder as `M365_Host_Web.ps1`. The host skips missing views and starts anyway,
so this fails quietly.

**Wrong tenant** — check the tenant name at the top before acting, not after.

---

## Files

Keep all scripts in one folder. `Common.ps1`, `M365Logic.psm1`, `M365Logic.psd1`,
and the `M365_View_*.ps1` files are shared components — the tools won't run
without them. `Tools.json` holds console button configuration.

Please raise problems with Glenn rather than editing scripts in place.
