# Search and Purge -- parked notes

Parked 2026-08-10. Blocked on the Purview **Search And Purge** role, which needs
PIM approval. Nothing here has ever run successfully end to end.

Files: `SearchAndPurge.ps1`, `Connect-SCC.ps1`. Both were built against
`AdminConsole\Common.ps1` and expect to sit in that folder when revived.

---

## Where it stopped

`Connect-SCC.ps1` completed a sign-in, then reported:

    'Get-ComplianceSearch' resolves only to an on-premises Exchange session
    in this process.

**Probable cause, unconfirmed: no compliance roles.** A Security & Compliance
session imports cmdlets according to the roles you hold. With none assigned the
session connects and imports essentially nothing, so no loaded module exports
`New-ComplianceSearchAction` and the resolver correctly finds only the
on-premises copy. If that is what happened, the code is right and the message is
accurate.

Run this immediately after connecting, once the role is held, before changing
any code:

```powershell
Get-ConnectionInformation | Select-Object ModuleName, ConnectionUri, TokenStatus
Get-Command Get-ComplianceSearch -All | Select-Object Name, ModuleName, Source
Get-Module | Where-Object { $_.ExportedCommands.Keys -contains 'New-ComplianceSearchAction' } | Select-Object Name
```

- Line 3 returns a module -> the resolver works; the failure was roles.
- Line 3 empty but line 2 shows a compliance module -> `Get-SccModule`'s export
  marker is wrong and needs widening.
- Line 1 shows no compliance connection -> the connect itself failed.

---

## Verified against Microsoft Learn, 2026-08-10

Re-verify before trusting any of this. Microsoft moved this surface twice in the
eighteen months before these notes were written.

**Still supported.** Classic eDiscovery retired 2025-08-31, but
`New-ComplianceSearch`, `Start-ComplianceSearch`, `Get-ComplianceSearch` and
`New-ComplianceSearchAction -Purge` survived it.

**Retired.** `-Export`, `-Report`, `-RetentionReport` (May 2025). `-Preview` is
flagged on-premises only. Consequence: **there is no way to see the actual
messages before deleting them from PowerShell** -- only per-mailbox hit counts.

**Prerequisites for purge.**

- ExchangeOnlineManagement **3.9.0 or later**
- `Connect-IPPSSession -EnableSearchOnlySession`
- The **Search And Purge** role

**Permissions trap worth the whole page.** Organization Management exists in
both Exchange Online and Purview and they are different role groups. Exchange
Online membership grants nothing here. When the role is missing the error is:

    A parameter cannot be found that matches parameter name 'Purge'

which reads as a broken script or a stale module, not as a permissions problem.

**Service limits.** 10 items per mailbox per purge action. 50,000 mailboxes per
search, enforced even when far fewer match. Unindexed items are never purged.
Purge behaves differently on mailboxes under legal hold. Recently delivered mail
may not be indexed yet, so a message that definitely arrived can return nothing.

**`-WhatIf` is accepted and does nothing** in Security & Compliance PowerShell.
It must not be used as a safety net. `-Confirm:$false` IS required -- the cmdlet
has a built-in pause and a WinForms process has no console to answer it in.

---

## Design decisions, and why

**Not OwnProcess.** An earlier version declared it. On-premises Exchange exports
its own `Get-ComplianceSearch`, so the bare name can resolve to the wrong
service in a console that has connected EMS -- and it fails as a false negative:
searches on-premises, finds nothing, reports "nothing found" while the message
sits in the cloud. Process isolation was the wrong fix. It cost the console
theme, and the console's launcher then ran `Connect-ExSE.ps1` into the fresh
process, importing the very on-premises session the isolation was avoiding.
`Invoke-SccCommand` pins each call to the compliance module instead.

**Connecting cannot happen inside the tool's window.** An interactive sign-in
needs a moment with no modal form in front of it -- see `Connect-ExO.ps1`, which
documents the ActiveX/synchronization-context reason in full. A Connect button
inside a `ShowDialog()` form draws and then silently does nothing. That is why
`Connect-SCC.ps1` exists as a separate console-level script and opens no window
of its own, and why the tool's button says Re-check rather than Connect.

**Gates.** Purge unlocks only on a completed search with hits AND the word PURGE
typed. HardDelete asks a second time; SoftDelete does not, because one is
undoable by the user and the other is undoable by nobody. A new search clears
prior results first, so a failed search cannot leave stale counts on screen
under a live button.

**The search object is kept**, named with date and operator. It is the only
record in the tenant of who removed what.

## Known defect to fix on revival

`SearchAndPurge.ps1` never calls `Set-ConsoleTheme -Form $form`. Every other tool
in the folder does, as the last thing before `ShowDialog()`. This is why the tool
appeared unthemed, and it was misread at the time as a consequence of running in
its own process -- it is not. Launching detached does not lose the theme, because
each tool themes itself. One missing line does.

---

## Depends on Common.ps1

Four functions were added and are still there, currently unreferenced:
`Get-SccModule`, `Get-SccCommand`, `Invoke-SccCommand`, `Test-SccConnectionState`.

They are additive -- nothing existing was modified, no name collisions -- so they
are harmless where they sit. If Common.ps1 is rebuilt before this is revived,
they are the only thing that needs carrying forward, and `Get-SccModule` is the
one to re-examine first.

The module is identified by what it exports, matching `Get-ExchangeModuleMap`.
The marker is a pair, because on-premises Exchange exports `Get-ComplianceSearch`
too: **exports `New-ComplianceSearchAction`, does not export `Get-RemoteMailbox`.**
That pairing is the untested assumption in all of this.

---

## Not built

A **Purge this message** hand-off from Message Trace was built and then reverted
at the same time as this was parked; Message Trace is untouched. The design was:
pass sender, subject and a UTC date range widened a day each side, launch the
purge tool pre-filled, delete nothing automatically.

It was a hand-off rather than a "delete these results" checkbox on the trace grid
on purpose. Trace reads the transport log; purge runs a fresh query against the
mailbox index. A checkbox would say "delete these 40 rows" and do something else
-- misstating scope on an irreversible action.

---

## Sources

- New-ComplianceSearchAction -- https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-compliancesearchaction
- Connect-IPPSSession -- https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/connect-ippssession
- Get-ComplianceSearch -- https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-compliancesearch
- Search for and delete email messages -- https://learn.microsoft.com/en-us/purview/ediscovery-search-for-and-delete-email-messages
- Retirement of legacy eDiscovery tools -- https://learn.microsoft.com/en-us/purview/ediscovery-legacy-retirement
