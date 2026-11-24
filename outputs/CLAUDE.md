# CLAUDE.md -- AdminConsole toolkit workspace (Glenn Armbruster)

Read this first, every session. Then read `Handoff/PROJECT-MEMORY.md`, and
`ScriptNotes/ARCHITECTURE.md` (sections 99-108 are the most recent work).
The authoritative scripting standard is `Handoff/PowerShell-WinForms-Scripting-Standards.md`
(Glenn's own document -- apply every section, finish against its Quick Reference Checklist).

This workspace was moved here from Claude Cowork on 2026-09-29. Nothing in the
toolkit depends on Cowork; every path below is relative to this folder.

---

## What this folder is

A PowerShell 5.1 WinForms admin toolkit for Microsoft 365 / Exchange hybrid,
shipped to THREE tenants from one code base:

| Tenant | Shape | Ship folder | File naming |
|---|---|---|---|
| EX (Express) | hybrid, slower drift to cloud | `shipEX/` | `*.ps1` unchanged |
| FH | hybrid, drifting to cloud; sent by EMAIL | `shipFH/` | `*.ps1.txt` (append .txt) |
| GC (gcearth) | cloud-only test tenant (no AD, no on-prem Exchange) | `shipGC/` | `*.ps1` unchanged |

Glenn builds on a Mac, copies ship folders to a Windows jump box
(`E:\Scripts\M365\` / `X:\...`), and runs the console there. He EMPTIES the ship
folders after shipping, so they are outboxes, not archives.

## Folder layout

- `AdminConsole/` -- THE MASTER COPY of every shared script (`Common.ps1`, `AdminConsole.ps1`, all tools).
  - `AdminConsole/shipEX|shipFH|shipGC/` -- tenant-ONLY scripts (archive; e.g. Disable-User is EX-only).
  - `AdminConsole/DisabledTools/` -- parked/superseded tools. `AdminConsole/Config/` -- NEVER shipped.
- `AdminConsoleFH/` -- the FH-only scripts (not a mirror).
- `shipEX/ shipFH/ shipGC/` -- outboxes for the next shipment.
- `ScriptNotes/ARCHITECTURE.md` -- design log. Every non-trivial change gets a section or an addendum.
- `undo/<date>-<what>/` -- pre-change copies of every file touched.
- `obsolete/`, `DL_MigrationPlan/`, `DevinPack/`, `MasterConsole/`, `TenantConfigs/` -- reference; don't ship.

## Standing rules (non-negotiable)

1. **Copy rule.** A change to a SHARED script: edit `AdminConsole/<file>`, then copy to
   `shipEX/<file>`, `shipGC/<file>`, `shipFH/<file>.txt`. A TENANT-ONLY script:
   `AdminConsole/shipXX/<file>` + `shipXX/<file>` only. Always top up from `AdminConsole/`,
   never from another ship folder. Verify all copies by SHA-256 after staging.
2. **Stage only files whose content changed.** A fresh timestamp must mean a change.
3. **Ship folders hold scripts only.** No docs, no README, no ARCHITECTURE, no subfolders.
4. **Never ship `Config/*.json`.** Only carve-outs: `ParkList.json`, `Additions-*.json`
   (content authored with the build, no local state on the receiving install).
5. **Banned words** in any .ps1 or shipped file (code AND comments), whole-word, case-insensitive:
   `firsthorizon`, `fran`, `franchise`, `HQ`, `HQTS`, `express`, `expresspros`, `expresspersonnel`,
   `SRG`, `ANZ`. Use `contoso.com` / `abcd` / `wxyz` placeholders.
   Run the scan on EVERY file before staging it -- it is Claude's gate, not Glenn's:
   `grep -o -i -E 'firsthorizon|franchise|\bfran\b|expresspros|expresspersonnel|\bexpress\b|\bSRG\b|\bANZ\b|\bHQTS\b|\bHQ\b' <files>`
   (Pre-existing hits in the EX-only scripts Disable-User, FranUser_Office, NotDeadYet, and in
   NewUser.ps1, are known and functional -- leave them.)
6. **Encoding:** UTF-8 WITH BOM, CRLF line endings, pure ASCII (no em-dashes, no smart quotes).
   Edit with a method that preserves CRLF + BOM (Python: read `utf-8-sig`, write with BOM, `newline=''`).
7. **Undo first.** Before changing or deleting any file in AdminConsole, AdminConsoleFH or
   ScriptNotes, copy the original to `undo/YYYY-MM-DD-<what>/`.
8. **Verify cmdlets against Microsoft Learn** (the Microsoft Learn MCP server) before delivering
   code -- parameter names, accepted values, applicability (on-prem vs Exchange Online).
   Do it without asking. If verification is unavailable, say so before delivering.
9. **Every Exchange/AD write from a WinForms tool passes `-Confirm:$false`** (a ShouldProcess
   prompt has no console and hangs the window).
10. **Proofpoint URL Defense rewrites URLs inside mailed files** (FH gets code by email).
    Anything coming BACK from another install: check for `urldefense` before trusting it.
    `FixCommon.ps1` carries the SHA-256 of `Common.ps1` -- RE-STAMP it (`$script:GoodHash`,
    `$script:GoodBytes`) every time Common.ps1 changes, and ship both together.
11. **Don't "fix" working code unasked.** Glenn: "I don't want you to fix bugs, because this has
    been working for MONTHS." Report what you notice; change it only when asked.
12. **Glenn's style:** terse, direct. Recommend before building anything consequential.
    Tell him what changed in chat, not in a new file.

## Verification gate (every change)

- Bracket/paren/brace deltas vs the undo copy must be balanced (`{`=`}`, `(`=`)`, `[`=`]`).
- **Control-before-handler ordering:** every WinForms control must be CREATED before any line
  that calls a method on it (`.Add_Click`, `.Add_CheckedChanged`, `.Enabled = ...`). This bug
  class (null-valued expression) hit this project four times.
- A multiline TextBox on a form with an `AcceptButton` needs `AcceptsReturn = $true`.
- Layout: no new control rectangles overlap on the same parent.
- Banned-word scan clean; encoding check (BOM, CRLF-only, ASCII).
- If `pwsh` is installed on this Mac (`brew install powershell`), a real parse check is possible:
  `pwsh -NoProfile -Command "[System.Management.Automation.Language.Parser]::ParseFile('<file>',[ref]$null,[ref]$e); $e"`

## Toolkit internals worth knowing

- Exchange calls go through `Invoke-EXOCommand` (cloud) / `Invoke-OnPremCommand` (on-prem),
  which pick the module explicitly via `Get-ExchangeModuleMap`. Bare Exchange cmdlet calls
  collide when both Exchanges are loaded.
- `TOOLKIT-REQUIRES: OwnProcess` tools run in a separate powershell.exe connected to
  ON-PREM Exchange only (Connect-ExSE). The console's Exchange Online session does NOT reach
  them; `Connect-ToolExchangeOnline` (Common.ps1) signs the tool's own process in on demand.
- Capability Ids are `ExO`, `SPO`, `EMS`, `AD`, `Graph` (not labels). `# TOOLKIT-CAPABILITY-OPTIONAL: ExO, SPO`
  lets a tool open without a connection and gate features itself.
- `Save-ToolkitSettings` REBUILDS Settings.json from its parameters -- a key it doesn't know is deleted.
- `$script:HeaderBandH = 50`: any top-level label with Top < 50 is recoloured by the theme band.
- PowerShell `-ne`/`-eq` are case-insensitive; `-cne` is case-sensitive.
- EXO REST cmdlets can fail with "a server side error has occurred" when `-ErrorAction Stop`
  is present; that can mask a real 403. An EXO session opened BEFORE PIM activation 403s --
  reconnect after activating.
- Hybrid Send-As: permission must exist where the SENDER's mailbox is. Cloud mailboxes need
  `Add-RecipientPermission` in Exchange Online on the synced group; AD `Add-ADPermission` only
  matters for on-prem mailboxes. On-prem DLs are always synced; senders are remote mailboxes.
- Domain completion: `Expand-TargetEntry` / `ConvertFrom-DialogList -CompleteDomain` append
  Tenant.json `PrimaryDomain` to bare names, verified by a Get-Recipient probe, falling back to
  the bare word.
- Search-UnifiedAuditLog (Exchange Online) covers SharePoint/OneDrive too -- no SPO connection.

## Open items (backlog)

- Test-Toolkit rules offered but not built: "handler attached before its control exists";
  "multiline box + AcceptButton without AcceptsReturn".
- Disable-User: NotifyOneDrive still sends if GrantOneDrive failed for a non-SPO reason (reported, not changed).
- MailAttributes.ps1 "New value" box has the Enter/AcceptButton trap (reported, not changed).
- ARCHITECTURE section 71's urldefense gate pattern `proofpoint\.com/v[23]` false-positives on a legit API URL; use `urldefense` alone.
- Group Editor: AD refusal (INSUFF_ACCESS_RIGHTS) on some franchise-domain DLs is unexplained; irrelevant while senders are cloud.
- Test-Toolkit.ps1 pulled from the FH drop pending a decision on its leak-detection regex.
- NewUser.ps1 in the AdminConsole root carries tenant markers -- Glenn to decide.
- Older: #183 park-time warning, #185 demo dry run, #298/#303 retire absorbed tools,
  #304 dark-theme colour sweep, #265-270 ServerHealth.
