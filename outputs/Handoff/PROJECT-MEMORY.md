# Project memory export (from the Claude.ai project "Ex Script Creator", 2026-09-29)

Work-related facts only, carried over so Claude Code has them. Everything here was stated by Glenn.

## Environment

- Two AD forests/domains: parent `expresspersonnel.com` and child `franchises.expresspersonnel.com`.
- Exchange Server Subscription Edition on-prem, hybrid with Microsoft 365; ~7,500 users; two Exchange admins.
- All production mailboxes in Exchange Online; MX to Microsoft; no mail flow through on-prem.
- All mailboxes, DLs, shared and resource mailboxes source-anchored on-prem via directory sync.
- On-prem groups (DLs) are always synced to Exchange Online; Send-As senders are remote (cloud) mailboxes.
- 800+ mail-enabled groups in the franchise domain.
- Franchise DCs: `fran-vm-dc01` .. `fran-vm-dc08.franchises.expresspersonnel.com`; dc03 was the historic default.
- Elevated admin account: `gwarmbrustera`. Admin access activated through PIM.
- Scripts at `E:\Scripts\M365\`; completed-task log `E:\Scripts\M365\Logs\Completed\CompletedTasks.csv`.
- Builds on a Mac; the Windows jump box runs the console (no Outlook/Teams installed there).
- Tenants: FH hybrid, EX hybrid, gcearth (GC) cloud-only test; ANZ and SRGPros cloud-only (older suites).
  Future tenants expected to be cloud-only.
- Graph app registrations: msgraph-teams `08baef7b-7bad-4253-b78f-ed5ccbd567ba` (Team.Create);
  msgraph-mails `d42871a4-a2cf-430d-b0e8-b6f59e917e49` (Mail.ReadWrite);
  msgraph-users `d40b59dd-6082-46ee-a0d8-736536baec7c` (User.ReadWrite.All);
  TenantID `f0e47f34-4b37-4ad9-86a2-433a50720bc3`.
- Colleagues referenced in scripts: Jeff Chancey, Nate Spencer, Daniel Bronco, Levi Dickie; Ihtesham (FH side).

## Preferences for this toolkit

- Unified Lookup is Glenn's favourite tool; the console's Look Up opens it with the text as typed -- protect that entry point.
- Mailbox Editor is the gold standard; Room and Group Editors should look and feel like it.
- ONE console for all tenants (reaffirmed 2026-09-21).
- Consolidate rather than add scripts.
- Never ship Settings.json or any JSON that can overwrite live config; new keys are absent-means-off.
- No BUILD-NOTES.md; durable notes go in ScriptNotes/ARCHITECTURE.md.
- Don't rezip the toolkit; Glenn rezips after testing.
- Undo copies before every change; stage only changed files; scripts only in ship folders.
- Devin (another tool) doesn't test scripts; the operator tests.
- Use the Microsoft Learn MCP server for cmdlet verification automatically.
- Glenn guards his scripting standards document -- share scripts with the team, not the instructions.

## Scripting standards (summary -- the full document is authoritative)

- Triage before code: (a) working script to update [preferred], (b) barebones console first,
  (c) full standards from scratch. One-sentence prompt format:
  `Create a WinForms tool to <task>, default <Franchise/HQ>, <barebones first / working script attached / full standards from scratch>`.
- Multi-environment scripts always support both environments; triage only asks the default.
- Lookup keys: PrimarySmtpAddress for mail objects (domain suffix optional, appended silently);
  sAMAccountName for AD users.
- Layout: inputs stacked vertically; buttons Parse -> Run -> Close; Run auto-parses; red bold
  connection banner first; StringBuilder/.Lines for output.
- Header block: Script, Author, Created, Revised, Purpose, Environment, Lookup key, Notes.
- BEFORE/AFTER state for modify scripts; one CSV log row per target object.
- `-BypassSecurityGroupManagerCheck -Confirm:$false` on DL membership writes; SKIP logic for already-member / not-member.
- Permission tools default to many-to-many inputs.
- UTF-8 BOM + CRLF + pure ASCII; surgical edits over rewrites.

## Hard-won technical lessons

- `Add-ADPermission -ExtendedRights 'Send-As'` needs the hyphen.
- `Set-ADServerSettings -PreferredServer` is session-persistent; a wrong DC pin is worse than no pin.
- `Set-Place` silently ignores several properties on hybrid-synced rooms.
- `Start-Process` doesn't set `$LASTEXITCODE`; use `-PassThru` and `.ExitCode`.
- MSOnline/AzureAD modules are retired; Graph only.
- EXO REST "a server side error has occurred" with `-ErrorAction Stop` can hide a 403.
- EXO session opened before PIM activation 403s on the new role -- reconnect.
- `Get-InboxRule` refuses View-Only Organization Management and Global Reader.
- `MessageCopyForSentAs` set on grant, never cleared on removal.
- Label-based parsing of ticket text beats positional parsing.
- Graph PowerShell: 2.25.0 in use; 2.38.1 caused TypeLoadException and was removed.

## Retirement-project decisions (Exchange hybrid)

- Split-group model for all mail-enabled security groups (on-prem security group stays; cloud DL created).
- SMTP relay for Postfix is direct-to-cloud. Franchise shared-mailbox conversion deferred.
- Mail contact export is baseline only. Dual-maintenance for split groups falls to group owners.
- Open: remote-mailbox management model; loss of Exchange RBAC/audit logging; timeline; unknown on-prem dependencies.
