# Performance and robustness backlog


Found by static scan on 2026-09-12. **None of these has been changed** -- every
one is a per-tool edit that needs a run against a live tenant to confirm, and
the operator is the test environment. Ordered by value.


All line numbers are as of this build.


---

## 1. N+1 calls -- one cmdlet round trip per object inside a loop


The single biggest win. A `foreach` over N objects that calls `Get-Mailbox` or
`Get-MailboxStatistics` per object makes N round trips to Exchange Online, at
roughly 0.5-2 s each. Fetch once with `-ResultSize Unlimited` (or
`Invoke-EXOBulkQuery` / `Invoke-SlicedExoQuery`, both in Common.ps1) and index
the result in a hashtable keyed on the identity.

**Honest sizing:** most of the hits below are resolving a list the OPERATOR typed --
a dozen names, not a directory. Those are fine as they are. The ones worth
changing are where the loop runs over a cmdlet result, and none of the report
tools do that, which is good news. Treat this section as "check the loop
bound", not "rewrite".


| Tool | Line | Call |
|---|---|---|
| `Bulk-DLMembership.ps1` | 147 | `Get-Recipient -Identity $t` |
| `DLtoSharedMB.ps1` | 653 | `Get-Recipient' @{ Identity = $who` |
| `Group-Manager-Editor.ps1` | 184 | `Get-Recipient' -Parameters @{ Identity = $t` |
| `MailboxOneDriveAccess.ps1` | 301 | `Get-MailboxPermission' -Parameters @{ Identity = $m; User = $u` |
| `MailboxOneDriveAccess.ps1` | 389 | `Get-MailboxPermission' -Parameters @{ Identity = $m` |
| `Resource-Delegates.ps1` | 115 | `Get-Recipient' -Parameters @{ Identity = $c` |
| `Resource-Delegates.ps1` | 145 | `Get-Mailbox' -Parameters @{ Identity = $c` |
| `StaleAccounts.ps1` | 75 | `Get-ADUser -Identity $dn` |
| `UserDefinedDelegates.ps1` | 93 | `Get-Mailbox' -Parameters @{ Identity = $cand` |


---

## 2. Bulk reads with no `-ResultSize` -- the silent 1000 cap

`Get-Mailbox`, `Get-Recipient` and friends return **at most 1000 objects by
default and do not warn**. A report that lists 1000 rows in a tenant of 4000
is not a slow report, it is a wrong one. Every bulk read needs `-ResultSize
Unlimited`. (ARCHITECTURE 14b in the reference toolkit covers both traps.)

**Reviewed 2026-09-12: nothing to fix.** The scan flagged four lines; on
reading, all four were a comment, an error-message string, a label caption
and a list of cmdlet names -- not reads. The only `-Filter` reads without
`-ResultSize` (`NewCloudGroup.ps1`, `DLMemberships.ps1`) look up one address
or one member's groups, where the cap cannot bite. Every genuine bulk read
already sets it. Left here so the next person does not re-run the same scan
and re-open the same question.

---

## 3. Waits on the UI thread


`Application.DoEvents()` and `Start-Sleep` inside a form handler freeze the
window or, worse, let a second click re-enter the handler mid-run. The toolkit
has the replacement: `Start-ToolJob` runs a cloud read on a worker with the
status bar showing progress, and `Invoke-ToolStep` names the step. Only
`MessageTrace.ps1` uses `Start-ToolJob` today. Every long-running tool should.


| Tool | Line | Call |
|---|---|---|
| `AuditSMB.ps1` | 429 | `[System.Windows.Forms.Application]::DoEvents()` |
| `AuditUser.ps1` | 360 | `[System.Windows.Forms.Application]::DoEvents()` |
| `Automap-Fix.ps1` | 327 | `[System.Windows.Forms.Application]::DoEvents()` |
| `Automap-Fix.ps1` | 328 | `Start-Sleep -Seconds 1` |
| `Bulk-DLMembership.ps1` | 569 | `$cb = { param([string]$m) $lblStatus.Text = $m; [System.Windows.Forms.Application]::DoEvents() }` |
| `HybridMessageTrace.ps1` | 1989 | `[System.Windows.Forms.Application]::DoEvents()` |
| `Member_BulkADD_CSV.ps1` | 189 | `[System.Windows.Forms.Application]::DoEvents()` |
| `OOFEditor.ps1` | 132 | `[System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 25; $n++` |
| `OOFEditor.ps1` | 487 | `$tabs.SelectedIndex = 1; [System.Windows.Forms.Application]::DoEvents()` |
| `OOFEditor.ps1` | 488 | `$tabs.SelectedIndex = 0; [System.Windows.Forms.Application]::DoEvents()` |
| `Protocols.ps1` | 87 | `Start-Sleep -Seconds $wait` |
| `SMB_Inventory.ps1` | 318 | `[System.Windows.Forms.Application]::DoEvents()` |
| `SMB_Inventory.ps1` | 334 | `[System.Windows.Forms.Application]::DoEvents()` |
| `Search_Purge.ps1` | 249 | `Start-Sleep -Seconds $IntervalSeconds` |
| `Search_Purge.ps1` | 250 | `[System.Windows.Forms.Application]::DoEvents()` |
| `Unified-Trace-EndUser.ps1` | 1016 | `[System.Windows.Forms.Application]::DoEvents()` |
| `Unified-Trace-EndUser.ps1` | 1017 | `Start-Sleep -Milliseconds $per` |
| `Unified-Trace-Mockup.ps1` | 999 | `[System.Windows.Forms.Application]::DoEvents()` |
| `Unified-Trace-Mockup.ps1` | 1000 | `Start-Sleep -Milliseconds $per` |


---

## 4. `Get-Mailbox` where `Get-EXOMailbox` would do


The EXO V3 REST cmdlets (`Get-EXOMailbox`, `Get-EXORecipient`,
`Get-EXOMailboxStatistics`) are several times faster than the remote-PowerShell
originals for bulk reads and do not hold a runspace open. **Not a blind
swap** -- they return fewer properties by default (`-Properties` or
`-PropertySets` must name what you want) and some property names differ.
Worth it for the report tools; not worth it for a single-object lookup.


| Tool | `Get-Mailbox` bulk calls |
|---|---|
| `ResourceDuplicator.ps1` | 2 |
| `Unified-Lookup.ps1` | 2 |
| `DLtoSharedMB.ps1` | 1 |


---

## 5. `$array += $item` in loops


Copies the whole array on every add, so building 5000 rows costs 12.5 million
element copies. Invisible below a few hundred rows; the reason a report
"hangs" above a few thousand. Replace with
`$list = New-Object System.Collections.Generic.List[object]` and
`$list.Add($item)`. Lowest priority -- fix it when touching the file anyway.


| Tool | `+=` sites |
|---|---|
| `ResourceDuplicator.ps1` | 16 |
| `Group-Manager-Editor.ps1` | 13 |
| `Dashboard.ps1` | 10 |
| `HybridMessageTrace.ps1` | 10 |
| `RestoreMailbox.ps1` | 10 |
| `DLtoSharedMB.ps1` | 9 |
| `CalendarReset.ps1` | 8 |
| `HoldRetention.ps1` | 8 |
| `InboxRules.ps1` | 8 |
| `MailboxPermDupe.ps1` | 8 |
| `AvePoint-Dashboard.ps1` | 7 |
| `UnifiedGroup.ps1` | 7 |
| `UnifiedGroupMembership.ps1` | 7 |
| `MailAttributes.ps1` | 6 |
| `NewCloudGroup.ps1` | 6 |
| `RecoverDeletedItems.ps1` | 6 |
| `Search_Purge.ps1` | 6 |


---

## 6. Things to fix while in each file


- `Write-Host` in a tool writes to a console window nobody can see when the
  tool is launched from the console. Route it to the status bar or the output
  box. `Diagnose-AdGate.ps1`, `Diagnose-ExoResolve.ps1` and `new-OOF.ps1` are
  the main offenders and were written to be run from a prompt.
- `ObjectTypes` is empty for every tool ported from the earlier build, so their
  buttons never grey out on object type. Fill it in per tool in `Tools.json`.
- `Invoke-BulkExoQuery` / `Invoke-EXOBulkQuery` and `Test-IppsConnectionState` /
  `Test-SccConnectionState` are duplicate pairs. Pick one of each and repoint --
  with a test run, not blind.
