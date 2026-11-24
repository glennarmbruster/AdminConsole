# Migration -- DLs and shared mailboxes to cloud-only

On-premises objects leave directory sync; the cloud copies are restored (mailboxes)
or recreated (lists) as cloud-only, then restamped from exports taken beforehand.
No Graph. Full reasoning: `ScriptNotes/ARCHITECTURE.md` section 40.

## Before you start

1. Fill in `DL-Input.csv` and `SMB-Input.csv` -- one column, `Name`, one object per row
   (alias, address or display name).
2. Tenant Manager: set **NonSync OU** on the directory the objects live in
   (canonical form, e.g. `domain.com/Accounts/NoSync`). Step 5 refuses to run until it is set.
3. Optional: `$DefaultPassword` at the top of `6-CLOUD-Restore.ps1`. Blank = random per
   mailbox, never shown, which is fine for a shared mailbox.

## Run order

| Step | Where | What | Changes anything? |
|---|---|---|---|
| `0-CLOUD-Preflight` | cloud | what a restore cannot bring back: security groups, transport rules, nesting, licence-needing mailboxes | no |
| `1-ONPREM-Export-DL` | on-prem | every DL attribute, owners, members, Send As | no |
| `2-CLOUD-Export-DL` | cloud | the same, cloud side (second LegacyExchangeDN, cloud-only members) | no |
| `3-ONPREM-Export-SMB` | on-prem | AD account, mail attributes, **MemberOf** | no |
| `4-CLOUD-Export-SMB` | cloud | mailbox attributes + a Perms file (FullAccess / SendAs / SendOnBehalf / Inbox / Calendar) | no |
| **`5-ONPREM-MoveToNonSyncOU`** | on-prem | moves the objects out of sync scope | **yes -- starts the clock** |
| *wait 75 minutes* | | sync every 30 min; two cycles plus slack | |
| `6-CLOUD-Restore` | cloud | `Undo-SoftDeletedMailbox` (user + mailbox, password set, no prompt); `New-DistributionGroup` for lists | yes |
| `7-CLOUD-Restamp` | cloud | everything from the exports that is missing; idempotent | yes (adds only) |
| `8-CLOUD-Verify` | cloud | cloud objects vs the pre-move exports, item by item | no |
| `9-ONPREM-Cleanup` | on-prem | deletes the parked AD objects; refuses without a clean, newer Verify file | **yes -- irreversible** |
| `5R-ONPREM-MoveBack` | on-prem | undo for 5 | yes |

Run Step 0 and clear every STOP before Step 5. Read the `MemberOf` warnings Steps 1 and 3 print. A mailbox that leaves sync scope is dropped from
every list that is still synced, and those lists cannot be edited in the cloud. Either add
those lists to `DL-Input.csv` (and re-run 1-2) or accept the loss -- before Step 5.

Batches over ~500 objects hit the Entra Connect export deletion threshold; the sync refuses
the deletes and Step 6 finds nothing. `Disable-ADSyncExportDeletionThreshold` on the sync
server first, or run smaller batches.

## Shared mailboxes only (the 2026-10-04 batch)

Leave `DL-Input.csv` empty. Run **3 -> 4 -> 5 -> wait -> 6 -> 7 -> 8**, then 9 some days
later. Steps 0, 1 and 2 are for lists and can be skipped. The on-premises steps (3, 5, 5R, 9)
open their own window connected to on-premises Exchange; the cloud steps (4, 6, 7, 8) run in
the console and need Exchange Online connected there.

2026-10-04 changes, before the first real run:
- Steps 3, 5, 5R and 9 now widen the Exchange recipient scope to the whole forest and pass
  `-Server` / `-DomainController` for the object's OWN domain (`Get-MigrationAdServer` in
  `_Shared.ps1`). Without that, a child-domain object answered "not found" from a jump box in
  the parent domain. Step 5 writes the controller it used to a `Server` column.
- Step 3 warns at 500+ rows (Entra Connect export deletion threshold).
- Step 6 can show EVERY soft-deleted mailbox in the tenant (tick box), with a `Deleted`
  column; click any header to sort. Rows outside the batch are listed unticked and can be
  restored, but Step 7 has nothing to restamp them from.

## Files

```
DL-Input.csv, SMB-Input.csv                input (ship empty)
Migration-DL-ONPREM_<ts>.csv               step 1
Migration-DL-CLOUD_<ts>.csv                step 2
Migration-SMB-ONPREM_<ts>.csv              step 3
Migration-SMB-CLOUD_<ts>.csv               step 4
Migration-SMB-CLOUD-Perms_<ts>.csv         step 4
Migration-Moved_<ts>.csv                   step 5  -- steps 6, 7 and 5R read the newest one
Migration-Restored_<ts>.csv                step 6
Migration-Restamp_<ts>.csv                 step 7  -- present / added / set / skipped / FAILED per item
Migration-Preflight_<ts>.csv               step 0  -- STOP / WARN / INFO per finding
Migration-Verify_<ts>.csv                  step 8  -- match / MISSING / DIFFERENT / skipped per item; step 9's gate
Migration-Cleaned_<ts>.csv                 step 9
Exports-Complete\                          move a finished batch here; empty before shipping
_Shared.ps1                                helpers; not a tool
Inventory-AllGroups.ps1                    the forest-wide group inventory this project started from
```

Every step is safe to re-run: exports overwrite nothing (new timestamp), 5 skips objects already
in the target OU, 6 skips objects already restored, 7 only adds what is missing.
