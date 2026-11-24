# Migration -- DLs and shared mailboxes to cloud-only

On-premises objects leave directory sync; the cloud copies are restored (mailboxes)
or recreated (lists) as cloud-only, then restamped from exports taken beforehand.
No Graph. Full reasoning: `ScriptNotes/ARCHITECTURE.md` section 40.

## Before you start

1. Fill in `DL-Input.csv` and `SMB-Input.csv` -- one column, `Name`, one object per row
   (alias, address or display name).
2. Tenant Manager: set **NonSync OU** on the directory the objects live in
   (canonical form, e.g. `domain.com/Accounts/NoSync`). Step 3 refuses to run until it is set.
3. Optional: `$DefaultPassword` at the top of `4-CLOUD-Restore.ps1`. Blank = random per
   mailbox, never shown, which is fine for a shared mailbox.

## Run order

| Step | Where | What | Changes anything? |
|---|---|---|---|
| `0-CLOUD-Preflight` | cloud | what a restore cannot bring back: security groups, transport rules, nesting, licence-needing mailboxes | no |
| `DL-1-ONPREM-Export` | on-prem | every DL attribute, owners, members, Send As | no |
| `DL-2-CLOUD-Export` | cloud | the same, cloud side (second LegacyExchangeDN, cloud-only members) | no |
| `1-ONPREM-Export-SMB` | on-prem | AD account, mail attributes, **MemberOf** | no |
| `2-CLOUD-Export-SMB` | cloud | mailbox attributes + a Perms file (FullAccess / SendAs / SendOnBehalf / Inbox / Calendar) | no |
| **`3-ONPREM-MoveToNonSyncOU`** | on-prem | moves the objects out of sync scope | **yes -- starts the clock** |
| *wait 75 minutes* | | sync every 30 min; two cycles plus slack | |
| `4-CLOUD-Restore` | cloud | `Undo-SoftDeletedMailbox` (user + mailbox, password set, no prompt); `New-DistributionGroup` for lists | yes |
| `5-CLOUD-Restamp` | cloud | everything from the exports that is missing; idempotent | yes (adds only) |
| `6-CLOUD-Verify` | cloud | cloud objects vs the pre-move exports, item by item | no |
| `7-ONPREM-Cleanup` | on-prem | deletes the parked AD objects; refuses without a clean, newer Verify file | **yes -- irreversible** |
| `3R-ONPREM-MoveBack` | on-prem | undo for 3 | yes |

Run Step 0 and clear every STOP before Step 3. Read the `MemberOf` warnings Steps DL-1 and 1 print. A mailbox that leaves sync scope is dropped from
every list that is still synced, and those lists cannot be edited in the cloud. Either add
those lists to `DL-Input.csv` (and re-run DL-1 and DL-2) or accept the loss -- before Step 3.

Batches over ~500 objects hit the Entra Connect export deletion threshold; the sync refuses
the deletes and Step 4 finds nothing. `Disable-ADSyncExportDeletionThreshold` on the sync
server first, or run smaller batches.

## Shared mailboxes only (the 2026-10-04 batch)

Step 3 reads SMB-Input.csv ONLY (changed 2026-10-04 -- lists are a separate project; DL-1/DL-2 and the
list branches of 4-7 are parked, not removed). Run **1 -> 2 -> 3 -> wait -> 4 -> 5 -> 6**, then 7 some days
later. Steps 0, DL-1 and DL-2 are for lists and can be skipped. The on-premises steps (1, 3, 3R, 7)
open their own window connected to on-premises Exchange; the cloud steps (2, 4, 5, 6) run in
the console and need Exchange Online connected there.

2026-10-04 changes, before the first real run:
- Steps 1, 3, 3R and 7 now widen the Exchange recipient scope to the whole forest and pass
  `-Server` / `-DomainController` for the object's OWN domain (`Get-MigrationAdServer` in
  `_Shared.ps1`). Without that, a child-domain object answered "not found" from a jump box in
  the parent domain. Step 3 writes the controller it used to a `Server` column.
- Step 1 warns at 500+ rows (Entra Connect export deletion threshold).
- Step 4 can show EVERY soft-deleted mailbox in the tenant (tick box), with a `Deleted`
  column; click any header to sort. Rows outside the batch are listed unticked and can be
  restored, but Step 5 has nothing to restamp them from.

## Files

```
DL-Input.csv, SMB-Input.csv                input (ship empty)
Migration-DL-ONPREM_<ts>.csv               step DL-1
Migration-DL-CLOUD_<ts>.csv                step DL-2
Migration-SMB-ONPREM_<ts>.csv              step 1
Migration-SMB-CLOUD_<ts>.csv               step 2
Migration-SMB-CLOUD-Perms_<ts>.csv         step 2
Migration-Moved_<ts>.csv                   step 3  -- steps 4, 5 and 3R read the newest one
Migration-Restored_<ts>.csv                step 4
Migration-Restamp_<ts>.csv                 step 5  -- present / added / set / skipped / FAILED per item
Migration-Preflight_<ts>.csv               step 0  -- STOP / WARN / INFO per finding
Migration-Verify_<ts>.csv                  step 6  -- match / MISSING / DIFFERENT / skipped per item; step 7's gate
Migration-Cleaned_<ts>.csv                 step 7
Exports-Complete\                          move a finished batch here; empty before shipping
_Shared.ps1                                helpers; not a tool
Inventory-AllGroups.ps1                    the forest-wide group inventory this project started from
```

Every step is safe to re-run: exports overwrite nothing (new timestamp), 5 skips objects already
in the target OU, 6 skips objects already restored, 7 only adds what is missing.
