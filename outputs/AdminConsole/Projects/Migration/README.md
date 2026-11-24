# Migration -- DLs and shared mailboxes to cloud-only

On-premises objects leave directory sync; the cloud copies are restored (mailboxes)
or recreated (lists) as cloud-only, then restamped from exports taken beforehand.
No Graph. Full reasoning: `ScriptNotes/ARCHITECTURE.md` sections 40, 111-114.

## Before you start

1. Fill in `SMB-Input.csv` -- one column, `Name`, one mailbox per row (primary address is
   safest). `DL-Input.csv` is for the parked list project; leave it empty.
2. Tenant Manager: set **NonSync OU** on the directory the objects live in
   (canonical form, e.g. `domain.com/Accounts/NoSync`). Step 3 refuses to run until it is set.
3. Batch size: keep it under the Entra Connect export deletion threshold (500 by default --
   and leavers deleted in the same cycle count). Two batches, or
   `Disable-ADSyncExportDeletionThreshold` on the sync server first.

## Run order (shared mailboxes)

| Step | Where | What | Changes anything? |
|---|---|---|---|
| `1-ONPREM-Export-SMB` | on-prem | AD account, mail attributes, MemberOf; THEN tags every DisplayName `SMB-<M-d-yyyy>-` | yes (display names) |
| *wait ~30 minutes* | | one sync, so the tag reaches the cloud | |
| `2-CLOUD-Export-SMB` | cloud | mailbox attributes (clean name) + a Perms file; says READY when every mailbox shows the tag in the cloud | no |
| **`3-ONPREM-MoveToNonSyncOU`** | on-prem | moves the accounts out of sync scope | **yes -- starts the clock** |
| *wait 75 minutes* | | two sync cycles plus slack | |
| **admin center** | cloud | Users > Deleted users > search the tag > select all > Restore (keep password). Exchange reconnects each mailbox within minutes. | yes |
| `4-CLOUD-Restamp` | cloud | converts back to Shared, puts the clean name back, re-adds anything missing from the exports; idempotent -- re-run until nothing is 'skipped' | yes (adds only) |
| `5-CLOUD-Verify` | cloud | cloud objects vs the pre-move exports, item by item | no |
| `6-ONPREM-Cleanup` | on-prem | deletes the parked AD objects; refuses without a clean, newer Verify file | **yes -- irreversible** |
| `3R-ONPREM-MoveBack` | on-prem | undo for 3 | yes |

The on-premises steps (1, 3, 3R, 6) open their own window connected to on-premises Exchange;
the cloud steps (2, 4, 5) run in the console and need Exchange Online connected there.
`0-CLOUD-Preflight`, `DL-1`, `DL-2` and the list branches of 4-6 belong to the parked list
project and are not part of this run.

Why the tag: without Graph the restore is done in the admin center by hand, and that page
searches display names. The tag turns a batch of hundreds into one search and one Restore.
Only DisplayName is tagged -- alias, addresses and UPN are never touched -- and Step 4 puts the
clean name back from the Step 1 export. The GAL shows the tagged name for the hour or so
between the sync that carries it and the sync that deletes the mailbox.

History: 2026-10-04 first live batch (2 mailboxes) found the domain-scope bug (fixed), the
empty EmailAddresses column (fixed), the "isn't unique" duplicate soft-deleted copy (restore by
GUID), and that Undo-SoftDeletedMailbox cannot restore a mailbox whose user the sync deleted.
2026-10-05: 4-CLOUD-Restore removed (obsolete\Migration-4-CLOUD-Restore_2026-10-05); the restore
is the admin center; tag added to Step 1, gate to Step 2, Shared conversion to Step 4.

## Files

```
DL-Input.csv, SMB-Input.csv                input (ship empty)
Migration-DL-ONPREM_<ts>.csv               step DL-1
Migration-DL-CLOUD_<ts>.csv                step DL-2
Migration-SMB-ONPREM_<ts>.csv              step 1
Migration-SMB-CLOUD_<ts>.csv               step 2
Migration-SMB-CLOUD-Perms_<ts>.csv         step 2
Migration-Moved_<ts>.csv                   step 3  -- steps 4, 5, 6 and 3R read the newest one
Migration-Restamp_<ts>.csv                 step 4  -- present / added / set / skipped / FAILED per item
Migration-Preflight_<ts>.csv               step 0  -- STOP / WARN / INFO per finding
Migration-Verify_<ts>.csv                  step 5  -- match / MISSING / DIFFERENT / skipped per item; step 6's gate
Migration-Cleaned_<ts>.csv                 step 6
Exports-Complete\                          move a finished batch here; empty before shipping
_Shared.ps1                                helpers; not a tool
Inventory-AllGroups.ps1                    the forest-wide group inventory this project started from
```

Every step is safe to re-run: exports overwrite nothing (new timestamp), 5 skips objects already
in the target OU, 6 skips objects already restored, 7 only adds what is missing.
