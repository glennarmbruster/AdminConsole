# SMBMigration -- shared mailboxes from on-premises to cloud-only (FH)

One batch at a time. Step 1 mints the batch id `M-d-NNN` (10 October, batch 1 = `10-10-001`)
and the input file `SMBMigration-10-10-001.csv`; every later step asks which batch and writes
`SMBMigration-<batch>-<Stage>.csv` beside it. Logs go to the toolkit's usual CompletedTasks /
Failed logs. Nothing is picked up by "newest file".

Reasoning: `ScriptNotes/ARCHITECTURE.md` section 119. Built 2026-10-10 from the Express
`Projects\Migration` scripts (sections 110-114), which were proven on a live batch.

## Before the first batch

- Tenant Manager: **Mail contact OU** = the non-syncing contacts OU (canonical form), and the
  onmicrosoft domain filled in. Step 5 refuses without them.
- Batches under the Entra Connect export deletion threshold (500 by default, other deletions
  in the same cycle count), or have the AD admins disable it for the sync after Step 4.
- Deleted users in the admin center is the restore (Step 6). Nothing here needs Graph.

## Run order

| Step | Where | What | Changes anything? |
|---|---|---|---|
| `1-ONPREM-Rename` | on-prem | asks the batch number, creates and opens the batch CSV, then prefixes each DisplayName `SMB-<M-d>-` | yes (display names) |
| *wait ~30 min* | | one sync, so the tag reaches the cloud | |
| `2-ONPREM-Export` | on-prem | AD account, mail attributes, on-prem LegacyExchangeDN, Send As, sender restrictions, custom attributes; DisplayName clean + tagged | no |
| `3-CLOUD-Export` | cloud | mailbox attributes + Perms file (FullAccess / SendAs / SendOnBehalf / Inbox / Calendar); says READY when the tag is in the cloud on every mailbox | no |
| **`4-ONPREM-Delete`** | on-prem | `Remove-RemoteMailbox`: deletes the AD user object. Soft gates: exports exist, tag synced, objects resolve, contact OU set. **No undo.** | **yes** |
| `5-ONPREM-Contacts` | on-prem | mail contact per mailbox in the contact OU: original Name/Alias/DisplayName, primary + external `<alias>@<tenant>.onmicrosoft.com`, every old address and `<alias>@<tenant>.mail.onmicrosoft.com` as proxies, X500s, hidden, policy off | yes |
| *wait ~75 min* | | two sync cycles; the cloud users become soft-deleted | |
| **Step 6 -- admin center** | cloud | Users > Deleted users > search `SMB-<M-d>-` > select all > Restore (keep password). Mailboxes reconnect within minutes. | yes |
| `7-CLOUD-Restamp` | cloud | converts back to Shared; re-adds proxies, X500 for both LegacyExchangeDNs, every permission, settings, cloud list memberships. Not DisplayName. Idempotent; re-run until nothing is 'skipped: not active yet' | yes (adds only) |
| `8-CLOUD-Rename` | cloud | takes the tag off DisplayName (clean name from the Step 2 export) | yes (display names) |
| `9-CLOUD-Verify` | cloud | every item vs the exports: match / MISSING / DIFFERENT / skipped | no |

Gates are SOFT everywhere: a dialog says what is wrong and offers Continue anyway.

## What a cloud-only shared mailbox LOSES (hostile review 2026-10-10 -- scripts cannot fix these)

- **Membership of lists synced from on-premises.** The account that held them is deleted; a
  cloud-only mailbox cannot be added to a synced list in the cloud. Step 3 lists them; Step 7
  reports them skipped. Only fix: make the list cloud-only too, or re-add on-premises via the
  contact (Step 5 offers it -- helps on-premises-expanded mail only).
- **Mail bounces between Step 4's sync and Step 6's restore** (~75-90 min). Do it when the
  mailboxes are quiet; listservs unsubscribe addresses that bounce.
- **Licences.** A mailbox with a litigation/in-place hold or an online archive needs one; if it
  came from an on-premises group, it ends with the account. Step 3 and Step 4 warn.
- **A person's mailbox in the list** would have its AD account DELETED. Steps 1 and 4 refuse
  anything that is not a Remote(Shared|Room|Equipment)Mailbox unless overridden.
- **The on-premises contact must route**: `<alias>@<tenant>.onmicrosoft.com` has to be a proxy on
  the restored mailbox. Step 9 checks it as its own item ('routing').

## Files per batch

```
SMBMigration-<batch>.csv                input: PrimarySmtpAddress, one per row (Step 1 creates it)
SMBMigration-<batch>-Renamed.csv        Step 1   old / new display name per mailbox
SMBMigration-<batch>-ONPREM.csv         Step 2   the only on-premises record after Step 4 -- keep it
SMBMigration-<batch>-Groups.csv         Step 2   one row per group membership (security groups too), for cleanup
SMBMigration-<batch>-CLOUD.csv          Step 3
SMBMigration-<batch>-CLOUD-Perms.csv    Step 3   Mailbox, Kind, User, Rights
SMBMigration-<batch>-Deleted.csv        Step 4
SMBMigration-<batch>-Contacts.csv       Step 5
SMBMigration-<batch>-Restamp.csv        Step 7   present / added / set / skipped / FAILED per item
SMBMigration-<batch>-RenamedCloud.csv   Step 8
SMBMigration-<batch>-Verify.csv         Step 9   match / MISSING / DIFFERENT / skipped per item
_Shared.ps1                             helpers; not a tool
```

Steps 1, 2, 4 and 5 open their own window connected to on-premises Exchange; 3, 7, 8 and 9 run
in the console and need Exchange Online connected there.

## 10 - CLOUD Add Contact Address (one-time, 2026-10-13)

After the batches: adds `<alias>@firsthorizon.onmicrosoft.com` (the on-premises contact's external address, from every `-Contacts.csv`; `-ONPREM.csv` Alias as fallback) as a secondary smtp: proxy on each migrated cloud mailbox. All batches in one run; already-present, not-found and still-synced mailboxes are skipped. Writes `SMBMigration-AddContactAddress-<timestamp>.csv`. Needs Exchange Online.
