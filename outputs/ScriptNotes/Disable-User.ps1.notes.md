# Disable-User.ps1

Termination / disable processing for HQ and Franchise in one tool. Built
2026-07-31, replacing `Disable-Intake.ps1`, `Disable-New.ps1` and
`Disable-New-HQ.ps1` (now in `obsolete/Superseded/`).

Cross-cutting context lives in `ARCHITECTURE.md` §21. This file is the detail.

---

## Why the three scripts became one

The two old runners differed by a handful of steps and were otherwise the same
code. That is duplication, not isolation — and it had already drifted: **HQ had
silently lost the forwarding-rule step that Franchise still had**, and nothing
anywhere made the difference visible. Nobody noticed because there was no place
where the two step lists could be compared.

Hence `$script:StepPlan`: two lists side by side, so a step cannot go missing
without somebody deleting it from a list they are looking at.

```
Franchise : ConvertShared, HideFromGAL, GrantDelegate, RemoveLicence,
            GrantOneDrive, NotifyOneDrive, DisableForwarding, ValidateLicence
HQ        : ConvertShared, HideFromGAL, GrantDelegate,
            GrantOneDrive, NotifyOneDrive, DisableForwarding
```

The old design also used `disablelist.csv` as a **message bus**: the intake
wrote it, then immediately invoked a script that read it back. One script means
records stay in memory, and the CSV reverts to what it is good at — optional
bulk input. It also had a live bug: the Franchise rows were written and run,
then the same file was **overwritten** with the HQ rows, so afterwards you could
not tell what the first run had been given.

## Deliberate HQ / Franchise differences

Confirmed with the team on 2026-07-31. These are policy, not drift, and should
not be "fixed":

- **`RemoveLicence` / `ValidateLicence` are Franchise-only.** HQ licences live
  in a different group that this team has no rights to modify.
- **The OneDrive notification body differs.** Franchise recipients are outside
  the IT organisation and get the full explanation — what happened, the 30-day
  window, what to do about it. HQ managers get the link only. See
  `Get-OneDriveNotificationBody`.
- **`DisableForwarding` now runs for both.** This was the one difference that
  *was* an oversight; HQ gained it on 2026-07-31. A terminated HQ user
  forwarding to a personal address is the same risk as a franchise one.

## The three rules the execution engine follows

### 1. Per user, not per step

The old scripts ran six separate `foreach ($row in $csv)` passes over the whole
batch. A failure part-way therefore left **every** user half-processed — all
converted to shared, none de-licensed. Here a user is finished before the next
one starts, so a failure leaves earlier users complete and later ones untouched,
and the resulting state is describable in a sentence.

### 2. Skip and report

One user failing must not stop the rest. The old fail-fast `catch` did
`Stop-Transcript; exit` — and the `CompletedTasks.csv` block sat *below* that.
**Every row in the old audit log was a success by construction**, not because
failures were rare but because the failure path exited before it could write.

### 3. Check then act

Every step asks the environment whether it is already done: mailbox type, the
GAL flag, a permission, group membership, a site collection admin, an inbox
rule. All queryable.

That makes retry correct regardless of what happened in between — **including a
manual fix**, which is the normal recovery path here. Stored per-step state
would have been a lie the moment somebody enabled a mailbox by hand.

`Skipped` is reported rather than folded into success, because "already shared"
and "converted just now" are different facts and the difference matters on a
retry.

**The exception is sending mail.** There is no `Get-DidIAlreadySendThis`, so
notification completion is recorded in the retry queue and honoured on retry.
That is the entire reason `GrantOneDrive` and `NotifyOneDrive` are separate
steps: the grant is verifiable and should be re-checked every time, the email is
not and must never repeat. A skipped notification is **announced** rather than
silent — if the mail genuinely never arrived, that line is what tells the
operator to send one by hand.

## Failure diagnosis is the point, not the error text

The dominant failure is not an Exchange problem. Onboarding is supposed to raise
a ticket to enable the remote mailbox and sometimes never does, so a user can
sit for a year with an AD account and no mailbox. The second cause is a typo in
the ticket. Those need opposite fixes, and a raw cmdlet error distinguishes them
for nobody.

`Resolve-DisableTarget` therefore reports which one it is and what to do:

| Condition | Reported as |
|---|---|
| No AD account | *"Most likely a typo in the ticket. Verify the identity and correct it, or remove the row."* |
| AD account, no remote mailbox | *"Onboarding almost certainly never raised the ticket to enable it. Enable the remote mailbox, wait for directory sync, then retry this row."* |
| Remote mailbox, no cloud mailbox | *"Directory sync has not caught up. Wait for the next sync cycle and retry."* |

`ConvertTo-LdapSafe` escapes the five characters an LDAP filter treats as
syntax. A UPN containing one of them would otherwise produce a filter that is
either invalid or — worse — matches something unintended.

## `disable-retry.csv`

Lives in the disable-request folder, **not** under `Logs\`. It is pending work
and rows get removed from it; an append-only, delete-denied audit folder (§20)
is the wrong home for a mutable queue.

```
Type, MB, Delegate, OD, Employee,
FirstFailed, LastAttempt, Attempts, FailedStep, StepsDone, LastError
```

`Type` is mandatory — without it a retry does not know which step list to run.

**Removing a row on success loses no history**, because every attempt has
already written a row to `CompletedTasks.csv` with the same target. Filter the
audit log by target and the whole saga is there, including how many tries it
took. The queue stays a list of outstanding work rather than becoming a second,
worse log.

An empty queue keeps its header rather than being deleted, so "no failures" and
"the file went missing" stay distinguishable.

## Ticket parsing

Pasted ticket text carries invisible characters — non-breaking spaces, figure
spaces, soft hyphens — that make a value fail to match for reasons nobody can
see on screen. `ConvertTo-NormalLine` strips them.

`NA`, `N/A`, `NONE`, `NULL`, `-` and `NO` all normalise to empty. Operators type
these to mean "nothing", and each one used to become a literal delegate name
that then failed to resolve.

**A `Delegate:` or `OneDrive:` line after the last `Logon Name:` applies to the
whole ticket.** Real tickets list ten people and then say `Delegate: NA` once at
the bottom; attaching that to whichever record happened to be open meant nine
people silently got nothing. Invisible when the value is `NA`, actively wrong
the first time it is a real name. Only **blank** values are filled, so a
per-user ticket (Logon Name / Delegate / Logon Name / Delegate) still works —
each person's own value was set before the next `Logon Name` started.

CSV import is read-only and optional. Thirty terminations in a spreadsheet is a
real case and pasting thirty ticket blocks is not a workflow. A file with no
`Type` column is treated as the old three-column format, which was Franchise.

## Notifications

Subject lines carry the **departing** user's address, because that identifies
which OneDrive this is about — the recipient already knows their own address.
The operator's copy appends the environment so it can be filtered on arrival:

```
to the new owner : OneDrive Ownership Change Notification - bob.smith@expresspros.com
your copy (Fran) : ... - bob.smith@expresspros.com - Fran
your copy (HQ)   : ... - bob.smith@expresspros.com - HQ
```

The copy goes to whoever **ran** the tool, via `Get-OperatorEmail`
(Settings.json, else the running account's AD `mail` attribute). It used to be
one hardcoded address, which meant Jeff's disables mailed Glenn — the copy had
stopped being a check on your own work and become mail about somebody else's.
Resolved once per run, not per user: it cannot change mid-batch, and a failed AD
probe should not be repeated for every row.

**No configured address means no copy is sent.** Falling back to a default would
quietly mail somebody who did not do the work, which is the problem this
replaced.

## Preflight

All four connections are checked **before** anything is touched, reading the
shared capability map so this cannot disagree with the console's own indicators.

This is not belt-and-braces. A batch started without the on-premises session
fails at `ConvertShared` for every single user, and each lands in the retry
queue with the same misleading error — so the queue, which is supposed to mean
"these need attention", instead means "you forgot to connect". One dialog up
front is worth more than thirty rows of the same message.

An *unproven* SharePoint session also stops the run. The console skips the SPO
probe on startup because it is expensive, and `Unknown` is normally not the same
as unavailable — but for this tool it is still a reason to stop, because the
OneDrive steps will fail if it turns out to be absent.

## Interface decisions

**840 wide, matching `AdminConsole.ps1`.** The first cut was 1160, which put the
grids beside the input boxes and made the whole thing hard to read — a window
wider than the console that launched it reads as a different application. Taller
is cheap; wider is not. The two paste boxes stack rather than sitting side by
side: at this width two columns would be about 380px each, narrower than the
lines people paste into them.

**One button per path.** `Process Requests` and `Retry Selected` each parse,
fill the plan grid, and go straight to the confirmation. An earlier "Build Plan"
step made you click twice to reach a confirmation you were going to get anyway.

**The plan grid has no Steps column** — the results grid names every step as it
runs, which is the same information at the point where it actually tells you
something. The results grid drops `Type` for the same reason: it is on the plan
row and never varies within a user.

**The retry grid drops `StepsDone` and `LastError`** into a detail pane below
it. Both are long enough to force horizontal scrolling that hides everything
else. The pane also computes **`Will run`** — the steps *not* in `StepsDone` —
which is the difference between guessing what a retry does and knowing.

**Retry buttons are declared before the grid.** Controls added to a container
earlier sit at a lower z-index, which in WinForms means drawn *on top*: a grid
added before its buttons hides them completely, which is exactly what happened.

The connection line is a **live status**, not a standing instruction. "Connect
before running" is advice nobody reads twice; "NOT connected: EMS, SPO" is a
fact. It is `NoTheme`-tagged so its red/green survives every palette.

Window placement is `BottomLeft` like every other tool — `TopLeft` belongs to
`AdminConsole.ps1`, and a tool opening there lands on top of the console that
launched it.

## Traps hit while building this

**`System.Collections.Generic.List[object]` returned through `@()`** throws
*"Argument types do not match"* from PowerShell 5.1's `PSToObjectArrayBinder`.
It bit in the most misleading way available: the disable itself **succeeded**,
the audit row was written, and only then did `Update-RetryQueue` hand an *empty*
List to `Export-RetryQueue` and blow up. Success, followed by a CLR crash dialog
with no script line in it. `ArrayList` is used instead — same in-place mutation,
which is required because `$results` is appended from inside a scriptblock where
`+=` would only update a local copy. Full write-up in ARCHITECTURE §14a.

**Every WinForms handler is wrapped.** An exception escaping a Click handler is
caught by .NET, not PowerShell, so the operator gets assembly versions instead
of a line number. `Show-HandlerError` puts the PowerShell view back and writes a
`Failed` row to the audit log, because an error somebody dismissed is otherwise
the one kind of failure that leaves no trace.

**`Disable-InboxRule` takes `-Identity` only.** Exchange Online rejects
`-Identity` and `-Mailbox` together, and the Identity returned by
`Get-InboxRule` is already scoped to its mailbox. Passing both cost two real
disables on 2026-07-31.

**`$upn`, not `$newId`, in the create failure path.** `$newId` is assigned
*inside* the try, so a throw from the create cmdlet itself — the failure most
worth recording — would log a blank target.

## Still open

- `Send-MailMessage` is formally deprecated by Microsoft and should move to
  Graph or an SMTP client eventually.
- The tenant constants at the top (licence group, OneDrive base URL, SMTP host,
  no-reply address) are in code rather than Settings — grouped and named so the
  move is mechanical.
- A sweep for *other* mailboxes forwarding **to** the leaver. The disable
  handles outbound; nothing yet finds inbound.


## Licence group and directory lookup (2026-08-09)

### 'E5 License Fran - Full' is gone

It was one tenant's group, in one directory, written into a script that ships to
six. Tenant.json has carried a LicenseGroups list -- Name, Sku, Directory --
that nothing read. `Resolve-TenantLicenseGroupName` now resolves it per record
from `$Ctx.Type`.

WORSE THAN NOT WORKING, IT REPORTED SUCCESS. The membership test was "is the
user in a group whose CN is this", and in a tenant without that group the honest
answer is no -- so the step returned **Skipped: not a member**, which is
indistinguishable from a user who genuinely has no licence. A red step gets
fixed. A green step that did nothing gets shipped, and turns up on an invoice.
An unconfigured group is now **Failed**, naming the directory and pointing at
Tenant Manager.

### "Could not ask" is not "not there"

`catch { $ad = $null }` sent an unreachable domain controller, an expired
session and a permissions error to the same place as a genuine miss -- and the
operator was told *"most likely a typo in the ticket"*. They would check the
ticket, find it correct, check again, and have nothing further to go on, because
the one fact that explained it was discarded a line earlier.

`Get-ADUser -Identity` throws ADIdentityNotFoundException for a real miss. Any
other exception is now kept and reported as "the directory did not answer, so
whether this account exists is UNKNOWN -- do not remove the row." Same shape as
the catch that reported "version unavailable" for a control-ordering bug.
