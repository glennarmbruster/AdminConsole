# ResourceDuplicator.ps1

Duplicates a room or equipment mailbox: reads every setting off an existing
resource, creates a new one, and applies the captured settings.

Two variants. **AdminConsole** creates on-premises with `New-RemoteMailbox` and
waits for directory sync. **FH-Agnostic** is cloud-only here and uses
`New-Mailbox`, because Express is the only tenant here that creates resources
on-prem first.

## Why it is three steps and not one

`New-RemoteMailbox` creates the object on-premises. The **cloud** mailbox does
not exist until directory sync publishes it, so `Set-CalendarProcessing`,
`Set-Place`, `Set-MailboxFolderPermission` and `Set-MailboxRegionalConfiguration`
have nothing to bind to for anywhere from minutes to half an hour.

A single-shot script would therefore have to sit and poll, and would fail in a
way that looked like a bug rather than a wait.

**Step 3 is deliberately self-contained.** It reads the capture *file* and a
target address typed into it — not anything left in memory by steps 1 and 2.
Directory sync takes long enough that you will close the window and go and do
something else, which is the entire reason the capture file exists. An earlier
version depended on in-memory state and the comment claiming otherwise was
wrong; Glenn caught it by trying exactly that.

The capture picker refreshes on `Form.Activated`, so a capture written by a
different run of this tool appears without a refresh button. A capture written
after Create remembers the target, so returning to finish the job needs no
retyping.

## Captures live under `reports\captures\`

They are output, so they belong under the output root — and being under it means
they inherit `ReportsPath`. Point reports at a share and a duplication started
on one machine can be finished on another. A subfolder rather than loose in
`reports\` because `ReportBrowser` lists `*.csv`, and a pile of JSON alongside
would make that view mixed-purpose for no gain.

## Set-Place and the documentation problem

Microsoft documents `Street`, `City`, `State`, `PostalCode`, `CountryOrRegion`,
`Phone` and `GeoCoordinates` as **not** working on a synchronized room mailbox.

Observed on 2026-07-30: **five of the seven do work.** Only `Phone` and
`GeoCoordinates` genuinely do not. `Phone` is therefore written on-premises to
`telephoneNumber` at create time instead.

Since neither the documentation nor a single observation is trustworthy alone,
**every applied value is read back** with `Get-Place` and compared. `Set-Place`
reports success either way — asking for the value afterwards is the only way to
know it landed. `Tags` is an array, so it is compared as a joined string to stop
ordering differences reading as failures.

Values written on-prem rather than in the cloud are **disclosed** in the result,
not silently dropped.

## Other details

Every cloud call goes through `Invoke-EXOCommand`, which forces resolution to
the temporary EXO module. This tool needs on-premises **and** cloud connections
in the same window, so an unqualified `Get-Mailbox` here is a coin toss (§2).
`*-RemoteMailbox` needs no such guard — it exists only on-premises.

`Get-Place` is room-only and returns nothing until the room reaches the Rooms
address list, so its absence is not an error.

**Room list membership is why a new room never shows up in Room Finder.** It is
a separate distribution group and nothing about creating the mailbox adds it.
The tool reports the source room's list membership so it can be matched by hand.

Booking behaviour is copied wholesale except `Identity` and the resource's own
address — copying those would either fail or point the new room at the old one.

`Set-MailboxFolderPermission` fails when no entry exists yet; `Add-` is the
other half, so both are tried.

Each settings group is applied independently and failures are **collected rather
than thrown**: a room with a correct booking policy and a missing tag is a far
better outcome than a room with neither.

## `Get-ResourceADServerParam`

A local wrapper, not the shared `Get-ADServerParam`. That function is a
FH-Agnostic-toolkit function and does not exist in the tenant-specific
`Common.ps1`, so calling it there fails with *"not recognized"*. Resolving it
locally is what lets this one file work unchanged in either folder.
