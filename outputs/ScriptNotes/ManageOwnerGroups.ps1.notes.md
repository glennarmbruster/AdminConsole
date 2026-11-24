# ManageOwnerGroups.ps1

Adds or removes franchise owners from the FranchiseOwners distribution groups:
the US group plus the alphabetical sub-group for their surname.

**AdminConsole only.** On-premises Exchange, Franchise. It is not in
`FH-Agnostic/` or `CloudOnly/` and should not be — the group names and the
four-band split are an Express process, not a pattern other tenants share.

Brought onto `Common.ps1` on 2026-08-03. Before that it was a standalone script
that happened to live in the same folder.

---

## What it does

Two groups per person, every time:

```
FranchiseOwnersUS@<domain>          everyone
FranchiseOwnersA-F@<domain>         surname A-F
FranchiseOwnersG-L@<domain>                 G-L
FranchiseOwnersM-R@<domain>                 M-R
FranchiseOwnersS-Z@<domain>                 S-Z
```

Surname is **everything after the last dot** in the local part, so
`mary.jo.smith` resolves to `smith` and lands in S-Z. A local part with no dot
has no surname to read; it is reported rather than guessed at, because a guess
puts a real person in the wrong group and nothing afterwards would show it. The
same applies to a non-A-Z initial — an accented surname or a leading digit gets
no band and says so.

Ten members per run. Add and Remove, neither preselected.

## What changed in the refactor

### The logging was the real problem

The old version carried its **own** `Write-CompletedLog` — the v1 two-column
`Task,DateTime` schema, writing to a hardcoded `E:\Scripts\M365\Logs\Completed`.
Launched by the console with `&`, that definition shadowed the v2 one from
`Common.ps1` inside this script's scope. So the tool worked, wrote a log, and
that log went to a different file in the wrong format.

Which means it was invisible. Not in LogViewer, not in the monthly reporting.
This is exactly the gap the audit log v2 work (§20) existed to close, sitting in
a tool that was doing real work every week.

**One row per member, not per run.** Ten owners produce ten rows, each with
`Target` set to that person's address. A single row listing ten people joined by
`+` answers "was this done" but never "was it done for *her*", and per-person
filtering is what LogViewer and Levi's reports key on. `Result` is per member:
`Success` when both groups took, `Partial` when one took and one was already in
the desired state, `Failed` when either group errored.

`-RunId` is deliberately **not** passed. `Write-CompletedLog` derives it from
`Get-LogScriptName`, which keys the per-process id on the script. Passing one
from here would have called `Get-ToolRunId` with no `-ScriptName`, filing every
row under the shared `unknown` key and silently grouping this tool's work with
anything else that made the same mistake.

### `-BypassSecurityGroupManagerCheck` on every call

The old version had it on the US group in both directions and on **neither**
alpha group call. That worked only because the alpha groups happen to have no
manager set today.

The first time one gets a manager, removals would fail there and succeed on the
US group — leaving the owner half-removed and still receiving mail from a group
nobody would think to check. Both cmdlets now go through one splat in
`Set-OwnerGroupMembership` that always includes it, so the asymmetry cannot come
back by editing one branch.

### No hardcoded domain

Group **stems** are constants; the domain comes from `Get-TargetSmtp`, which
already implements "append the domain if there isn't one". Reusing it rather
than repeating the rule is the point — repeating it is how the old copy drifted
into a literal `expresspros.com` in eight places.

The stems and bands stay in code rather than `Settings.json` because they are a
process decision, not a deployment one. The team chose four bands and these
names; a config file would invite changing them without changing the process
that depends on them. They are grouped and named so the move stays mechanical if
it is ever wanted — the same treatment the tenant constants in `Disable-User.ps1`
get.

### The DC list came from Common

The old file carried its own copy of the eight franchise controllers. It was
already drifting: a DC added to `Common.ps1` would not have appeared here. Now
sourced from `$script:FranchiseDCs`, defaulting to `$script:FranchiseAutoDetectDC`.

**The selector stays visible**, unlike the console's, which is hidden. Group
membership replicates, and being able to pin a controller and then re-check on
that same one has real operational value when replication is lagging. The hint
beside it says so, because "which DC did I write to" is the question people ask
ten minutes later.

### Encoding

The file was **Windows-1252 with no BOM**, while its own line 2 declared "UTF-8
with BOM marker — do not remove this line". Thirteen em-dashes were byte `0x97`,
seven of them inside runtime UI strings.

It worked because PowerShell 5.1 falls back to the system codepage when a script
has no BOM. It would have mangled the first time anything treated it as UTF-8.
Now UTF-8 with a single BOM and CRLF, and no non-ASCII characters at all —
em-dashes replaced with `--` like the rest of the toolkit.

### Everything else

- **Every handler wrapped** in `Show-HandlerError`. The old `Parse` handler had
  no `try`, so an exception there would have produced a CLR dialog with no
  script line.
- **`Set-ConsoleTheme` + `BottomLeft`** placement, matching every other tool.
  `CenterScreen` put it over whatever was underneath; `TopLeft` belongs to the
  console.
- **Connection state is a live label, not an instruction.** "Connect first" is
  advice nobody reads twice; "NOT connected to on-premises Exchange" is a fact.
  Re-checked on Run as well, because a batch started without a session fails on
  every member with the same misleading error, and thirty identical errors read
  as a broken tool rather than a missing connection.
- **Over-limit entries are reported.** The old parser broke out of the loop
  quietly, so pasting twelve names ran ten and never mentioned the other two.
  They now appear in the preview and in the confirmation dialog.
- **`[string[]]` cast on the preview.** `TextBox.Lines` is `string[]` and
  `ArrayList.ToArray()` returns `object[]`. PowerShell usually coerces that, but
  "usually" inside a WinForms property setter is how you get a CLR dialog with
  no line number (§14a).
- **`Parse-Members` → `ConvertTo-OwnerList`.** Approved verb, and it now returns
  both the kept and dropped lists rather than just an array.

## Two interface decisions worth keeping

**Neither radio button is preselected**, and that is not an oversight. Add and
Remove are opposite and equally plausible on any given ticket, so a default is a
trap: the operator confirms a dialog that agrees with what they meant to do
while the radio says the other thing. Run refuses to proceed until one is
chosen.

**The DC hint sits beside the selector, not in the header.** "Membership changes
replicate -- re-check on the SAME controller you wrote to" answers the question
people actually ask ten minutes later, when the change appears to have not
taken. Putting it next to the control that causes it is the only place it gets
read.

## `Set-ADServerSettings` was added to the EMS capability patterns

Without it this tool scanned as needing **no connection at all**, so its button
would never grey out.

The EMS patterns only matched `*-RemoteMailbox`, `Get-ExchangeServer` and
`Update-Recipient`. This tool's working cmdlets are
`Add-`/`Remove-DistributionGroupMember`, which are exported by both Exchange
Online and on-premises and therefore identify neither.
`Set-ADServerSettings` is on-premises exclusive, which makes it the honest
marker — the same reasoning that made `*-RemoteMailbox` the marker in the first
place.

## Adding it to the console

New section, then the tool. In `ToolManager.ps1`:

- **Sections…** → add `OnPrem Scripts`, with however many button rows you want.
- **Add** → Name `Franchise Owner Groups`, Script `.\ManageOwnerGroups.ps1`,
  Section `OnPrem Scripts`.

The button greys itself out when the Exchange Management Shell is not loaded,
via the capability scan above. No wiring beyond the two entries.

## Still open

- The four alpha bands are uneven — S-Z and A-F carry more surnames than G-L in
  most English-speaking populations. Nobody has complained, and rebalancing
  means moving live members between groups, so it is noted rather than proposed.
- Error matching is on English message text (`already a member|not a member of
  the group`). A localised Exchange would report those as failures instead of
  skips. Not a concern in this environment, but it is why that match is fragile.
