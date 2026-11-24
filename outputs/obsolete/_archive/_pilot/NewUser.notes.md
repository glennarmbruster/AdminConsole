# NewUser.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Enables on-prem Remote Mailboxes (with archive) in Exchange Hybrid, with
    automatic HQ vs Franchise detection.
.DESCRIPTION
    REWRITTEN 2026-07-24 from the original two-box version. What changed, and
    why:

    ONE INPUT BOX INSTEAD OF TWO. The original made the operator know in
    advance whether each account was HQ or Franchise and type it into the
    matching box. That is no longer needed: every account is looked up in HQ
    first, then Franchise, and whichever forest answers is where it acts.
    Mixed batches are fine -- HQ and Franchise accounts can sit in the same
    list, in any order.

    DETECTION IS BY sAMAccountName, NOT BY EMAIL -- and this distinction is
    load-bearing, so do not "simplify" it to reuse Find-TargetEnvironment.
    That existing function resolves an object by its MAIL ATTRIBUTE, which is
    right for every other tool in this suite because they all act on things
    that are already mail-enabled. The accounts THIS tool targets are not
    mail-enabled yet -- that is the entire reason it is being run -- so their
    mail attribute is empty, and a mail-based lookup would report "not found"
    for every valid account. This uses Common.ps1's Find-UserEnvironmentBySam
    instead, which resolves by sAMAccountName: an attribute that exists from
    the moment the account is created.

    "ACCOUNT DOES NOT EXIST YET" IS NOW A FIRST-CLASS OUTCOME rather than a
    confusing Exchange error mid-run. Every account is resolved BEFORE
    anything is changed; any that exist in neither forest are collected and
    reported in a single popup, and nothing is attempted for them. One popup
    listing them all, not one per account -- this is a batch tool.

    RESULTS STAY IN THE WINDOW. The original hid the form and wrote to the
    console, where output scrolled away. There is a results box now, and the
    form no longer hides.

    THE FRANCHISE DC DROPDOWN REMAINS, but is now consulted only for accounts
    that actually resolve to Franchise. HQ accounts still run forest-wide with
    no DC pinned. Note detection itself uses Common.ps1's own
    $script:FranchiseAutoDetectDC; the dropdown governs the WRITE, so the
    operator can still aim changes at a specific DC and verify against that
    same DC afterwards.

    UNCHANGED FROM THE ORIGINAL, deliberately:
      - Uses only *-RemoteMailbox cmdlets, which are on-prem-only and
        unambiguous. It never calls Get-Mailbox / Set-Mailbox, so it is immune
        to the on-prem/cloud cmdlet collision that makes the DL tools risky to
        run alongside an Exchange Online session.
      - Enables the archive with a second Enable-RemoteMailbox call, and
        treats an archive failure as a WARNING on an otherwise-successful
        enable rather than failing the whole operation.
      - Pre-checks with Get-RemoteMailbox and SKIPS accounts that already have
        one, showing their existing state.
      - RemoteRoutingAddress is <sam>@expresspros.mail.onmicrosoft.com.
      - HQ uses Set-ADServerSettings -ViewEntireForest; Franchise pins
        -PreferredServer and passes -DomainController.
      - Does NOT touch CN, UPN, sAMAccountName, display name, email address
        policy, or licensing. GUID stamping still intentionally omitted (the
        cloud mailbox is not provisioned quickly enough at enable time).
      - Logs one CompletedTasks row per verified enable.
.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-06-17
    Revised         : 2026-06-24 0900
    Revised         : 2026-07-24 -- single box + HQ/Franchise auto-detect by
                       sAMAccountName; missing accounts reported explicitly;
                       results shown in-window. See REWRITTEN note above.
    Revised         : 2026-07-27 -- archive state now derived from ArchiveGuid
                       instead of raw ArchiveStatus, which read "None" on
                       archives that had in fact been created (the status is
                       cloud-side and had not replicated back yet). See
                       Format-ArchiveState. Also: an archive failure is now
                       recorded in the CompletedTasks row, not just in the
                       results box.
    Run from        : On-prem Exchange Management Shell
    Required modules: ActiveDirectory (via EMS host); no ExchangeOnlineManagement
```

- **L83-86** Common.ps1 supplies Find-UserEnvironmentBySam, Get-ADServerString, Format-ErrorDetail, Write-CompletedLog, $script:FranchiseDCs and $script:FranchiseAutoDetectDC -- so this tool no longer keeps its own copies of the DC list or the logging helper.

## region Configuration

- **L90** Specific to this tool, so it stays local rather than moving to Common.ps1.

## function Get-OnPremState

- **L108-110** ArchiveGuid is selected alongside ArchiveStatus deliberately -- see Format-ArchiveState below for why ArchiveStatus alone is misleading here. empty $DC = HQ (no preferred server)

## function Format-ArchiveState

<!-- line 118 -->

```text
Renders the archive line for a RemoteMailbox.

        WHY THIS EXISTS: reading raw ArchiveStatus here produced a false
        negative -- it showed "None" immediately after a successful
        Enable-RemoteMailbox -Archive, on accounts whose archive had in fact
        been created. ArchiveStatus is a CLOUD-side property; on an on-prem
        RemoteMailbox object it reflects what has replicated back from
        Exchange Online, which lags the enable by minutes to hours, so
        checking it seconds later will essentially always read "None". This is
        the same behaviour Microsoft KB4486848 describes.

        ArchiveGuid is the reliable indicator, because Enable-RemoteMailbox
        -Archive stamps it on the on-prem AD object as part of the enable
        itself -- nothing has to round-trip to the cloud for it to be true.
        Every other tool in this suite (MailboxHealth, MailboxPermissions,
        Mailbox-Group-Inspector, M365Logic) already derives ArchiveEnabled
        this way; this tool was the last one reading ArchiveStatus raw.

        So: GUID present = archive is enabled, full stop. ArchiveStatus is
        still shown in parentheses, but only as replication detail -- never as
        the verdict.
```

- **L149** The exact case that prompted this: enabled on-prem, cloud not caught up.

## function Invoke-MailboxEnable

<!-- line 167 -->

```text
Enables one account's remote mailbox + archive. The environment is
        already resolved by the caller, so this simply acts on what it's told.
        Returns report text; the caller accumulates it.
```

- **L174-175** '' = HQ (forest-wide); populated = Franchise 'HQ' or 'Franchise', for display + logging
- **L183-185** Pin the DC for Franchise writes; go forest-wide for HQ. Unchanged from the original -- see its note about allowing replication before verifying HQ changes.
- **L197** Pre-check: already mail-enabled? Doubles as the BEFORE state.
- **L206** 1. Enable the remote mailbox.
- **L228-229** 2. Enable the archive. A failure here is a warning, not a failure of the enable above -- the mailbox exists either way.
- **L241-244** 3. Verify, then log only on a verified success. The archive outcome is carried into the log line: verification passes on the mailbox alone, so without this an archive failure would produce a CSV row that reads as a clean success once the window is closed.

## region WinForms UI

- **L279** Forest context, mirroring the tenant label the other consoles now carry.
- **L303** ONE box now -- see header. Mixed HQ/Franchise batches are fine.
- **L345-346** Results stay in the window -- the original hid the form and wrote to the console, where output scrolled away.
- **L358** AcceptButton intentionally NOT set -- Enter inserts newlines in the input box.
- **L376-379** --- Pass 1: resolve EVERY account before changing anything. Two --- reasons: the confirmation can then state exactly what will happen --- and where, and missing accounts surface up front instead of --- halfway through a batch.
- **L390-391** HQ writes go forest-wide (''); Franchise writes are pinned to the DC the OPERATOR picked, not the detection DC.
- **L404** --- The requested popup: accounts that exist in neither forest.
- **L432-433** --- Confirm, naming the environment each account resolved to, so a --- wrong detection is visible BEFORE anything is written.
- **L453-454** --- Pass 2: act. HQ first, then Franchise -- same order as the --- original, keeping DC-pinned Franchise writes grouped together.
- **L507** ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
- **L509** Console, which opens top-left. Helper is in Common.ps1.
