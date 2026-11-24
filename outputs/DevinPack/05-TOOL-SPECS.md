# TOOL SPECS — build order, with the traps that apply to each

Build in this order. Read-only tools first, so every write tool stands on a read
path already proven against the live tenant.

Codes in the **Traps** rows refer to `01-HARD-CONSTRAINTS.md`. They are the
specific ones that have actually bitten these tools — not a general reminder to
be careful.

**Every tool below has a working exemplar in `reference/`**, named at the top of
its section. Read the exemplar first, then write the tool from the spec. The
exemplar answers *how* — layout, guards, logging, error handling; the spec
answers *what*. Do not edit a copy of the exemplar: the source tenant had two AD
domains where this one has one, and copying carries assumptions that stay
invisible until they are wrong.

Every tool, without exception:

- dot-sources `Common.ps1` by upward search (never a counted `..`)
- wraps every handler (**B1**) and turns word wrap on in output boxes (**B6**)
- writes an audit row with the **primary SMTP address** as Target, never a
  display name
- opens **bottom-left** so it never covers the console (only the console uses
  top-left)
- carries a **Help button** and a filled-in `TOOLKIT-HELP` block -- no exceptions
- follows the fixed end-of-form order: target bar, theme, corner, help button
- passes `Test-Toolkit.ps1` clean before it is considered done

---

## Phase 2 — with the console

### 0. Activate (PIM)

> **Exemplar: `reference/activate.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Opens Privileged Identity Management so the operator can turn their admin
access on.*

Built with the console rather than among the tools, because **nothing else works
until it has been used.** Admin rights at this tenant are *eligible, not
standing*: until an operator activates, they are signed in with no admin access
and every other tool fails in ways that look like faults rather than like a
missing activation.

| | |
|---|---|
| **Input** | None |
| **Does** | Opens the PIM portal in the admin browser profile; records the activation time locally |
| **Writes** | Nothing in the tenant. A local state file only |
| **Traps** | **A4** — `[int]` rounds half-to-even, so the remaining-time display must `[Math]::Floor` or a 8h59m remainder renders as "9h 59m". This exact bug shipped once. **B2** — the portal opens in a browser, i.e. a separate process, so it is unaffected by the console's modal loop. |

**ORDER MATTERS, and it is the part people get wrong:**

> 1. **Activate FIRST**
> 2. **Connect SECOND** (Exchange Online, SharePoint, on-premises)

A connection made *before* activating keeps the token it was issued with, and
**that token does not contain access activated afterwards.** The symptom is a
tool that authenticates perfectly and then refuses every operation, which reads
as a permissions problem and is not one. If the operator has already connected,
they must activate and then reconnect. Say this in the tool's help text.

**The console header countdown** reads `Pim.DurationHours` from `Tenant.json` —
**9 at this tenant** — and shows remaining time via `Get-PimStatusLine` and
`Get-PimStatusColour` (contract Tier 1b). It is not a timeout the toolkit
enforces; it is a display of how long Entra granted. A wrong duration produces a
confidently wrong countdown, which is worse than none: an operator who trusts it
starts a long job believing they have hours in hand.

**Done when:** the portal opens in the admin browser profile (not whichever
profile was last used), the activation time is recorded locally, and the console
header counts down from **9 hours**. Check the countdown against a clock after a
few minutes -- a rounding error here reads as a longer policy than you have.

**Out of scope: programmatic self-activation.** PIM for Groups is a Graph-only
API needing admin-consented delegated permissions
(`PrivilegedEligibilitySchedule.Read.AzureADGroup` and
`PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup`). Unless that consent
exists, open the portal. Do not attempt a Graph implementation without
confirming consent first — and if you do confirm it, ask before building.

---

## Phase 3, then Phase 4a — read-only

> **Tool 1 alone is Phase 3, and it ends at a hard stop** — see
> `00-BUILD-PROMPT.md` §6. Tools 2 and 3 below are the start of **Phase 4** and
> must not be written until the operator has run tool 1 on Windows.

### 1. Mailbox Health

> **Exemplar: `reference/MailboxHealth.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*The first place to look when someone says "my email is doing something
strange".*

One mailbox, one screen, read-only. Answers in a single view what would
otherwise be six commands: size and quota, mailbox type, forwarding, delegates
and permissions, auto-reply state, litigation hold, archive status, and whether
the account is licensed.

| | |
|---|---|
| **Input** | One SMTP address or alias |
| **Reads** | `Get-Mailbox`, `Get-MailboxStatistics`, `Get-CASMailbox`, `Get-MailboxPermission`, `Get-MailboxAutoReplyConfiguration` |
| **Writes** | Nothing |
| **Traps** | **C9** — archive state comes from `ArchiveGuid`, never `ArchiveStatus`, which lags replication and reports `None` on an archive that exists. **C14** — `EwsEnabled` is often `$null` meaning *inherit*; resolve against `Get-OrganizationConfig` and show the effective value. **C12** — do not classify recipient type by substring. |

Build this first even though it is not the most useful tool, because it
exercises the whole read stack — resolution, connection gating, form layout,
audit logging of a read — with nothing at risk.

**Done when:** every field is either populated or explicitly says why it is not.
A blank field that means "not applicable" and a blank field that means "the
lookup failed" must not look the same.

---

### 2. Message Trace

> **Exemplar: `reference/MessageTrace.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Follow a message through Exchange Online — where it went, whether it was
delivered, what happened if it was not.*

| | |
|---|---|
| **Input** | Sender, recipient, subject, date range, status filter |
| **Reads** | `Get-MessageTrace`, `Get-MessageTraceDetail`, quarantine cmdlets |
| **Writes** | Quarantine release only |
| **Traps** | **C3** — the trace enumerates; pass `ResultSize = 'Unlimited'`. **A4** — any latency arithmetic must `[Math]::Floor`, or 119 seconds renders as "2m 59s". **D6** — never write a zero-row export. |

Two design points that matter more than they look:

- **Filter the results without re-running the trace.** A search box over the
  returned rows, not a new query. Re-running is slow and the operator has
  usually already got what they need on screen.
- **Say what the tool does not do.** Message trace answers "what happened to
  this message". It is not threat hunting — no URL verdicts, no campaign views,
  no bulk remediation. Put that in the help text and link out to the right
  place, because the alternative is somebody concluding the data is missing.

Trace data is retained for roughly ten days in the detailed view. State that in
the UI rather than returning an empty grid for an older date.

**Done when:** a trace returns rows for a message you can verify independently,
the search box narrows them **without re-running the query**, and a date outside
the retention window says so rather than returning an empty grid.

---

### 3. Unified Lookup

> **Exemplar: `reference/Unified-Lookup.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Find any object without knowing where it lives.*

The most complex read tool and the one worth the most. One search box; searches
Exchange Online **and** the on-premises directory in the same pass, reports what
each side found, classifies each result, and opens the right editor for that
combination.

| | |
|---|---|
| **Input** | A name fragment, alias, or SMTP address |
| **Reads** | `Get-Recipient`, `Get-Mailbox`, `Get-DistributionGroup`, `Get-UnifiedGroup` in the cloud; AD in the directory |
| **Writes** | Nothing directly — it routes to editors |
| **Traps** | **C12** — `GroupMailbox` contains "mailbox" and is a Microsoft 365 Group. **C3** — every one of these enumerates. **A1** — the result filter returns a possibly-empty selection; do not also use `$null` for cancelled. **C4** — carry the GUID from the search into whatever tool is opened; never re-resolve by the display Identity. |

Two behaviours worth specifying explicitly:

- **Cap the result set, and say so.** A common word matches thousands. Cap at a
  few hundred, and when the cap is hit **say the results are incomplete** —
  silently truncating leaves the operator filtering a set that was already
  wrong, which looks like a tidy complete answer.
- **A name existing in two places is not an ambiguity to resolve.** It is the
  answer to the ticket. Show both.

**Done when:** an object that exists only in the cloud, only on-premises, and in
both, each produce a correct and visibly different result.

---

## Phase 4b — low-risk writes

From here on, every tool writes. Each needs: a confirmation naming exactly what
will change, an audit row on success **and** on failure, and the guard from
**C2** if it touches on-premises.

### 4. Mailbox Permissions

> **Exemplar: `reference/MailboxPermissions.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Grant, view or remove access to somebody else's mailbox.*

Two things that are constantly confused and are granted separately:

- **Full Access** — open and read the mailbox
- **Send As** — send mail appearing to come from it

*"I can see the shared mailbox but can't send from it"* is Full Access without
Send As, and it is the most common ticket in this area. Handle both in one tool
and label them unmistakably.

| | |
|---|---|
| **Reads** | `Get-MailboxPermission`, `Get-RecipientPermission` |
| **Writes** | `Add-/Remove-MailboxPermission`, `Add-/Remove-RecipientPermission` |
| **Traps** | **C12** — a Microsoft 365 Group is not a mailbox; access to one comes from *membership* and these cmdlets do not apply. Detect and refuse with an explanation rather than failing obscurely. **C4** — bind on GUID. |

Note on AutoMapping: it can be set when granting Full Access but there is **no
documented way to read it back** in Exchange Online. Do not display a value you
cannot verify — offer it as a write-time choice and say that current state
cannot be shown.

Access granted here does not expire. Say so in the UI.

**Done when:** Full Access and Send As can each be granted, listed and removed
**independently** -- granting one must not grant the other -- and pointing the
tool at a Microsoft 365 Group produces a clear refusal explaining that access
there comes from membership, not from these cmdlets.

---

### 5. Calendar Permissions

> **Exemplar: `reference/CalendarPermissions.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*See and change who can view or edit someone's calendar.*

Answers "my assistant can't see my calendar" and "someone can see more than they
should".

| | |
|---|---|
| **Reads** | `Get-MailboxFolderPermission` on the calendar folder |
| **Writes** | `Add-/Set-/Remove-MailboxFolderPermission` |
| **Traps** | The calendar folder path is localised — resolve it rather than hardcoding `:\Calendar`. `Default` and `Anonymous` are pseudo-users that always exist and cannot be removed, only set to `None`. |

Show current permissions for everyone who has one, and offer the standard access
levels rather than raw rights strings.

**Done when:** the current list matches what Outlook shows for the same mailbox,
`Default` and `Anonymous` appear and can be set to `None` but not removed, and a
non-English calendar folder still resolves.

---

### 6. Out-of-Office Editor

> **Exemplar: `reference/OOFEditor.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Read and set someone's auto-reply without needing their mailbox open.*

For the classic ticket: somebody left, or is off sick, and their auto-reply
needs setting or clearing on their behalf.

| | |
|---|---|
| **Reads** | `Get-MailboxAutoReplyConfiguration` |
| **Writes** | `Set-MailboxAutoReplyConfiguration` |
| **Traps** | There are **two** messages — internal and external — and they are usually different, with only one of them wrong. Show both, always. The scheduled window matters: `Scheduled` with a past end date reads as "on" in some views and sends nothing. |

**Done when:** both messages are visible without scrolling or switching tabs, a
scheduled window shows its start and end, and auto-reply can be turned off on a
mailbox you do not own.

---

### 7. Inbox Rules

> **Exemplar: `reference/InboxRules.ps1`.** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*View, disable, enable and delete rules on a mailbox.*

The usual reasons to look: mail vanishing into a folder nobody created, mail
forwarded somewhere it should not be, or a rule left behind by a compromise.

| | |
|---|---|
| **Reads** | `Get-InboxRule` |
| **Writes** | `Enable-/Disable-/Remove-InboxRule` |
| **Traps** | **Deleting cannot be undone; disabling is reversible.** Default to disable and make delete the deliberate choice. **C4** — bind on the rule's identity carefully; rule names are not unique and are frequently blank. |

A rule forwarding externally is a security finding, not just a nuisance —
highlight it.

**Done when:** rules list with their conditions readable, **disable is the
default action and delete requires a separate deliberate confirmation**, a rule
with a blank name is still addressable, and external forwarding is visibly
flagged.

---

## Phase 4c — groups

### 8. Group Manager, and Group Manager Editor

> **Exemplar: `reference/Group-Manager.ps1` (front door) and `reference/Group-Manager-Editor.ps1` (editor).** Read it before writing.
>
> Match its shape and conventions; write the tool from this spec rather than
> editing a copy of it.

*Start here for any group; it works out what the group is and opens the right
editor.*

Two tools. The front door classifies; the editor edits.

**The front door** looks the group up and determines two things — **what** it is
and **where** it lives:

| What | Where | Opens |
|---|---|---|
| Distribution list / mail-enabled security group | synced from on-premises | Editor, against on-premises |
| Distribution list / mail-enabled security group | cloud-native | Editor, against Exchange Online |
| Microsoft 365 Group | (always cloud) | Group editor — members and owners |

**A Microsoft 365 Group is a different animal** and that is why it gets its own
destination. Access comes from **membership**, not mailbox permissions, and it
is managed by an entirely separate cmdlet family (`*-UnifiedGroup`) that the
distribution-list editor cannot speak.

| | |
|---|---|
| **Reads** | `Get-Recipient` to classify, then `Get-DistributionGroup` / `Get-UnifiedGroup` |
| **Writes** | Membership, owners, delegation, addresses |
| **Traps** | **C12** — the classification hinges on not matching `GroupMailbox` as a mailbox. **C18** — `IsDirSynced` decides which side to write; say so *before* the attempt, since the cloud refuses most changes to a synced object. **C4** — resolve once, bind on GUID; a group's `Identity` renders as the AD CN and re-binding it is a fresh forest-wide lookup. **C10/C11** — address edits rewrite the whole proxy collection; carry X500 across or replies to historical mail break. **C1/C2** — this is the tool family most exposed to the cmdlet collision. |

> **`Group-Manager-Editor.ps1` MUST declare `# TOOLKIT-REQUIRES: OwnProcess` in
> its own header, in the first 40 lines.** Its on-premises path uses
> `Add-`/`Remove-ADPermission`, and those names exist in both Exchange editions
> — in one process the last import owns them (C1/C2). This is the only tool in
> the first pass that needs the marker, and it is the most exposed file in it.
> Two more declare it at Stage F; see `09`.
>
> **The front door has a console button; the editor does not.** The editor
> refuses to open standalone — it is reached only by routing from Group Manager,
> whose window title reads **DL Manager**, the name operators use for it.

Why the routing matters operationally: during a migration, groups move
on-premises → cloud a batch at a time, and **nothing about the group's name or
address changes when it does.** There is no way to tell by looking, which is
precisely why the tool should determine it rather than the operator guessing.

**Done when:** a synced distribution list, a cloud-native distribution list and a
Microsoft 365 Group each route to the correct editor, **and the front door says
which it found and why** before opening anything. A silent correct route and a
silent wrong one look identical.

---

## After these

**The remaining forty-eight tools are specified in
`09-FULL-TOOL-CATALOGUE.md`**, in eight lettered stages. Do not open it as a
build instruction until this first pass has been run on Windows and accepted per
`07`.

They are variations on patterns established above: attribute editors (apply
**C10** and **C11** exactly — this is where the X500 loss happens), reports
(**C3** and **C15** — slice large queries, never retry an oversized one), room
administration, and on-premises provisioning.

**A leaver / termination workflow is NOT in the catalogue** and is not
authorised work. If it is ever asked for, it is the highest-risk tool in any
toolkit of this kind: multi-step, partially reversible, and the step that fails
is usually not the step that reports. It would need a retry queue — a record of
what could not be finished, so a half-completed run is visible rather than lost.
The console's Dashboard already surfaces a retry-queue panel for exactly that
shape of tool.
