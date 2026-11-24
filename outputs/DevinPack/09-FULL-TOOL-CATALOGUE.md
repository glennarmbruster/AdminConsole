# FULL TOOL CATALOGUE — every remaining tool in the finished toolkit

`05-TOOL-SPECS.md` specifies the first pass in depth — Activate plus eight
tools, ten scripts — and stops there. **This document specifies everything else:
forty-eight further scripts in eight stages, plus additions to three files
already built** (`AdminConsole.ps1`, `ToolManager.ps1`, `Test-Toolkit.ps1` — see
§5). Together they describe the whole toolkit, and the whole toolkit is the
destination.

**A note on names.** `Group-Manager.ps1` is called **Group Manager** in `00`,
`02`, `04` and `05`. Its window title, and what operators call it, is **DL
Manager**. They are the same tool. This document uses *Group Manager* for
consistency with the rest of the pack and notes the window title where it
matters.

---

## 0. READ THIS FIRST — how this document relates to the rest of the pack

**`00-BUILD-PROMPT.md` §5 defines the FIRST PASS. This document defines the
DESTINATION. They are not in conflict, and neither one cancels the other.**

Several documents in this pack say things are "out of scope" — reports,
SharePoint tools, provisioning, theme editing. **Read every one of those as "not
in the first pass."** They are all in this catalogue and all get built. The
sequencing is deliberate: the first pass proves the conventions on eight tools,
and everything here is built on top of conventions that have already been
confirmed on a real tenant.

Specifically, these read as *deferred* rather than *excluded*:

| Where | Deferred thing | Where it lands |
|---|---|---|
| `00` §5 | provisioning | Stage F |
| `00` §5 | reporting suite | Stage D |
| `00` §5 | SharePoint tools | Stage A (Connect-SPO) and Stage D |
| `00` §5 | theming beyond one palette | Stage H |
| `04` | "`Tools.json` … most of which are deferred out of the first pass" | Sections and `Group` field: §2 and §3 below |
| `08` | `Add-ReportFileControls` | Stage D, with the first report |
| `08` Tier 3 | theme editing | Stage H |

**Genuinely excluded, and NOT re-authorised by this document:** everything in §6
below, and **a leaver / termination workflow**, which is not in this catalogue,
is named as unauthorised in `05`, and would be the highest-risk tool in the
toolkit. If it is ever wanted, it is new work with its own conversation.

**`00` §7 tells you to stop and ask when you are about to build something not in
§5. That instruction still stands for anything not in THIS document.** Work
listed here is authorised; work listed nowhere is not.

**Do not begin any stage below until the first pass has been run on Windows and
accepted** per `07-ACCEPTANCE-CHECKLIST.md`. Everything here compounds the
conventions established there — including the mistakes.

### Stage numbering

This document uses **lettered stages (A–H)** deliberately, because
`00-BUILD-PROMPT.md` §6 already numbers Phases 0–5 and its **Phase 5 is
handover**. Stages A–H sit between `00`'s Phase 4 and its Phase 5. When you
report progress, say "Stage C" or "Phase 4" — never a bare number.

---

## 1. What "done" means

The finished toolkit is **55 console buttons over 65 `.ps1` files.** Every entry
exists because somebody needed it on a real ticket.

That reconciles as:

```
55 console buttons
  – 1  the "Rooms" button, which fronts 5 scripts (see §3)
  + 5  those five scripts
= 59  scripts reachable from a console button
  + 1  Group-Manager-Editor.ps1   opened BY Group Manager, no button of its own
  + 2  Common.ps1, AdminConsole.ps1        infrastructure
  + 2  Deploy.ps1, Publish-Toolkit.ps1     run from a shell, never a button
  + 1  UnBlock.ps1                shipped inside handouts; not a tool
= 65  .ps1 files
```

**Build in stage order.** Do not build breadth-first across all of them. Each
stage ends where a human can run what exists and find out whether the
conventions are holding, and a convention that has drifted across forty files is
far more expensive to correct than one that has drifted across five.

**Piecemeal delivery is expected and fine. A partial tool is not.** A tool is
delivered or it is not started — no stubs, no `TODO`, no half-wired buttons. A
button that opens something broken is worse than a button that is absent,
because the operator has already stopped looking for another way to do the job.

---

## 2. The finished console

```
Connect        1 row  ×  5   =  5 slots     4 buttons
Admin Tools    5 rows ×  5   = 25 slots    24 buttons
Daily Scripts  2 rows ×  5   = 10 slots     5 buttons
Reporting      3 rows ×  5   = 15 slots    12 buttons
Toolkit        2 rows ×  5   = 10 slots    10 buttons
                               ---------   ----------
                               65 slots    55 buttons
```

Admin Tools' 24 = **8 from `05`** (Mailbox Health, Message Trace, Unified
Lookup, Mailbox Permissions, Calendar Permissions, Out of Office, Inbox Rules,
Group Manager) **+ 16 from this document**.

`Sections.json` for the finished state:

```jsonc
{
  "ButtonsPerRow": 5,
  "Sections": [
    { "Name": "Connect",       "Rows": 1 },
    { "Name": "Admin Tools",   "Rows": 5 },
    { "Name": "Daily Scripts", "Rows": 2 },
    { "Name": "Reporting",     "Rows": 3 },
    { "Name": "Toolkit",       "Rows": 2 }
  ]
}
```

**Ship sections as they fill.** A section with no tools renders as an empty box,
so add each one when it has three or more tools, not before. The table above is
the destination, not the starting `Sections.json`.

> **Two deliberate moves from the source toolkit.** `Cert Check`
> (`CertificateExpiry.ps1`) and `Tester` (`Test-Toolkit.ps1`) sit in **Admin
> Tools** in the source `Tools.json`; here they move to **Toolkit**, which is a
> new section. Neither is an operator tool for a ticket. This move is what keeps
> Admin Tools at 24 of 25 slots — if you put them back, Admin Tools overflows
> and the console silently drops two buttons.

### The height check is now load-bearing

Thirteen rows of buttons plus five section headers plus the lookup controls plus
the output box is a tall window. On a 1080p screen with a taskbar it is close to
the limit.

**At console startup, compute the total required height from `Sections.json` and
compare it against `WorkingArea` — not `Bounds`, which includes the taskbar
(B4; verifier rule `WF004`).** If it does not fit, say so plainly in the output
box and name the section that overflowed. Do not silently draw buttons below the
bottom edge, where they are unreachable and invisible, and do not silently
shrink them.

This did not matter with fifteen buttons. It matters at fifty-five.

---

## 3. Grouping — one button, several scripts

**A `Group` is several tools behind one console button.** Clicking it opens a
small chooser listing the members by name; picking one runs it. This is how
fifty-nine button-backed scripts fit into fifty-five buttons without a wall of
near-identical controls.

`Tools.json` gains one optional field:

```jsonc
{
  "Name": "Room Regional Settings",
  "Script": ".\\RoomRegional.ps1",
  "Section": "Admin Tools",
  "Group": "Rooms",          // <-- same string on every member
  "Order": 22,
  "ObjectTypes": [],
  "NewProcess": false
}
```

Rules, all of which the reference console already implements:

- **Same `Group` string on every member.** Blank means the tool stands alone,
  which is the default and what most tools should stay as.
- **The group button takes the position of the member with the lowest `Order`**,
  so the group lands where you expect rather than wherever the alphabet puts it.
- **Members are still created as buttons — just hidden.** Two reasons, both
  real: the console's search filter can then reveal a matching member directly,
  so a grouped tool stays reachable by name; and the grey-out pass can compute
  the group button's state from members it already has in front of it.
- **The chooser repeats each member's availability.** A member that cannot run
  is disabled *in the chooser too*, with the same tooltip reason the console
  would have shown. Otherwise the group button becomes the place where greying
  goes to die, and the operator learns the reason one dialog too late.
- **Capacity counts buttons, not scripts.** A group of five occupies **one**
  slot. Get this wrong and §2's arithmetic is wrong.

> **Fully implemented in `reference/AdminConsole.ps1`** — `Get-ToolGroupName`,
> `Show-ToolGroupChooser`, and the button-build loop — and editable via the
> `Group` column in `reference/ToolManager.ps1`, which carries its own
> explanation. Take both verbatim rather than reimplementing.

**Groups in the finished toolkit: `Rooms` only** (five scripts, Admin Tools).
Others may be introduced if a section overflows, but grouping tools that are not
genuinely siblings makes them harder to find, not easier.

---

## 4. How to read an entry

Each entry heads with:

```
`Section` · ObjectTypes · capability · variant · Stage · ~lines
```

- **`ObjectTypes: []`** means the button is **never greyed on the looked-up
  object** — the tool is not about one object. It does **not** mean the button
  is never greyed at all: capability gating still applies, per
  `02-ARCHITECTURE-SPEC.md`. A tool marked `ObjectTypes: []` · *needs EXO* is
  greyed when Exchange Online is unavailable and enabled regardless of what was
  looked up.
- **capability** names what must be available for the button to be live —
  *needs EXO*, *needs EMS* (on-premises Exchange), *needs AD*, *needs SPO*, or
  *none*. Absent capability ⇒ greyed, with a tooltip saying which one.
- **variant** is the `# TOOLKIT-VARIANT:` marker the script carries.

**Where an entry names a constraint (C4, C11, …), that constraint has bitten
this specific tool before.** It is not a general reminder.

**Where an entry names a cmdlet, verify it against Microsoft Learn before use**,
per `00` §4A, and record the citation at the call site. This matters more here
than in `05`: these tools reach into corners of Exchange — `Set-Place`,
`Set-CalendarProcessing`, `New-RemoteMailbox`, blocked-sender release — where
parameter names and accepted values are easy to get plausibly wrong and the
failure is silent.

**Where you cannot tell what a tool should do from its entry, ask.** Do not
invent behaviour, and do not build a thinner version and call it done. Every
entry summarises something that already exists and is in daily use, so a missing
detail is knowable rather than a matter of taste.

### What every tool carries

The standard header block; the upward search for `Common.ps1`; the connection
guard; wrapped handlers (B1); bottom-left placement (B3a); the `?` help button
added **last** (B3b); UTF-8 with a single BOM and CRLF (A8).

### Audit `Target` — use `Get-LogTarget`, and stop thinking about it

`02`, `08` and `03`'s `LOG002` all require the **primary SMTP address** as
`Target`. That is correct for every tool acting on a mail-enabled object, and it
is most of them.

**About half the tools in this catalogue have no mailbox in play** — every
report (whose target is the report path), the three viewers, Header Analyzer,
Transport Rules, Connect-SPO, Dashboard and the whole Toolkit section.

**`Common.ps1` already solves this. Use it:**

```powershell
Write-CompletedLog -Action 'Grant' -Result 'Success' `
    -Target (Get-LogTarget $mbx.PrimarySmtpAddress $mbx.Name)
```

`Get-LogTarget` takes candidates in preference order and returns the first
non-empty one. **The fallback is deliberate, not lazy** — a plain security
group, a room list before Exchange assigns an address, a failed create: some
things genuinely have no address, and a name is a worse identifier than an
address but a far better one than nothing.

**Where there is no object at all, pass a literal scope name:** `'SharePoint
Online'`, `'Toolkit'`, `'Config'`, the report path. `reference/Connect-ExO.ps1`
is the pattern.

**The reference verifier already exempts literal targets** — it flags
`.Name` / `.DisplayName` / `.SamAccountName` but deliberately skips
`-Target '...'`, because those are the non-object cases where the literal *is*
the identifier. **Do not "fix" that exemption, and do not fabricate an address
to satisfy a rule.** A made-up address in an audit log is worse than a plain
label, because it looks searchable.

**A tool that reads nothing and changes nothing need not log at all.** Quick
Links and M365 Reports open a browser; there is no event worth a row.

---

## 5. Stages

Each stage ends with a **stop**: report what was built, what the verifier says,
and what went into `UNVERIFIED.md`. Do not roll into the next stage unprompted.

| Stage | What | Scripts | Why here |
|---|---|---|---|
| **A** | Read-only tools, plus Connect-SPO | 6 | Lowest risk in the set. Establishes the query-into-a-grid pattern that C–D reuse. SPO must exist before the SharePoint reports in D. |
| **B** | Cloud group tools | 2 | Depends on the classification logic built for Group Manager in Phase 4. |
| **C** | Cloud writes on a single object | 4 | First writes beyond the first pass. C10/C11 live here. |
| **D** | The reporting suite | 12 | Highly patterned and near-identical to one another. The volume is here and the risk is low — build one properly, then the rest as a batch. |
| **E** | Rooms | 5 | Self-contained, cloud-only, and the first `Group`. |
| **F** | On-premises and hybrid writes | 6 | **Highest risk in the toolkit.** C1/C2/C7/C8 apply to every one. Do not start before Stage C's audit rows have been confirmed on a real tenant. |
| **G** | Console furniture | 3 | Dashboard reports on everything above, so it goes after them. |
| **H** | Toolkit maintenance and the publish chain | 10 | Needed to keep and hand on the toolkit, not to use it. |

48 scripts. **Stage F is the one to slow down on.** Every tool in it can write
to the wrong directory silently, and "no exception was thrown" is not evidence
that it wrote where you intended (C1, C10).

### Work that is NOT a new script, and when it happens

Three pieces of this catalogue land in files built during the first pass. They
have no stage of their own and are easy to miss:

| What | Where | When |
|---|---|---|
| **The `Group` mechanism** — `Get-ToolGroupName`, `Show-ToolGroupChooser`, the grouped button-build loop (§3) | `AdminConsole.ps1` | **Before Stage E.** Take verbatim from the reference |
| **`Group` column + over-capacity warning** (§2, §3) | `ToolManager.ps1` | **Group column before Stage E; the capacity warning before Stage A**, which is where a section first overflows |
| **`WF010` / `WF011` + the `LOG002` literal exemption** | `Test-Toolkit.ps1` | **Before Stage A** |

**`Sections.json` `Rows` must grow with each stage, in the same commit.** The
first pass ships `Admin Tools: Rows 2` — ten slots, eight tools. Stage A adds
six Admin Tools buttons. **Fourteen buttons in ten slots does not error; four
buttons are simply absent.** Raise `Rows` to 4 at Stage A and to 5 by Stage E,
and add the `Toolkit` section at Stage H. §2 has the destination.

### Own-process isolation, and what it is not

Three tools here declare `# TOOLKIT-REQUIRES: OwnProcess`. That mechanism is
**the console's**, not a `Common.ps1` function: `AdminConsole.ps1` reads each
script's own declaration via `Test-ToolDeclaresOwnProcess` and launches it with
`Start-ToolScriptInNewProcess`. Both are in `reference/AdminConsole.ps1` and are
in scope.

`08-COMMON-CONTRACT.md` tells you **not** to port `Start-ToolInOnPremProcess`.
That is a different thing — it is multi-directory *routing*, which this
single-domain tenant does not need. **That prohibition stands. Own-process
isolation is still required**, because it defends against the cmdlet collision
(C1), not against a second directory.

---

# THE CATALOGUE

---

## Connect

### SPO — `Connect-SPO.ps1`
`Connect` · `ObjectTypes: []` · none · **Stage A** · ~167 lines

Connects the console to SharePoint Online. **Same pattern as `Connect-ExO.ps1`:**
the console closes, the sign-in happens, and the console reopens with the
connection live — because a sign-in window cannot accept typing while the
console's own window is in front of it (B2).

Needed by the OneDrive and SharePoint reports at Stage D, and by Mailbox /
OneDrive Perms at Stage F. **The Exchange tools do not use it and must not be
greyed on its state.**

**Flip `Capabilities.SharePointOnline` to `true` in `Tenant.json` as part of
this stage.** `02` ships it `false` because nothing in the first pass uses it,
and an unticked capability makes its tools **absent rather than greyed** — which
reads as a build defect.

> The exemplar is `reference/Connect-ExO.ps1`. **There is no
> `reference/Connect-SPO.ps1`** — only the ExO and ExSE connect scripts ship in
> the reference set. Follow ExO's shape.

**The admin URL is a TENANT FACT and lives in `Tenant.json`, at
`SharePoint.AdminUrl` — never in `Settings.json`.** This is a corrected bug, not
a preference: the source read `Settings.json` and then guessed, never looking at
`Tenant.json` at all, so the value Tenant Manager wrote and displayed was inert.
An engineer could type the correct admin centre into the tool whose entire job is
tenant facts, save it, and watch Connect SPO use a derived host that had never
existed. `Settings.json` no longer carries the key at all.

**Derivation is the fallback, not the source.** When `SharePoint.AdminUrl` is
empty, derive `https://<label>-admin.sharepoint.com` from the **session
organisation** — the SharePoint tenant label follows the `*.onmicrosoft.com`
prefix, **not** the vanity mail domain. They match at most tenants; they do not
always, and a tenant that renamed its mail domain would otherwise get a URL that
never existed.

Logs with `Target` = `SharePoint Online`.

---

## Admin Tools

### Dashboard — `Dashboard.ps1`
`Admin Tools` · `ObjectTypes: []` · none · **Stage G** · ~919 lines

The start-of-day page, and `Order: 0` — the first button on the console.

**Six panels**, each answering a question worth asking before touching anything:

| Panel | Question |
|---|---|
| **ENVIRONMENT** | Am I connected to what I need, and is PIM still active? |
| **TOOLKIT HEALTH** | Is my copy complete and current? |
| **TODAY** | What has run today, and did any of it fail? |
| **TOP TOOLS** | What actually gets used. |
| **FAILURE HOTSPOTS** | What keeps going wrong, so it can be fixed. |

> **The source has a sixth panel, RETRY QUEUE, and you should NOT build it.** It
> reports on a leaver/termination workflow, which is not in this catalogue and is
> not authorised work (`05`). Nothing in scope writes a retry queue, so the panel
> would be permanently empty and would read as a defect rather than as good news.
> Build five panels. If a multi-step workflow is ever added, the panel comes with
> it.

**TOOLKIT HEALTH reads `Version.json`, which only `Publish-Toolkit.ps1` writes
(Stage H).** Until then it must say "not stamped" rather than "incomplete" — an
unstamped working copy is the normal state, not a fault.

**READ-ONLY.** Nothing here changes a mailbox, a group or a file. Its buttons
open other tools; they do not act.

**The environment and toolkit-health panels must call the SAME functions as the
console header and footer**, not work it out again. Otherwise the two can
disagree, and **a second opinion about your own state is worse than one.** This
is a hard requirement, not a tidiness preference.

**Never fold an unknown result into `Success`.** Any audit row without a
`Result` counts as `Unknown` and is displayed as such. A dashboard that
overstates what worked is worse than one with a visible gap.

**It is built at Stage G** because five of its six panels report on tools built
in Stages A–F. Building it early produces a page reporting on nothing.

> Dashboard can launch tools that have no console button. `ToolStatus.ps1`
> recognises that as a legitimate state — see its entry and §7.

### AD Lookup — `AD-Lookup.ps1`
`Admin Tools` · `ObjectTypes: []` · **needs EMS** (degrades to AD-only) · **variant: OnPremAD** · **Stage F** · ~1,383 lines

Type part of a name, get back every matching on-premises object with **what it
is** and **where it lives**.

**It sits in Stage F, not with the read-only tools, because its attribute editor
writes to on-premises AD with `Set-ADObject`.** That puts it behind the same
gate as the other on-premises writers.

**How it reads, which is the thing most easily got wrong:** the search calls
`Get-Recipient` **unqualified against on-premises Exchange**, routed through
`Invoke-OnPremCommand`. That routing is not optional — an Exchange Online
session in the same process would answer for the cloud and **the OU column would
be empty on every row** (C1).

**When there is no on-premises Exchange in the process at all**, it falls back
to Active Directory and derives the type from attributes via
`Resolve-AdObjectType` in `Common.ps1`. That is a **degraded path**, not the
primary one.

> **Do not build this on `Get-ADObject` alone.** `msExchRecipientTypeDetails`
> comes back from AD as an integer, and the source explicitly considered and
> rejected decoding that bitmask. Guessing it produces a Type column that is
> confidently wrong, which is worse than an empty one.

**The Type column is the point.** These look alike in a name list and behave
nothing alike:

| Type | What delegation means |
|---|---|
| `UserMailbox` / `SharedMailbox` | Full Access and Send As both work |
| `MailUniversalDistributionGroup` | Send As works; Full Access is meaningless |
| `MailUniversalSecurityGroup` | A distribution list that also grants rights |
| `MailContact` | No mailbox anywhere; nothing to delegate |
| `RemoteSharedMailbox` | The mailbox is in the cloud; this is the on-premises stub that syncs it |

Getting this wrong is not a typo — it is twenty minutes and a confused reply.

**The OU column** answers the other half: which business unit owns this, and
which directory it came from.

**Double-click puts the address on the clipboard**, ready to paste into whichever
tool does the actual work. That is the primary workflow.

**It works on GROUPS**, which the older attribute tools did not — they used
`Get-ADUser` and simply failed on anything that was not a user. **Half of what
this finds is a distribution list.**

**"Also search Active Directory"** is a checkbox that widens the Exchange search
to include objects with **no mailbox at all** — the plain security group sitting
next to the shared mailbox with almost the same name. It cannot be delegated to,
but seeing it explains why the name looked ambiguous. **Off by default, because
it is slower.**

#### The attribute editor

Opens the selected object. Everything populated is shown; **46 attributes are
editable** — names, job and organisation, every phone field, the full address,
and all fifteen `extensionAttribute` fields.

**Five classes are READ-ONLY, and the list is short on purpose. All five must be
enforced:**

| Class | Fields | Why |
|---|---|---|
| **addresses** | `mail`, `proxyAddresses` | They belong to Exchange. Editing them behind its back **is how an object gets two primaries** (and see C11). Send the operator to Mail Attributes |
| **identity** | `sAMAccountName`, `userPrincipalName`, `cn`, `name` | Changing one of these is a migration, not an edit |
| **relationships** | `member`, `memberOf`, `managedBy`, `manager` | The value is a distinguished name, and a text box invites somebody to type a person's name into it |
| **state** | `userAccountControl`, the timestamps | Enabling an account by hand-editing a bitmask is a way to be wrong quietly |
| **multi-valued** | anything shown as `multi (n)` | **Writing one value back would silently discard the rest** |

Reading is safe. **The attribute editor is the only thing that writes**; it
confirms every change and records it in the audit log.

Escape every operator-typed value before it reaches an LDAP filter — RFC 4515
hex escaping, **not** the OPath doubling used for Exchange filters (C19a). The
two are not interchangeable, and using the wrong one is a silent wrong-results
bug rather than an error.

### Header Analyzer — `HeaderAnalyzer.ps1`
`Admin Tools` · `ObjectTypes: []` · none · **Stage A** · ~824 lines

Paste raw internet headers, get a readable account of what happened to the
message:

- the delivery path, **oldest hop first**, with the time spent at each hop
- SPF, DKIM and DMARC results, plus Microsoft's composite verdict
- the Exchange Online filtering result — spam confidence, bulk level, and *why*
  it was filtered, in words rather than numeric codes
- anything worth a second look, such as a `Reply-To` pointing somewhere other
  than the `From` address

Answers "why did this go to junk", "why was this blocked", and "is this actually
from who it says it is".

**Everything happens on this machine and nothing is uploaded.** That is the main
reason the tool exists — headers carry internal server names, addresses and
subject lines that must not be pasted into a public analyser website. Say so in
the help text.

Connects to nothing; parses text. No connection guard. Nothing worth auditing.

### DL Memberships — `DLMemberships.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **variant: Cloud** · **Stage A** · ~328 lines

Every distribution list a person belongs to. Answers "why is this person getting
this mail" and "what will they stop receiving when they leave" — the second is
worth checking *before* a termination rather than after.

**One filtered query, not an enumeration.** It should return in a second or two
however many groups the tenant has. Fetching all groups and filtering
client-side is the wrong shape (C15).

**Two limitations, both stated in the window and not only in the help:**

1. **DIRECT MEMBERSHIPS ONLY.** If someone is in "All Staff" because "Accounts"
   is in "All Staff", this lists Accounts and not All Staff. That is the reason
   a person can still receive mail from a list not shown here.
2. **Microsoft 365 groups and dynamic distribution groups are not included.**
   Dynamic groups have **no stored membership to query** — they are evaluated
   when mail is sent. Do not attempt to enumerate them.

Read-only. Point the operator at Group Manager or Bulk DL Membership to change
anything.

### Group Editor — `GroupEditor.ps1`
`Admin Tools` · `ObjectTypes: ["UnifiedGroup","OnpremGroup"]` · needs EXO · **variant: Cloud** · **Stage B** · ~601 lines

View and change **owners and members** of Microsoft 365 groups and of
on-premises groups synced to the cloud.

**Owners matter more than they look.** A group with no owner can be managed by
nobody except an admin, and nobody is notified when the last one leaves. Surface
the owner count prominently and warn before removing the last one.

**C18 applies and must be handled before the attempt, not after it.** `OnpremGroup`
means `IsDirSynced` is true — the object is mastered on-premises and **the cloud
will refuse most changes to it.** Check `IsDirSynced` and say plainly that the
change must be made on-premises, rather than issuing a write that appears to
succeed. Reporting success on a change that did not happen is the exact failure
mode `00` §1 names as the reason this pack exists.

Changes that do go through take effect immediately and are written to the audit
log.

Reached from Group Manager's routing as well as from its own button.

### UnifiedGroup Editor — `UnifiedGroup.ps1`
`Admin Tools` · `ObjectTypes: ["UnifiedGroup"]` · needs EXO · **Stage B** · ~344 lines

Membership of a Microsoft 365 group specifically.

**A Microsoft 365 Group is not a distribution list**, even though it receives
mail like one. It also carries a Teams team, a SharePoint site and a shared
mailbox — so adding somebody here gives them considerably more than an email
feed, **and the confirmation must say that.**

Owners and members are separate lists and both matter; ownership changes belong
in Group Editor.

Its recipient type is `GroupMailbox`, which **contains the substring "mailbox"
and is not one** (C12). Any code deciding "is this a mailbox" by substring
classifies this wrongly.

### Mail Attributes — `MailAttributes.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **variant: Cloud** · **Stage C** · ~739 lines

Addresses, display name, alias, and how the person appears in the address book.
**Mostly used for renames** — someone marries, changes name, or was created with
a typo.

**This is the most damaging tool in the toolkit if written naively, and it is
where C11 lives.** Writing a bare `EmailAddresses` list **REPLACES the whole
proxy-address collection.** The X500 addresses that keep old Outlook replies and
calendar items resolving are in that collection, they are invisible in most UIs,
and losing them is not obvious until weeks later when replies start bouncing.

**Add and remove individual entries. Never assign the whole collection.**

**Adding the new address while keeping the old one is the safe move**, because
mail to the old address keeps arriving. Changing the *primary* changes what
recipients see and what appears in the directory; the old address should normally
be kept as a secondary rather than deleted, and the tool should default to that.

**C10 applies with force.** `Set-Mailbox -PrimarySmtpAddress` is on-premises
only, and `Set-MailUser -PrimarySmtpAddress` is **accepted and silently ignored**
in Exchange Online. A tool that calls it, sees no exception and reports success
is lying. **Read the value back and compare before reporting anything.**

### All Attributes — `AllAttributes.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **variant: Cloud** · **Stage C** · ~449 lines

A wider view of the directory attributes on an account, beyond the mail ones.
For when something is set that should not be, or missing when it should be
there, and the operator needs the **raw value** rather than what a UI reports.

**Editing here writes directly to the attribute with no validation beyond what
the directory itself enforces.** Say that in the confirmation dialog, not only in
the help text.

Point the operator at Mail Attributes first if the field is a mail one — that
tool understands addresses properly and this one does not.

### Audit UserMB — `AuditUser.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **Stage A** · read-only · ~441 lines

Reports the **mailbox audit settings** on a user mailbox — which actions are
recorded, separately for the owner, for delegates, and for administrators.

Used when someone asks what can be proven about access to a mailbox, or before
an investigation, to establish whether the events of interest were being logged
**at the time**. That is the tool's whole value: it answers a question about the
past, and the answer is often no.

### Audit SMB — `AuditSMB.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **Stage A** · read-only · ~497 lines

The same report for a **shared** mailbox, and worth being a separate tool.

Shared mailboxes usually have several delegates, so **delegate actions are the
events that matter most** — and they are the ones most often not being recorded.
Order and emphasise the output accordingly rather than mirroring the
user-mailbox layout.

### Blocked Senders — `BlockedSenders.ps1`
`Admin Tools` · `ObjectTypes: []` · needs EXO · **variant: Cloud** · **Stage C** · ~481 lines

Every account the service has blocked from sending mail, and the ability to
release one.

An account lands here after sending mail the service classified as spam. In
practice that is usually a deliberate bulk send that went out the wrong way —
**but it is also exactly what a compromised mailbox looks like, and the two are
indistinguishable from this list alone. Find out which before releasing.** That
sentence belongs next to the release button, not in a footnote.

Columns shown before any action:

| Column | Why |
|---|---|
| **Reason** | why the service blocked them, in its own words |
| **Blocked** | when it happened |
| **Released here** | how many times *this toolkit* has released this address, and when it last did |

The last column is computed from the toolkit's own audit log. A second release is
a coincidence; a fourth is a pattern — usually a team that keeps forgetting the
bulk-send tool. Releasing again fixes today and not next week.

**Microsoft limits how many times an account can be released.** Exceed it and
the release fails and only a support case clears it. **The count is a budget, not
just a habit tracker.**

**The ticket reference the operator types is recorded with the release** — in the
audit log and in the service — so "who released this and why" is answerable
later.

Releasing does not fix the cause. **If the account was compromised, releasing it
hands the mailbox back to whoever took it** — put that in the confirmation.

### Transport Rules — `TransportRules.ps1`
`Admin Tools` · `ObjectTypes: []` · needs EXO · **variant: Cloud** · **Stage A** · ~427 lines

Every mail-flow (transport) rule in the organisation, in one list.

**It does not change, enable or reorder any rule.** The only thing it writes is a
snapshot file (below).

**Why not edit here, and the reasoning belongs in the help text.** A transport
rule applies to every message the organisation sends or receives, and
`Set-TransportRule` carries well over a hundred conditions and actions. A form
covering all of them would be the admin centre rebuilt worse; a form covering
some of them could only edit the rules it understood and **would quietly drop the
parts it did not.** The admin centre's condition builder is the one thing it
genuinely does well. Author there. If asked to add editing, say why not.

**Four capabilities, and the last two are why the tool earns a button:**

| | |
|---|---|
| **Search** | one box, matched against the **whole rule** — name, description, every condition and every action. "Which rule mentions this domain" is one keystroke here and a long afternoon in the admin centre |
| **Priority** | the running order in a single list. **Which rule wins is decided by this, and it is the thing the admin centre shows least clearly** |
| **Detail** | every property the rule actually **sets**, with the couple of hundred empty ones left out |
| **Snapshot** | the whole collection to a dated XML file. **Mail flow has no change history, so a snapshot before anyone edits is the only way to answer "what did this look like last week"** |

**"Take a snapshot BEFORE changing rules, not after. It is the only version of
yesterday you will ever have."** Put that sentence in the tool.

The snapshot is the one write, and it is to a file. Audit it with the snapshot
path as `Target`.

### Change Organizer — `ChangeMeetingOrganizer.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **variant: Cloud** · **Stage C** · ~382 lines

Moves a meeting or series from one organizer to another. The usual case is
someone leaving who owns a recurring meeting nobody else can edit.

Flow: look up the **current** organizer → type the subject → name the new
organizer → **Preview** → **Transfer**.

**Four warnings. All four must appear BEFORE the transfer, not after.** An
operator who learns any of these afterwards has made things worse than before
they started.

1. **Teams.** Changing the organizer does **not** change the organizer of the
   Teams online meeting attached to it. The new organizer will not be able to
   manage or start that Teams meeting and must replace the online meeting details
   manually afterwards.
2. **Who notices.** Attendees *in this tenant* see nothing — silently updated, no
   re-RSVP. **External and on-premises attendees get a cancellation from the old
   organizer and a new invitation from the new one.** They must re-RSVP, and any
   Reminder, Category or Show As changes they made are lost.
3. **One way, once.** After a transfer you cannot transfer the part of the series
   left behind on the original calendar again.
4. **Scope.** **User mailboxes only, default calendar only.** Shared and room
   mailboxes are **not supported** — give the new person access to the mailbox
   instead.

Point 4 is a functional refusal, not advice. `ObjectTypes: ["Mailbox"]` admits
shared and room mailboxes, so **the tool must check the recipient type itself and
refuse** with that explanation.

### Rooms — group of five
`Admin Tools` · `Group: "Rooms"` · **Stage E** · one button, five scripts

Room administration splits into five genuinely different things that get
confused with one another constantly. The chooser naming them is most of the
value.

All five: `ObjectTypes: []` · needs EXO · **variant: Cloud**.

#### Room Lists — `RoomList.ps1` · ~429 lines

*Why a room does not appear in Outlook's Room Finder.*

**The Room Finder does not list rooms. It lists ROOM LISTS, and then the rooms
inside them.** A room mailbox belonging to no room list is invisible there no
matter how correct its mailbox and place record are — and the symptom people
report is "the room does not exist", which points nowhere near the cause.

A room list is an ordinary distribution group flagged as a room list. **It cannot
be flagged after the fact through this tool**: a normal group converted to a room
list keeps behaving oddly in the Room Finder, so **New is the only route
offered.** That is a deliberate omission, not a missing feature.

Membership changes take effect immediately, **but Outlook and Teams cache the
Room Finder aggressively — allow up to 24 hours** before concluding it did not
work. Verify in this tool rather than in Outlook.

#### Room Places — `RoomPlace.ps1` · ~493 lines

Capacity, floor, building, devices, accessibility, tags — the room metadata
behind the Room Finder.

**This is not the mailbox.** A room has a mailbox (name, address, permissions)
*and* a "place" record (what the room is like), and **only the place record
drives room suggestions.** A room that never appears in a search usually has an
empty place record rather than a broken mailbox.

Writes with `Set-Place` immediately. **No staging and no undo** — say so.

**In a hybrid organisation some fields cannot be set here at all.** If the room
mailbox originated on-premises and syncs up, Microsoft blocks `Set-Place` from
changing `City`, `CountryOrRegion`, `GeoCoordinates`, `Phone`, `PostalCode`,
`State` and `Street`. Those come from the on-premises directory and must be
changed there with `Set-User` or `Set-Mailbox`.

**Mark those rows in the UI rather than letting the operator discover it from a
failure.** This tenant is hybrid, so this is the normal case, not the edge case.

#### Room Regional — `RoomRegional.ps1` · ~522 lines

The timezone a room thinks it is in, and the hours it considers bookable.

A room created without a timezone inherits one, and that is how a room accepts a
9am booking that shows as 3am to the people in it. **Nothing errors; the meeting
is simply at the wrong time.**

**Two timezones live here and they are not the same setting:**

- **Regional TimeZone** — how the mailbox stores and displays times
- **Working hours TimeZone** — the zone the working hours are stated in

They should almost always match. When they do not, "only during working hours"
means something nobody intended, **and this tool must say so** rather than
displaying both without comment.

`ScheduleOnlyDuringWorkHours` in Calendar Processing depends entirely on the
working hours set here. Turning that on without checking these is how a room
starts declining bookings for no visible reason — cross-reference the two tools
in both directions.

#### Calendar Processing — `CalendarProcessing.ps1` · ~498 lines

How a room decides what to do with a meeting request: accept, decline, hold for
approval, and what the reply says. This is where "the room accepts everything",
"the room declines recurring meetings" and "nobody ever approves the bookings"
are actually decided.

**`AutomateProcessing` is the master switch** and must be understood before
anything else on the form is touched:

| Value | Behaviour |
|---|---|
| `AutoAccept` | the room books itself by the rules below |
| `AutoUpdate` | requests are marked tentative; a delegate decides |
| `None` | nothing is automatic; the mailbox behaves like a person |

**Most of the other settings only take effect under `AutoAccept`.** Lay the form
out so that switch reads as governing the rest, because it does, and grey or mark
the settings that are inert under the current value.

**The same "list alone does nothing" trap as Booking Policy applies here.**
`BookInPolicy`, `RequestInPolicy` and `RequestOutOfPolicy` hold lists of people
and interact with the `AllBookInPolicy` / `AllRequestInPolicy` switches.
**Setting a list while the matching `All-` switch is still true does nothing
visible.** Set them together and say so.

Writes with `Set-CalendarProcessing` immediately. **No undo.**

#### Booking Policy — `Set-BookingPolicy.ps1` · ~548 lines

**Who** is allowed to book a room. Everyone not on the list has their request
declined automatically.

Add by typing an address or alias and pressing Enter. **Bare aliases work** — if
`jsmith` does not resolve on its own, retry against the tenant's own primary
domain before giving up. Remove by selecting a row and pressing Delete, or via
the Remove button.

**THE LIST ALONE DOES NOTHING.** A room also carries a master switch,
`AllBookInPolicy`, and while it is `True` the room accepts bookings from everyone
and **the list is ignored entirely.** Restricting a room means setting both. This
tool sets them together and says so before it writes.

**An empty list with restriction on means nobody can book the room at all.** A
valid thing to want and a terrible thing to do by accident — confirm it
separately, with different wording from the normal confirmation.

### Resource Duplicator — `ResourceDuplicator.ps1`
`Admin Tools` · `ObjectTypes: []` · **needs EMS *and* EXO** · **Stage F** · ~819 lines

Copies a room or equipment mailbox: reads every setting off an existing resource,
creates a new one, and applies the same settings — booking rules, delegates,
capacity, auto-accept behaviour. More reliable than setting a dozen options by
hand and hoping you remembered them all.

**This is a HYBRID, THREE-STEP tool and the hardest thing in this catalogue to
get right. It is not an ordinary Admin Tools button.** The form is laid out as
three labelled groups, in this order:

```
Step 1  --  source resource
Step 2  --  new resource  (created on-premises)
Step 3  --  apply settings  (after directory sync; works in a later session)
```

**THE CAPTURE FILE IS THE MECHANISM, and it is the thing you would not guess.**
Step 1 reads every setting off the source and **writes them to a capture file on
disk**. Step 3 reads that file — **not the live session**. That is what makes
"come back after the next sync cycle" workable: the operator can close the
window, go home, and finish tomorrow. The tool says so out loud:

> *"You can close this window: Step 3 reads the capture file, not this session."*

Show the available captures in Step 3, with `(no captures yet -- run Step 1)`
when there are none. **Without the capture file the operator must leave a
WinForms window open across a directory-sync cycle, which is not a workflow.**

**Step 2 — create, on-premises.** Calls `New-RemoteMailbox` with `-Room` or
`-Equipment`. That cmdlet exists **only in the on-premises Exchange Management
Shell**; check for it and fail with that explanation rather than a
`CommandNotFoundException`:

```powershell
if (-not (Get-Command -Name 'New-RemoteMailbox' -ErrorAction SilentlyContinue)) {
    throw 'New-RemoteMailbox is not available. This step needs the on-premises Exchange Management Shell.'
}
```

**Ask the cmdlet what it declares before using a switch** (C13) — `-Room` and
`-Equipment` are not present on every build:

```powershell
$declared = Get-CmdletParameterNames -Name 'New-RemoteMailbox'
if ($declared -notcontains 'Room') { throw 'New-RemoteMailbox does not declare -Room on this build.' }
```

**Step 3 — configure, in the cloud, AFTER directory sync.** The new mailbox does
not exist in Exchange Online until sync publishes it. The tool must **detect
that and stop cleanly**, not fail obscurely:

```powershell
throw "'$Identity' is not in Exchange Online yet. Directory sync has not published it -- wait for the next cycle and click Apply again."
```

**The operator clicks Apply again after the next sync cycle, possibly from a
fresh launch of the toolkit.** A single-pass implementation fails on every run.
This and the capture file are the two most important things to get right here,
and neither is visible from the tool's own help text.

**Every cloud call goes through `Invoke-EXOCommand`** — `Get-Mailbox`,
`Get-CalendarProcessing`, `Set-CalendarProcessing`, `Set-Mailbox`,
`Get-MailboxFolderPermission`, `Get-MailboxRegionalConfiguration`,
`Get-DistributionGroup`, `Get-Place` / `Set-Place`. Several of those names exist
in **both** editions (C1); the wrapper is what stops the on-premises session
answering a call meant for the cloud in a process that has both loaded.

**Its mistakes get copied too.** Show the source configuration for review before
creating anything, and make the operator confirm against what they can see rather
than against a room name.

Stands outside the `Rooms` group deliberately: the other five configure an
existing resource, this one creates one.

### Console — `console.ps1`
`Admin Tools` · `ObjectTypes: ["Mailbox"]` · needs EXO · **Stage G** · ~158 lines

A PowerShell prompt that **shares the console's live connections.** For the times
a tool does not cover what is needed and the operator wants to run cmdlets by
hand without connecting to everything again in a new window.

Whatever the console is connected to, this is connected to — which necessarily
includes the ability to change things, **with none of the previews or
confirmations the tools provide.** State that in the help text.

**Must run in the console's own process**, since sharing the connections is the
entire point (C17). It is the one tool for which `OwnProcess` would defeat the
purpose.

---

## Daily Scripts

### New User — `NewUser.ps1`
`Daily Scripts` · `ObjectTypes: ["Mailbox"]` · needs EMS · **Stage F** · ~584 lines

Creates the mailbox for a new starter in the hybrid setup. Enables a **remote
mailbox** on-premises — the on-premises object that tells the directory this
person's mailbox lives in the cloud — along with their archive.

**The mailbox appears in the cloud after the next directory sync, not
immediately.** That gap is normal and is the usual reason a brand-new account
looks broken for a while. **The completion message must say this**, or the tool
generates its own support tickets.

**C7 applies and is the trap here.** This tool sets Exchange's AD session scope.
`Set-ADServerSettings` is **process-wide and outlives the tool that set it** — a
tool that sets it and does not restore it poisons every subsequent tool in the
console's process. Either restore the previous value in a `finally`, or pass
scope per call and never set it globally. Note which you chose in
`UNVERIFIED.md`.

**C8 also applies:** Exchange's `-DomainController` takes a **server**, not a
domain. Passing a domain name returns error `0x52`, which reads like a typo and
is not.

**Derive the set of directories from what was actually resolved; never hardcode
a list.** The source looped over two literal names, so an account resolving to
anything else matched no iteration and was **silently never written** — no error,
no audit row, no mailbox.

> **This tenant has ONE directory, and `04-REFERENCE-FILES.md` and `00` §8 both
> tell you not to port multi-directory machinery. Both still apply.** Do not
> build cross-directory routing, per-directory controller selection, or
> side-routing. The point here is narrower: where the code iterates over
> resolved objects, iterate over **what it resolved** rather than a literal —
> which with one directory is a one-element set and costs nothing.

### Mailbox / OneDrive Perms — `MailboxOneDriveAccess.ps1`
`Daily Scripts` · `ObjectTypes: []` · **needs EXO *and* SPO** · **Stage F** · ~808 lines

Grants or removes access to somebody's mailbox **and** their OneDrive in one
place. The usual case is someone leaving or off unexpectedly when their manager
needs both — **and doing them separately is how one of the two gets forgotten.**

- **Mailbox** — Full Access and Send As
- **OneDrive** — site collection administrator on their personal site

Gate on both connections and **say which one is missing**, not "not connected".

**OneDrive access is visible to the person if they return.** Appropriate for a
departure, less so for a temporary absence, and either way it should be removed
when no longer needed. Put that in the confirmation.

**Audit each half separately.** A run that granted mailbox access and failed on
OneDrive is a **Partial** — not a Success and not a Failure. If the audit row
cannot distinguish them, the log cannot reconstruct what actually happened, which
is the one job it has.

### Manage Mail Object — `ManageMailObject.ps1`
`Daily Scripts` · `ObjectTypes: []` · needs EMS · **variant: OnPrem** · **`OwnProcess`** · **Stage F** · ~645 lines

Creates and removes shared mailboxes and distribution groups on-premises.

Flow: type the address (with or without domain) → pick object type → pick Add or
Remove → **Preview** → **Apply**.

**Declares `# TOOLKIT-REQUIRES: OwnProcess` in its own header, within the first
40 lines.** `New-DistributionGroup` and `Remove-DistributionGroup` exist in
**both** Exchange editions, and in one process the last import owns the name
(C1/C2). Getting this wrong creates or deletes the object in the wrong place with
no error at all.

**On Add** it needs a directory, because that decides which OU the object lands
in. Directories with an OU configured in Tenant Manager use it without asking;
those without prompt for it.

**Shared mailboxes are created as REMOTE mailboxes** — the AD account is created
on-premises, the mailbox lives in the cloud. That is why they appear after a sync
cycle rather than immediately.

**On Remove it looks the object up first, shows what it found, and refuses if the
object is not the type selected.** Typing a person's address into "Remove Shared
Mailbox" would otherwise delete their mailbox. **This guard is the single most
important line of code in the tool.**

- **Remove** deletes the AD account. The cloud mailbox is soft-deleted and
  recoverable for a limited window.
- **Disable** keeps the AD account and strips its mail attributes. **Disable is
  the reversible one**, and is offered as a checkbox on Remove.

**Owner is optional.** On a shared mailbox it grants FullAccess and SendAs; on a
distribution group it sets `ManagedBy`, which also lets that person edit
membership from Outlook.

Distribution groups reject external senders unless the box is ticked.

### Bulk DL Membership — `Bulk-DLMembership.ps1`
`Daily Scripts` · `ObjectTypes: []` · needs EXO or EMS · **`OwnProcess`** · **Stage F** · ~692 lines

Adds or removes a set of people across a set of distribution lists in one go.
Distribution lists in one box, users in the other.

**The operation is every user against every list.** Ten users across four lists
is forty changes, **and the preview says that number before anything runs.**

- **Check** — resolves everything and shows exactly what will happen
- **Run** — does it, after a confirmation

**Anything that will not resolve is named individually and skipped**, rather than
failing the whole batch. A batch that aborts on item three of forty is worse than
one that completes thirty-seven and names the three it could not.

**One audit row per person**, not one per batch. There is no undo: a removal is
recoverable only by adding people back, and **the audit log is the only record of
who was in what.**

**Which environment it writes to depends on how it was opened.** Opened from **Group Manager**, it follows the group — a synced group is edited on-premises in its own
window, a cloud-managed one in Exchange Online in the console's session. **The
banner at the top of the window always names the environment it is about to write
to.** Read it before pressing Run.

**It refuses to run in a process where both environments are loaded** (C1). Not a
warning — a refusal.

### Quick Links — `QuickLaunch.ps1`
`Daily Scripts` · `ObjectTypes: []` · none · **Stage G** · ~354 lines

The operator's own set of links, as buttons. Each opens in the Edge profile it
was configured with, so admin portals open with the elevated account and everyday
links do not.

Links marked as admin render **dark blue**, so which is which is visible before
clicking. A Find box narrows a long list.

Reads `Config\QuickLaunch.json`. Edited by Quick Launch Tool (Stage H).

Opens links in a browser. **Connects to nothing and must not be greyed on any
connection state.** Nothing worth auditing.

`Start-Process -ArgumentList` does not quote for you — build the profile argument
carefully (A11).

---

## Reporting

Twelve buttons. **Nine of them produce a CSV; three are viewers.** The rules
below are stated separately for each, because applying the CSV rules to the
viewers produces requirements that cannot be met.

### The nine CSV reports

`Protocols`, `DisabledLicensed`, `StaleAccounts`, `DLInventory`,
`ExternalForwarding`, `PermissionsAudit`, `OrphanedOneDrive`, `ExternalSharing`,
`OwnerlessGroups`. All `ObjectTypes: []`, all read-only, all **Stage D**.

Shape: run a query, write a timestamped CSV to the reports folder, show results
in a grid, offer to open the file.

- **Use the shared report-file controls** — `Add-ReportFileControls` from `08`
  Tier 2 — rather than each tool inventing its own save path and naming.
  **`08` says to add it with the first report. This is that point.**
- **Slice large queries; never retry an oversized one** (C15). A query that
  failed for size fails again the same way.
- **`ResultSize` explicitly on every enumerating *Exchange* call** (C3). The
  default is 1000 and the short answer arrives silently. **`DLInventory` queries
  Active Directory, where `ResultSize` does not exist** — use `-ResultPageSize`
  / paging there instead, and do not bolt an Exchange parameter onto an AD
  cmdlet.
- **Never write a zero-row file.** An empty CSV and a failed run look identical
  in a folder listing a week later.
- **Label a partial run as partial**, in the file and in the audit row. A report
  covering 80% of mailboxes that does not say so is worse than no report, because
  it will be used as though it were complete (D6).
- `Target` in the audit row is the report name or output path.

### The three viewers

`ReportBrowser`, `M365Reports`, `LogViewer`. These run no Exchange query, write
no CSV, and the four rules above do not apply to them. Specified individually
below.

---

### Protocols — `Protocols.ps1`
needs EXO · **Stage D** · ~404 lines

Which client protocols are enabled per mailbox — POP, IMAP, ActiveSync, Outlook
on the web and the rest.

Security review material: **legacy protocols are the ones most used in
password-spray attacks**, and this shows where they are still on.

This is the tool that taught **C15** (slice the query) and **C13** (ask the
cmdlet what it declares rather than assuming the property set). It is in
`reference/` — **treat it as the reference implementation for every bulk query
in this section.**

### Disabled/Licensed — `DisabledLicensed.ps1`
needs EXO + AD · **Stage D** · ~345 lines

Accounts disabled in Active Directory but still in a licence group — **still
costing money.** Straightforward cleanup list, and usually longer than expected.

**The licence group name comes from config, never a literal.** Removing the
licence is a separate, deliberate step and is not offered here.

### MB w No Activity — `StaleAccounts.ps1`
needs EXO · **Stage D** · ~348 lines

Accounts enabled and licensed that have not signed in for a while. Two reasons to
care: a licence paid for and not used, and an account nobody is watching that
still works.

**Not every result is dead** — shared and service accounts appear here
legitimately. Present it as a list to **review**, not a list to disable, and say
so in the tool rather than only in the help.

### DL Inventory — `DLInventory.ps1`
needs AD · **variant: OnPremAD** · **Stage D** · ~404 lines

Inventories the mail-enabled groups in **on-premises Active Directory** for
migration planning: what exists, how big it is, who manages it, in one sortable
file.

> **The source carries `# TOOLKIT-VARIANT: Cloud` while reading on-premises AD.
> That marker is wrong in the source.** Write `OnPremAD` and note the correction
> in `UNVERIFIED.md` rather than propagating it.

### External Forwarding — `ExternalForwarding.ps1`
needs EXO · **Stage D** · ~405 lines

Mail leaving the organisation, from **both** places it can be configured:

1. **mailbox forwarding**, set by an admin or by the user
2. **inbox rules**, which users create themselves — and which are the ones
   nobody thinks to check

**The classic sign of a compromised mailbox is a quiet rule forwarding everything
to an outside address. This is the report that finds it.** Checking only the
first source is the common mistake, and it produces a clean report on a
compromised tenant.

Also the reference for the shared report controls and partial-run labelling. It
ships in `DEVIN-REFERENCE-CODE-2.md`; after Step 0 it is at
`reference/ExternalForwarding.ps1`.

### Permissions Audit — `PermissionsAudit.ps1`
needs EXO · **Stage D** · ~393 lines

Every mailbox delegation across the organisation — Full Access, Send As, Send on
Behalf.

Answers "who can get into what", **a question that only ever gets asked under
time pressure** — which is the argument for it being a saved report rather than
something assembled live. Access granted for a good reason years ago is still
access.

It reports; it removes nothing.

### Orphaned OneDrive — `OrphanedOneDrive.ps1`
needs SPO + AD · **Stage D** · ~361 lines

OneDrive sites whose owner is disabled or no longer in Active Directory.

When someone leaves, their OneDrive does not leave with them. It sits there
holding whatever they had, consuming storage, with nobody responsible for it.

**Read-only on purpose** — deciding what to do with the content is a
conversation, not a script.

### External Sharing — `ExternalSharing.ps1`
needs SPO · **Stage D** · ~436 lines

How each SharePoint site is configured for external sharing, **and** which guest
accounts actually have access.

**Both halves matter and neither is enough alone:** a site permitting external
sharing with no guests on it is a **setting**, while a guest with access to a
site is a **person who can read your files.** Report both and keep them visually
distinct.

### Ownerless Groups — `OwnerlessGroups.ps1`
needs EXO · **Stage D** · ~349 lines

Microsoft 365 groups with no effective owner.

An ownerless group cannot be managed by anyone but an admin — nobody can add
members, change settings or delete it — **and it happens quietly whenever the
last owner leaves. Nobody is notified.**

### Report Browser — `ReportBrowser.ps1`
`ObjectTypes: []` · none · **Stage D** · ~279 lines

Opens any CSV the reporting tools have written, without going to find it in
Explorer. Lists everything in the reports folder, **newest first, across every
report type**, so last month's can be re-read without re-running it.

**Depends on every report using the shared save path.** A report that invents its
own location is invisible here — which is the practical argument for
`Add-ReportFileControls`.

Reads the filesystem. Connects to nothing.

### M365 Reports — `M365Reports.ps1`
`ObjectTypes: []` · none · **Stage D** · ~304 lines

Buttons that open Microsoft 365 admin portal reports — usage, activity,
licensing, security — in the browser.

**These are reports Microsoft already computes from data a script cannot get at
directly. Rather than reproduce them badly, this puts them one click away.**

Opens each in the elevated Edge profile, so the operator lands signed in as the
right account.

**Deep links need no Graph consent** (C19), which is what makes this viable in a
tenant that will not consent to Graph (C20). **Reads nothing, changes nothing,
opens links.** Nothing to audit.

### Log Viewer — `LogViewer.ps1`
`ObjectTypes: []` · none · **Stage D** · ~518 lines

A read-only view of the audit log — every change the toolkit's tools have made,
who made it, when, and to whom. **The answer to "when did that happen and who did
it".**

Search and filter across everything.

**Never fold a row with no `Result` into `Success`.** Count it as `Unknown` and
show it as such. A log that overstates what worked is worse than one with a
visible gap.

> The source handles a legacy two-column log schema. **That schema has never
> existed at this tenant** — see §6 — so do not build conversion or dual-schema
> reading. The `Unknown` rule above still stands as defensive handling.

Nothing here changes anything. The log is the record; this reads it.

---

## Toolkit

Maintenance and configuration. Needed to *keep* and *hand on* the toolkit, not to
use it. All **Stage H** unless noted.

### Tenant Manager — `TenantManager.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~536 lines

Edits `Tenant.json` — the facts about this tenant: name, domains, which services
it has, its directories. Every other tool reads it, and it is **the only file in
the toolkit allowed to contain a domain or server name.**

**Why this tool exists:** capabilities decide which buttons appear at all, and an
absent `Tenant.json` means every capability reads as false. A new tenant opened a
console with the on-premises tools simply missing and nothing saying why. Every
other config file had an editor; **the one that decides what the console IS had
none.**

**Two modes, chosen automatically:**

| Condition | Mode |
|---|---|
| No `Tenant.json` | first-run setup, started from the shipped template |
| `Tenant.json` present | editor, showing what is there now |

**`Detect`** fills in what it can from the connections the window already has, so
most of the work is confirming rather than typing.

**CAPABILITY IS A CLAIM, NOT A PROBE.** Ticking Exchange Online says this tenant
*has* it, not that you are connected right now. Unticking something the tenant
genuinely does not have makes its tools **disappear** rather than sit greyed.

**Preserve keys it does not understand.** The source dropped `_comment_*` keys on
save — it carried the root `_comment` and lost `_comment_Capabilities`,
`_comment_Directories` and the rest. An editor that silently discards fields it
was not taught about will eventually delete a live setting. **Round-trip
unknown keys unchanged.**

**Also owns two keys other tools depend on and nothing else surfaces:**
`SharePoint.AdminUrl` (see Connect-SPO) and `FeatureRequest.Recipient`, which
Report a Bug and Feature Request share. Both are in `02`'s schema; both fail
quietly if this editor does not expose them.

### Tool Manager — `ToolManager.ps1`
`Toolkit` · `ObjectTypes: []` · none · **built in Phase 2 — additions only, no stage** · ~948 lines

Already required infrastructure in `00` §5. Listed here for **two additions the
finished toolkit needs:**

1. **The `Group` column** from §3, with its explanation.
2. **An over-capacity warning.** Report a section that has more buttons than
   `Rows × ButtonsPerRow` rather than letting the console silently drop the
   overflow — **and count buttons, not scripts.**

### Theme Manager — `ThemeManager.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~604 lines

The console's colours, and which theme each tenant gets.

**COLOUR IS NOT DECORATION HERE.** One console folder per tenant, all looking
identical, is how a change lands in the wrong tenant. A distinct theme per tenant
means the operator can tell at a glance which one is in front of them, **before
clicking anything.**

**Add shipped themes** brings in the bundled palettes. Pick one, Save, relaunch
the console.

**Tenant binding** ties a theme to a tenant so the right colours load
automatically. The console reads `Theme` from `Tenant.json`; this is where that
value comes from.

**CONTRAST MATTERS MORE THAN TASTE.** A palette where disabled text is nearly
the same as enabled text **hides the grey-out that tells you a tool cannot run.**
**The shipped themes are checked for that; hand-rolled ones are not** — say
exactly that in the help text. This is a stated limitation, not a feature: the
source tool does no contrast checking and you should not invent one.

The logo is cosmetic and optional. It is the one thing here that is.

**`Themes.json` is per-deployment, like the rest of `Config\`, and is NOT
shipped.** Note the tension with Deploy's rule that "everything else in `Config`
is tenant fact and does travel" — `Themes.json` is the documented exception, and
the source carries both statements. Resolve it in favour of not shipping it, and
record the decision in `UNVERIFIED.md`.

### Tool Status — `ToolStatus.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~725 lines

Lists every `.ps1` beside the console and says **what loads it:**

| State | Colour | Meaning |
|---|---|---|
| **AdminConsole** | green | a console button points at it |
| **Dashboard** | green | the dashboard can launch it |
| **Not loaded** | yellow | nothing opens it — safe to park |
| **Parked** | grey | already moved to `DisabledTools` |

**Yellow does not mean useless.** A script may be deliberately shelved, or called
by another script rather than opened by a button. **Read it before parking it** —
the `?` column, or a double-click, shows that tool's own help text without
leaving the window.

**`Common.ps1`, `AdminConsole.ps1`, `Publish-Toolkit.ps1` and this tool are not
listed.** They are infrastructure and must never be parked.

**Parking MOVES a file to `DisabledTools`.** It deletes nothing, and Restore puts
it back. **This deployment keeps only what it uses; the master copy keeps
everything.**

**Last run comes from the audit log**, so it is when the tool last did something
worth recording — **not when it was last opened.** Label the column so that is
obvious.

**After parking anything, re-run Publish-Toolkit.** The version manifest lists
files by name, so a parked script reads as *missing* until the toolkit is
re-stamped.

### Quick Launch Tool — `QuickLaunchTool.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~371 lines

Edits the links Quick Links shows.

- **Section** — any name; links group under it, and sections appear in the order
  they first appear in the list
- **Move Up / Move Down** — controls both section order and order within a
  section
- **Edge profile** — picked from the profiles that actually exist on the machine.
  Blank means Edge's Default profile
- **Admin** — makes the button dark blue, for links needing an elevated account
- **Test Selected** — opens the highlighted row exactly as Quick Links would, so
  a wrong profile is found here rather than in front of a customer

**Nothing is saved until Save is pressed.** Say so.

### Config Backup — `ConfigBackup.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~307 lines

Copies `Config\*.json` somewhere safe, and puts them back. These are the files
that are **the operator's** rather than the toolkit's: `Tools.json`,
`Settings.json`, `Sections.json`, `QuickLaunch.json`, `Themes.json`,
`Tenant.json`.

**Copying a new console folder over an old one replaces every one of them, and
nothing warns you.** The loss shows up later as blank paths and missing buttons,
long after the cause.

**RUN A BACKUP BEFORE COPYING A NEW CONSOLE OVER THIS ONE.** That is the whole
reason the tool exists, and it belongs at the top of its help text.

**Backups are written OUTSIDE the console folder, beside it**, so that replacing
or renaming the console cannot take the backup with it. **Each backup is a
timestamped folder** holding the files exactly as they were.

**Restore copies a backup back over `Config\` — and takes a fresh backup of
what is there first**, so restoring the wrong snapshot is not final either.

### Cert Check — `CertificateExpiry.ps1`
`Toolkit` · `ObjectTypes: []` · needs EMS · read-only · **Stage H** · ~459 lines

Certificates on every on-premises Exchange server and how long each has left.
Worth running monthly.

**An expired certificate breaks hybrid quietly.** The first symptom is usually
somebody reporting that free/busy has stopped working between on-premises and the
cloud, **which does not look like a certificate problem** — hence the tool.

**Separate certificates in use from ones merely present**, since only the first
kind breaks something.

**Takes `-Environment` as a parameter and must USE it.** The source declared the
parameter and then passed a hardcoded value at the call site, so the parameter did
nothing. If you accept a parameter, use it or do not declare it.

### Tester — `Test-Toolkit.ps1`
`Toolkit` · `ObjectTypes: []` · none · **built in Phase 0 — additions only, no stage** · ~1,950 lines

The verifier, given a console button so it can be run without a shell.

By the finished state it carries **every rule in `03-SELF-TEST-SPEC.md`** —
including `PS7002`, `WF002` (control overlap) and `LOG001` (a write with no audit
call), all of which are already specified there. Do not treat those as extra
rules to invent.

**Three additions belong to this catalogue, and all three are needed BEFORE
Stage A:**

| Rule | What |
|---|---|
| `WF010` | Console height exceeds `WorkingArea` — §2. **Not `WF004`, which is already taken** and asks a different question (which API you call, not whether the result fits). |
| `WF011` | A section holds more buttons than `Rows × ButtonsPerRow`, a group counting as one — §3. |
| `LOG002` exemption | The rule must not flag a literal `-Target 'Toolkit'` — §4. The reference verifier already behaves this way; **check before changing anything.** |

> **The reference verifier uses its own rule codes** and several collide with
> `03`'s. See the note at the top of `03`: run and extend the reference, do not
> renumber it, and take the next free number in a family rather than reusing one.

### Report a Bug — `BugFix.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~449 lines · *sends one email; changes nothing else*

Reports something broken.

**One send. The operator is the `To:`, and the real destination is `Bcc:`.** That
is how the reporter gets their own copy without a second message, and it is why
there is no direct-reply path — **requests arrive through these buttons, not as
ad-hoc mail.**

**THE FIELD THAT MATTERS MOST IS "Steps to reproduce".** A bug that can be
reproduced is usually fixed the same day; one that cannot may never be. Even "it
happens on shared mailboxes but not user ones" halves the search. **Weight the
form accordingly** — that field must not look like one of several equal boxes.

Also ask for:

- **The exact error text, pasted rather than described.** "It said something
  about a domain controller" and the actual `0x52` message are different amounts
  of information.
- **What you EXPECTED.** Half of all bug reports turn out to be the tool doing
  what it was built to do, and knowing what the reporter expected instead is what
  turns that into a real answer rather than "working as designed".
- **Whether it happens every time.** An intermittent fault and a constant one are
  chased completely differently.

**"Blocking me right now" is a checkbox**, and it is the only thing on the form
that jumps the queue.

"You do not need to know what caused it. Report it anyway."

### Feature Request — `FeatureRequest.ps1`
`Toolkit` · `ObjectTypes: []` · none · **Stage H** · ~363 lines · *sends one email; changes nothing else*

Asks for something the console does not do yet. Three boxes, Send, and the same
`To:` operator / `Bcc:` destination pattern.

**The frequency question decides what gets built first.** "Twice a day" and "once
a quarter" are both fine answers — a small annoyance that happens constantly is
usually worth more than a large one that almost never does, and there is no way
to know which is which unless the requester says.

> **Both tools share ONE recipient key: `FeatureRequest.Recipient` in
> `Config\Tenant.json`.** Do not add a separate `BugFix.Recipient` — a second
> thing to configure is a second thing to get wrong, and the source deliberately
> collapsed them.
>
> **Scrub that address out of any error surfaced to the operator.** A failed send
> should not print the destination into a window the operator may screenshot.

---

## Infrastructure — no console button

### Deploy — `Deploy.ps1`
**Stage H** · ~633 lines · run from a shell

Builds a handout copy of the toolkit and zips it. **One build; everyone gets the
same one.**

Creates a staging folder beside the current one —
`adminconsole3.1_2026-08-13_0930`, version and timestamp — so what shipped and
when is visible at a glance, and an older build is still there to go back to.
Zips it and drops a `README.txt` inside for the recipient.

**No name prompt.** An earlier version asked who the copy was for and baked the
name into the folder. Wrong twice over: the build is identical whoever receives
it, and a name baked into a folder is stale the moment it lands on their machine.
**The timestamp is a BUILD label, not an install path.**

**Excluded from the handout, and why:**

| Excluded | Reason |
|---|---|
| `Logs`, `Reports`, `RetryQueue` | your history, not theirs — they start empty |
| `DisabledTools` | tools deliberately out of service *here* |
| `Config\Settings.json` | per-engineer paths and email |
| `Config\Themes.json` | **per-deployment, like `Settings.json`.** Everything else in `Config` — `Tenant.json`, `Tools.json`, `Sections.json`, `QuickLaunch.json` — is tenant or layout fact and does travel |
| `Publish-Toolkit.ps1` | stamping is the master copy's job |
| **`Deploy.ps1`** | **copies do not make copies** |
| `*Cache*.json` | stale scan data is worse than none — theirs rebuilds on first use |

**`UnBlock.ps1` IS included** — it is the first thing the recipient needs,
because files out of a downloaded zip arrive blocked by Windows.

**NOTHING IS EVER OVERWRITTEN.** Every run makes a new timestamped folder, so a
build can be reissued, compared, or gone back to. Old ones are deleted by hand.

**The manifest (`Version.json`) is REBUILT for the copy, not copied** — otherwise
the recipient's self-test reports the excluded files as missing **forever.**

### Publish Toolkit — `Publish-Toolkit.ps1`
**Stage H** · ~605 lines · run from a shell · **the last thing before handing the toolkit on**

**RUN THIS ON THE MASTER COPY ONLY.** On someone else's copy it re-stamps their
edits as official, which is exactly backwards. **This is the only thing that
writes `Version.json`.**

In order:

1. **UNBLOCK every file.** Anything downloaded or copied from another machine
   carries a mark-of-the-web tag that makes PowerShell refuse to run it. The
   files actually blocked are listed — that list should match what was just
   downloaded, and anything surprising on it is worth a look.
2. **SHOW what changed** since the last stamp.
3. **CONFIRM WHICH TENANT** this build is for, and **refuse to continue unless
   `Config\Tenant.json` already says the same thing.**

   > This is the one thing nothing else can catch. `Config\` is excluded from the
   > version manifest, and `Test-Toolkit` reads the tenant names *out of*
   > `Tenant.json` — so a toolkit built against one tenant and published for
   > another would ship the first one's domains and servers **past every other
   > guard, and work perfectly, at the wrong company.**

4. **LIST tools this tenant can never run**, so they can be removed before
   shipping rather than greyed out on someone else's screen.
5. **ASK minor or major** and derive the number from the last one — 1.03 becomes
   1.04, or 2.0. Nothing to type, so nothing to mistype.
6. **STAMP `Version.json`** with the version, the tenant, and a SHA-256 per
   script.

   **Unblocking happens BEFORE hashing, on purpose.** Whatever clearing the tag
   does or does not do to a file, the manifest then describes the bytes exactly
   as they will be handed over.

---

## 6. Deliberately excluded

Do not build these. If you think one is needed, ask first.

| Script | Why not |
|---|---|
| `PimActivation.ps1` | **Programmatic PIM self-activation is impossible in this tenant.** PIM for Groups is a Graph-only API, and the tenant will not consent to `PrivilegedEligibilitySchedule.Read.AzureADGroup` / `PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup` (both admin-consent delegated permissions). The retired AzureADPreview cmdlets covered directory *roles*, not groups. **`activate.ps1`, which opens the PIM portal, is the supported route and is already in scope** (C19/C20). |
| `Convert-LogHistory.ps1` | A one-time migration from a two-column audit-log schema **that has never existed at this tenant.** Nothing to convert. |
| `UnifiedGroupMembership.ps1` | Superseded by `UnifiedGroup.ps1`, which covers the same ground. Building both produces two buttons that disagree. |
| `Settings.ps1` | A 52-line shim. Settings are edited through Tool Manager. |
| Anything Graph-based | C20. Consent is the blocker, not capability. Deep links (C19) are the workaround and need no consent. |

**`UnBlock.ps1`** (11 lines — note the capital B in the filename) is a helper
shipped *inside* handouts, not a tool. Include it in Deploy's output; it needs no
button, no help block and no spec.

---

## 7. Definition of done — the whole toolkit

Beyond the per-script definition in `00-BUILD-PROMPT.md` §10:

- [ ] Every button in `Tools.json` points at a script that exists
- [ ] Every `.ps1` beside the console is one of: **registered in `Tools.json`**;
      **launchable from Dashboard**; **opened by another tool**
      (`Group-Manager-Editor.ps1` is the one — Group Manager routes to it and it
      refuses to open standalone); **infrastructure** (`Common.ps1`,
      `AdminConsole.ps1`, `Publish-Toolkit.ps1`, `Deploy.ps1`, `UnBlock.ps1`,
      `Test-Toolkit.ps1`); or **listed in `UNVERIFIED.md`** with a reason
- [ ] **Tool Status shows no unexplained yellow.** It reports "nothing loads
      this — safe to park" for anything not on a button, which is correct for
      `Group-Manager-Editor.ps1`, `Deploy.ps1` and `UnBlock.ps1` and **wrong to
      act on**. Parking the group editor silently breaks Group Manager
- [ ] `Sections.json` capacity ≥ button count **per section**, counting a group
      as **one** button
- [ ] Total console height fits `WorkingArea` on a 1920×1080 screen with a
      taskbar
- [ ] The verifier passes, **and can still be made to fail** on a planted defect
- [ ] Every tool that writes has an audit row on success **and** on failure
- [ ] Every audit `Target` is a primary SMTP address, or a documented scope name
      for the tools with no mailbox in play (§4)
- [ ] `UNVERIFIED.md` lists every cmdlet used that could not be confirmed against
      Microsoft Learn, and every assumption made about tenant shape
- [ ] No tenant name, domain or server in any `.ps1` — `TEN001`/`TEN002` clean

**A toolkit where fifty buttons work and five are honestly listed as unbuilt is a
success. A toolkit where fifty-five buttons exist and five of them quietly do the
wrong thing is not, and is considerably harder to recover from.**
