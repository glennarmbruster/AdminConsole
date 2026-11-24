# Graph Backlog

Things that need Microsoft Graph, have been recommended, and are **not built**.

Started 2026-09-19. This is a running list, not a plan — nothing here is
committed to. Add to it whenever a Graph-shaped answer comes up and gets
deferred, so the decision to build is made once, deliberately, against the
whole list rather than one tool at a time.

**Standing position:** no new Graph tools yet (Glenn, 2026-09-19). The
capability system already hides Graph tools in tenants that do not declare
Graph, so building against it does not disturb a tenant that has not
consented.

---

## 1. Already built — do not re-recommend

These exist and declare their own scopes in a `TOOLKIT-GRAPH-SCOPES:` header
line, which `Connect-MSGraph.ps1` reads and asks for as a union.

| Tool | Scopes it declares |
|---|---|
| `SignInActivity.ps1` | `AuditLog.Read.All`, `User.Read.All` |
| `RiskyUsers.ps1` | `IdentityRiskyUser.Read.All`, `IdentityRiskEvent.Read.All` |
| `AuthMethods.ps1` | `UserAuthenticationMethod.ReadWrite.All`, `User.Read.All` |
| `LicenseAssignment.ps1` | `LicenseAssignment.ReadWrite.All`, `Organization.Read.All`, `User.ReadWrite.All`, `Group.ReadWrite.All` |
| `ServiceHealth.ps1` | `ServiceHealth.Read.All`, `ServiceMessage.Read.All` |
| `DisabledLicensed.ps1` | *(no scopes declared — worth checking)* |
| `Connect-MSGraph.ps1` | the connector itself |

---

## 2. Recommended and not built

### 2.1 Find a message — where did it go?

**Raised** 2026-09-19, from "Bob says Sally sent him an email and he doesn't
have it. Message trace says it was delivered. Where is it now?"

Message trace answers *did it arrive*. It does not answer *which folder is it
in now*. The Inbox Rules tab only lets you infer — "a rule matching this
would have moved it to X" is a guess, not an observation.

- **Call:** `GET /users/{upn}/messages?$search="..."`, then resolve each hit's
  `parentFolderId` to a folder name.
- **Permission:** `Mail.ReadBasic.All` (application). Learn marks it *least
  privileged* for this call. Its description, verbatim: "Includes all
  properties except body, previewBody, attachments and any extended
  properties." So the tool structurally cannot read message content.
- **Admin consent:** required.
- **Open question before building:** Learn attaches an application-access-policy
  note to `Mail.Read` and `Mail.ReadWrite` ("Administrators can configure
  application access policy to limit app access to specific mailboxes") but
  attaches **no such note** to `Mail.ReadBasic.All`. That is an absence of a
  statement, not a statement of absence — but it means the safer-looking
  permission is the one whose scoping story is undocumented. Test against
  gcearth before trusting it tenant-wide.
- **Non-Graph alternative worth checking first:** the mailbox audit log.
  Move, MoveToDeletedItems, SoftDelete and HardDelete are owner actions and
  the records carry folder paths. If that holds up it answers a *better*
  question — not just where it is, but how it got there and when — with no new
  permission and no message content. Record shape and retention window are
  unverified.

### 2.2 Person Editor — sign-in state, lockout, password expiry

**Raised** 2026-09-19, while making the Person Editor work in a cloud-only
tenant.

Checked the full `Set-User` parameter list: there is **nothing** for account
enabled/disabled, lockout, or password expiry. Those live in Entra. In a
cloud-only tenant the Account tab therefore hides Unlock / Enable / Disable
entirely and offers only "must change password at next sign-in", which
`Set-User -ResetPasswordOnNextLogon` does carry (Applicable: Exchange Online).

- **Needs:** Graph user account state — read and write.
- **Blocks:** no starter/leaver process is possible for a cloud-only tenant
  without this.

### 2.3 Person Editor — Groups tab

**Raised** 2026-09-19.

The tab currently reads nothing. Its note says so plainly and points at the
Group Editor for mail-enabled groups in the meantime.

- Mail-enabled membership is reachable **without Graph** via
  `Get-DistributionGroupMember` / `Add-DistributionGroupMember` — that part
  could be built today and is tracked as ordinary work, not here.
- **Entra-only group membership needs Graph** and is the part that belongs on
  this list.

### 2.4 Person Editor — Licence tab

**Raised** 2026-09-19.

Reports "not available" with a reason in a cloud-only tenant. Previously said
**"NO"**, which read as *unlicensed* when it meant *cannot see* — fixed, but
the tab is still report-only and has no writer in any mode.

Note `LicenseAssignment.ps1` already exists and holds Graph licence logic, so
this may be absorbing an existing tool rather than new work.

### 2.5 Entra tab (Mailbox / Person editors)

**Raised** before 2026-09-19; long-standing deferral. No design yet.

### 2.6 Soft-deleted mailbox restore

**Raised** before 2026-09-19.

`Get-MgDirectoryDeletedItemAsUser` / `Restore-MgDirectoryDeletedItem`. Noted
as a good practical test of the Graph connection after the module version
mess below, because it is a small, self-contained, high-value operation.

### 2.7 Shared mailbox cloud restoration — hybrid retirement Phase 4

**Raised** before 2026-09-19, in the Exchange hybrid retirement plan.

Phase 4 of the retirement project depends on recreating shared mailboxes in
the cloud. Listed here so the retirement plan's Graph dependency is visible
from this side too.

---

## 3. Permissions this list would add, if all of it were built

| Permission | Type | Needed by | Admin consent |
|---|---|---|---|
| `Mail.ReadBasic.All` | Application | 2.1 Find a message | Yes |
| User account read/write (exact scope TBD) | Application | 2.2 Account state | Yes |
| `Group.Read.All` or `GroupMember.Read.All` | Application | 2.3 Groups tab | Yes |
| Directory object restore (exact scope TBD) | Application | 2.6 Soft-deleted restore | Yes |

Nothing here has been verified as the *minimum* scope except
`Mail.ReadBasic.All`. Each should be checked against the "least privileged"
row on its own API page before anyone is asked to consent — Learn puts that
designation on the **API method pages**, not on the permissions reference.

---

## 4. Known gotchas

- **Module versions.** 2.25.0 is the version in use. 2.38.1 caused a
  `TypeLoadException` on `GetTokenAsync` and was uninstalled; 2.29.0 was also
  installed at one point. Three versions side by side is what caused it. A
  successful connection after the uninstall was never confirmed — confirm it
  before building anything new.
- **App registrations.** Three exist for this tenant, with narrow single
  scopes each (Teams create, mail read/write, user read/write). Their client
  IDs are in the project notes, deliberately not copied here. A new tool
  should say which registration it expects rather than assuming one.
- **Least-privilege designations live on API pages.** The permissions
  reference itself makes no least-privilege claims. Always read the specific
  method page.
- **Capability gating already works.** `Get-ScriptCapabilityRequirement` reads
  a script's cmdlets, and the `Graph` capability hides tools in tenants that
  do not declare it. A new Graph tool will be invisible in gcearth-style
  tenants until Graph is ticked in Tenant Manager — no extra work needed to
  keep it out of the way.
