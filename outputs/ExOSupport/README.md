# ExO Support

A narrow, cloud-only console for support staff who manage **shared mailboxes** and **distribution groups** once they live in Exchange Online. Separate from the Admin Console; it shares no files with it.

## What's in the folder

| File | Purpose |
|---|---|
| `ExOSupport.cmd` | Double-click to start |
| `ExOSupport.ps1` | The console window: Activate, Connect ExO, Cloud Lookup, status strip |
| `ExOSupport.Common.ps1` | Config, tenant, connection, authorization and visibility, PIM, logging |
| `ExOSupport.Lookup.ps1` | Cloud Lookup panel in the main window |
| `ExOSupport.Editors.ps1` | Shared Mailbox Editor and Group Editor |
| `Config\ExOSupport.json` | Which tenant, its primary domain, authorization group, PIM |
| `Logs\` | Created on first change: `ExOSupport-yyyy-MM.csv` |

The black PowerShell window that `ExOSupport.cmd` opens minimises itself once the console is up, and comes back only while you sign in. Closing it closes the console.

Requirements: Windows PowerShell 5.1 and the `ExchangeOnlineManagement` module (`Install-Module ExchangeOnlineManagement -Scope CurrentUser`). Nothing else.

## Operator flow

One window. On launch it asks for **your email address** (your normal mailbox, not your admin account), pre-filled with last time's. Membership of the authorization group is checked against that address, and both the address and the admin sign-in go into every log line. Use **Change my email** (top right) to correct it.

- **Top:** **Activate (PIM)** opens the PIM portal. Activate there, then press OK in the console: that records the activation time. Activate and Connect are independent. **Connect ExO** closes the console, signs in from the PowerShell window behind it, and reopens the console where it was, re-running your last search.
- **Middle, with the focus:** **Cloud Lookup.** Type a name, address or part of either, then press Enter. `*` on its own lists everything you can see. Select a row and click **Shared Mailbox Editor** or **Group Editor**, or double-click the row.
- **Bottom status strip:** PIM activation time, time left and when it ends (amber under 30 minutes, red when expired), and who is working: your email and the admin account signed in to Exchange Online. The top bar shows only whether you are connected.

The lookup returns only shared mailboxes, distribution groups and mail-enabled security groups that are **managed in the cloud**. Directory-synced objects are always hidden. User and resource mailboxes, Microsoft 365 groups and dynamic groups are never returned.

## Editors

- **Shared mailbox:** Add Full Access, or Full Access + Send As. Automapping is always on. Remove selected removes whatever those people have.
- **Group** (distribution or mail-enabled security): Add members, Remove selected.
- Every typed line must resolve to exactly one recipient before anything is written. A bare alias gets the tenant's primary domain.
- The object is re-checked (cloud-managed, right type, its attribute value) when the editor opens and before every write.

## Config

```json
"ActiveTenant": "ExpressPros"
```

`ActiveTenant` picks the tenant block. Leave it blank with several tenants listed and the console asks at launch (1, 2, ...). **To add a tenant**, copy the tenant block and change `Id`, `DisplayName`, `PrimaryDomain`, and the authorization group. Optionally set `TenantId` (GUID): the console then refuses a sign-in to any other tenant.

### Access (one authorization group)

```json
"Scope": {
  "Enabled": false,
  "Attribute": "CustomAttribute10",
  "Group": "ExOSupport-HQ@...",
  "Value": "HQ"
}
```

- With filtering on, the operator's email must be a member of `Group` (a **mail-enabled** group; nested groups are followed, 5 levels). A member sees and edits every cloud object whose `Attribute` equals `Value`. A non-member sees and edits nothing.
- **Replace the placeholder group address** before turning `Enabled` on.
- **To lock filtering on for delivery:** in `ExOSupport.Common.ps1`, set `$script:ExsScopeFilterLocked = $true`. The config switch is then ignored.
- Don't give the authorization group itself the `Attribute` value, or operators could edit their own access.

### Setting the attribute on cloud objects

Cloud-only objects are stamped in Exchange Online PowerShell:

```powershell
Set-Mailbox           -Identity sharedbox@domain -CustomAttribute10 HQ    # shared mailboxes
Set-DistributionGroup -Identity group@domain     -CustomAttribute10 HQ    # distribution + mail-enabled security groups
```

Objects still synced from on-premises must be stamped on-premises (`Set-RemoteMailbox` / `Set-DistributionGroup -CustomAttribute10`, i.e. `extensionAttribute10` in AD) and then synced. Exchange Online rejects the change on a synced object.

## Exchange permissions (set up separately)

The console filter is convenience. Enforcement is Exchange RBAC.

- Operators need to read recipients, manage mailbox permissions (Full Access, Send As) and manage group membership.
- **Group membership as a non-owner** needs `-BypassSecurityGroupManagerCheck`, which requires the **Security Group Creation and Membership** role (Microsoft Learn, Add-DistributionGroupMember). Without it, an operator can only change groups they own. The console uses the switch when the session offers it.
- Scoped enforcement: a management scope with `-RecipientRestrictionFilter "CustomAttribute10 -eq 'HQ'"`, applied with `-CustomRecipientWriteScope` on the role assignments for the HQ role group.
- **PIM caution:** if operators PIM into an org-wide Entra role (for example Exchange Recipient Administrator), that role is not scoped by the attribute and overrides the role-group scoping. For enforced scoping, the activated access should be the scoped Exchange role group (for example through a PIM-for-Groups group that is a member of it), not the org-wide role.

## Log

One CSV per month in `Logs\` (`ExOSupport-v2-yyyy-MM.csv`). Columns: WhenLocal, WhenUtc, Operator (your email), AdminAccount (Exchange sign-in), WindowsUser, Machine, Tenant, Action, Target, Result, Detail. Point the folder at a share later for one combined trail.
