# CloudOnly/ — the cloud-only toolkit

Forked from `FH-Agnostic/` on 2026-08-03 for SRG and ANZ, which have no
on-premises directory at all.

Cross-cutting context lives in `ARCHITECTURE.md` §23. This file is the per-script
detail for the eight scripts that were actually rewritten. Every other file in
`CloudOnly/` behaves exactly like its `FH-Agnostic/` namesake — its existing
notes file still applies — except that it no longer takes `-DomainController`.

---

## Why a fork rather than a flag

`FH-Agnostic/` was written cloud-first and AD-where-it-helps, which is the
right shape for a hybrid tenant. Over time seven of its 32 scripts grew a real AD
dependency: both attribute editors on `Get-ADUser`, `DLInventory` on a bulk read
of the `member` attribute, `DLMemberships` on the `memberOf` backlink,
`GroupEditor` on `Set-ADGroup`, `DL_SendToAs` on `Get-ADPermission`, and
`Common.ps1` on `Find-ADObjectBySmtp` underneath all of them.

None of that works where there is no forest. The failure is not an error either —
the window opens, the grid is empty, and nothing on screen says why. That is the
worst available outcome: it looks like the tenant has no data rather than like
the tool cannot ask the question.

A runtime `if (Test-AdAvailable)` was the obvious alternative and is worse. The
AD branch would never execute in the environment this build exists for, so it
would never be exercised, never be noticed when it broke, and would sit there
implying a capability that is not present. Removing the path means the absence is
enforced rather than intended — and it is checked mechanically on every
verification pass.

## The absence is a test, not a convention

The verification sweep fails the build on any of these appearing in code (as
opposed to a comment):

```
*-AD*            Get-ADServerParam    Find-ADUserBySam
Find-ADObjectBySmtp   Get-DomainController   Test-AdAvailable
-DomainController
```

`Import-Module ActiveDirectory` too. The comment exemption matters — several
functions explain what their on-premises equivalent *was*, and that is worth
keeping, so the check strips comments and strings before it looks.

## `Resolve-ToolTarget` fails loudly on `-DomainController`

The first cut kept the parameter and ignored it, on the reasoning that ten-odd
callers passed it and a shared launcher should not break.

That was wrong. Silently accepting it means a hybrid console can launch these
scripts, get cloud-only answers for hybrid objects, and never be told. A binding
error on an unexpected parameter is the correct outcome: it says "this is the
wrong toolkit for this environment" at the only moment anybody would act on it.
So the parameter is gone from `Resolve-ToolTarget` and from all twenty-one
scripts that passed it through.

## The capability map lost `EMS` and `AD`

Two of the four probed capabilities described things this build cannot reach.
Leaving them in would grey tools out for a condition no one on the target machine
can satisfy — an outage dressed as a feature. `Test-EmsAvailable` and the `EMS`
CapabilityDef went with them, so the map now probes exactly the two dependencies
that exist: Exchange Online and SharePoint Online.

---

## `Common.ps1`

`Resolve-DirectoryAccountBySmtp` replaces `Find-ADObjectBySmtp`. The one fact
callers actually wanted from AD was whether the account is disabled, and
`Get-User` carries `AccountDisabled` in the cloud — so the shape of the answer is
unchanged and only the source moved.

`Resolve-Target` gained a **third** attempt: Mailbox → UnifiedGroup →
DistributionGroup. Dropping the on-premises group branch without adding this left
a hole, because a plain cloud distribution group answers to neither of the first
two and would have resolved as `NotFound`. `Get-DistributionGroup` is tried
**last** deliberately: it also returns mail-enabled security groups and would
happily claim objects that `Get-Mailbox` or `Get-UnifiedGroup` describe better.

`Get-MailDomain` lost its AD fallback. The remaining chain is Settings →
`$script:MailDomain` → the tenant's primary accepted domain. The AD DNS domain
was never a good guess anyway — a forest called `corp.example.net` routinely
sends as `example.com` — and here it does not exist to guess from.

## `DLMemberships.ps1`

Exchange Online has no reverse-membership cmdlet. `memberOf` was a backlink the
directory maintained; the cloud simply does not offer the equivalent, and no
amount of filtering produces one.

So this is a **scan**: enumerate every distribution group, call
`Get-DistributionGroupMember` on each, test for the address. The window says so
in as many words, because a tool that takes four minutes without explaining why
reads as hung.

**Cancel stops between groups, never inside one.** Partial results from a whole
number of groups are true — "these are memberships, and there may be more".
Stopping mid-group would produce a membership list that is wrong in a way that
looks exactly like a right one.

`Scanned` reports how many groups were actually checked, not how many exist, so a
cancelled run says what it covered.

## `GroupEditor.ps1`

Handles two group kinds now — Microsoft 365 and distribution — where it used to
handle Microsoft 365 and on-premises-synced. `Get-GroupState` normalises both
into one shape rather than letting every caller branch, because two display paths
drift apart and then disagree about the same group.

`Set-DistributionGroup -ManagedBy` **replaces the whole owner list**. There is no
remove-one-owner operation, so Remove Owner throws with that explanation instead
of appearing to work and clearing every other owner. Microsoft 365 groups keep
their `Remove-UnifiedGroupLinks` path, which does remove one.

A directory-synced group is flagged in the state display before you try to change
it. Exchange Online will refuse the write, and "mastered somewhere else" is a
different problem from "this tool is broken".

## `DL_SendToAs.ps1`

Send As moved from `Get-/Add-/Remove-ADPermission -ExtendedRights Send-As` to
`Get-/Add-/Remove-RecipientPermission -AccessRights SendAs`. `NT AUTHORITY\SELF`
appears on every object and is the platform's own grant, so it is filtered rather
than reported as somebody's delegation. **Deny entries are kept and flagged** — a
Deny outranks an Allow, so hiding one would mean reporting access that does not
actually work.

Send-To keys on **primary SMTP address**, not distinguished name. Cloud DLs
return `AcceptMessagesOnlyFromSendersOrMembers` as names rather than DNs, so
every entry is re-resolved; the address is the only stable handle to compare on,
and comparing on anything else lets a set test miss a match that is really there
and add a duplicate.

An entry that will not re-resolve is **carried through** the rewrite rather than
dropped. It is usually a deleted recipient still on the list, and silently
removing it during an unrelated change is a modification nobody asked for.

The whole `Resolve-DLEnvironment` domain-controller apparatus is gone; it now
just confirms the DL exists, so a typo produces one clear message instead of two
half-finished Send-To / Send-As passes.

## `MailAttributes.ps1`

Reads `Get-User` for the person fields and `Get-Mailbox` for the mail fields,
because neither carries all of them. Both are keyed on the same address, so a
mailbox that resolves in one and not the other is a real fault and is surfaced
rather than hidden behind a half-populated grid.

The name mapping:

```
displayName → DisplayName      mailNickname      → Alias
givenName   → FirstName        msExchArchiveName → ArchiveName
sn          → LastName         proxyAddresses    → EmailAddresses
mail        → WindowsEmailAddress
```

**`Set-Mailbox -PrimarySmtpAddress` does not exist in Exchange Online.** Verified
against Microsoft Learn on 2026-08-03: the parameter is documented "available
only in on-premises Exchange". The supported cloud route is to rewrite the whole
`EmailAddresses` collection with exactly one uppercase `SMTP:` entry — which is
exactly what `Set-PrimarySmtpInProxyAddresses` already produced, so that helper
and the Keep-or-Remove-old-primary dialog ported across untouched and only the
write cmdlet changed.

There is no `-Clear` in the cloud. A blank is written as `$null`. `Alias` and
`PrimarySmtpAddress` refuse a blank outright — an aliasless or addressless
mailbox is broken, not empty.

Three things became read-only rather than disappearing:

- **`targetAddress`** is dropped entirely. It was the on-premises pointer *to*
  the cloud mailbox; here the mailbox is the destination and there is nothing to
  point at. `Get-MailUser -ExternalEmailAddress` is the equivalent for mail
  users, which is a different object and a different tool.
- **`Name`.** `Set-User -Name` renames the mailbox, not an RDN — a different
  operation with different consequences, so it is not quietly wired to the old
  edit path under the old label.
- **`mail` / `WindowsEmailAddress`.** `Set-Mailbox -WindowsEmailAddress` has no
  effect while the mailbox is subject to an email address policy, which is the
  normal state. An edit button that silently does nothing is worse than no
  button.

`ArchiveName` carries a warning of its own: it names an archive that already
exists and does not create one, and Outlook desktop ignores the value entirely
(it shows `Online Archive - <primary address>` regardless). Outlook on the web is
where it is visible.

## `AllAttributes.ps1`

`Get-User` properties appear plain; `Get-Mailbox` properties are prefixed
`Mailbox.` and are always read-only. The prefix is provenance: a merged list with
no source would leave the operator guessing which cmdlet changes what, and
`Set-Mailbox` work belongs in `MailAttributes.ps1` where the consequences are
spelled out.

The editable set is exactly the `Set-User` parameters that describe a *person* —
contact details and organisation placement. Nothing structural: `Name`,
`UserPrincipalName`, `SamAccountName`, `DisplayName`, `FirstName`, `LastName` and
`Manager` all change how the object is identified or routed, and a grid of two
hundred rows is the wrong place to do that by accident.

Four attributes just changed name — `State` → `StateOrProvince`, `POBox` →
`PostOfficeBox`, `OfficePhone` → `Phone`, `HomePage` → `WebPage`. Five have no
`Set-User` equivalent and are **read-only rather than missing**, so their values
stay visible: `Description`, `Division`, `EmployeeID`, `EmployeeNumber`,
`Organization`.

`Set-AllAttribsAttribute` re-checks the editable list before writing. The grid
already gates the button, but a function that will splat any property name onto
`Set-User` is one refactor away from being a way to set anything.

## `DLInventory.ps1`

Two cmdlets, because the cloud splits what AD kept in one place.
`Get-DistributionGroup` returns classic distribution groups and mail-enabled
security groups; `Get-UnifiedGroup` returns Microsoft 365 groups, which
`Get-DistributionGroup` does not see **at all**. Leaving the second one out would
under-report the estate and quietly make the migration plan wrong.

Membership costs one call per group. There is no cloud equivalent of reading
`member` off a bulk directory query, so this is a scan and is labelled as one on
the window.

New column: **`IsDirSynced`**. It is the most useful thing this report can tell a
cloud-only tenant. A synced group is mastered on-premises, cannot be edited here,
and — more to the point — **in a genuinely cloud-only tenant the count should be
zero**. Anything else means a directory sync is still running somewhere, which is
a finding rather than a rounding error.

`GroupType` distinguishes Distribution / MailSecurity / Microsoft365, and the
summary groups on it, because "how many of each" is the first question anybody
asks of this report.

Dropped: `GroupScope` and the AD `info` attribute, neither of which the cloud
exposes. `Description` covers the notes field for both group kinds.

The staleness caveat is unchanged and still the most important line in the
summary: `DaysSinceChange` is when the **object** last changed, not when the
group was last **emailed**. Exchange holds no last-used data and message trace
reaches back about ten days, so a group can look stale and still be in daily use.

## `AdminConsole.ps1`

The lookup used to run an AD probe first and then resolve the type. Now Exchange
Online *is* the directory, so existence and type are one question answered by one
chain instead of two that could disagree with each other.

`NotFound` is still split into two messages, because "the tenant refused" and
"that address does not exist" send you to different places. The test is the same
one `Resolve-ToolTarget` uses: an error mentioning `couldn't be found` is
absence; anything else is refusal.

`OnpremGroup` left `$script:KnownObjectTypes`, so no tool can be gated on a type
that can never be detected.
