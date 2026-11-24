# Server names -- what used to be hardcoded, and where it lives now

**This file is a deliberate exception to the rule.** Everywhere else in this
toolkit, a server, domain or organisation name in anything other than
`Config\Tenant.json` or `Config\Settings.json` is a bug that `Test-Toolkit`
fails the build for. This file names servers on purpose, because the point of
it is to be a record of what was removed from the code on 2026-09-12 -- and a
record with the names taken out would be no record at all.

**It ships to this tenant only. Do not copy it to the reference toolkit**, and
do not copy anything out of it back into a `.ps1`.

---

## The four

| Host | Where it lives now |
|---|---|
| `dc3pmexapp01` | `Tenant.json` -> `OnPremExchange.Servers` |
| `dc3pmexapp02` | same |
| `dc4pmexapp01` | same |
| `dc4pmexapp02` | same |

`Tenant.json` -> `OnPremExchange.ServerUri` is `dc3pmexapp01.ftbco.ftn.com`
and was already there; it is what the EMS connection targets and is separate
from the list above.

## Excluded, and already in Tenant.json before today

`OnPremExchange.ExcludedServers`: `DC3pmexapp03`, `dc3pmexapp04`,
`dc4pmexapp03`, `dc4pmexapp04`. These are subtracted from whichever list the
toolkit ends up with, including one it discovers for itself.

---

## Where the names used to be

Five copies of the same fact, written as literals in the scripts themselves:

| File | Form |
|---|---|
| `Clear-IISLogs.ps1` | `$script:Servers = @(...)`, lower case |
| `Get-Queues.ps1` | `$script:Servers = @(...)`, lower case |
| `Exchange-HealthCheck.ps1` | `$script:Servers = @(...)`, UPPER case, one line |
| `onprem-mailtest.ps1` | `$script:Servers = @(...)`, UPPER case, one line |
| `Unified-Trace.ps1` | the list written inline in a `foreach` |

All five now call `Get-TenantExchangeServerList` in `Common.ps1`.

**One detail worth keeping.** The list as originally supplied to the author of
Check Queues named `dc4pmexapp01` twice and `dc4pmexapp02` not at all. That
looked like a typo, so the duplicate was collapsed and the missing server
added -- a judgement call, not a transcription. If the intent really was three
servers rather than four, this is the paragraph that explains why there are
four.

---

## How the list is resolved now

`Get-TenantExchangeServerList`, first answer wins:

1. `OnPremExchange.Servers` in `Tenant.json` -- an explicit list, which is the
   right answer when only some servers are in scope.
2. `Get-ExchangeServer`, when an on-premises session is open. The directory
   knows its own servers, and one added last week is included without an edit.
3. Nothing -- and the four tools say so plainly rather than scanning zero
   servers and reporting success.

`ExcludedServers` is subtracted from 1 and 2 alike. Matching ignores case and
any domain suffix, so `DC3PMEXAPP01` and `dc3pmexapp01.ftbco.ftn.com` are one
server.

**Consequence worth knowing:** drop this toolkit on a machine whose
`Tenant.json` has no `Servers` key, connect OnPrem, and the four tools work
against that tenant's real servers with nothing to configure. That is the
behaviour the hardcoded lists were preventing.
