# CertificateExpiry.ps1

Certificate expiry across every on-premises Exchange server. Read-only.

Built 2026-08-03, first piece of the on-premises health report. Certificates
first because an expired one breaks hybrid silently — free/busy stops working
and the first symptom is a user complaining about calendars.

**AdminConsole only.** `Get-ExchangeCertificate` is documented "available only
in on-premises Exchange", so this has no place in the cloud builds.

---

## Three traps, all verified against Microsoft Learn before a line was written

### There is no `Expired` status

`Status` has eight documented values: `DateInvalid`, `Invalid`,
`PendingRequest`, `RevocationCheckFailure`, `Revoked`, `Unknown`, `Untrusted`,
`Valid`.

**`Expired` is not one of them.** An expired certificate reports `DateInvalid`.

A report written the obvious way — `Where-Object { $_.Status -eq 'Expired' }` —
returns **nothing at all**, on every run, forever. For a tool whose entire
purpose is finding expiry, that is the worst available failure: it looks like
good news.

Expiry is therefore computed from `NotAfter` and never inferred from `Status`.
`Status` is still reported verbatim, because `Revoked` and `Untrusted` matter
too, but nothing branches on it.

### `-Server` cannot be combined with `-Identity`

Three parameter sets. `-Server` appears in **Instance** and **Thumbprint**, not
in **Identity**. Microsoft states it plainly: *"You can't use this parameter
with the Identity parameter."*

It also does **not** accept pipeline-by-property-name, so
`Get-ExchangeServer | Get-ExchangeCertificate -Server ...` does not work. It has
to be an explicit loop passing the name.

### The certificate object has no server property

Nothing in the output identifies which server a certificate came from. The
closest is `Identity`, documented as `ServerFQDN\Thumbprint` — parseable, but
parsing a composite key to recover something you already knew is the kind of
cleverness that breaks quietly.

The server is carried from the loop variable instead.

## Property names came from prose, not a schema

Worth recording, because it changes how much to trust them: the
`Get-ExchangeCertificate` page has **no output properties table**. The `Outputs`
section is boilerplate that says nothing. Every property name used here —
`NotAfter`, `NotBefore`, `Subject`, `Thumbprint`, `Services`, `Status`,
`IsSelfSigned`, `Issuer`, `CertificateDomains` — comes from the **Description
prose**, which lists what appears under `Format-List`.

So every property is read defensively via `PSObject.Properties[...]`. A name
that turns out wrong yields a blank column rather than a crash, and a blank
column is visible.

`Get-ExchangeServer` is worse: Learn documents **no output properties at all**.
`Name`, `Fqdn` and `AdminDisplayVersion` are real in practice and unverifiable
from the docs, so those are read defensively too, with `Name` falling back to
the object's own string form.

## `Services` is deliberately stringified

Documented values are `None`, `Federation`, `IIS`, `IMAP`, `POP`, `SMTP`, `UM`,
`UMCallRouter` — but the docs never say whether the property is a **flags enum,
an array, or a formatted string**. That determines whether you test it with
`-band`, `-contains`, or `-like`, and getting it wrong gives silently wrong
answers rather than an error.

`"$($c.Services)"` behaves identically for all three. `ServicesStringForm` and
`IISServices` also exist but are entirely undocumented, so neither is used.

## In-use versus not

A certificate with `Services = None` is not doing anything. Self-signed
leftovers and the `WMSvc-<ServerName>` certificate for the IIS Web Management
Service both show up that way, and Microsoft calls that out explicitly.

Expiry only matters for certificates **bound to a service**, so every severity
count is reported twice: total, and how many of those are in use. The status bar
counts only in-use expired-or-critical, because that is the number that means
"do something today".

Nothing is hidden, though. An unused certificate about to expire is still worth
seeing — it might be unused *by mistake*.

## The OS access probe

Optional checkbox, off by default, and nothing in this report needs it.

It exists because the **full** health report does: disk space on the log volume
is the most common Exchange outage, and reading it needs WMI or WinRM on the
servers, which is a different permission model from Exchange RBAC and may be
gated by PIM.

Rather than guess later or write a throwaway probe script, it asks now: one
`Get-CimInstance Win32_LogicalDisk` per server — read-only, and exactly the call
the disk check would make. **Failure is an answer, not an error**, and it is
reported per server.

## Other decisions

**Thresholds on the form, not in Settings.** Critical 30 days, warning 60,
adjustable per run. A report you eyeball wants the knob next to the output, and
it avoided another round of Settings plumbing. Warning below Critical is
rejected — it means the operator has them backwards.

**One unreadable server does not fail the run.** It is collected into `Failed`,
the report is marked `PARTIAL` in both the summary and the audit log, and every
other server's dates still appear. A server being down is exactly when you most
want to see the rest.

**Sorted by urgency, then days.** Expired first, then critical, then warning.
The top of the file is the thing to act on.

## Still open

- **Untested.** Written against verified documentation, not against a live
  Exchange server. The likeliest failure is a property name that reads blank —
  visible, not silent, by design. First run should be eyeballed against
  `Get-ExchangeCertificate -Server <name> | Format-List *` on one server.
- **No thumbprint-level history.** Each run is a snapshot; nothing diffs against
  the last. That is the Drift Watch idea and belongs there, not here.
- **Certificates only.** Component state, disk, queues, databases and DAG
  replication are the rest of the health report, pending the answer the OS probe
  collects.
