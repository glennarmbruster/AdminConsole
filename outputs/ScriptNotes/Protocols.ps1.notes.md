# Protocols.ps1

Client-access protocol state for every mailbox, exported to CSV. Shared report
conventions are in `ARCHITECTURE.md` §16; this covers what is specific.

## `$null` does not mean False

`EwsEnabled` on a mailbox is very often `$null`, which does **not** mean
disabled — it means no mailbox-level setting exists and the organisation value
applies. Reporting the raw property produces an empty column that reads as
"disabled". The org default is resolved once up front so every row can state the
**effective** value.

## The module surface varies, so it is interrogated rather than assumed

`Get-CmdletParameterName` returns the parameters a cmdlet actually declares.
Written after guessing `-PageSize` and `-RecipientTypeDetails` onto
`Get-EXOCasMailbox` and having both rejected: the ExchangeOnlineManagement
module's parameter set differs between versions and tenants, so the only
reliable source is the command object in front of you, not the documentation.

Nothing is passed that the installed build does not declare. A hardcoded list is
a guaranteed failure on somebody else's machine.

`Get-EXOCASMailbox` is used rather than `Get-CASMailbox`: it is the REST-backed
cmdlet Microsoft recommends, far quicker over thousands of mailboxes, and
**EXO-only**, so unlike `Get-CASMailbox` it cannot be shadowed by the on-prem
cmdlet of the same name (§2).

## Why the query is sliced

`RecipientTypeDetails` is **not** requested as a property — on
`Get-EXOCasMailbox` it is a *filter* parameter, not a returnable one, and asking
for it fails the call outright. Querying one recipient type at a time then
solves three problems at once: each call is a fraction of the size so a dropped
connection costs less, progress is reportable, and the type is known without
asking for it.

**Slicing only happens if the cmdlet accepts the filter.** If it does not, the
filter would be silently dropped and each of the seven passes would return every
mailbox — seven duplicates apiece, each stamped with the wrong type. Silently
wrong is far worse than unsliced.

A type still too large for one response is subdivided again by alias initial.
**Subdivide rather than retry**: retrying an oversized request reproduces the
same oversized response; splitting it into ~36 small ones is what actually
changes the outcome.

Alphabet slicing keys on `Alias`, so a mailbox whose alias starts outside a-z /
0-9 would be missed and overlapping slices could double-count. Rows are deduped
on SMTP address and the count is reported so the total can be sanity-checked
against the tenant.

## Failure handling

If the service rejects any requested property it returns *"Some of requested
properties are not valid. InvalidProperties = X, Y"*. Because it names them, the
offending properties are dropped and the call retried rather than failing the
whole report.

Transport faults — a dropped connection, a throttling disconnect, a proxy
timeout — are transient and retried with backoff rather than losing a run that
may already be minutes in.

Incompleteness is surfaced prominently: **a partial report that looks complete
is worse than one that failed outright.**
