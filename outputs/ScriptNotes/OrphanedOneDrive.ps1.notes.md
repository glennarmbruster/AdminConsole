# OrphanedOneDrive.ps1

OneDrive sites whose owner is disabled or gone from AD. Shared report
conventions in `ARCHITECTURE.md` §16.

## Finding the sites

`Get-SPOSite -Filter` is server-side and only supports `Owner`, `Template`,
`LockState`, `Url` and `IsAuthoritative`. `Url` is the one that identifies a
personal site.

If `-Filter` turns out to be unavailable the enumeration returns every site in
the tenant, so the list is narrowed client-side as well — belt and braces,
because reporting team sites as orphaned OneDrives would be worse than slow.

The `Owner` value on a personal site is the UPN of the person it belongs to.
**Blank means SharePoint itself has lost the owner**, which is a finding in its
own right.

## Not inventing orphans

A directory that could not answer is **not** evidence the person left. Those
sites are skipped and the count disclosed in the summary, so the report stays
free of invented orphans.

## Storage: the same `-Limit` trap as ExternalSharing

Storage is the whole point — an orphaned OneDrive matters in proportion to how
much data is sitting in it — but `StorageUsageCurrent` is **not populated** when
`-Limit` or `-Filter` is used, which the enumeration above requires. So it is
fetched per site with `-Identity`, and only for sites that are already findings.

`StorageUsedMB` and `LastModified` start as `$null` and are filled in by that
pass. `$null` prints as an empty cell, which is honest — a `0` would read as
*"this OneDrive is empty, ignore it"*.

Biggest first: that is the order in which they are worth dealing with.
