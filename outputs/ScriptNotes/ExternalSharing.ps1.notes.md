# ExternalSharing.ps1

External sharing posture of every SharePoint site, plus the guest accounts that
actually have access. Shared report conventions in `ARCHITECTURE.md` §16.

## The trap that shapes the whole thing

Microsoft documents that when `-Limit` or `-Filter` is used, `Get-SPOSite` does
**not** populate `SharingCapability` — it returns a default value. But
`-Limit ALL` is the only way to enumerate the tenant.

So the obvious one-liner:

```powershell
Get-SPOSite -Limit ALL | Select Url, SharingCapability     # WRONG
```

returns a full site list with a **fabricated** sharing column, and it looks
entirely correct. The enumeration is therefore used only to collect URLs, and
each site is re-fetched with `-Identity` to get its real posture.

## What this report deliberately does not do

**It does not enumerate individual sharing links.** A link lives on an item, so
finding every "Anyone with the link" URL means walking every file in every
library in every site — millions of items across ~5,000 OneDrives. That is not a
slow report, it is a different project. This reports the *posture* (what a site
permits) and the *guests* (who actually has access), which is what you act on.

## Guests

`Get-SPOExternalUser` refuses a `PageSize` above 50, so paging with `-Position`
is not optional. A short page is the end of the collection.

There is a hard per-site stop so a runaway site cannot hang the report, and it
is **disclosed when hit** rather than silently truncating that site's guest
list — a truncated list that looks complete is the failure mode.

Guests are only enumerated where external sharing is actually allowed. A site
with sharing `Disabled` cannot have acquired new guests, and asking anyway would
multiply the runtime for rows that cannot be findings.

"No external users" is commonly surfaced as a *terminating error* rather than an
empty collection, so a first-page failure with nothing collected is not
necessarily a fault.

## Ordering

`SharingCapability` values are ranked worst-first —
`ExternalUserAndGuestSharing` means anonymous "Anyone" links are permitted,
which is the strongest finding. Then by guest count: the order they matter in.
