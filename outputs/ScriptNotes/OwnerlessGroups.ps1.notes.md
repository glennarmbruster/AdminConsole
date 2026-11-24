# OwnerlessGroups.ps1

Microsoft 365 groups with no effective owner. Shared report conventions in
`ARCHITECTURE.md` §16.

## Two findings in one report, with a Reason column

**"Zero owners"** is the case everyone thinks of and the rarer one — Microsoft
blocks removing the last owner through most surfaces.

**"All owners disabled"** is what actually happens here: the owner leaves, the
disable script runs, and the group keeps an owner who can no longer sign in.
Nothing flags it, because from the group's point of view it still has one.

## Enumeration

`ResultSize` defaults to 1000. Omitting it silently truncates the tenant — and
**a truncated ownerless-groups report reads as good news.**

Soft-deleted groups are deliberately excluded. They are already on their way
out, and listing them as unmanaged findings pads the report with rows nobody
should act on.

## Resolving owners

`ManagedBy` gives owner identities but not their SMTP addresses, and an AD
lookup needs an address. `Get-UnifiedGroupLinks` returns recipient objects, so it
is the resolvable source — but it is one call per group, so it only runs where
`ManagedBy` says there is somebody to check.

An owner who could not be looked up at all is **not** counted as dead. Treating
a directory timeout as "this owner is gone" would invent findings.

`GroupMemberCount` is not guaranteed to be populated on every module build. A
blank cell is honest; a `0` would read as "empty group".

Teams first, then most members: an unmanaged Team with 40 people in it is a
bigger problem than an empty group nobody uses.
