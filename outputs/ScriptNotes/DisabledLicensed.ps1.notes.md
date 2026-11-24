# DisabledLicensed.ps1

Disabled AD accounts still sitting in a licence group. AdminConsole only.
Shared report conventions in `ARCHITECTURE.md` §16.

## Licence groups live in Common.ps1

`$script:LicenseGroups` is shared because more than one report needs it. A
second copy here would mean editing a group rename in two places and finding out
about the miss from an **empty report** rather than an error.

A group found only after normalising had a lookalike character in its name —
that is reported, because a silent fuzzy match hides a real inconsistency
between AD and this config.

`Get-ADGroupMember -Recursive` so a nested group cannot hide licensed members;
it returns only leaf user objects, which is what is wanted.

## Rows are people, not memberships

An account can sit in more than one licence group. It is reported **once**,
naming every group, so the row count equals the number of people.

## Recent creations are flagged, not hidden

An account created recently and still disabled is very likely a future starter
rather than a cleanup miss. Flagged rather than filtered, so the list stays
complete and the noise is sortable.

Worst first: never-logged-on and longest-dormant at the top, likely future hires
pushed to the bottom where they are easy to ignore.
