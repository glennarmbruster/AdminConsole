# LogViewer.ps1

Read-only view over `CompletedTasks.csv`. The log's schema and the reasoning
behind it are in `ARCHITECTURE.md` §20; this is about the viewer.

## Read-only on purpose

There is no delete, no edit, no save. Not because a button would be hard to add,
but because the file this reads is an audit trail: **a tool that can quietly
change it is worth less than one that cannot.** Stopping deletion is an NTFS job
— deny Delete on the folder — not something a script can enforce.

## It reads both schema versions

A v1 file has only `Task` and `DateTime`. Rather than showing those as two
lonely columns, the same recovery the *writer* performs is applied here:
`Split-LegacyTask` pulls `Action`, `Target` and `Environment` out of the
sentence where it can. An old log and a new one can then be read side by side in
one grid, which is the only way to look at a change spanning the cutover.

**v1 rows show a blank `Result`, not `Success`.** Those rows never recorded
whether the action worked, and filling in "Success" would invent a fact the file
never held — on exactly the rows where a silent failure could be hiding.

Both timestamp formats are parsed to real `DateTime` objects so the grid sorts
chronologically rather than alphabetically and the date filter works across the
cutover. v1 wrote `yyyy-MM-dd HHmm`, which no standard parser accepts, so it is
handled explicitly. Undated rows sort last rather than throwing.

## Interface

Filter dropdowns are built from the data actually loaded, not a fixed list — a
hardcoded script list would silently stop offering a tool the day it was
renamed.

Search is a plain substring, not a regex: an operator typing a mailbox address
should not have to think about what a dot means.

Only the `Result` **cell** is coloured. Tinting the whole row fights the theme's
own background and turns unreadable on the dark palettes.

**Double-clicking a row filters to its `RunId`** — the whole operation rather
than the one step you happened to click. This is the reason RunId exists.

Export writes **what is on screen**, not the whole file. The use for it is
handing somebody the six rows that matter, and a filtered export is the
difference between an answer and homework.

The detail pane uses an explicit field order rather than column order: `Detail`
is the only field that can wrap to several lines, so it goes last and cannot
push anything out of view if the columns have been dragged around.

This is a **viewer**, so it is deliberately wider than the console (§9).

## Traps

`$null -eq $when`, not `-not $when`: PowerShell will not coerce a `[datetime]`
to a bool, so `-not` throws once the first row parses successfully.

`@()` at the **assignment**, not just inside the function. `return @(...)`
unrolls a single-element array back to a scalar on the way out, and `.Count` on
a lone `PSCustomObject` is `$null` — so the row total silently went blank
whenever exactly one row matched.


## Tenant column (2026-08-09)

Added alongside schema v3. It sits next to Operator rather than next to
Environment on purpose: Environment is the DIRECTORY within a tenant (HQ /
Franchise / EXO), and putting the two side by side invited reading one as the
other.

Blank for v2 rows, and deliberately left blank. `Import-Csv` yields `$null` for
a column that is not in the file, which renders as empty -- and empty reads as
"logged before this column existed", which is true. Filling it in from whatever
Tenant.json this viewer happened to launch under would be a guess written into
an audit trail.

### The grid is positional

`$grid.Rows.Add(...)` takes bare values in the order `Set-GridColumns` declared.
There are no names in that call, so a column added in one place and not the
other shifts every value after it one column left. Nothing errors -- the grid
simply starts showing the wrong data under the right headings, which is worse
than a crash. Both lists now carry a comment saying so, and the verification
script asserts they match.
