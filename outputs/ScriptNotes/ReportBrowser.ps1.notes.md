# ReportBrowser.ps1

Every CSV in the reports folder, across all report types.

## Why it exists

Each report window can already reach its own history. This answers the other
question — *"what have I got?"* — when the answer spans report types, or when
the report that produced a file is not the one you happen to have open.

**It needs no connection to anything**, which is unusual in this toolkit and
means it never greys out.

## Row counts and the folder signature

Row counts are read here rather than in `Get-ReportFileList` because this window
is explicitly a "look at my files" view — the cost is the point of the screen.
Very large files are reported by size instead, so one 400 MB CSV cannot stall
the whole list.

Refreshing on `Form.Activated` is what makes a file written by another report
appear without pressing Refresh. But rebuilding the list re-counts rows in every
file, so doing it on every focus change would stall the window for no reason.

A cheap **folder signature** — file count, total size, newest write time —
distinguishes "nothing has changed" from "a report just landed" without opening
anything, so the common case costs nothing.

## Guards

Opening is capped, because opening a dozen spreadsheets by accident is its own
kind of bad day.

Delete is **the only destructive control in the reporting suite** and names the
files in the prompt rather than just a count. *"Delete 7 reports?"* tells you
nothing about whether they were the seven you meant.

It is the one report window that benefits from being made bigger — scanning a
list is the entire job — so it uses anchoring rather than a resize handler. See
§9 for why viewers are allowed to be wider than action tools.
