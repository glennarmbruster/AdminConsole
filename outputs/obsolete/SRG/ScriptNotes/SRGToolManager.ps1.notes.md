# SRGToolManager.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Editor for SRGTools.json -- add, remove, rename, reorder, and repoint the
    tools SRGAdminConsole.ps1 displays, with no code changes required.
.DESCRIPTION
    SRGAdminConsole.ps1 builds its button list entirely from SRGTools.json
    (see Build-ToolButtons there). This utility is the supported way to
    change that file -- it never touches SRGAdminConsole.ps1's code, and
    SRGAdminConsole.ps1 never needs to be edited to add, remove, or reorder
    a tool.

    Ported from the original suite's ToolManager.ps1, with one structural
    difference: NO Section column. SRGAdminConsole.ps1 uses a single "Tools"
    button panel, not the original's Admin Tools / Daily Scripts split (the
    operator said most of the original suite's scripts won't carry over to
    SRG, so there's no second category to sort into yet -- if that changes,
    add a Section column back the same way the original's REVISED
    2026-07-21 note describes, and split SRGAdminConsole.ps1's single
    $flowTools panel into two).

    Each row in the grid is one SRGTools.json entry:
      Order        - read-only, shown for reference only. Derived from the
                     row's top-to-bottom position in the grid at Save time --
                     NOT typed in by you. Use Move Up / Move Down to reorder;
                     the Order column updates automatically to match.
      Name         - button text
      Script       - relative path to the tool .ps1 (must live alongside
                     SRGAdminConsole.ps1 / SRGCommon.ps1)
      ObjectTypes  - comma-separated: Mailbox, UnifiedGroup, DistributionGroup
                     -- controls which detected object types show this button
                     (SRG has no OnpremGroup type -- cloud-only tenant, no
                     synced on-prem AD groups to find)
      Tier         - free-text label (Tier1/Tier2/Tier3 or your own scheme),
                     for your own organization/documentation -- SRGAdminConsole
                     ignores it
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21 (ported from ToolManager.ps1, Section column removed)
```

- **L49-52** Matches SRGAdminConsole.ps1's $flowTools panel sizing (5-wide x 4-tall) -- see that file's header for the pixel math this number comes from. Same "warn, don't block" treatment as the original ToolManager.ps1's $sectionCapacity check, just against one panel instead of two.

## function Read-ToolsJson

- **L61-66** @() here matters two ways: if the JSON root is an ARRAY, ConvertFrom-Json already returns an array and @() is a harmless no-op; if the JSON root is a single OBJECT (e.g. someone hand-edited the file down to one entry without keeping the [ ] wrapper), ConvertFrom-Json returns exactly one PSCustomObject, and @() wraps it into a 1-element array so every caller below can treat the result uniformly as a collection either way.

## function Write-ToolsJson

- **L74-76** ConvertTo-Json collapses a single-item array to a bare object -- force array bracket syntax so SRGTools.json is always a JSON array, even with 0 or 1 entries.

## region UI

- **L132-135** Recomputes the read-only Order column from each row's current top-to-bottom position. Called after every load/reload/move so the displayed Order always matches what Get-ToolsFromGrid will actually write out -- there is no other, independently-typed source of truth for it.

## function Set-OrderColumn

- **L145-146** Clears and repopulates the grid from a parsed SRGTools.json array (the PSCustomObject[] Read-ToolsJson returns).

## function Update-ToolsGrid

- **L157-161** Reads the grid back into an array of tool objects, in top-to-bottom row order, skipping the trailing "type here to add a row" row and any fully blank row. Order is assigned from position here too, independent of whatever Set-OrderColumn last displayed, so Save is correct even if this is ever called without a preceding grid refresh.

## function Get-ToolsFromGrid

- **L240-243** Swaps the selected row's cell VALUES with the adjacent row's, rather than trying to physically reorder rows in the grid (DataGridView has no direct "move row" API for an unbound grid) -- the net visual effect is identical. Set-OrderColumn afterward keeps the read-only Order column in sync.

## function Move-SelectedRow

- **L327-329** Capacity is a layout nicety, not a hard rule (SRGAdminConsole's panel still shows every button via AutoScroll past this point) -- warn, don't block Save over it.
