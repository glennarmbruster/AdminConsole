# ToolManager.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Editor for Tools.json -- add, remove, rename, reorder, and repoint the
    tools AdminConsole.ps1 displays, with no code changes required.
.DESCRIPTION
    AdminConsole.ps1 builds its button list entirely from Tools.json (see
    Update-ToolButtons there). This utility is the supported way to change
    that file -- it never touches AdminConsole.ps1's code, and
    AdminConsole.ps1 never needs to be edited to add, remove, or reorder a
    tool.

    This is also the mechanism for the tiered-delegation goal: to hand a
    junior admin only Tier 1 tools, copy AdminConsole.ps1 + Common.ps1 +
    only the Tier 1 .ps1 tool scripts + a Tools.json trimmed (via this
    utility) to just those entries. AdminConsole.ps1 doesn't need to know
    tiers exist -- Tools.json only contains what that admin is allowed to
    see, so nothing is hidden by permission checks, only by what's actually
    present in the folder.

    Each row in the grid is one Tools.json entry:
      Order        - read-only, shown for reference only. Derived from the
                     row's top-to-bottom position in the grid at Save time --
                     NOT typed in by you. Use Move Up / Move Down to reorder;
                     the Order column updates automatically to match.
      Name         - button text
      Script       - relative path to the tool .ps1 (must live alongside
                     AdminConsole.ps1 / Common.ps1)
      ObjectTypes  - comma-separated: Mailbox, UnifiedGroup, OnpremGroup --
                     controls which detected object types show this button
      Section      - "Admin Tools", "Daily Scripts", or "Connect Scripts" --
                     which button grid in AdminConsole.ps1 the tool's button
                     renders in. Defaults to "Admin Tools" if left blank or
                     missing (older Tools.json files). AdminConsole sizes
                     Admin Tools for a 5-wide x 4-tall grid (20 buttons),
                     Daily Scripts for 5-wide x 3-tall (15 buttons), and
                     Connect Scripts for a single row, 5-wide x 1-tall (5
                     buttons); Save below warns, but does not block, if any
                     section's row count exceeds its grid's nominal capacity
                     -- AdminConsole still shows every button either way (it
                     scrolls), it just stops being a clean rectangle past
                     that count.
      Tier         - free-text label (Tier1/Tier2/Tier3 or your own scheme),
                     for your own organization/documentation when deciding
                     what to distribute to whom -- AdminConsole ignores it

    REVISED 2026-07-23 (operator request -- third Section value, "Connect
    Scripts"): $validSections and $sectionCapacity both gained a "Connect
    Scripts" entry (capacity 5, matching AdminConsole's new single-row
    panel -- see that file's own header for the pixel math), and
    Get-NormalizedSection now recognizes "Connect Scripts" as its own value
    instead of folding it into the "anything unrecognized defaults to Admin
    Tools" catch-all. No other behavior changed -- the Section column's
    ComboBox picks up the new value automatically since it's populated from
    $validSections, and Save's over-capacity warning already generalizes to
    any number of sections in $sectionCapacity without further changes.

    REVISED 2026-07-21 (operator request -- AdminConsole split into two
    button sections): added the Section column above/described here, a
    ComboBox-style grid column so the value is always exactly "Admin Tools"
    or "Daily Scripts", never a typo. New rows default to "Admin Tools" (see
    the DefaultValuesNeeded handler near the grid setup below) since that's
    where every existing tool currently lives; Save's capacity check is new
    too (Section 21 spirit: warn before the operator finds out the hard way
    that a section grew past what its grid was sized for).

    REVISED 2026-07-11 (field report: "there's no add button"). The grid had
    AllowUserToAddRows = $true, so a new tool could technically be added by
    scrolling to the always-present blank row at the bottom and typing into
    it -- but that's a stock DataGridView affordance with no visual cue, easy
    to miss entirely. Added explicit "Add Tool" (jumps to and opens that
    blank row for editing) and "Delete Selected" (removes the selected row,
    with a confirm prompt) buttons so both actions are discoverable instead
    of implicit. Neither changes what Get-ToolsFromGrid/Save already did --
    they're just explicit entry points into behavior the grid already
    supported.

    REWRITTEN 2026-07-10b (field report: "Tool Manager only shows one script,
    even though Tools.json has 10 tools in it"). The previous version bound
    the grid to a System.Data.DataTable through a
    System.Windows.Forms.BindingSource (DataGridView.DataSource = BindingSource,
    BindingSource.DataSource = DataTable) -- a standard WinForms pattern, but
    one this file is the only place in the whole toolset that used it. Every
    other grid in this project (grdRules in InboxRules.ps1, grdTrace in
    MessageTrace.ps1, grdCal in CalendarPermissions.ps1, etc.) populates an
    UNBOUND DataGridView directly via Columns.Add + Rows.Add, and none of
    those have ever shown this symptom. Rather than keep chasing the DataTable/
    BindingSource path blind (this environment can't run WinForms to
    reproduce it interactively), this file now uses that same
    proven-elsewhere unbound-grid pattern instead: Update-ToolsGrid populates
    it directly from the parsed Tools.json array, and Get-ToolsFromGrid reads
    it back out at Save time. This also removes the Order column's old
    failure mode entirely (a typed-in Order value that could silently
    disagree with the row's actual position) by making Order purely a
    read-only, auto-computed reflection of row position -- Set-OrderColumn
    recalculates it after every load, reload, move, or row add/delete.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (rewritten 2026-07-10b -- see REFACTOR-DESIGN.md)
    Revised : 2026-07-21 1357 (added Section column -- see REVISED note above)
    Revised : 2026-07-23 (added "Connect Scripts" as a third Section value --
              see REVISED note above)
```

- **L113-116** Matches AdminConsole.ps1's $flowAdminTools (5x4) / $flowDailyScripts (5x3) / $flowConnectScripts (5x1) panel sizing -- see the REVISED 2026-07-21 and 2026-07-23 notes in AdminConsole.ps1's header for the pixel math these numbers come from.
- **L119-122** Normalizes a raw Section cell value the same permissive way AdminConsole.ps1's own Get-ToolSection does -- blank/missing/unrecognized all default to 'Admin Tools' so this file and AdminConsole never disagree about where a tool with no explicit Section lands.

## function Read-ToolsJson

- **L137-142** @() here matters two ways: if the JSON root is an ARRAY, ConvertFrom-Json already returns an array and @() is a harmless no-op; if the JSON root is a single OBJECT (e.g. someone hand-edited the file down to one entry without keeping the [ ] wrapper), ConvertFrom-Json returns exactly one PSCustomObject, and @() wraps it into a 1-element array so every caller below can treat the result uniformly as a collection either way.

## function Write-ToolsJson

- **L150-154** ConvertTo-Json collapses a single-item array to a bare object -- force array bracket syntax so Tools.json is always a JSON array, even with 0 or 1 entries (Get-ToolDefinitions in AdminConsole.ps1 always wraps its read in @(...), and Read-ToolsJson above now tolerates a bare object too, but keeping the file itself unambiguous is still cleaner).

## region UI

- **L172-174** Unbound DataGridView, populated directly via Columns.Add/Rows.Add -- see the .DESCRIPTION header above for why this replaced the previous DataTable+BindingSource approach.
- **L219-222** Recomputes the read-only Order column from each row's current top-to-bottom position. Called after every load/reload/move so the displayed Order always matches what Get-ToolsFromGrid will actually write out -- there is no other, independently-typed source of truth for it.

## function Set-OrderColumn

- **L232-233** Clears and repopulates the grid from a parsed Tools.json array (the PSCustomObject[] Read-ToolsJson returns).

## function Update-ToolsGrid

- **L245-250** DataGridView's built-in blank trailing row (AllowUserToAddRows = $true) starts every ComboBox cell unset, which would otherwise let a brand-new row reach Get-ToolsFromGrid with a blank Section -- Get-NormalizedSection there would still default it to 'Admin Tools' correctly, but pre-filling it here means the operator SEES that default immediately while typing, instead of it only appearing after Save/Reload.
- **L256-260** Reads the grid back into an array of tool objects, in top-to-bottom row order, skipping the trailing "type here to add a row" row and any fully blank row. Order is assigned from position here too, independent of whatever Set-OrderColumn last displayed, so Save is correct even if this is ever called without a preceding grid refresh.

## function Get-ToolsFromGrid

- **L295-299** Row 1: row-level edits (Add / Delete). The grid already supports adding a row via its built-in blank trailing row (AllowUserToAddRows = $true) and deleting via the Delete key (AllowUserToDeleteRows = $true) -- these buttons don't add new capability, they just make both actions discoverable instead of relying on stock DataGridView affordances with no visual cue.
- **L311** Row 2: row-order + file-level actions.
- **L323-327** Re-reads Tools.json from disk and rebuilds the grid from it, discarding any unsaved edits. Useful any time the file on disk might have changed since this window opened (a newer Tools.json copied in, another admin's edit, etc.) -- also the fastest way to confirm "what does Tools.json on disk actually contain right now" if the grid ever looks wrong.
- **L352-355** Swaps the selected row's cell VALUES with the adjacent row's, rather than trying to physically reorder rows in the grid (DataGridView has no direct "move row" API for an unbound grid) -- the net visual effect is identical. Set-OrderColumn afterward keeps the read-only Order column in sync.

## function Move-SelectedRow

- **L377-380** Jumps to and opens the grid's always-present blank trailing row (present because AllowUserToAddRows = $true) so a new tool can be typed in immediately, instead of the operator having to notice and scroll to it themselves.
- **L389-390** Removes the selected row from the grid (not from Tools.json until Save is clicked). Confirms first since there's no undo besides Reload from Disk.
- **L445-447** Capacity is a layout nicety, not a hard rule (AdminConsole's panels still show every button via AutoScroll past this point) -- warn, don't block Save over it.

## Sections dialog: live size preview (2026-08-09)

Buttons per row used to be a number you changed, saved, relaunched the console
to look at, and then judged -- and if it was wrong, round again. Now that the
console derives button width by measuring captions, the resulting window width
is knowable in the dialog, before saving. The preview updates on
`ValueChanged` and warns when the window would exceed the screen working area.

It reads Tools.json from disk rather than the grid sitting behind the dialog.
The console measures the FILE, so previewing anything else would predict a width
the console is not going to use.

### Two numbers claiming to be the same setting

The instruction label said "Every section is N buttons across", with N baked in
at dialog creation and never updated. The moment the spinner was touched there
were two different numbers on screen for one setting. The label no longer
carries a number; the live one is in the preview.

### ButtonWidth is carried through on save

The save handler rebuilt the config object from scratch, which would have
dropped `ButtonWidth` -- an override this dialog does not edit. Merely
reordering a section would have wiped it, and the symptom would have been "the
console resized itself for no reason", which is not a clue that leads anywhere.


### The preview's scrollbar warning is knowingly pessimistic

The console decides whether to reserve scrollbar width by asking the panels,
after tenant capabilities have hidden whatever they are going to hide. The
dialog cannot do that — there are no panels — so it predicts by counting every
tool in Tools.json against the row counts currently in the grid.

That can predict an overflow the console will not have. It is wrong in the safe
direction: it warns about a window wider than you may actually get, never a
narrower one. The section rows come from the GRID rather than the saved file, so
the number moves as you edit rather than freezing at whatever was last saved.
