# Admin Console — customisation

*September 2026*

Almost everything about how the console looks and how its buttons are arranged can be changed
without touching a script. This covers what can be changed, where each change is made, and —
the part that matters most with several people sharing one folder — who it changes it for.

---

## The question every save asks

Four windows ask **"just me, or everyone?"** before they write anything. The answer decides
which file is written:

| What you changed | "Just me" | "Everyone" |
|---|---|---|
| Theme (Theme Manager) | your profile | the tenant file |
| Sections, columns, rows (Tool Manager) | your profile | `Sections.json` |
| Parked scripts and sections (Park In Place) | your profile | the tenant file |
| Button arrangement (Save Arrangement) | your profile | `Sections.json` + `Tools.json` |

Your profile is `Config\Users\<your windows account>.json`, created the first time you save
something. It holds a `Default` block that applies everywhere plus an optional block per
tenant that wins over it — so you can keep one arrangement generally and a different one for a
particular tenant.

**Nothing personal changes what a tool does.** A profile says where a button sits, what colour
the window is, and whether a button is visible. What the tool *is*, where it points, and what
it is allowed to do live in `Tools.json` and `Tenant.json` and are shared by everyone.

Config is read once at startup. Every one of these windows ends by saying "reopen the Admin
Console to see it", and means it.

---

## Themes

**Theme Manager** (Dashboard → Theme Manager) edits `Themes.json`. A theme is a named set of
nineteen colours — window background and text, panel background, group heading text, button
background/text/border/hover/pressed, input background and text, an optional gradient end,
selected radio and checkbox text, the ServiceNow button, the primary (Save/Apply) button and
its text and hover, and the danger (Delete/Remove) colour.

- **Add shipped themes** brings in the bundled palettes if the file has none.
- **Set as active theme** picks the one in use, then asks just-me-or-everyone.
- A theme with `Enabled: false` is ignored. This is the quiet failure: the save succeeds and
  the console opens in the old colours with nothing saying why.

If a theme does not appear to take, run `Diagnose-Theme.ps1` from a prompt. It prints every
step of resolution — your profile, then the tenant's, then the fallbacks — and names the
winner. That is faster than guessing, and it knows about per-user themes.

The order is: **your profile beats the tenant's theme, which beats `ActiveTheme`.**

---

## The coloured line at the top of every window

Every window in the toolkit carries a 6-pixel band across the top in the active tenant's
colour, and the window title gets the tenant's label appended in brackets.

It is not decoration. Two consoles open side by side look identical otherwise, and the tools
they launch write to real mailboxes in different organisations. The band is how you know which
one you are typing into.

Set the colour in **Tenant Manager** — there is a colour picker beside the *Band colour* field,
and an Automatic button. Automatic derives a stable colour from the tenant key, so a tenant
that has never been given a colour still gets a consistent one rather than none.

Single-tenant installs have no band. There is nothing to disambiguate.

---

## Moving buttons around

Click **Arrange Off** in the console header; it becomes **Arrange On** and buttons can be
dragged. Drop a button where you want it — within a section or into another one. The grid
shows where it will land as you drag.

While Arrange is on, buttons do not launch. That is deliberate: dragging and clicking are the
same gesture at the start, and a tool that fired mid-drag would be a bad surprise.

Click **Save Arrangement** when done and answer just-me-or-everyone. What is saved is
placement only — which section each tool is in, its order, and its group. A tool renamed,
repointed or retired in `Tools.json` is still renamed, repointed or retired for everybody,
profile or no profile.

**Refresh** (beside Arrange) redraws the console from the current config: it re-measures the
buttons, reopens any collapsed sections and clears leftover scrollbars. Use it after a change
rather than restarting, where the change did not need a restart.

---

## Sections, rows and columns

**Tool Manager → Sections…** edits the bands the buttons sit in. Each section has a name, a
number of rows, and an optional accent colour; `ButtonsPerRow` applies to the whole window. A
live preview shows the resulting window size as you change the numbers, because a section five
columns wide and four rows deep is a different window from six by three.

The accent tints the section heading and, faintly, the panel behind its buttons — enough to
read as "this band is Reporting", not enough to fight the buttons sitting on it.

Sections hold buttons in the order they are given. A tool whose section does not exist lands in
the first section rather than disappearing.

---

## Grouping tools behind one button

Give several tools the same word in the **Group** column in Tool Manager and they collapse
behind a single console button carrying that word. Clicking it opens a chooser listing the
members; clicking one runs it.

**A group button is outlined in white** so it is obvious at a glance that more sits behind it —
a group called *Calendar* would otherwise look exactly like a tool called *Calendar*. The ring
is drawn on both button styles, and a greyed-out group keeps it: the ring says what the button
*is*, not whether it can be clicked right now.

Searching the console with **Ctrl+F** looks past the group — type a member's name and the
member itself appears, so grouping never hides a tool from search. Clearing the group word
takes a tool back out again.

---

## Parking: hiding tools without deleting them

**Park In Place** (Dashboard → Park In Place) hides a tool's button while leaving the script
in the folder. It also stops `Add-ConsoleAdditions` offering the script again on the next
publish, which is usually the real reason to park something.

Two scopes, chosen with the radio buttons at the top:

- **Just me** — hidden for your Windows account only.
- **Everyone (this tenant)** — hidden for every operator in this tenant.

The tick boxes always mean *the scope you have selected*, and the save writes exactly what is
ticked to that scope. The **HiddenBy** column tells you which list is doing the hiding, so a
tool hidden by the tenant shows as such even when you are looking at your own list.

**Tenant parking always wins.** A tool the tenant has parked stays hidden whatever your profile
says — un-ticking it in your own scope changes nothing, and the window tells you so.

Whole **sections** can be parked the same way, from the checklist at the bottom. A parked
section disappears entirely and the window shrinks to suit. Project scripts under
`Projects\<name>\` can be parked individually; they are keyed by their path rather than their
file name, since project folders reuse names.

Parking everything is refused. An empty console is never what somebody meant — use
`DisabledTools\` if the whole toolkit is being retired.

---

## The look of the buttons themselves

These live in `Settings.json` and apply to the whole install, not per person. Edit them through
**Tool Manager → Settings…**

- **`ConsoleButtonStyle`** — `Raised` gives buttons a bevelled face; `Enhanced` pushes the same
  effect further with a deeper edge and an inner highlight; anything else (blank) is flat.
- **`ConsoleCorners`** — `Rounded` gives 6px corners, anything else is square. Rounded applies
  to raised buttons only, which is a rendering fact rather than a preference: a flat button is
  drawn by WinForms and there is no way to round one.
- **`ConsoleIcons`** — `On` draws a small glyph on each button, and needs a glyph font present
  to do it. Blank is off, which is what most installs run.
- **`ConsoleDialogStyle`** — `Themed` draws message boxes in the theme's colours; anything else
  uses the system box.
- **`ConsoleTitle`** and **`ConsoleTitleFont`** — the words across the top of the console, and
  the face they are set in.

Safe mode (`-SafeMode`) forces all of these off: flat, square, no icons, system dialogs, no
theme. That is the point of it — when the console will not draw, safe mode gives you a plain
Windows window you can still work in.

---

## Carrying a layout somewhere else

**Save / Load Layout** on the Dashboard writes your profile out to a file and puts one back.
Useful for moving an arrangement to another machine, or handing a new starter a sensible
starting point instead of the default grid.

Load offers *this tenant* or *all tenants*. Reset returns you to the shared layout; resetting
this tenant only removes that tenant's block, and resetting the whole profile renames the file
rather than deleting it — so an afternoon of arranging is recoverable.

---

## Two things that catch people out

**A theme saved as "just me" does not show in a console that was already open.** The console
re-reads your profile when the file changes, but one opened *before* the save keeps what it
had. Close it and open it again.

**Parking and theme changes made in a tool window do not reach a running console.** Tenant
Manager, Theme Manager, Tool Manager and Park In Place all run in their own process. The
console picks the change up next time it starts.

---

## Views: the console can look several ways

The **View** button in the console header opens the same tools drawn differently
--- a large-text list, a command box you type into, raised tiles, a board of what
is connected right now. They are windows over the console, not modes: every one
of them launches the same scripts from the same `Tools.json`, and closing one
puts you back on the button grid.

The button grid is a view too. It is the console window itself, and it is what
everyone gets until they choose otherwise.

**Dashboard -> View Manager** says which view *you* open into. Pick one and the
console opens straight into it next launch. The choice is written to your own
profile beside your theme and your groups --- nobody else's console changes ---
and "Back to the button grid" undoes it.

The **A+ / A-** buttons inside a view are remembered too, per view and per
person. Make the tool list readable once and it stays readable; the tile view
keeps its own, separate size.

**The console itself is the first entry, called *Console*.** It is the window --
the grid plus the lookup box and the connect buttons -- and it is not a file, so
it cannot be copied. **Button Layout** is the grid *as* a view: the same tools in
the same sections, drawn as buttons, in a file you can copy and change. That is
the one to start from if you want the grid to look different.

To make your own: View Manager -> pick a view -> **Copy to a new view...** ->
name it (`Bobsview` becomes `View_Bobsview.ps1`; the prefix is added for you).
Edit the file, then set it as the one you open into. Note that the FILE is
shared -- everyone sees it on their View menu -- but nobody else opens into it
unless they choose to.

Each view is a single `View_*.ps1` in `Views\`. A new one is one file dropped in
that folder --- it appears on the View menu and in View Manager with nothing else
to edit. If the View menu says *"Some views did not load"*, hover it: the tooltip
names the file and the reason, which is almost always that `Views\` was not
copied with an update.
