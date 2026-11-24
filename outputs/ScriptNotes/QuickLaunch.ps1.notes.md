# QuickLaunch.ps1 / QuickLaunchTool.ps1

A URL launcher and its editor. Buttons open links in a named Edge profile;
`QuickLaunchTool.ps1` is where the links, their order, their sections and their
profiles are set.

Built 2026-08-06. Modelled on `M365Reports.ps1`, which does the same job with a
hardcoded link list -- the difference here is that everything is configuration.

---

## Why the profile is per link

Some destinations need the elevated account and some must not have it. A single
selector at the top of the launcher would have been simpler, and would also
have been one stray click away from opening an admin portal under the everyday
identity.

So the profile is a property of the link, set once in the editor, and there is
**no runtime override**. The failure a session-wide override enables is silent:
the URL opens, the page loads, and the account is wrong.

**Blank means `Default`, spelled out.** That is Edge's own directory name for
the default profile, and it is not the same as passing no `--profile-directory`
at all -- that gives the *last used* profile, which depends on what happened to
be open. `Start-UrlInBrowserProfile`'s whole reason for existing is to stop
"whichever profile Edge had open" from deciding an identity, so the launcher
names the profile explicitly even when the link does not.

## The dropdown is populated from Edge, not typed

`Get-EdgeProfiles` reads Edge's `Local State` JSON, so the editor lists profiles
that provably exist. A typo'd profile name resolves to nothing, and
`Start-UrlInBrowserProfile` then falls back to the last-used profile *without an
error* -- for an elevated link that is a sign-in prompt at best.

A profile that has since been renamed or deleted in Edge is **kept** in the
dropdown for that session rather than being blanked. Blanking it would change
where the link opens without saying so, which is the same silent-change problem
one step further along.

## Sections are just a property

There is no `Sections.json` equivalent and no `Order` field. One list, in file
order:

- a section exists because a link is in it
- section order is where that section first appears
- order within a section is row order
- Move Up / Move Down in the editor is the entire ordering model

`Tools.json` and `Sections.json` are split, and the cost of that split is states
that have to be handled every time either file is read: a section with no tools,
a tool naming a section that does not exist. None of those states are
representable here.

## Colour

Two roles this iteration:

| Role | Colour | Meaning |
|---|---|---|
| Admin | dark blue, white text | opens with an elevated identity |
| everything else | the ServiceNow green already in AdminConsole | ordinary link |

Both read from the active theme (`QuickLaunchAdminBack` and friends, added to
every palette in `Themes.json`) with defaults in code, so an older `Themes.json`
still renders correctly rather than falling back to grey.

The point of the Admin colour is that it is visible **before** the click. A
tooltip saying "elevated" is only read by someone already unsure.

**Theme is applied first, link buttons painted after.** `Set-ConsoleTheme`
repaints every `Button` it finds, and its opt-out is `Tag = 'NoTheme'` -- which
is unavailable here because `Tag` is holding the link object. Ordering solves it
without needing a second channel for the data.

## Search

Matches against label, section, note **and URL**. Searching "sharepoint" should
find a button called "Sites" that points at sharepoint.com; matching only the
label would mean knowing what you called it, which is the thing you have
forgotten if you are searching.

The panel is rebuilt rather than filtered-in-place, because a filtered view has
to close the gaps and moving existing controls means tracking their original
positions -- more state than these controls are worth. Old controls are
`Dispose()`d, not just removed: this runs on every keystroke.

## What the editor refuses

Save is blocked, with a list, when a link has a URL but no text, text but no
URL, or a URL that is not `http`/`https`.

The scheme check is not pedantry. `Start-Process` will hand `file:` and anything
else straight to the shell, so without it a JSON file becomes a way to launch
local executables from a button.

Validation lives in the editor rather than the launcher on purpose: the launcher
skips unusable links quietly, which is right when you are trying to *use* it and
useless when you are trying to *fix* it.

**Test Selected** opens the highlighted row exactly as the launcher would, so a
wrong profile is discovered here rather than in front of a sign-in prompt.

## Adding the buttons

Both are registered in `Tools.json` under **Admin Tools**:

| Name | Script | Own window |
|---|---|---|
| `QuickLaunch` | `.\QuickLaunch.ps1` | no |
| `QuickLaunch Tool` | `.\QuickLaunchTool.ps1` | no |

Neither touches Exchange, so neither needs `TOOLKIT-REQUIRES: OwnProcess` and
neither is affected by which directories are connected.

`Config\QuickLaunch.json` is seeded with the three sections discussed --
AvePoint, M365 Quick Links, Express -- as a shape to edit rather than a
recommendation. The URLs in it are the obvious portal addresses and should be
checked before they are trusted.

## Still open

- **A per-link colour override**, deliberately deferred. It wants a colour
  picker in the editor and a decision about whether an override beats the Admin
  role -- and the same override was asked for on the main console's tool
  buttons, so it should be designed once for both rather than twice.
- **Open a whole section at once.** Wanted, not this iteration.
- **No duplicate detection.** Two links to the same URL in different sections is
  legitimate, so this is only worth doing as a warning, not a block.
- **Untested against a live Edge.** The profile plumbing is the same
  `Start-UrlInBrowserProfile` the ServiceNow button and M365 Reports already
  use, so the risk is in the new UI rather than the launching.
