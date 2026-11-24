# ThemeManager.ps1

Editor for `Config\Themes.json` — palettes, tenant bindings, and the logo.

## Showing what would actually load

The window states which theme the console would load **right now**. Without it,
a tenant binding silently overrides `ActiveTheme` and "Set as active theme"
looks broken — the setting takes, and nothing changes.

Picking a theme explicitly therefore also switches the binding override off,
otherwise the binding would win and the choice would appear to do nothing.

## Logos are copied, not referenced

A path into somebody's Desktop breaks the moment the file moves, and is
meaningless on a colleague's machine. The image is copied into `Config\Images\`
(§19).

## Cache invalidation

Saving clears **both** caches, including the globals that child scripts read
from. Without that, an already-open console keeps serving the previous theme and
the change appears not to have worked.

The preview panel is tagged `NoGradient`: it is a Panel, and a gradient painted
on the parent form would not show through it anyway.
