# TenantManager.ps1

Creates and edits `Config\Tenant.json`.

## Why it exists

Every other config file had an editor. `Tools.json`, `Sections.json` and
`Settings.json` in ToolManager, `Themes.json` in ThemeManager, `QuickLaunch.json`
in QuickLaunchTool. The one file that decides what the console **is** had none.

Sixteen scripts read tenant config; `Common.ps1` alone reads it thirty-eight
times. Nothing wrote it. A new tenant hand-authored JSON, with the same
backslash-escaping hazard that was designed out of the `ScriptOnPrem` field the
same day.

### The failure it prevents

`Test-TenantCapability` returns `$false` when config is absent, and the
capability gate **hides** tools that are off rather than greying them -- which
is the correct behaviour, and which means:

> Ship the toolkit to a tenant with no `Tenant.json` and they open a console
> with every capability-gated tool **invisible**. No error, no warning, no
> prompt. It looks broken and gives no clue why.

That is the worst first-run experience the toolkit can produce, and it was
reachable by simply forgetting one file.

## Two modes, chosen automatically

    no Tenant.json  -> first-run setup, skeleton from Config\Tenant.TEMPLATE.json
    Tenant.json     -> editor

The template now SHIPS in `Config\`. It previously sat in `TenantConfigs\`
outside the toolkit folder, so it never reached the tenant that needed it.

It is deliberately named `Tenant.TEMPLATE.json` and not `Tenant.json`. A blank
file called `Tenant.json` would be worse than none: capabilities all false is
exactly the invisible-console state, except now it looks configured.

## Design decisions

**Detect fills only what is EMPTY.** It reads accepted domains from Exchange
Online and the forest root from AD, and never overwrites a value somebody set.
A config tool that second-guesses the operator is one they stop trusting, and
this is the file that can least afford to be distrusted.

**Unknown fields are carried through, not rebuilt.** `_comment` and the `Pim`
URLs pass straight from what was loaded. An editor that silently drops the
fields it does not understand destroys data the next version needs -- and
`_comment` is the only documentation of this file that travels with it.

**Validate names the consequence, not the rule.** "No capabilities ticked --
every capability-gated tool will be HIDDEN and the console will look empty" is
actionable. "Capabilities is required" is not.

**UTF-8 with BOM and CRLF**, matching every other config file. Without the BOM
PowerShell 5.1 reads it through the ANSI codepage, and a domain containing a
non-ASCII character silently becomes a different domain.

**Saving refreshes the cache in THIS process only.** Windows already open keep
their old values, and the tool says so rather than leaving somebody wondering
why the console did not change.

## Documentation this corrected

Two comments were actively wrong and were found by grepping for what claimed to
edit `Tenant.json`:

- `AdminConsole.ps1` help said "Tenant.json decides that, and **ToolManager >
  Settings edits it**". That dialog writes `Settings.json` only. The help sent
  operators somewhere that could not do the thing.
- `ToolManager.ps1` said "Tenant constants stay hardcoded in Common.ps1 -- this
  is the expresspros toolkit". True before the tenant-agnostic refactor, false
  since, and carrying a tenant name in a comment. It survived the leakage rules
  only because TEN002 now strips comments before matching.

Both are the same failure as a stale variable: documentation that describes a
version of the code that no longer exists, and that nothing checks.

## Not covered

- **PIM portal URLs** are carried through but not editable. They are long,
  rarely change, and a text box for each would dominate the form.
- **`PreferredDomainController`** under `OnPremExchange` is preserved but not
  exposed -- see 14f on why pinning a named controller is usually wrong.
- **No Themes.json equivalent.** A missing `Themes.json` degrades to unthemed
  but fully working tools, and ThemeManager reports its absence in a message
  box. Loud and harmless, unlike the Tenant.json case. Shipping a default theme
  file is still worth doing, but it is cosmetic rather than structural.
