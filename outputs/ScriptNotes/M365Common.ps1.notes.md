# M365Common.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    WinForms UI-layer helpers for the M365_ toolkit. Imports M365Logic (all
    the actual logic) and adds only the WinForms-specific pieces on top.
.DESCRIPTION
    REVISED 2026-07-24 -- GUTTED, deliberately. This file used to hold both
    the logic AND the UI helpers for the tenant-agnostic toolkit. Per the
    operator's "logic module + UI layer" direction, every piece of logic
    that was here has MOVED to M365Logic.psm1, unchanged:

        Get-EXOCommand, Invoke-EXOCommand, Format-ErrorDetail,
        Get-M365DefaultDomain, Get-TargetSmtp, Resolve-Target,
        Get-InboxRulesForMailbox, Write-CompletedLog

    Those are all still available to any script that dot-sources THIS file,
    because this file imports the module -- so no existing tool's calls had
    to change. What's left below is only what is genuinely WinForms and
    genuinely UI:

        Show-TargetPromptDialog  -- the modal "enter a mailbox/group" dialog
        Resolve-ToolTarget       -- decides whether that dialog is needed

    WHY THOSE TWO STAYED BEHIND: both build a System.Windows.Forms.Form.
    They cannot run in the PS7 browser host (WinForms is Windows-only), they
    cannot run in a scheduled job (nothing there to click OK), and they'd
    deadlock a web request if a server thread ever called them. Keeping them
    out of M365Logic.psm1 is what lets that module load and work identically
    in WinForms, PS7, and headless contexts.

    Note the split inside what used to be one concept: Get-TargetSmtp (pure
    string normalization -- append the tenant's default domain if the
    operator didn't type one) is LOGIC and moved to the module.
    Resolve-ToolTarget (if we have no usable address, pop a dialog and ask)
    is UI and stayed here. It calls the module's Get-TargetSmtp.

    HOW TO USE (unchanged for every existing WinForms tool):
        . (Join-Path $PSScriptRoot 'M365Common.ps1')
    That single dot-source still yields everything a WinForms tool needs --
    the module's functions plus the two dialog helpers.

    The web layer does NOT dot-source this file. M365WebCommon.ps1 imports
    M365Logic directly and never touches WinForms at all.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24 (as the tenant-agnostic sibling of Common.ps1 /
              SRGCommon.ps1 / ANZCommon.ps1)
    Revised : 2026-07-24, later -- logic extracted to M365Logic.psm1; this
              file is now the WinForms UI layer only. See REVISED note.
```


## region Logic module import

- **L53-56** -Force so edits to the .psm1 are picked up without restarting the process (without it, a second Import-Module in the same session is a no-op). Importing the .psd1 manifest rather than the .psm1 directly so the manifest's explicit FunctionsToExport applies.
- **L61-63** Fall back to the raw module if the manifest isn't alongside it, so a partially-copied folder fails with a clear message rather than a confusing "command not found" later.

## function Show-TargetPromptDialog

<!-- line 74 -->

```text
.SYNOPSIS
        Modal WinForms dialog asking for a mailbox or group address.
    .DESCRIPTION
        WinForms-only, interactive-only -- see this file's header for why
        this did not move into M365Logic.psm1.

        No Environment radio group and no Domain Controller ComboBox, unlike
        the on-prem suite's equivalent: this toolkit is cloud-only, so there
        is nothing for them to select. The returned object carries only
        .Smtp.

        The label deliberately names no specific domain -- it can't, since
        which domain applies isn't known until Connect-ExchangeOnline has
        run against some tenant. It says the domain can be omitted once
        connected instead.
    .OUTPUTS
        PSCustomObject with a .Smtp property, or $null if cancelled.
```

- **L144** Get-TargetSmtp now comes from M365Logic (imported above).

## function Resolve-ToolTarget

<!-- line 161 -->

```text
.SYNOPSIS
        Returns a usable target for a WinForms tool, prompting only if
        needed.
    .DESCRIPTION
        UI-layer decision function -- it may pop a modal dialog, so it
        belongs here rather than in M365Logic.psm1.

        When a tool is launched with a usable -SmtpAddress (e.g. from a
        console that already looked one up), this returns immediately with no
        UI at all. When it's launched with nothing -- double-clicked
        directly, or handed a bare name with no tenant connected yet to
        resolve the domain against -- it pops Show-TargetPromptDialog.
    .OUTPUTS
        PSCustomObject with a .Smtp property, or $null if the operator
        cancelled.
```

- **L185** operator cancelled
