# M365Logic.psd1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 1 -->

```text
Module manifest for M365Logic -- the cloud-side (Exchange Online /
    SharePoint Online) logic module for the M365_ toolkit.

    ADDED 2026-07-24. See M365Logic.psm1's own header for the architecture
    this belongs to (logic module + UI layer split) and for why it's a
    module rather than another dot-sourced .ps1.

    Import either way:
        Import-Module (Join-Path $PSScriptRoot 'M365Logic.psd1') -Force
        Import-Module (Join-Path $PSScriptRoot 'M365Logic.psm1') -Force
    Importing the .psd1 (this file) is preferred -- it applies
    FunctionsToExport below, so only the intended public surface is exposed.
    Importing the .psm1 directly still works but exports every function in
    it.

    -Force matters for consumers: without it, a second Import-Module in the
    same process is a no-op, so edits to the .psm1 wouldn't be picked up
    until the process restarted. Every UI layer in this toolkit uses -Force
    for that reason.

    PowerShellVersion = '5.1' is the deliberate floor, NOT a target: this
    module must load in Windows PowerShell 5.1 (the WinForms suite) AND
    PowerShell 7 (the browser host). Nothing in the .psm1 uses PS7-only
    syntax. Do not raise this without checking both consumers.

    RequiredModules is deliberately EMPTY rather than listing
    ExchangeOnlineManagement. Two reasons: (1) RequiredModules forces an
    auto-import at module-load time, which would make importing this module
    fail outright on a machine that doesn't have EXO installed -- even for
    a caller that only wants the pure helpers (Format-ErrorDetail,
    ConvertTo-NormalizedUserList, Get-OneDrivePersonalSiteUrl) and never
    touches EXO; (2) the EXO cmdlets this module calls come from the
    temporary tmpEXO_* module that Connect-ExchangeOnline generates per
    session, not from statically-declarable exports, so declaring the
    dependency wouldn't actually guarantee the cmdlets are present anyway.
    Get-EXOCommand throws a clear, actionable message when they're missing,
    which is the honest place for that check. Same reasoning for
    Microsoft.Online.SharePoint.PowerShell, which only two functions need.
```

- **L51-52** Explicit public surface. Everything here is intended to be called by a UI layer or a job; nothing is exported by accident.
- **L54** --- Connection & EXO plumbing ---
- **L62** --- Error formatting ---
- **L65** --- Address / input parsing ---
- **L71** --- Object-type resolution ---
- **L74** --- Admin portal deep links ---
- **L79** --- Suite identity + tenant guard ---
- **L84** --- Logging ---
- **L87** --- Inbox rules ---
- **L94** --- Mailbox audit ---
- **L100** --- Mailbox health ---
- **L104** --- Mailbox permissions ---
- **L113** --- Message trace / quarantine ---
- **L125** --- Calendar permissions ---
- **L132** --- Out of office ---
- **L138** --- OneDrive / SharePoint Online ---
- **L142** --- Mailbox + OneDrive access (bulk) ---
- **L147** --- Disable / termination ---
