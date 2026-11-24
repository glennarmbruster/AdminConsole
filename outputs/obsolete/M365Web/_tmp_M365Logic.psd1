<#
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
#>
@{
    RootModule        = 'M365Logic.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b7e4c1a2-9f3d-4e58-a6b1-2c7d8e9f0a13'
    Author            = 'Glenn Armbruster'
    Description       = 'Cloud-side (Exchange Online / SharePoint Online) administration logic for the M365_ toolkit. No UI -- consumed by WinForms tools, the browser host, scheduled jobs, and the ad-hoc console alike.'
    PowerShellVersion = '5.1'

    RequiredModules   = @()

    # Explicit public surface. Everything here is intended to be called by a
    # UI layer or a job; nothing is exported by accident.
    FunctionsToExport = @(
        # --- Connection & EXO plumbing ---
        'Get-EXOCommand'
        'Invoke-EXOCommand'
        'Test-M365ExoConnection'
        'Test-M365SpoConnection'
        'Test-M365AadConnection'
        'Get-M365ConnectionSummary'

        # --- Error formatting ---
        'Format-ErrorDetail'

        # --- Address / input parsing ---
        'Get-M365DefaultDomain'
        'Clear-M365DefaultDomainCache'
        'Get-TargetSmtp'
        'ConvertTo-NormalizedUserList'

        # --- Object-type resolution ---
        'Resolve-Target'

        # --- Admin portal deep links ---
        'Get-M365PortalUrl'
        'Set-M365PortalUrlTemplate'
        'Get-M365EntraObjectId'

        # --- Suite identity + tenant guard ---
        'Set-M365SuiteIdentity'
        'Get-M365SuiteName'
        'Test-M365ExpectedTenant'

        # --- Logging ---
        'Write-CompletedLog'

        # --- Inbox rules ---
        'Get-InboxRulesForMailbox'
        'Get-InboxRuleSummaries'
        'Enable-TargetInboxRule'
        'Disable-TargetInboxRule'
        'Remove-TargetInboxRule'

        # --- Mailbox audit ---
        'Get-M365MailboxAuditOperations'
        'Invoke-M365MailboxAudit'
        'ConvertTo-M365AuditRecord'
        'Get-M365FullAccessHolders'

        # --- Mailbox health ---
        'Get-MailboxHealthSnapshot'
        'Format-MailboxHealthReport'

        # --- Mailbox permissions ---
        'Get-MailboxState'
        'Format-MailboxState'
        'Grant-FullAccessPermission'
        'Revoke-FullAccessPermission'
        'Grant-SendAsPermission'
        'Revoke-SendAsPermission'
        'Invoke-MailboxPermissionEdit'

        # --- Message trace / quarantine ---
        'Invoke-MessageTrace'
        'Invoke-MessageTraceDetail'
        'Find-QuarantineMatch'
        'Get-QuarantineMessageDetail'
        'Format-QuarantineMessageDetail'
        'Get-QuarantineMessagePreview'
        'Invoke-QuarantineRelease'
        'Get-TraceTimeZone'
        'ConvertTo-TraceDisplayTime'
        'Format-TraceDetail'

        # --- Calendar permissions ---
        'Get-CalendarFolderIdentity'
        'Get-CalendarPermissions'
        'Set-CalendarPermission'
        'Remove-CalendarPermission'
        'Get-CalendarPermissionRoles'

        # --- Out of office ---
        'Get-OutOfOfficeConfig'
        'Set-OutOfOfficeConfig'
        'Get-OutOfOfficeStates'
        'Get-OutOfOfficeAudiences'

        # --- OneDrive / SharePoint Online ---
        'Get-OneDrivePersonalSiteUrl'
        'Set-M365OneDriveSiteAdmin'

        # --- Mailbox + OneDrive access (bulk) ---
        'Get-M365AccessFlags'
        'Get-M365AccessTypeText'
        'Invoke-M365MailboxOneDriveAccess'

        # --- Disable / termination ---
        'Disable-M365User'
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    PrivateData = @{
        PSData = @{
            Tags         = @('Exchange', 'ExchangeOnline', 'Microsoft365', 'SharePointOnline', 'Administration')
            ReleaseNotes = 'Initial extraction: cloud-side logic pulled out of the M365_ tool scripts so WinForms, the browser host, scheduled jobs, and the ad-hoc console can all call the same functions. Adds Test-M365ExoConnection for long-running-host session checking.'
        }
    }
}
