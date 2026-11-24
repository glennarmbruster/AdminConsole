<#
    M365Logic.psd1
    (see ScriptNotes)

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

@{
    RootModule        = 'M365Logic.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b7e4c1a2-9f3d-4e58-a6b1-2c7d8e9f0a13'
    Author            = 'Glenn Armbruster'
    Description       = 'Cloud-side (Exchange Online / SharePoint Online) administration logic for the M365_ toolkit. No UI -- consumed by WinForms tools, the browser host, scheduled jobs, and the ad-hoc console alike.'
    PowerShellVersion = '5.1'

    RequiredModules   = @()

    FunctionsToExport = @(
        'Get-EXOCommand'
        'Invoke-EXOCommand'
        'Test-M365ExoConnection'
        'Test-M365SpoConnection'
        'Test-M365AadConnection'
        'Get-M365ConnectionSummary'

        'Format-ErrorDetail'

        'Get-M365DefaultDomain'
        'Clear-M365DefaultDomainCache'
        'Get-TargetSmtp'
        'ConvertTo-NormalizedUserList'

        'Resolve-Target'

        'Get-M365PortalUrl'
        'Set-M365PortalUrlTemplate'
        'Get-M365EntraObjectId'

        'Set-M365SuiteIdentity'
        'Get-M365SuiteName'
        'Test-M365ExpectedTenant'

        'Write-CompletedLog'

        'Get-InboxRulesForMailbox'
        'Get-InboxRuleSummaries'
        'Enable-TargetInboxRule'
        'Disable-TargetInboxRule'
        'Remove-TargetInboxRule'

        'Get-M365MailboxAuditOperations'
        'Invoke-M365MailboxAudit'
        'ConvertTo-M365AuditRecord'
        'Get-M365FullAccessHolders'

        'Get-MailboxHealthSnapshot'
        'Format-MailboxHealthReport'

        'Get-MailboxState'
        'Format-MailboxState'
        'Grant-FullAccessPermission'
        'Revoke-FullAccessPermission'
        'Grant-SendAsPermission'
        'Revoke-SendAsPermission'
        'Invoke-MailboxPermissionEdit'

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

        'Get-CalendarFolderIdentity'
        'Get-CalendarPermissions'
        'Set-CalendarPermission'
        'Remove-CalendarPermission'
        'Get-CalendarPermissionRoles'

        'Get-OutOfOfficeConfig'
        'Set-OutOfOfficeConfig'
        'Get-OutOfOfficeStates'
        'Get-OutOfOfficeAudiences'

        'Get-OneDrivePersonalSiteUrl'
        'Set-M365OneDriveSiteAdmin'

        'Get-M365AccessFlags'
        'Get-M365AccessTypeText'
        'Invoke-M365MailboxOneDriveAccess'

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
