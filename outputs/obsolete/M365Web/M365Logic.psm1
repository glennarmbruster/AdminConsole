#Requires -Version 5.1
<#
    M365Logic.psm1
    M365Logic -- all cloud-side (Exchange Online / SharePoint Online)
    administration LOGIC for the M365_ toolkit, with no UI of any kind.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

#region Module-private state -----------------------------------------------------
$script:M365DefaultDomainCache = $null
#endregion

#region EXO cmdlet resolution ----------------------------------------------------
function Get-EXOCommand {
    param([Parameter(Mandatory)][string]$Name)
    $all = @(Get-Command -Name $Name -All -ErrorAction SilentlyContinue)
    if ($all.Count -eq 0) {
        throw "No command named '$Name' exists in this session. Is Exchange Online connected (Connect-ExchangeOnline) to the target tenant? Note that Connect-ExchangeOnline is per-PROCESS -- connecting in a different window/process does not carry over to this one."
    }
    $cmd = $all | Where-Object { $_.ModuleName -like 'tmpEXO_*' -or $_.Source -like 'tmpEXO_*' } | Select-Object -First 1
    if (-not $cmd) {
        $cmd = $all | Where-Object { $_.ModuleName -eq 'ExchangeOnlineManagement' -or $_.Source -eq 'ExchangeOnlineManagement' } | Select-Object -First 1
    }
    if (-not $cmd) {
        $sources = ($all | ForEach-Object { "$($_.Name) [ModuleName=$($_.ModuleName); Source=$($_.Source)]" }) -join '  |  '
        throw "Found '$Name' but none of the loaded copies look like Exchange Online: $sources -- run 'Get-Command $Name -All | Select Name,ModuleName,Source' to confirm which one this should be; an on-prem Exchange module/snap-in may be shadowing it."
    }
    return $cmd
}

function Invoke-EXOCommand {
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    $cmd = Get-EXOCommand -Name $Name
    & $cmd @Parameters
}

function Test-M365ExoConnection {
    param()
    $result = [ordered]@{
        Connected     = $false
        CmdletsLoaded = $false
        Detail        = ''
        Organization  = $null
    }

    $probe = @(Get-Command -Name 'Get-Mailbox' -All -ErrorAction SilentlyContinue |
        Where-Object { $_.ModuleName -like 'tmpEXO_*' -or $_.Source -like 'tmpEXO_*' -or
                       $_.ModuleName -eq 'ExchangeOnlineManagement' -or $_.Source -eq 'ExchangeOnlineManagement' })
    if ($probe.Count -eq 0) {
        $result.Detail = 'No Exchange Online cmdlets are present in this PowerShell process. Run Connect-ExchangeOnline in THIS process (it is per-process -- connecting elsewhere does not carry over).'
        return [PSCustomObject]$result
    }
    $result.CmdletsLoaded = $true

    $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
    if (-not $gci) {
        $result.Connected = $true
        $result.Detail    = 'Exchange Online cmdlets are loaded. This module version has no Get-ConnectionInformation, so live session state could not be verified -- assuming usable.'
        return [PSCustomObject]$result
    }

    try {
        $conns = @(& $gci -ErrorAction Stop)
        if ($conns.Count -eq 0) {
            $result.Detail = 'Exchange Online cmdlets are loaded, but there is no active connection (the session may have expired or been disconnected). Run Connect-ExchangeOnline again in this process.'
            return [PSCustomObject]$result
        }
        $healthy = $conns | Where-Object {
            (-not $_.PSObject.Properties['State']       -or "$($_.State)"       -eq 'Connected') -and
            (-not $_.PSObject.Properties['TokenStatus'] -or "$($_.TokenStatus)" -ne 'Expired')
        }
        $chosen = $conns | Select-Object -First 1
        if ($chosen.PSObject.Properties['Organization']) { $result.Organization = "$($chosen.Organization)" }
        if (@($healthy).Count -gt 0) {
            $result.Connected = $true
            $orgNote = ''
            if ($result.Organization) { $orgNote = " (organization: $($result.Organization))" }
            $result.Detail = "Connected to Exchange Online$orgNote."
        } else {
            $result.Detail = 'Exchange Online cmdlets are loaded, but the existing connection reports itself expired/disconnected. Run Connect-ExchangeOnline again in this process.'
        }
    } catch {
        $result.Connected = $true
        $result.Detail    = "Exchange Online cmdlets are loaded. Get-ConnectionInformation could not be read ($($_.Exception.Message)) -- assuming usable."
    }
    return [PSCustomObject]$result
}
#endregion

#region SPO / Entra connection probes ---------------------------------------------

function Test-M365SpoConnection {
    param()
    $r = [ordered]@{ Connected = $false; CmdletsLoaded = $false; Detail = '' }

    if (-not (Get-Command -Name 'Get-SPOTenant' -ErrorAction SilentlyContinue)) {
        $r.Detail = 'SharePoint Online cmdlets are not loaded in this process. Import Microsoft.Online.SharePoint.PowerShell and run Connect-SPOService. Note this module is not natively supported on PowerShell 7 -- in PS7 it needs Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell.'
        return [PSCustomObject]$r
    }
    $r.CmdletsLoaded = $true
    try {
        $null = Get-SPOTenant -ErrorAction Stop
        $r.Connected = $true
        $r.Detail    = 'Connected to SharePoint Online.'
    } catch {
        $r.Detail = "SharePoint Online cmdlets are loaded but not connected (or the connection has expired): $($_.Exception.Message)"
    }
    return [PSCustomObject]$r
}

function Test-M365AadConnection {
    param()
    $r = [ordered]@{ Connected = $false; Provider = $null; Detail = ''; Account = $null }
    $tried = @()

    if (Get-Command -Name 'Get-AzureADCurrentSessionInfo' -ErrorAction SilentlyContinue) {
        try {
            $s = Get-AzureADCurrentSessionInfo -ErrorAction Stop
            $r.Connected = $true
            $r.Provider  = 'AzureAD'
            if ($s.Account)    { $r.Account = "$($s.Account)" }
            elseif ($s.TenantId) { $r.Account = "tenant $($s.TenantId)" }
            $r.Detail = "Connected via the AzureAD module$(if ($r.Account) { " as $($r.Account)" }). Note: AzureAD is deprecated -- some cmdlets that call the retired Azure AD Graph endpoint will fail even though the connection itself is fine."
            return [PSCustomObject]$r
        } catch {
            $tried += 'AzureAD (cmdlets present, no active session)'
        }
    } else {
        $tried += 'AzureAD (module not loaded -- it is Windows PowerShell 5.1 only, so it cannot load under PowerShell 7)'
    }

    if (Get-Command -Name 'Get-MgContext' -ErrorAction SilentlyContinue) {
        try {
            $ctx = Get-MgContext -ErrorAction Stop
            if ($ctx) {
                $r.Connected = $true
                $r.Provider  = 'Graph'
                if ($ctx.Account) { $r.Account = "$($ctx.Account)" }
                $r.Detail = "Connected via Microsoft Graph$(if ($r.Account) { " as $($r.Account)" })."
                return [PSCustomObject]$r
            }
            $tried += 'Graph (module present, no context)'
        } catch {
            $tried += 'Graph (context read failed)'
        }
    } else {
        $tried += 'Graph (module not loaded)'
    }

    if (Get-Command -Name 'Get-EntraContext' -ErrorAction SilentlyContinue) {
        try {
            $ctx = Get-EntraContext -ErrorAction Stop
            if ($ctx) {
                $r.Connected = $true
                $r.Provider  = 'Entra'
                if ($ctx.Account) { $r.Account = "$($ctx.Account)" }
                $r.Detail = "Connected via the Microsoft.Entra module$(if ($r.Account) { " as $($r.Account)" })."
                return [PSCustomObject]$r
            }
            $tried += 'Entra (module present, no context)'
        } catch {
            $tried += 'Entra (context read failed)'
        }
    } else {
        $tried += 'Entra (module not loaded)'
    }

    $r.Detail = "No directory connection found. Tried: $($tried -join '; ')."
    return [PSCustomObject]$r
}

function Get-M365ConnectionSummary {
    param(
        [switch]$IncludeExpensive,
        [switch]$IncludeDirectory
    )

    $exo = Test-M365ExoConnection

    if ($IncludeExpensive) {
        $spo = Test-M365SpoConnection
    } else {
        $spo = [PSCustomObject]@{ Connected = $false; CmdletsLoaded = $false; Detail = 'Not checked on this pass (checking SharePoint costs a service call -- use Refresh).'; Skipped = $true }
    }

    if (-not $IncludeDirectory) {
        $aad = [PSCustomObject]@{ Connected = $false; Provider = $null; Detail = 'Directory connection is not monitored (nothing in this toolkit uses it).'; Account = $null; Skipped = $true; NotMonitored = $true }
    } elseif ($IncludeExpensive) {
        $aad = Test-M365AadConnection
    } else {
        $aad = [PSCustomObject]@{ Connected = $false; Provider = $null; Detail = 'Not checked on this pass (use Refresh).'; Account = $null; Skipped = $true }
    }

    return [PSCustomObject]@{
        Exo = $exo
        Spo = $spo
        Aad = $aad
    }
}
#endregion

#region Error formatting ---------------------------------------------------------
function Format-ErrorDetail {
    param(
        [Parameter(Mandatory)] $ErrorRecord,
        [string]$Context = ''
    )
    $ex   = $ErrorRecord.Exception
    $cmd  = if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.MyCommand) {
                $ErrorRecord.InvocationInfo.MyCommand.Name
            } else { '(cmdlet unknown)' }
    $line = if ($ErrorRecord.InvocationInfo) { $ErrorRecord.InvocationInfo.ScriptLineNumber } else { '?' }
    $sb = New-Object System.Text.StringBuilder
    $prefix = if ($Context) { "ERROR [$Context]" } else { 'ERROR' }
    [void]$sb.AppendLine("$prefix -- $cmd (script line $line)")
    [void]$sb.AppendLine("    Message  : $($ex.Message)")
    [void]$sb.AppendLine("    Type     : $($ex.GetType().FullName)")
    if ($ErrorRecord.CategoryInfo) {
        [void]$sb.AppendLine("    Category : $($ErrorRecord.CategoryInfo.Category) ($($ErrorRecord.CategoryInfo.Reason))")
    }
    if ($ex.InnerException) {
        [void]$sb.AppendLine("    Inner    : $($ex.InnerException.Message)")
    }
    return $sb.ToString().TrimEnd()
}
#endregion

#region Address parsing ----------------------------------------------------------
function Get-M365DefaultDomain {
    if ($script:M365DefaultDomainCache) { return $script:M365DefaultDomainCache }
    try {
        $accepted = @(Invoke-EXOCommand -Name 'Get-AcceptedDomain' -Parameters @{ ErrorAction = 'Stop' })
        $default = $accepted | Where-Object { $_.Default -eq $true } | Select-Object -First 1
        if ($default) {
            $script:M365DefaultDomainCache = $default.DomainName
            return $script:M365DefaultDomainCache
        }
    } catch {
    }
    return $null
}

function Clear-M365DefaultDomainCache {
    $script:M365DefaultDomainCache = $null
}

function Get-TargetSmtp {
    param([string]$Text)
    if (-not $Text) { return $null }
    $t = $Text.Trim()
    if ($t -eq '') { return $null }
    if ($t -like '*@*') { return $t }
    $domain = Get-M365DefaultDomain
    if (-not $domain) { return $null }
    return "$t@$domain"
}

function ConvertTo-NormalizedUserList {
    param([string]$Text)
    if (-not $Text) { return @() }
    $raw = @($Text -split '[\r\n,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $result = @()
    foreach ($u in $raw) {
        if ($seen.Add($u.ToLowerInvariant())) { $result += $u }
    }
    return $result
}
#endregion

#region Object-type resolution (EXO only -- no on-prem branch) -------------------
function Resolve-Target {
    param([string]$Smtp)
    $result = [ordered]@{
        Type             = 'NotFound'
        Smtp             = $Smtp
        ResolutionErrors = @()
    }
    try {
        Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'Mailbox'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (EXO) resolution attempt')
    }
    try {
        Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'UnifiedGroup'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroup (EXO) resolution attempt')
    }
    try {
        Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'DistributionGroup'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroup (EXO) resolution attempt')
    }
    return [PSCustomObject]$result
}
#endregion

#region Admin portal deep links ---------------------------------------------------
$script:M365PortalUrlTemplates = @{
    EntraUser      = 'https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/{objectId}'
    EntraUsers     = 'https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/AllUsers'
    ThreatExplorer = 'https://security.microsoft.com/threatexplorerv3'
    Quarantine     = 'https://security.microsoft.com/quarantine?viewid=Email'
    ExchangeAdmin  = 'https://admin.exchange.microsoft.com/'
}

function Set-M365PortalUrlTemplate {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('EntraUser', 'EntraUsers', 'ThreatExplorer', 'Quarantine', 'ExchangeAdmin')]
        [string]$Kind,
        [Parameter(Mandatory)][string]$Template
    )
    $script:M365PortalUrlTemplates[$Kind] = $Template
}

function Get-M365PortalUrl {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('EntraUser', 'EntraUsers', 'ThreatExplorer', 'Quarantine', 'ExchangeAdmin')]
        [string]$Kind,
        [string]$ObjectId,
        [string]$Sender,
        [string]$Recipient,
        [string]$MessageId,
        [switch]$Detailed
    )

    if ($Kind -eq 'EntraUser' -and [string]::IsNullOrWhiteSpace($ObjectId)) {
        $Kind = 'EntraUsers'
    }

    $template = $script:M365PortalUrlTemplates[$Kind]
    $url = $template
    $didSubstitute = $false

    $subs = @(
        @{ Token = '{objectId}';  Value = $ObjectId }
        @{ Token = '{sender}';    Value = $Sender }
        @{ Token = '{recipient}'; Value = $Recipient }
        @{ Token = '{messageId}'; Value = $MessageId }
    )
    foreach ($s in $subs) {
        if ($url -like "*$($s.Token)*") {
            $val = if ([string]::IsNullOrWhiteSpace($s.Value)) { '' } else { [uri]::EscapeDataString($s.Value) }
            $url = $url.Replace($s.Token, $val)
            if ($s.Token -ne '{objectId}' -and -not [string]::IsNullOrWhiteSpace($s.Value)) {
                $didSubstitute = $true
            }
        }
    }

    if ($Detailed) {
        return [PSCustomObject]@{ Url = $url; Prepopulated = $didSubstitute }
    }
    return $url
}

function Get-M365EntraObjectId {
    param([Parameter(Mandatory)][string]$Smtp)
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        if ($mbx.ExternalDirectoryObjectId) { return "$($mbx.ExternalDirectoryObjectId)" }
    } catch {
    }
    return $null
}
#endregion

#region Suite identity + tenant guard ---------------------------------------------
$script:M365SuiteName    = 'M365'
$script:M365LogFileName  = 'M365-CompletedTasks.csv'
$script:M365ExpectedDomain = $null

function Set-M365SuiteIdentity {
    param(
        [Parameter(Mandatory)][string]$SuiteName,
        [string]$LogFileName,
        [string]$ExpectedDomain
    )
    $script:M365SuiteName = $SuiteName
    if ($LogFileName)   { $script:M365LogFileName = $LogFileName }
    if ($ExpectedDomain) { $script:M365ExpectedDomain = $ExpectedDomain }
    Clear-M365DefaultDomainCache
}

function Get-M365SuiteName { return $script:M365SuiteName }

function Test-M365ExpectedTenant {
    param()
    $r = [ordered]@{ Armed = $false; Match = $true; Expected = $script:M365ExpectedDomain; Actual = $null; Detail = '' }

    if (-not $script:M365ExpectedDomain) {
        $r.Detail = 'No expected tenant declared -- this host is tenant-agnostic and will work against whatever is connected.'
        return [PSCustomObject]$r
    }
    $r.Armed = $true

    $actual = Get-M365DefaultDomain
    $r.Actual = $actual
    if (-not $actual) {
        $r.Detail = "Expecting $($script:M365ExpectedDomain), but the connected tenant is not known yet (connect Exchange Online first)."
        return [PSCustomObject]$r
    }
    if ($actual -ieq $script:M365ExpectedDomain) {
        $r.Detail = "Connected tenant matches this host: $actual."
    } else {
        $r.Match  = $false
        $r.Detail = "TENANT MISMATCH -- this host is for $($script:M365ExpectedDomain), but the connected tenant's default domain is $actual. Every tool here will act on $actual. Close this host, or reconnect to the right tenant."
    }
    return [PSCustomObject]$r
}
#endregion

#region Completed-tasks logging ---------------------------------------------------
function Write-CompletedLog {
    param([string]$Task)
    $folder = Join-Path $PSScriptRoot 'Logs\Completed'
    $file   = Join-Path $folder $script:M365LogFileName
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $file)) {
        "Task,DateTime" | Out-File -FilePath $file -Encoding utf8
    }
    $dt = Get-Date -Format "yyyy-MM-dd HHmm"
    if ($Task -and $Task.Trim() -ne "") {
        Add-Content -Path $file -Value ('"{0}","{1}"' -f $Task.Trim(), $dt)
    }
}
#endregion

#region Inbox rules --------------------------------------------------------------
function Get-InboxRulesForMailbox {
    param([string]$Smtp)
    try {
        $rules = @(Invoke-EXOCommand -Name 'Get-InboxRule' -Parameters @{ Mailbox = $Smtp; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ Rules = $rules; Error = $null }
    } catch {
        return [PSCustomObject]@{ Rules = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-InboxRule') }
    }
}
#endregion

#region Mailbox audit (Search-UnifiedAuditLog) ------------------------------------
$script:M365MailboxOperations = [ordered]@{
    'Accessed mailbox attachments'                                  = 'AttachmentAccess'
    'Accessed Mailbox folder'                                       = 'FolderBind'
    'Accessed mailbox items'                                        = 'MailItemsAccessed'
    'Accessed message'                                              = 'MessageBind'
    'Added delegate mailbox permissions'                            = 'Add-MailboxPermission'
    'Added or removed user with delegate access to calendar folder' = 'UpdateCalendarDelegation'
    'Added permissions to folder'                                   = 'AddFolderPermissions'
    'Copied messages to another folder'                             = 'Copy'
    'Created mailbox item'                                          = 'Create'
    'Created new inbox rule in Outlook web app'                     = 'New-InboxRule'
    'Delete priority cleanup label item'                            = 'PriorityCleanupDelete'
    'Deleted messages from Deleted Items folder'                    = 'SoftDelete'
    'Labeled item as a record'                                      = 'ApplyRecord'
    'Labeled message as a record'                                   = 'ApplyRecordLabel'
    'Modified folder permission'                                    = 'ModifyFolderPermissions'
    'Modified inbox rule from Outlook web app'                      = 'Set-InboxRule'
    'Moved messages to another folder'                              = 'Move'
    'Moved messages to Deleted Items folder'                        = 'MoveToDeletedItems'
    'Preserve Mailbox item'                                         = 'PreservedMailItemProactively'
    'Priority cleanup label applied'                                = 'ApplyPriorityCleanup'
    'Purged messages from the mailbox'                              = 'HardDelete'
    'Removed delegate mailbox permissions'                          = 'Remove-MailboxPermission'
    'Removed permissions from folder'                               = 'RemoveFolderPermissions'
    'Search items in a mailbox'                                     = 'SearchQueryInitiated'
    'Sent message'                                                  = 'Send'
    'Sent message using Send As permissions'                        = 'SendAs'
    'Sent message using Send On Behalf permissions'                 = 'SendOnBehalf'
    'Update item label'                                             = 'UpdateComplianceTag'
    'Updated inbox rules from Outlook client'                       = 'UpdateInboxRules'
    'Updated message'                                               = 'Update'
    'User signed in to mailbox'                                     = 'MailboxLogin'
}
$script:M365OperationFriendlyNames = @{}
foreach ($key in $script:M365MailboxOperations.Keys) {
    $script:M365OperationFriendlyNames[$script:M365MailboxOperations[$key]] = $key
}
$script:M365AuditResultCeiling = 50000
$script:M365AuditPageSize      = 5000

function Get-M365MailboxAuditOperations {
    return $script:M365MailboxOperations
}

function Invoke-M365MailboxAudit {
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][datetime]$StartUtc,
        [Parameter(Mandatory)][datetime]$EndUtc,
        [string[]]$Operations,
        [int]$MaxPages = 15
    )
    $sessionId  = "M365Audit_$([Guid]::NewGuid().ToString())"
    $all        = New-Object System.Collections.Generic.List[object]
    $hitCeiling = $false
    $iteration  = 0

    while ($iteration -lt $MaxPages) {
        $iteration++
        $params = @{
            StartDate      = $StartUtc
            EndDate        = $EndUtc
            UserIds        = $Mailbox
            SessionId      = $sessionId
            SessionCommand = 'ReturnLargeSet'
            ResultSize     = $script:M365AuditPageSize
            ErrorAction    = 'Stop'
        }
        if ($Operations -and $Operations.Count -gt 0) { $params['Operations'] = $Operations }
        $page = @(Invoke-EXOCommand -Name 'Search-UnifiedAuditLog' -Parameters $params)
        if ($page.Count -eq 0) { break }
        foreach ($row in $page) { $all.Add($row) }
        if ($all.Count -ge $script:M365AuditResultCeiling) { $hitCeiling = $true; break }
        if ($page.Count -lt $script:M365AuditPageSize) { break }
    }
    return [PSCustomObject]@{
        Records    = $all.ToArray()
        HitCeiling = $hitCeiling
        Pages      = $iteration
    }
}

function ConvertTo-M365AuditRecord {
    param($RawRecord)
    $parsed = $null
    try { $parsed = $RawRecord.AuditData | ConvertFrom-Json -ErrorAction Stop } catch { $parsed = $null }
    $opValue  = if ($parsed -and $parsed.Operation) { "$($parsed.Operation)" } else { "$($RawRecord.Operations)" }
    $friendly = if ($script:M365OperationFriendlyNames.ContainsKey($opValue)) { $script:M365OperationFriendlyNames[$opValue] } else { $opValue }
    $creationUtc = if ($parsed -and $parsed.CreationTime) {
        try {
            [datetime]::Parse($parsed.CreationTime, [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
        } catch { $RawRecord.CreationDate }
    } else { $RawRecord.CreationDate }

    [PSCustomObject]@{
        TimeUtc           = if ($creationUtc) { ([datetime]$creationUtc).ToUniversalTime().ToString('o') } else { $null }
        Operation         = $opValue
        OperationFriendly = $friendly
        Actor             = if ($parsed) { "$($parsed.UserId)" } else { "$($RawRecord.UserIds)" }
        MailboxOwnerUPN   = if ($parsed) { "$($parsed.MailboxOwnerUPN)" } else { '' }
        ClientIP          = if ($parsed) { "$($parsed.ClientIP)" } else { '' }
        ClientInfo        = if ($parsed) { "$($parsed.ClientInfoString)" } else { '' }
        Subject           = if ($parsed -and $parsed.Item -and $parsed.Item.Subject) { "$($parsed.Item.Subject)" } else { '' }
        Folder            = if ($parsed -and $parsed.Item -and $parsed.Item.ParentFolder -and $parsed.Item.ParentFolder.Name) { "$($parsed.Item.ParentFolder.Name)" } else { '' }
        ResultStatus      = if ($parsed) { "$($parsed.ResultStatus)" } else { '' }
        RawJson           = if ($parsed) { ($parsed | ConvertTo-Json -Depth 8) } else { "$($RawRecord.AuditData)" }
    }
}

function Get-M365FullAccessHolders {
    param([Parameter(Mandatory)][string]$Smtp)
    try {
        $fa = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.User -notlike 'NT AUTHORITY\SELF' -and
                $_.IsInherited -eq $false -and
                $_.Deny -eq $false -and
                ($_.AccessRights -contains 'FullAccess')
            })
        return [PSCustomObject]@{ Holders = @($fa | ForEach-Object { "$($_.User)" }); Error = $null }
    } catch {
        return [PSCustomObject]@{ Holders = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access holders)') }
    }
}
#endregion

#region Inbox rule actions --------------------------------------------------------
function Enable-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Enable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Disable-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Disable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Remove-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}

function Get-InboxRuleSummaries {
    param([string]$Smtp)
    $res = Get-InboxRulesForMailbox -Smtp $Smtp
    if ($res.Error) { return [PSCustomObject]@{ Rules = @(); Error = $res.Error } }

    $rules = @($res.Rules | ForEach-Object {
        $actions = @()
        if ($_.ForwardTo -or $_.ForwardAsAttachmentTo) { $actions += 'Forward' }
        if ($_.RedirectTo) { $actions += 'Redirect' }
        if ($_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage) { $actions += 'Delete' }
        $suspicious = ([bool]$_.Enabled) -and ($actions.Count -gt 0)

        [PSCustomObject]@{
            Name        = "$($_.Name)"
            Identity    = "$($_.Identity)"
            Enabled     = [bool]$_.Enabled
            Priority    = $_.Priority
            Description = if ($_.Description) { ($_.Description -replace '\s+', ' ').Trim() } else { '' }
            FullDescription = "$($_.Description)"
            Actions     = ($actions -join ', ')
            Suspicious  = $suspicious
        }
    })
    return [PSCustomObject]@{ Rules = $rules; Error = $null }
}
#endregion

#region Mailbox health -----------------------------------------------------------
function Get-MailboxHealthSnapshot {
    param([string]$Smtp)
    $health = [ordered]@{
        LastLogonTime              = $null
        ItemCount                  = $null
        MailboxSize                = $null
        ProhibitSendQuota          = $null
        ProhibitSendReceiveQuota   = $null
        ArchiveEnabled             = $null
        ArchiveSize                = $null
        AccountDisabled            = $null
        ForwardingSmtpAddress      = $null
        ForwardingAddress          = $null
        DeliverToMailboxAndForward = $null
        LitigationHoldEnabled      = $null
        RetentionHoldEnabled       = $null
        InPlaceHoldsCount          = $null
        FullAccessDelegates        = @()
        SendAsDelegates            = @()
        RuleCount                  = $null
        SuspiciousRules            = @()
        OWAEnabled                 = $null
        ActiveSyncEnabled          = $null
        PopEnabled                 = $null
        ImapEnabled                = $null
        MAPIEnabled                = $null
        EwsEnabled                 = $null
        Errors                     = @()
    }
    try {
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.LastLogonTime = $stats.LastLogonTime
        $health.ItemCount     = $stats.ItemCount
        $health.MailboxSize   = $stats.TotalItemSize
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (health)')
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.ArchiveEnabled             = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        $health.AccountDisabled            = $mbx.AccountDisabled
        $health.ProhibitSendQuota          = $mbx.ProhibitSendQuota
        $health.ProhibitSendReceiveQuota   = $mbx.ProhibitSendReceiveQuota
        $health.ForwardingSmtpAddress      = $mbx.ForwardingSmtpAddress
        $health.ForwardingAddress          = $mbx.ForwardingAddress
        $health.DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
        $health.LitigationHoldEnabled      = $mbx.LitigationHoldEnabled
        $health.RetentionHoldEnabled       = $mbx.RetentionHoldEnabled
        $health.InPlaceHoldsCount          = if ($mbx.InPlaceHolds) { @($mbx.InPlaceHolds).Count } else { 0 }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (quota/forwarding/hold/account status)')
    }
    if ($health.ArchiveEnabled) {
        try {
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $health.ArchiveSize = $astats.TotalItemSize
        } catch {
            $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (health)')
        }
    }
    try {
        $fa = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.User -notlike 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false -and $_.Deny -eq $false -and ($_.AccessRights -contains 'FullAccess') })
        $health.FullAccessDelegates = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access delegates)')
    }
    try {
        $sa = @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.Trustee -notlike 'NT AUTHORITY\SELF' -and $_.AccessControlType -eq 'Allow' })
        $health.SendAsDelegates = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As delegates)')
    }
    try {
        $rules = Get-InboxRulesForMailbox -Smtp $Smtp
        if ($rules.Error) {
            $health.Errors += $rules.Error
        } else {
            $health.RuleCount = $rules.Rules.Count
            $health.SuspiciousRules = @(
                $rules.Rules | Where-Object {
                    $_.Enabled -and (
                        $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or
                        $_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage
                    )
                } | ForEach-Object {
                    $actions = @()
                    if ($_.ForwardTo -or $_.ForwardAsAttachmentTo) { $actions += 'Forward' }
                    if ($_.RedirectTo) { $actions += 'Redirect' }
                    if ($_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage) { $actions += 'Delete' }
                    "$($_.Name)  [$($actions -join ', ')]"
                }
            )
        }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-InboxRule (rule count / suspicious rule scan)')
    }
    try {
        $cas = Invoke-EXOCommand -Name 'Get-CASMailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.OWAEnabled        = $cas.OWAEnabled
        $health.ActiveSyncEnabled = $cas.ActiveSyncEnabled
        $health.PopEnabled        = $cas.PopEnabled
        $health.ImapEnabled       = $cas.ImapEnabled
        $health.MAPIEnabled       = $cas.MAPIEnabled
        $health.EwsEnabled        = $cas.EwsEnabled
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-CASMailbox (protocol access)')
    }
    return [PSCustomObject]$health
}

function Format-MailboxHealthReport {
    param([string]$Smtp, $H)
    $sb = New-Object System.Text.StringBuilder
    $issues = @()
    if ($H.AccountDisabled -eq $true) { $issues += "Account sign-in is disabled (AccountDisabled=True) -- unverified as a Get-Mailbox field, confirm before relying on it." }
    if ($H.ForwardingSmtpAddress -or $H.ForwardingAddress) {
        $fwdTarget = if ($H.ForwardingSmtpAddress) { $H.ForwardingSmtpAddress } else { $H.ForwardingAddress }
        $fwdMode = if ($H.DeliverToMailboxAndForward) { "copy kept in mailbox AND forwarded" } else { "forwarded ONLY -- nothing kept in this mailbox" }
        $issues += "Mail forwarding is configured to '$fwdTarget' ($fwdMode)."
    }
    if ($H.LitigationHoldEnabled -eq $true) { $issues += "Litigation Hold is enabled." }
    if ($H.RetentionHoldEnabled -eq $true) { $issues += "Retention Hold is enabled." }
    if ($H.InPlaceHoldsCount -gt 0) { $issues += "$($H.InPlaceHoldsCount) In-Place/eDiscovery hold(s) applied." }
    if ($H.SuspiciousRules.Count -gt 0) { $issues += "$($H.SuspiciousRules.Count) enabled Inbox rule(s) forward, redirect, or delete mail -- see 'Suspicious Inbox Rules' below." }
    [void]$sb.AppendLine("Mailbox                : $Smtp")
    [void]$sb.AppendLine("=== Potential Issues (review before escalating) ===")
    if ($issues.Count -eq 0) { [void]$sb.AppendLine("  None of the checks below flagged anything.") }
    else { $issues | ForEach-Object { [void]$sb.AppendLine("  - $_") } }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Basics ---")
    [void]$sb.AppendLine("Last Logon             : $($H.LastLogonTime)")
    [void]$sb.AppendLine("Item Count              : $($H.ItemCount)")
    [void]$sb.AppendLine("Mailbox Size            : $($H.MailboxSize)")
    [void]$sb.AppendLine("ProhibitSendQuota       : $($H.ProhibitSendQuota)  (compare to Mailbox Size above)")
    [void]$sb.AppendLine("ProhibitSendReceiveQuota: $($H.ProhibitSendReceiveQuota)")
    [void]$sb.AppendLine("Archive Enabled         : $($H.ArchiveEnabled)")
    if ($H.ArchiveEnabled) { [void]$sb.AppendLine("Archive Size            : $($H.ArchiveSize)") }
    [void]$sb.AppendLine("Account Disabled        : $($H.AccountDisabled)  (unverified Get-Mailbox field -- see Get-MailboxHealthSnapshot help)")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Forwarding & Holds ---")
    [void]$sb.AppendLine("Forwarding SMTP Address : $($H.ForwardingSmtpAddress)")
    [void]$sb.AppendLine("Forwarding Address      : $($H.ForwardingAddress)")
    [void]$sb.AppendLine("Deliver And Forward     : $($H.DeliverToMailboxAndForward)")
    [void]$sb.AppendLine("Litigation Hold         : $($H.LitigationHoldEnabled)")
    [void]$sb.AppendLine("Retention Hold          : $($H.RetentionHoldEnabled)")
    [void]$sb.AppendLine("In-Place Holds          : $($H.InPlaceHoldsCount)")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Delegates ---")
    [void]$sb.AppendLine("Full Access grantees ($($H.FullAccessDelegates.Count)):")
    if ($H.FullAccessDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
    else { $H.FullAccessDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
    [void]$sb.AppendLine("Send As grantees ($($H.SendAsDelegates.Count)):")
    if ($H.SendAsDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
    else { $H.SendAsDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Inbox Rules ---")
    [void]$sb.AppendLine("Inbox Rule Count        : $($H.RuleCount)")
    [void]$sb.AppendLine("Suspicious Inbox Rules (enabled, forward/redirect/delete):")
    if ($H.SuspiciousRules.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
    else { $H.SuspiciousRules | ForEach-Object { [void]$sb.AppendLine("  $_") } }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Protocol Access (enabled/disabled -- not traffic volume) ---")
    [void]$sb.AppendLine("  OWA        : $($H.OWAEnabled)")
    [void]$sb.AppendLine("  ActiveSync : $($H.ActiveSyncEnabled)")
    [void]$sb.AppendLine("  POP        : $($H.PopEnabled)")
    [void]$sb.AppendLine("  IMAP       : $($H.ImapEnabled)")
    [void]$sb.AppendLine("  MAPI       : $($H.MAPIEnabled)")
    [void]$sb.AppendLine("  EWS        : $($H.EwsEnabled)")
    if ($H.Errors.Count -gt 0) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("--- Field errors (partial data above is still valid) ---")
        $H.Errors | ForEach-Object { [void]$sb.AppendLine($_) }
    }
    return $sb.ToString()
}
#endregion

#region Mailbox permissions ------------------------------------------------------
function Get-MailboxState {
    param([string]$Smtp)
    $state = [ordered]@{
        DisplayName          = $null
        PrimarySmtpAddress   = $null
        Alias                = $null
        RecipientTypeDetails = $null
        Hidden               = $null
        LitigationHold       = $null
        ArchiveEnabled       = $null
        ArchiveStatus        = $null
        MailboxSize          = $null
        MailboxItemCount     = $null
        ArchiveSize          = $null
        ArchiveItemCount     = $null
        FullAccess           = @()
        SendAs               = @()
        Errors               = @()
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.DisplayName          = $mbx.DisplayName
        $state.PrimarySmtpAddress   = $mbx.PrimarySmtpAddress
        $state.Alias                = $mbx.Alias
        $state.RecipientTypeDetails = $mbx.RecipientTypeDetails
        $state.Hidden               = $mbx.HiddenFromAddressListsEnabled
        $state.LitigationHold       = $mbx.LitigationHoldEnabled
        $state.ArchiveEnabled       = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        $state.ArchiveStatus        = $mbx.ArchiveStatus
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (core properties)')
    }
    try {
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.MailboxSize      = $stats.TotalItemSize
        $state.MailboxItemCount = $stats.ItemCount
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (mailbox size)')
    }
    if ($state.ArchiveEnabled) {
        try {
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $state.ArchiveSize      = $astats.TotalItemSize
            $state.ArchiveItemCount = $astats.ItemCount
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (archive size)')
        }
    }
    try {
        $fa = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.User -notlike 'NT AUTHORITY\SELF' -and
                $_.IsInherited -eq $false -and
                $_.Deny -eq $false -and
                ($_.AccessRights -contains 'FullAccess')
            }
        $state.FullAccess = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access)')
    }
    try {
        $sa = Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.Trustee -notlike 'NT AUTHORITY\SELF' -and
                $_.AccessControlType -eq 'Allow'
            }
        $state.SendAs = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As)')
    }
    return [PSCustomObject]$state
}

function Format-MailboxState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading (Mailbox) ---")
    [void]$sb.AppendLine("  Display Name        : $($State.DisplayName)")
    [void]$sb.AppendLine("  Primary SMTP        : $($State.PrimarySmtpAddress)")
    [void]$sb.AppendLine("  Alias               : $($State.Alias)")
    [void]$sb.AppendLine("  Mailbox Type        : $($State.RecipientTypeDetails)")
    [void]$sb.AppendLine("  Hidden From GAL     : $($State.Hidden)")
    [void]$sb.AppendLine("  Litigation Hold     : $($State.LitigationHold)")
    [void]$sb.AppendLine("  Mailbox Size        : $($State.MailboxSize)  ($($State.MailboxItemCount) items)")
    [void]$sb.AppendLine("  Archive Enabled     : $($State.ArchiveEnabled)  (ArchiveStatus: $($State.ArchiveStatus))")
    if ($State.ArchiveEnabled) {
        [void]$sb.AppendLine("  Archive Size        : $($State.ArchiveSize)  ($($State.ArchiveItemCount) items)")
    }
    [void]$sb.AppendLine("  Full Access grantees:")
    if ($State.FullAccess.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.FullAccess | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    [void]$sb.AppendLine("  Send As grantees    :")
    if ($State.SendAs.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.SendAs | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    if ($State.Errors.Count -gt 0) {
        [void]$sb.AppendLine("  --- Field errors (partial data above is still valid) ---")
        $State.Errors | ForEach-Object { [void]$sb.AppendLine("  $_") }
    }
    return $sb.ToString()
}

function Grant-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Grant-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}

function Invoke-MailboxPermissionEdit {
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$Grantee,
        [Parameter(Mandatory)][string]$ActionLabel,
        [Parameter(Mandatory)][string]$Verb,
        [Parameter(Mandatory)][scriptblock]$Action
    )
    $out = New-Object System.Text.StringBuilder
    $before = Get-MailboxState -Smtp $Smtp
    [void]$out.Append((Format-MailboxState -Heading 'BEFORE' -State $before))
    & $Action
    [void]$out.AppendLine("$ActionLabel : SUCCESS ($Grantee)")
    $after = Get-MailboxState -Smtp $Smtp
    [void]$out.Append((Format-MailboxState -Heading 'AFTER' -State $after))
    Write-CompletedLog -Task "$Verb $Smtp (grantee: $Grantee)"
    return $out.ToString()
}
#endregion

#region Message trace ------------------------------------------------------------
function Invoke-MessageTrace {
    param(
        [string]$SenderAddress,
        [string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [string]$Subject,
        [string]$Status
    )
    $params = @{
        StartDate   = $StartDate
        EndDate     = $EndDate
        ResultSize  = 5000
        ErrorAction = 'Stop'
    }
    if ($SenderAddress)    { $params.SenderAddress    = $SenderAddress }
    if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }
    if ($Subject) {
        $params.Subject           = $Subject
        $params.SubjectFilterType = 'Contains'
    }
    if ($Status) { $params.Status = $Status }
    Invoke-EXOCommand -Name 'Get-MessageTraceV2' -Parameters $params
}

function Invoke-MessageTraceDetail {
    param(
        [Parameter(Mandatory)][guid]$MessageTraceId,
        [Parameter(Mandatory)][string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )
    Invoke-EXOCommand -Name 'Get-MessageTraceDetailV2' -Parameters @{
        MessageTraceId   = $MessageTraceId
        RecipientAddress = $RecipientAddress
        StartDate        = $StartDate
        EndDate          = $EndDate
        ErrorAction      = 'Stop'
    }
}

function Find-QuarantineMatch {
    param(
        [string]$MessageId,
        [string]$RecipientAddress,
        [datetime]$StartReceivedDate,
        [datetime]$EndReceivedDate
    )
    if (-not $MessageId) {
        throw "This trace result has no Message-ID value to match against quarantine -- cannot look this message up automatically."
    }
    @(Invoke-EXOCommand -Name 'Get-QuarantineMessage' -Parameters @{
        MessageId         = $MessageId
        RecipientAddress  = $RecipientAddress
        StartReceivedDate = $StartReceivedDate
        EndReceivedDate   = $EndReceivedDate
        ErrorAction       = 'Stop'
    })
}

function Get-QuarantineMessageDetail {
    param([Parameter(Mandatory)][string]$Identity)
    $d = [ordered]@{
        Identity         = $Identity
        Subject          = $null
        SenderAddress    = $null
        RecipientAddress = $null
        ReceivedTime     = $null
        Expires          = $null
        Type             = $null
        QuarantineTypes  = $null
        PolicyName       = $null
        PolicyType       = $null
        Direction        = $null
        ReleaseStatus    = $null
        Released         = $null
        ReleasedUsers    = $null
        MessageId        = $null
        Size             = $null
        SpamConfidence   = $null
        Error            = $null
    }
    try {
        $q = Invoke-EXOCommand -Name 'Get-QuarantineMessage' -Parameters @{ Identity = $Identity; ErrorAction = 'Stop' }
        if (-not $q) { $d.Error = "No quarantined message found with identity '$Identity' (it may have been released, deleted, or expired)."; return [PSCustomObject]$d }

        $get = {
            param($obj, $name)
            if ($obj.PSObject.Properties[$name]) { return $obj.$name }
            return $null
        }
        $d.Subject          = & $get $q 'Subject'
        $d.SenderAddress    = & $get $q 'SenderAddress'
        $d.RecipientAddress = @(& $get $q 'RecipientAddress') -join ', '
        $d.ReceivedTime     = & $get $q 'ReceivedTime'
        $d.Expires          = & $get $q 'Expires'
        $d.Type             = & $get $q 'Type'
        $d.QuarantineTypes  = & $get $q 'QuarantineTypes'
        $d.PolicyName       = & $get $q 'PolicyName'
        $d.PolicyType       = & $get $q 'PolicyType'
        $d.Direction        = & $get $q 'Direction'
        $d.ReleaseStatus    = & $get $q 'ReleaseStatus'
        $d.Released         = & $get $q 'Released'
        $d.ReleasedUsers    = @(& $get $q 'ReleasedUsers') -join ', '
        $d.MessageId        = & $get $q 'MessageId'
        $d.Size             = & $get $q 'Size'
        $d.SpamConfidence   = & $get $q 'SpamConfidenceLevel'
    } catch {
        $d.Error = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-QuarantineMessage (detail)'
    }
    return [PSCustomObject]$d
}

function Format-QuarantineMessageDetail {
    param($Detail)
    if ($Detail.Error) { return $Detail.Error }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Subject          : $($Detail.Subject)")
    [void]$sb.AppendLine("From             : $($Detail.SenderAddress)")
    [void]$sb.AppendLine("To               : $($Detail.RecipientAddress)")
    [void]$sb.AppendLine("Received         : $($Detail.ReceivedTime)")
    [void]$sb.AppendLine("Expires from qtn : $($Detail.Expires)")
    [void]$sb.AppendLine("Direction        : $($Detail.Direction)")
    [void]$sb.AppendLine("Quarantine type  : $($Detail.Type)  $(if ($Detail.QuarantineTypes) { "($($Detail.QuarantineTypes))" })")
    [void]$sb.AppendLine("Policy           : $($Detail.PolicyName)  $(if ($Detail.PolicyType) { "[$($Detail.PolicyType)]" })")
    [void]$sb.AppendLine("Spam confidence  : $($Detail.SpamConfidence)")
    [void]$sb.AppendLine("Size             : $($Detail.Size)")
    [void]$sb.AppendLine("Release status   : $($Detail.ReleaseStatus)   Released: $($Detail.Released)")
    if ($Detail.ReleasedUsers) { [void]$sb.AppendLine("Released to      : $($Detail.ReleasedUsers)") }
    [void]$sb.AppendLine("Message-ID       : $($Detail.MessageId)")
    [void]$sb.AppendLine("Quarantine ID    : $($Detail.Identity)")
    return $sb.ToString()
}

function Get-QuarantineMessagePreview {
    param(
        [Parameter(Mandatory)][string]$Identity,
        [string]$RecipientAddress
    )
    $p = @{ Identity = $Identity; ErrorAction = 'Stop' }
    if ($RecipientAddress) { $p.RecipientAddress = $RecipientAddress }
    try {
        $result = Invoke-EXOCommand -Name 'Preview-QuarantineMessage' -Parameters $p
        return [PSCustomObject]@{ Preview = $result; Error = $null }
    } catch {
        return [PSCustomObject]@{
            Preview = $null
            Error   = (Format-ErrorDetail -ErrorRecord $_ -Context 'Preview-QuarantineMessage (note: previewing is not supported for Safe Attachments files from SharePoint/OneDrive/Teams)')
        }
    }
}

function Invoke-QuarantineRelease {
    param([Parameter(Mandatory)]$QuarantineIdentity)
    Invoke-EXOCommand -Name 'Release-QuarantineMessage' -Parameters @{
        Identity     = $QuarantineIdentity
        ReleaseToAll = $true
        Confirm      = $false
        ErrorAction  = 'Stop'
    } | Out-Null
}

function Get-TraceTimeZone {
    param([string]$Label)
    $ids = @{
        'Eastern'  = 'Eastern Standard Time'
        'Central'  = 'Central Standard Time'
        'Mountain' = 'Mountain Standard Time'
        'Pacific'  = 'Pacific Standard Time'
        'UTC'      = 'UTC'
    }
    $id = $ids[$Label]
    if (-not $id) { $id = 'UTC' }
    try {
        return [System.TimeZoneInfo]::FindSystemTimeZoneById($id)
    } catch {
        return [System.TimeZoneInfo]::Utc
    }
}

function ConvertTo-TraceDisplayTime {
    param($UtcValue, [System.TimeZoneInfo]$TimeZone)
    if (-not $UtcValue) { return $UtcValue }
    try {
        $forced = [DateTime]::SpecifyKind([DateTime]$UtcValue, [DateTimeKind]::Utc)
        return [System.TimeZoneInfo]::ConvertTimeFromUtc($forced, $TimeZone)
    } catch {
        return $UtcValue
    }
}

function Format-TraceDetail {
    param($Details, [System.TimeZoneInfo]$TimeZone)
    $sb = New-Object System.Text.StringBuilder
    $rows = @($Details)
    if ($rows.Count -eq 0) {
        [void]$sb.AppendLine("(no detail records returned)")
        return $sb.ToString()
    }
    foreach ($d in $rows) {
        $displayDate = if ($TimeZone) { ConvertTo-TraceDisplayTime -UtcValue $d.Date -TimeZone $TimeZone } else { $d.Date }
        [void]$sb.AppendLine("$displayDate  [$($d.Event)]  $($d.Action)")
        if ($d.Detail) { [void]$sb.AppendLine("    $($d.Detail)") }
    }
    return $sb.ToString()
}
#endregion

#region Calendar permissions ------------------------------------------------------
function Get-CalendarFolderIdentity {
    param([string]$Smtp)
    $folders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{ Identity = $Smtp; FolderScope = 'Calendar'; ErrorAction = 'Stop' })
    if ($folders.Count -eq 0) { throw "No Calendar-scope folder found for $Smtp." }
    $primary = $folders | Where-Object { $_.FolderPath -eq '/Calendar' } | Select-Object -First 1
    if (-not $primary) { $primary = $folders | Select-Object -First 1 }
    return "$($Smtp):\$($primary.Name)"
}

function Get-CalendarPermissions {
    param([string]$Smtp)
    try {
        $calId = Get-CalendarFolderIdentity -Smtp $Smtp
        $perms = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $calId; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ CalendarIdentity = $calId; Permissions = $perms; Error = $null }
    } catch {
        return [PSCustomObject]@{ CalendarIdentity = $null; Permissions = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxFolderPermission (Calendar)') }
    }
}

function Set-CalendarPermission {
    param([string]$CalendarIdentity, [string]$Grantee, [string]$Role)
    $existing = @()
    try {
        $existing = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $CalendarIdentity; User = $Grantee; ErrorAction = 'Stop' })
    } catch {
        $existing = @()
    }
    if ($existing.Count -gt 0) {
        Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Updated'
    } else {
        Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Added'
    }
}

function Remove-CalendarPermission {
    param([string]$CalendarIdentity, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxFolderPermission' -Parameters @{
        Identity = $CalendarIdentity; User = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}

function Get-CalendarPermissionRoles {
    return @('Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author',
             'NonEditingAuthor', 'Reviewer', 'Contributor', 'AvailabilityOnly',
             'LimitedDetails', 'None')
}
#endregion

#region Out of office -------------------------------------------------------------
function Get-OutOfOfficeConfig {
    param([string]$Smtp)
    Invoke-EXOCommand -Name 'Get-MailboxAutoReplyConfiguration' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
}

function Set-OutOfOfficeConfig {
    param(
        [string]$Smtp,
        [string]$AutoReplyState,
        [string]$ExternalAudience,
        [string]$InternalMessage,
        [string]$ExternalMessage,
        [Nullable[DateTime]]$StartTime,
        [Nullable[DateTime]]$EndTime
    )
    $params = @{
        Identity         = $Smtp
        AutoReplyState   = $AutoReplyState
        ExternalAudience = $ExternalAudience
        InternalMessage  = $InternalMessage
        ExternalMessage  = $ExternalMessage
        Confirm          = $false
        ErrorAction      = 'Stop'
    }
    if ($AutoReplyState -eq 'Scheduled') {
        if ($StartTime) { $params.StartTime = $StartTime }
        if ($EndTime)   { $params.EndTime   = $EndTime }
    }
    Invoke-EXOCommand -Name 'Set-MailboxAutoReplyConfiguration' -Parameters $params | Out-Null
}

function Get-OutOfOfficeStates    { return @('Disabled', 'Enabled', 'Scheduled') }
function Get-OutOfOfficeAudiences { return @('None', 'Known', 'All') }
#endregion

#region OneDrive / SharePoint Online ---------------------------------------------
function Get-OneDrivePersonalSiteUrl {
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$SpoTenantPrefix
    )
    $slug = ($Smtp -replace '\.', '_') -replace '@', '_'
    return "https://$SpoTenantPrefix-my.sharepoint.com/personal/$slug"
}

function Set-M365OneDriveSiteAdmin {
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$LoginName,
        [Parameter(Mandatory)][bool]$IsSiteCollectionAdmin
    )
    Set-SPOUser -Site $SiteUrl -LoginName $LoginName -IsSiteCollectionAdmin $IsSiteCollectionAdmin -ErrorAction Stop | Out-Null
}
#endregion

#region Mailbox + OneDrive access (bulk) ------------------------------------------
function Get-M365AccessFlags {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Full', 'SendAs', 'OneDrive', 'FullSendAs', 'FullOneDrive')]
        [string]$AccessType
    )
    switch ($AccessType) {
        'Full'         { return [PSCustomObject]@{ Full = $true;  SendAs = $false; OneDrive = $false } }
        'SendAs'       { return [PSCustomObject]@{ Full = $false; SendAs = $true;  OneDrive = $false } }
        'OneDrive'     { return [PSCustomObject]@{ Full = $false; SendAs = $false; OneDrive = $true  } }
        'FullSendAs'   { return [PSCustomObject]@{ Full = $true;  SendAs = $true;  OneDrive = $false } }
        'FullOneDrive' { return [PSCustomObject]@{ Full = $true;  SendAs = $false; OneDrive = $true  } }
    }
}

function Get-M365AccessTypeText {
    param([Parameter(Mandatory)][string]$AccessType)
    switch ($AccessType) {
        'Full'         { return 'Full Access' }
        'SendAs'       { return 'Send-As' }
        'OneDrive'     { return 'OneDrive' }
        'FullSendAs'   { return 'Full Access & Send-As' }
        'FullOneDrive' { return 'Full Access & OneDrive' }
        default        { return $AccessType }
    }
}

function Invoke-M365MailboxOneDriveAccess {
    param(
        [Parameter(Mandatory)][string[]]$Mailboxes,
        [Parameter(Mandatory)][string[]]$Users,
        [Parameter(Mandatory)]
        [ValidateSet('Full', 'SendAs', 'OneDrive', 'FullSendAs', 'FullOneDrive')]
        [string]$AccessType,
        [Parameter(Mandatory)][ValidateSet('Grant', 'Remove')][string]$Mode,
        [string]$SpoTenantPrefix
    )

    $flags   = Get-M365AccessFlags -AccessType $AccessType
    $isGrant = ($Mode -eq 'Grant')
    $verb    = if ($isGrant) { 'Granted' } else { 'Removed' }
    $actionText = "$Mode $(Get-M365AccessTypeText -AccessType $AccessType)"

    $log = New-Object System.Text.StringBuilder
    [void]$log.AppendLine("=== Mailbox & OneDrive Access ===")
    [void]$log.AppendLine("Action      : $actionText")
    [void]$log.AppendLine("Mailbox(es) : $($Mailboxes -join ', ')")
    [void]$log.AppendLine("User(s)     : $($Users -join ', ')")
    [void]$log.AppendLine("")

    if ($flags.OneDrive -and [string]::IsNullOrWhiteSpace($SpoTenantPrefix)) {
        [void]$log.AppendLine("STOPPED -- this Access Type includes OneDrive but no SharePoint tenant prefix was supplied, and it cannot be guessed safely. Nothing was changed.")
        return [PSCustomObject]@{ Text = $log.ToString(); OneDriveGrantedFor = @(); Succeeded = 0; Failed = 0; Skipped = 0 }
    }

    if ($flags.Full -or $flags.SendAs) {
        $notFound = @()
        foreach ($m in $Mailboxes) {
            try { Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $m; ErrorAction = 'Stop' } | Out-Null }
            catch { $notFound += $m }
        }
        foreach ($u in $Users) {
            try { Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $u; ErrorAction = 'Stop' } | Out-Null }
            catch { $notFound += $u }
        }
        if ($notFound.Count -gt 0) {
            $list = (($notFound | Sort-Object -Unique) -join "`r`n    ")
            [void]$log.AppendLine("PRE-FLIGHT FAILED -- these could not be resolved:")
            [void]$log.AppendLine("    $list")
            [void]$log.AppendLine("")
            [void]$log.AppendLine("No changes were made. Correct the above and try again.")
            return [PSCustomObject]@{ Text = $log.ToString(); OneDriveGrantedFor = @(); Succeeded = 0; Failed = 0; Skipped = 0 }
        }
        [void]$log.AppendLine("Pre-flight  : all identities resolved OK.")
        [void]$log.AppendLine("")
    }

    $succeeded = 0; $failed = 0; $skipped = 0
    $oneDriveOk = New-Object System.Collections.Generic.List[string]

    foreach ($m in $Mailboxes) {

        if ($flags.SendAs -and $isGrant) {
            try {
                Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{ Identity = $m; MessageCopyForSentAs = $true; ErrorAction = 'Stop' } | Out-Null
                [void]$log.AppendLine("  [OK]   MessageCopyForSentAs = true on $m")
            } catch {
                [void]$log.AppendLine("  [WARN] Could not set MessageCopyForSentAs on $m : $($_.Exception.Message)")
            }
        }

        $siteUrl = $null
        if ($flags.OneDrive) {
            $siteUrl = Get-OneDrivePersonalSiteUrl -Smtp $m -SpoTenantPrefix $SpoTenantPrefix
            [void]$log.AppendLine("  OneDrive site for $m : $siteUrl")
        }

        foreach ($u in $Users) {

            if ($flags.Full) {
                if ($isGrant) {
                    try {
                        Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
                            Identity = $m; User = $u; AccessRights = 'FullAccess'; AutoMapping = $true; ErrorAction = 'Stop'
                        } | Out-Null
                        [void]$log.AppendLine("  [OK]   Granted Full Access : $u -> $m")
                        $succeeded++
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Full Access : $u -> $m : $($_.Exception.Message)")
                        $failed++
                    }
                } else {
                    try {
                        $fa = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $m; User = $u; ErrorAction = 'SilentlyContinue' } |
                              Where-Object { $_.AccessRights -contains 'FullAccess' }
                        if ($fa) {
                            Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{
                                Identity = $m; User = $u; AccessRights = 'FullAccess'; InheritanceType = 'All'; Confirm = $false; ErrorAction = 'Stop'
                            } | Out-Null
                            [void]$log.AppendLine("  [OK]   Removed Full Access : $u -> $m")
                            $succeeded++
                        } else {
                            [void]$log.AppendLine("  [SKIP] No Full Access found : $u -> $m")
                            $skipped++
                        }
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Remove Full Access : $u -> $m : $($_.Exception.Message)")
                        $failed++
                    }
                }
            }

            if ($flags.SendAs) {
                if ($isGrant) {
                    try {
                        Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
                            Identity = $m; Trustee = $u; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
                        } | Out-Null
                        [void]$log.AppendLine("  [OK]   Granted Send-As : $u -> $m")
                        $succeeded++
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Send-As : $u -> $m : $($_.Exception.Message)")
                        $failed++
                    }
                } else {
                    try {
                        $sa = Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $m; Trustee = $u; ErrorAction = 'SilentlyContinue' } |
                              Where-Object { $_.AccessRights -contains 'SendAs' }
                        if ($sa) {
                            Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{
                                Identity = $m; Trustee = $u; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
                            } | Out-Null
                            [void]$log.AppendLine("  [OK]   Removed Send-As : $u -> $m")
                            $succeeded++
                        } else {
                            [void]$log.AppendLine("  [SKIP] No Send-As found : $u -> $m")
                            $skipped++
                        }
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Remove Send-As : $u -> $m : $($_.Exception.Message)")
                        $failed++
                    }
                }
            }

            if ($flags.OneDrive) {
                try {
                    Set-M365OneDriveSiteAdmin -SiteUrl $siteUrl -LoginName $u -IsSiteCollectionAdmin $isGrant
                    [void]$log.AppendLine("  [OK]   $verb OneDrive admin : $u -> $m")
                    $succeeded++
                    if ($isGrant -and -not $oneDriveOk.Contains($m)) { $oneDriveOk.Add($m) }
                } catch {
                    [void]$log.AppendLine("  [FAIL] OneDrive admin : $u -> $m : $($_.Exception.Message)")
                    $failed++
                }
            }
        }
    }

    [void]$log.AppendLine("")
    [void]$log.AppendLine("Summary : $succeeded succeeded, $failed failed, $skipped with nothing to do.")

    if ($succeeded -gt 0) {
        $userList = ($Users -join '+')
        foreach ($m in $Mailboxes) {
            Write-CompletedLog -Task "$verb MailboxOneDriveAccess [$actionText] $m (users: $userList)"
        }
    }

    return [PSCustomObject]@{
        Text               = $log.ToString()
        OneDriveGrantedFor = @($oneDriveOk)
        Succeeded          = $succeeded
        Failed             = $failed
        Skipped            = $skipped
    }
}
#endregion

#region Disable / termination ----------------------------------------------------
function Disable-M365User {
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [string]$ManagerSmtp,
        [bool]$GrantMailboxAccess = $false,
        [bool]$GrantOneDriveAccess = $false,
        [string]$SpoTenantPrefix
    )
    $out = New-Object System.Text.StringBuilder
    [void]$out.AppendLine("=== $Smtp ===")

    $objectId = Get-M365EntraObjectId -Smtp $Smtp

    $statusHideShared = 'FAILED'
    try {
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $Smtp; HiddenFromAddressListsEnabled = $true; Type = 'Shared'; ErrorAction = 'Stop'
        } | Out-Null
        $statusHideShared = 'OK'
        [void]$out.AppendLine("  Hide + convert to shared mailbox: OK")
    } catch {
        [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context 'Hide + convert to shared'))
    }

    $statusFullAccess = 'SKIPPED'
    if ($ManagerSmtp -and $GrantMailboxAccess) {
        try {
            Grant-FullAccessPermission -Smtp $Smtp -Grantee $ManagerSmtp
            $statusFullAccess = 'OK'
            [void]$out.AppendLine("  Manager Full Access ($ManagerSmtp): OK")
        } catch {
            $statusFullAccess = 'FAILED'
            [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context 'Manager Full Access'))
        }
    } else {
        $reason = if (-not $ManagerSmtp) { 'no manager supplied' } else { 'mailbox access not requested' }
        [void]$out.AppendLine("  Manager Full Access: SKIPPED ($reason)")
    }

    $statusOneDrive = 'SKIPPED'
    if ($ManagerSmtp -and $GrantOneDriveAccess) {
        if (-not $SpoTenantPrefix) {
            $statusOneDrive = 'FAILED'
            [void]$out.AppendLine("  Manager OneDrive access: FAILED -- OneDrive access was requested but no SharePoint tenant prefix was supplied, and it cannot be guessed safely (see Get-OneDrivePersonalSiteUrl help).")
        } else {
            $odUrl = Get-OneDrivePersonalSiteUrl -Smtp $Smtp -SpoTenantPrefix $SpoTenantPrefix
            try {
                Set-M365OneDriveSiteAdmin -SiteUrl $odUrl -LoginName $ManagerSmtp -IsSiteCollectionAdmin $true
                $statusOneDrive = 'OK'
                [void]$out.AppendLine("  Manager OneDrive access ($ManagerSmtp -> $odUrl): OK")
            } catch {
                $statusOneDrive = 'FAILED'
                [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Manager OneDrive access ($odUrl)"))
            }
        }
    } else {
        $reason = if (-not $ManagerSmtp) { 'no manager supplied' } else { 'OneDrive access not requested' }
        [void]$out.AppendLine("  Manager OneDrive access: SKIPPED ($reason)")
    }

    [void]$out.AppendLine("  License removal: NOT ATTEMPTED -- handle manually.")
    [void]$out.AppendLine("  Sign-in block: NOT ATTEMPTED -- handle manually.")

    $managerLabel = if ($ManagerSmtp) { $ManagerSmtp } else { '(none)' }
    Write-CompletedLog -Task "Disabled $Smtp (manager: $managerLabel; HideShared=$statusHideShared; FullAccess=$statusFullAccess; OneDrive=$statusOneDrive; LicenseRemoved=NOT ATTEMPTED; SignInBlocked=NOT ATTEMPTED)"

    return [PSCustomObject]@{
        Smtp     = $Smtp
        ObjectId = $objectId
        Text     = $out.ToString()
        Statuses = @{
            HideShared = $statusHideShared
            FullAccess = $statusFullAccess
            OneDrive   = $statusOneDrive
        }
    }
}
#endregion
