#Requires -Version 5.1
<#
    Audit.ps1
    Audit-log search for a user mailbox, a shared mailbox, and SharePoint /
    OneDrive files, in one window with a tab for each.

    Author   : Glenn Armbruster
    Created  : 2026-09-20

    MERGED FROM AuditUser.ps1 AND AuditSMB.ps1 on 2026-09-20. Both are on the
    park list; nothing they did was dropped.

    AUDIT SHAREPOINT TAB added 2026-09-29. Same unified audit log, same
    Exchange Online connection -- no SharePoint Online connection is needed.
#>
# TOOLKIT-HELP-RUNS: searches the Microsoft 365 audit log (mailboxes, and SharePoint/OneDrive files), and on the Audit SMB tab reads Full Access permissions. Changes nothing.
# TOOLKIT-HELP-BEGIN
# Searches the audit log for what actually HAPPENED to a mailbox -- who opened
# it, what they did, from where, and when. Two tabs, because a user mailbox and
# a shared mailbox are different questions.
#
# THIS IS NOT A SETTINGS REPORT. It does not tell you what is being recorded;
# it tells you what was recorded. "Who read that message", "did anyone else
# send as this account", "what did the delegate do on Tuesday".
#
#   AUDIT MAILBOX   For one person's mailbox. The search is filtered by
#                   UserIds, so it returns what THAT ACCOUNT did.
#
#   AUDIT SMB       For a shared mailbox. The search is a FreeText search on
#                   the address instead, so it returns everything mentioning
#                   the mailbox whoever did it -- which is the only useful
#                   question on a mailbox six people share. It also lists
#                   everyone holding Full Access before you search, and flags
#                   every result row against that list.
#
#   AUDIT SHAREPOINT  For a SharePoint or OneDrive file, folder or site. Paste
#                   the URL -- the path from the address bar or the file's
#                   Details pane, not a 'Copy link' sharing link. 'This exact
#                   file' finds everything recorded against that one item;
#                   'Anything under this folder or site' finds every file and
#                   sharing event beneath it. Optionally narrow to one person.
#                   Shows who opened, previewed, downloaded, synced, changed,
#                   moved, deleted or shared it, from which IP and app.
#                   Needs only Exchange Online -- the SharePoint connection is
#                   NOT used; this is the same audit log as the other tabs.
#
# THE TWO MAILBOX SEARCHES ARE NOT THE SAME SEARCH and are deliberately not merged.
# Running the shared-mailbox tab against a user mailbox is not wrong, just
# noisier; running the mailbox tab against a shared mailbox will quietly miss
# everything done by the people who share it.
#
# Pick a date range, optionally narrow to particular operations, and search.
#
# THREE THINGS THAT MAKE IT LOOK BROKEN WHEN IT IS NOT:
#
#   * Entries can take 30+ MINUTES to appear. An action performed a moment ago
#     legitimately will not be there yet, which is exactly when somebody tests
#     it and concludes the tool does not work.
#   * Default retention is 180 days, unless Audit (Premium) or a custom
#     retention policy extends it. Older than that is gone, not hidden.
#   * It needs the 'View-Only Audit Logs' or 'Audit Logs' role in Microsoft
#     Purview, on top of an Exchange Online connection.
#
# A wide date range on a busy mailbox takes minutes. It pages through the
# result set and can be stopped part way -- you keep whatever came back, and
# the window says so rather than pretending that was all of it.
#
# Each row keeps its raw JSON alongside the parsed columns. The columns answer
# most questions; the JSON answers the one they do not.
#
# ONE ADDRESS BOX SERVES BOTH TABS. Look Up retargets both and clears both,
# because results left on screen under a new address are results about the
# wrong mailbox.
#
# Read-only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

# NO PROMPT BEFORE THE WINDOW. Each tab carries its OWN address box -- see
# Add-TabTargetRow in Common.ps1. A single form-level bar was tried first and
# was wrong twice: it sat far from the controls it drove, and one box cannot
# serve two tabs asking about two different objects. Resolve-ToolTarget is not
# used either: its job was to decide which DIRECTORY an address belongs to, and
# this tool reads the unified audit log in Microsoft 365, which does not care.
#
# An empty address is fine. The window opens and asks.
# ONE TARGET PER TAB. The two tabs ask about two different objects -- an
# ACCOUNT on one, a SHARED MAILBOX on the other -- so a single shared address
# would silently change meaning when the operator switched tab. Both are
# seeded from the launch parameter; after that they move independently.
$script:TargetMbx = Get-TargetSmtp -Text $SmtpAddress
$script:TargetSmb = $script:TargetMbx

#region Constants ----------------------------------------------------------------
# ONE COPY, SHARED BY BOTH TABS. These were byte-identical in the two scripts
# this file replaces, and two copies of a 31-entry operation map is two places
# for Microsoft to add an operation and only one of them to get it.
$script:MailboxOperations = [ordered]@{
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
$script:OperationFriendlyNames = @{}
foreach ($key in $script:MailboxOperations.Keys) {
    $script:OperationFriendlyNames[$script:MailboxOperations[$key]] = $key
}
$script:AuditResultCeiling = 50000
$script:AuditPageSize = 5000

# SHAREPOINT / ONEDRIVE OPERATIONS. Names taken from Microsoft Learn,
# 'Audit log activities' -- file and page activities, and sharing and access
# request activities (checked 2026-09-29). The friendly names are ours, kept
# short enough for a checklist.
$script:SpoOperations = [ordered]@{
    'Accessed file'                         = 'FileAccessed'
    'Previewed file'                        = 'FilePreviewed'
    'Downloaded file'                       = 'FileDownloaded'
    'Downloaded to computer (sync app)'     = 'FileSyncDownloadedFull'
    'Modified file'                         = 'FileModified'
    'Uploaded file'                         = 'FileUploaded'
    'Uploaded from computer (sync app)'     = 'FileSyncUploadedFull'
    'Copied file'                           = 'FileCopied'
    'Moved file'                            = 'FileMoved'
    'Renamed file'                          = 'FileRenamed'
    'Checked out file'                      = 'FileCheckedOut'
    'Recycled file'                         = 'FileRecycled'
    'Deleted file'                          = 'FileDeleted'
    'Deleted file from recycle bin'         = 'FileDeletedFirstStageRecycleBin'
    'Restored file'                         = 'FileRestored'
    'Shared file, folder or site'           = 'SharingSet'
    'Sharing invitation (external person)'  = 'SharingInvitationCreated'
    'Created anyone link'                   = 'AnonymousLinkCreated'
    'Used anyone link'                      = 'AnonymousLinkUsed'
    'Created company-wide link'             = 'CompanyLinkCreated'
    'Used company-wide link'                = 'CompanyLinkUsed'
    'Created specific-people link'          = 'SecureLinkCreated'
    'Person added to specific-people link'  = 'AddedToSecureLink'
    'Unshared (access removed)'             = 'SharingRevoked'
    'Accepted access request'               = 'AccessRequestAccepted'
}
$script:SpoFriendlyNames = @{}
foreach ($key in $script:SpoOperations.Keys) {
    $script:SpoFriendlyNames[$script:SpoOperations[$key]] = $key
}
#endregion

#region Data layer -----------------------------------------------------------------
# THESE NAMES ARE DELIBERATELY NOT Common.ps1's NAMES.
#
# Common.ps1 exports Invoke-MailboxAuditSearch and ConvertTo-MailboxAuditRecord,
# and this file dot-sources it. A local definition wins, so calls here would
# reach the local one -- but the shadowing was a trap with two edges. A fix
# made in Common would never arrive here, and anyone tidying up an apparent
# duplicate would delete these and leave the calls binding silently to
# Common's versions, which do not take the same parameters. That is a run-time
# failure at the click, on a Saturday.
#
# They are genuinely different functions that happened to share a name, so
# they carry different names. Renamed 2026-09-12; behaviour unchanged, and
# unchanged again by the 2026-09-20 merge.

function Invoke-AuditUserSearch {
    <#
        THE USER-MAILBOX SEARCH: filtered by UserIds.

        UserIds asks 'what did this ACCOUNT do'. That is the right question for
        one person's mailbox and the wrong one for a mailbox several people
        share -- see Invoke-SharedMailboxAuditSearch, which asks the other.
    #>
    param(
        [string]$Mailbox,
        [datetime]$StartUtc,
        [datetime]$EndUtc,
        [string[]]$Operations,
        [scriptblock]$ProgressCallback
    )
    $sessionId = "MailboxAudit_$([Guid]::NewGuid().ToString())"
    $all = New-Object System.Collections.Generic.List[object]
    $hitCeiling = $false
    $iteration = 0
    while ($iteration -lt 15) {
        $iteration++
        $params = @{
            StartDate      = $StartUtc
            EndDate        = $EndUtc
            UserIds        = $Mailbox
            SessionId      = $sessionId
            SessionCommand = 'ReturnLargeSet'
            ResultSize     = $script:AuditPageSize
            ErrorAction    = 'Stop'
        }
        if ($Operations -and $Operations.Count -gt 0) { $params['Operations'] = $Operations }
        $page = @(Search-UnifiedAuditLog @params)
        if ($page.Count -eq 0) { break }
        foreach ($row in $page) { $all.Add($row) }
        if ($ProgressCallback) { & $ProgressCallback $all.Count }
        if ($script:AuditCancelRequested) { break }
        if ($all.Count -ge $script:AuditResultCeiling) { $hitCeiling = $true; break }
        if ($page.Count -lt $script:AuditPageSize) { break }
    }
    return [PSCustomObject]@{
        Records    = $all.ToArray()
        HitCeiling = $hitCeiling
    }
}

function ConvertTo-AuditUserRecord {
    param($RawRecord)
    $parsed = $null
    try { $parsed = $RawRecord.AuditData | ConvertFrom-Json -ErrorAction Stop } catch { $parsed = $null }
    $opValue = if ($parsed -and $parsed.Operation) { "$($parsed.Operation)" } else { "$($RawRecord.Operations)" }
    $friendly = if ($script:OperationFriendlyNames.ContainsKey($opValue)) { $script:OperationFriendlyNames[$opValue] } else { $opValue }
    $creationUtc = if ($parsed -and $parsed.CreationTime) {
        try { [datetime]::Parse($parsed.CreationTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal) } catch { $RawRecord.CreationDate }
    } else { $RawRecord.CreationDate }
    [PSCustomObject]@{
        TimeLocal         = $creationUtc.ToLocalTime()
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

function Get-FullAccessHolders {
    param([string]$Smtp)
    $fa = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
        Where-Object {
            $_.User -notlike 'NT AUTHORITY\SELF' -and
            $_.IsInherited -eq $false -and
            $_.Deny -eq $false -and
            ($_.AccessRights -contains 'FullAccess')
        }
    return @($fa | ForEach-Object { $_.User.ToString() })
}

function Invoke-SharedMailboxAuditSearch {
    <#
        THE SHARED-MAILBOX SEARCH: a FreeText search, NOT a UserIds filter.

        This is the one difference that matters between the two tabs, and the
        reason they are two tabs rather than one with a checkbox. UserIds
        returns what the ACCOUNT did; a shared mailbox almost never does
        anything itself -- six people do things to it, each under their own
        account. FreeText finds the records that mention the address whoever
        the actor was.

        Swapping these would silently return almost nothing on a shared
        mailbox, with no error to say why.
    #>
    param(
        [string]$Mailbox,
        [datetime]$StartUtc,
        [datetime]$EndUtc,
        [string[]]$Operations,
        [scriptblock]$ProgressCallback
    )
    $sessionId = "SharedMailboxAudit_$([Guid]::NewGuid().ToString())"
    $all = New-Object System.Collections.Generic.List[object]
    $hitCeiling = $false
    $iteration = 0
    while ($iteration -lt 15) {
        $iteration++
        $params = @{
            StartDate      = $StartUtc
            EndDate        = $EndUtc
            FreeText       = $Mailbox
            SessionId      = $sessionId
            SessionCommand = 'ReturnLargeSet'
            ResultSize     = $script:AuditPageSize
            ErrorAction    = 'Stop'
        }
        if ($Operations -and $Operations.Count -gt 0) { $params['Operations'] = $Operations }
        $page = @(Search-UnifiedAuditLog @params)
        if ($page.Count -eq 0) { break }
        foreach ($row in $page) { $all.Add($row) }
        if ($ProgressCallback) { & $ProgressCallback $all.Count }
        if ($script:AuditCancelRequested) { break }
        if ($all.Count -ge $script:AuditResultCeiling) { $hitCeiling = $true; break }
        if ($page.Count -lt $script:AuditPageSize) { break }
    }
    return [PSCustomObject]@{
        Records    = $all.ToArray()
        HitCeiling = $hitCeiling
    }
}

function ConvertTo-SharedMailboxAuditRecord {
    # Same shape as ConvertTo-AuditUserRecord plus HolderStatus, which the
    # shared tab fills in after the search by comparing Actor against the Full
    # Access list. Kept as its own function rather than a switch on the other
    # one: they are the record shapes of two different grids.
    param($RawRecord)
    $parsed = $null
    try { $parsed = $RawRecord.AuditData | ConvertFrom-Json -ErrorAction Stop } catch { $parsed = $null }
    $opValue = if ($parsed -and $parsed.Operation) { "$($parsed.Operation)" } else { "$($RawRecord.Operations)" }
    $friendly = if ($script:OperationFriendlyNames.ContainsKey($opValue)) { $script:OperationFriendlyNames[$opValue] } else { $opValue }
    $creationUtc = if ($parsed -and $parsed.CreationTime) {
        try { [datetime]::Parse($parsed.CreationTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal) } catch { $RawRecord.CreationDate }
    } else { $RawRecord.CreationDate }
    [PSCustomObject]@{
        TimeLocal         = $creationUtc.ToLocalTime()
        Operation         = $opValue
        OperationFriendly = $friendly
        Actor             = if ($parsed) { "$($parsed.UserId)" } else { "$($RawRecord.UserIds)" }
        MailboxOwnerUPN   = if ($parsed) { "$($parsed.MailboxOwnerUPN)" } else { '' }
        ClientIP          = if ($parsed) { "$($parsed.ClientIP)" } else { '' }
        ClientInfo        = if ($parsed) { "$($parsed.ClientInfoString)" } else { '' }
        Subject           = if ($parsed -and $parsed.Item -and $parsed.Item.Subject) { "$($parsed.Item.Subject)" } else { '' }
        Folder            = if ($parsed -and $parsed.Item -and $parsed.Item.ParentFolder -and $parsed.Item.ParentFolder.Name) { "$($parsed.Item.ParentFolder.Name)" } else { '' }
        ResultStatus      = if ($parsed) { "$($parsed.ResultStatus)" } else { '' }
        HolderStatus      = ''
        RawJson           = if ($parsed) { ($parsed | ConvertTo-Json -Depth 8) } else { "$($RawRecord.AuditData)" }
    }
}
#endregion

#region SharePoint data layer ------------------------------------------------------
function ConvertTo-SpoAuditUrl {
    <#
        A pasted SharePoint address, in the form the audit log stores it.

        The log keeps the path DECODED -- 'Shared Documents', with a space --
        as Microsoft Learn's own Search-UnifiedAuditLog example shows. A path
        copied from a browser arrives as 'Shared%20Documents', which matches
        nothing. Query strings (?web=1, ?csf=1 ...) and a trailing slash are
        dropped for the same reason.
    #>
    param([string]$Text)
    $u = "$Text".Trim().Trim('"').Trim("'")
    if (-not $u) { return '' }
    $q = $u.IndexOfAny([char[]]@('?', '#'))
    if ($q -ge 0) { $u = $u.Substring(0, $q) }
    try { $u = [System.Uri]::UnescapeDataString($u) } catch { }
    return $u.TrimEnd('/')
}

function Invoke-UalPagedPass {
    <#
        One paged Search-UnifiedAuditLog query, up to the 50,000 ceiling.

        The same loop the two mailbox searches use, with the filter supplied
        by the caller rather than fixed -- the SharePoint tab needs ObjectIds,
        FreeText, RecordType and UserIds in different combinations, and a
        loop per combination is how two copies drift.
    #>
    param(
        [hashtable]$Filter,
        [datetime]$StartUtc,
        [datetime]$EndUtc,
        [scriptblock]$ProgressCallback,
        [int]$AlreadyFetched = 0
    )
    $sessionId = "SpoAudit_$([Guid]::NewGuid().ToString())"
    $all = New-Object System.Collections.Generic.List[object]
    $hitCeiling = $false
    $iteration = 0
    while ($iteration -lt 15) {
        $iteration++
        $params = @{
            StartDate      = $StartUtc
            EndDate        = $EndUtc
            SessionId      = $sessionId
            SessionCommand = 'ReturnLargeSet'
            ResultSize     = $script:AuditPageSize
            ErrorAction    = 'Stop'
        }
        foreach ($k in $Filter.Keys) { $params[$k] = $Filter[$k] }
        $page = @(Search-UnifiedAuditLog @params)
        if ($page.Count -eq 0) { break }
        foreach ($row in $page) { $all.Add($row) }
        if ($ProgressCallback) { & $ProgressCallback ($AlreadyFetched + $all.Count) }
        if ($script:AuditCancelRequested) { break }
        if ($all.Count -ge $script:AuditResultCeiling) { $hitCeiling = $true; break }
        if ($page.Count -lt $script:AuditPageSize) { break }
    }
    return [PSCustomObject]@{ Records = $all.ToArray(); HitCeiling = $hitCeiling }
}

function Invoke-SpoAuditSearch {
    <#
        THE SHAREPOINT SEARCH.

        EXACT FILE: -ObjectIds with the full path. That is Microsoft's own
        example for 'everything about one document', it is not limited to one
        record type, and it is precise. It is also an EXACT string match -- so
        if it finds nothing, the search is repeated as a text search on the
        FILE NAME and the results are kept only where the recorded path is the
        same file. A path that differs by one character in how it was copied
        should not read as 'nobody touched it'.

        UNDER A FOLDER OR SITE: a text search on the path, run once per record
        type -- file operations, then sharing -- because Search-UnifiedAuditLog
        takes ONE record type per call (Microsoft Learn, 'Export, configure and
        view audit log records'). The record type is what keeps mail that
        merely mentions the link out of the results. The caller then keeps
        only rows whose path starts with the one asked about.

        Records seen in more than one pass are kept once.
    #>
    param(
        [string]$Url,
        [bool]$Under,
        [string]$UserId,
        [datetime]$StartUtc,
        [datetime]$EndUtc,
        [string[]]$Operations,
        [scriptblock]$ProgressCallback
    )
    $common = @{}
    if ("$UserId".Trim()) { $common['UserIds'] = @("$UserId".Trim()) }
    if ($Operations -and $Operations.Count -gt 0) { $common['Operations'] = $Operations }

    $passes = @()
    if (-not $Under) {
        $passes += ,(@{ ObjectIds = @($Url) } + $common)
    } else {
        foreach ($rt in @('SharePointFileOperation', 'SharePointSharingOperation')) {
            $passes += ,(@{ FreeText = $Url; RecordType = $rt } + $common)
        }
    }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $out  = New-Object System.Collections.Generic.List[object]
    $ceiling = $false
    $runPasses = {
        param($list)
        foreach ($f in $list) {
            if ($script:AuditCancelRequested) { break }
            $r = Invoke-UalPagedPass -Filter $f -StartUtc $StartUtc -EndUtc $EndUtc -ProgressCallback $ProgressCallback -AlreadyFetched $out.Count
            if ($r.HitCeiling) { $script:SpoPassCeiling = $true }
            foreach ($rec in $r.Records) {
                $key = "$($rec.Identity)"
                if (-not $key) { $key = "$($rec.CreationDate)|$($rec.UserIds)|$($rec.Operations)|$($rec.AuditData.GetHashCode())" }
                if ($seen.Add($key)) { $out.Add($rec) }
            }
        }
    }
    $script:SpoPassCeiling = $false
    & $runPasses $passes

    $usedFallback = $false
    if (-not $Under -and $out.Count -eq 0 -and -not $script:AuditCancelRequested) {
        $leaf = ($Url -split '/')[-1]
        if ($leaf) {
            $usedFallback = $true
            $fb = @()
            foreach ($rt in @('SharePointFileOperation', 'SharePointSharingOperation')) {
                $fb += ,(@{ FreeText = $leaf; RecordType = $rt } + $common)
            }
            & $runPasses $fb
        }
    }
    return [PSCustomObject]@{
        Records      = $out.ToArray()
        HitCeiling   = [bool]$script:SpoPassCeiling
        UsedFallback = $usedFallback
    }
}

function ConvertTo-SpoAuditRecord {
    param($RawRecord)
    $parsed = $null
    try { $parsed = $RawRecord.AuditData | ConvertFrom-Json -ErrorAction Stop } catch { $parsed = $null }
    $opValue = if ($parsed -and $parsed.Operation) { "$($parsed.Operation)" } else { "$($RawRecord.Operations)" }
    $friendly = if ($script:SpoFriendlyNames.ContainsKey($opValue)) { $script:SpoFriendlyNames[$opValue] } else { $opValue }
    $creationUtc = if ($parsed -and $parsed.CreationTime) {
        try { [datetime]::Parse($parsed.CreationTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal) } catch { $RawRecord.CreationDate }
    } else { $RawRecord.CreationDate }
    $objectId = if ($parsed) { ConvertTo-SpoAuditUrl -Text "$($parsed.ObjectId)" } else { '' }
    $file = if ($parsed -and "$($parsed.SourceFileName)".Trim()) { "$($parsed.SourceFileName)" } elseif ($objectId) { ($objectId -split '/')[-1] } else { '' }
    $sharedWith = ''
    if ($parsed) {
        foreach ($f in @('TargetUserOrGroupName', 'UserSharedWith')) {
            if ($parsed.PSObject.Properties[$f] -and "$($parsed.$f)".Trim()) { $sharedWith = "$($parsed.$f)"; break }
        }
    }
    [PSCustomObject]@{
        TimeLocal         = $creationUtc.ToLocalTime()
        Operation         = $opValue
        OperationFriendly = $friendly
        Actor             = if ($parsed) { "$($parsed.UserId)" } else { "$($RawRecord.UserIds)" }
        File              = $file
        Folder            = if ($parsed) { "$($parsed.SourceRelativeUrl)" } else { '' }
        Site              = if ($parsed) { "$($parsed.SiteUrl)" } else { '' }
        SharedWith        = $sharedWith
        ClientIP          = if ($parsed) { "$($parsed.ClientIP)" } else { '' }
        UserAgent         = if ($parsed) { "$($parsed.UserAgent)" } else { '' }
        Workload          = if ($parsed) { "$($parsed.Workload)" } else { '' }
        ObjectId          = $objectId
        RawJson           = if ($parsed) { ($parsed | ConvertTo-Json -Depth 8) } else { "$($RawRecord.AuditData)" }
    }
}
#endregion

#region Shared state and helpers ---------------------------------------------------
# ONE SET OF RESULTS PER TAB. Sharing $script:AuditResults between two grids
# would mean Export and the raw-record double-click acting on whichever tab
# searched last, not the one being looked at.
$script:AuditResultsMbx = @()
$script:AuditResultsSmb = @()
$script:AuditResultsSpo = @()
$script:TargetSpo       = ''
$script:CurrentHolders  = @()
# ONE CANCEL FLAG IS ENOUGH: only one search can be running, because the
# search blocks the UI thread that would have to start the other.
$script:AuditCancelRequested = $false

function Show-AuditRawDetail {
    param([string]$Json, [string]$Title)
    $detailForm = New-Object System.Windows.Forms.Form
    $detailForm.Text = $Title
    $detailForm.Size = New-Object System.Drawing.Size(700, 560)
    $detailForm.StartPosition = "CenterParent"
    $txtDetail = New-Object System.Windows.Forms.TextBox
    $txtDetail.Multiline = $true
    $txtDetail.ReadOnly = $true
    $txtDetail.ScrollBars = "Both"
    $txtDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtDetail.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtDetail.Text = $Json
    $detailForm.Controls.Add($txtDetail)
    [void]$detailForm.ShowDialog()
}

function New-AuditOpsGroup {
    # The operations checklist, built twice with different widths. The list is
    # the same 31 operations on both tabs; the group is not shared because a
    # control belongs to one parent and the two tabs are two parents.
    # -Operations picks the list: the mailbox map by default, the SharePoint
    # map for the SharePoint tab.
    param($Parent, [int]$X, [int]$Y, [int]$W, [System.Collections.IDictionary]$Operations = $script:MailboxOperations)
    $grp = New-Object System.Windows.Forms.GroupBox
    $grp.Text = "Operations (none checked = all operations)"
    $grp.Location = New-Object System.Drawing.Point($X, $Y)
    $grp.Size = New-Object System.Drawing.Size($W, 260)
    $Parent.Controls.Add($grp)
    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(8, 20)
    $clb.Size = New-Object System.Drawing.Size(($W - 18), 200)
    $clb.CheckOnClick = $true
    foreach ($friendlyName in $Operations.Keys) { [void]$clb.Items.Add($friendlyName) }
    $grp.Controls.Add($clb)
    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = "Select All"
    $btnAll.Location = New-Object System.Drawing.Point(8, 226)
    $btnAll.Size = New-Object System.Drawing.Size(90, 26)
    $grp.Controls.Add($btnAll)
    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = "Clear All"
    $btnNone.Location = New-Object System.Drawing.Point(102, 226)
    $btnNone.Size = New-Object System.Drawing.Size(90, 26)
    $grp.Controls.Add($btnNone)
    $btnAll.Add_Click({ for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) } }.GetNewClosure())
    $btnNone.Add_Click({ for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) } }.GetNewClosure())
    return $clb
}

function New-AuditRangeGroup {
    # Start and end pickers. Returns both so the caller can read them.
    param($Parent, [int]$X, [int]$Y)
    $grp = New-Object System.Windows.Forms.GroupBox
    $grp.Text = "Date Range (local time -- converted to UTC for the search)"
    $grp.Location = New-Object System.Drawing.Point($X, $Y)
    $grp.Size = New-Object System.Drawing.Size(400, 66)
    $Parent.Controls.Add($grp)
    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = "Start:"
    $lblStart.Location = New-Object System.Drawing.Point(10, 26)
    $lblStart.AutoSize = $true
    $grp.Controls.Add($lblStart)
    $dtpStart = New-Object System.Windows.Forms.DateTimePicker
    $dtpStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $dtpStart.CustomFormat = 'yyyy-MM-dd HH:mm'
    $dtpStart.Value = (Get-Date).AddDays(-7)
    $dtpStart.Location = New-Object System.Drawing.Point(56, 22)
    $dtpStart.Size = New-Object System.Drawing.Size(150, 22)
    $grp.Controls.Add($dtpStart)
    $lblEnd = New-Object System.Windows.Forms.Label
    $lblEnd.Text = "End:"
    $lblEnd.Location = New-Object System.Drawing.Point(220, 26)
    $lblEnd.AutoSize = $true
    $grp.Controls.Add($lblEnd)
    $dtpEnd = New-Object System.Windows.Forms.DateTimePicker
    $dtpEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $dtpEnd.CustomFormat = 'yyyy-MM-dd HH:mm'
    $dtpEnd.Value = Get-Date
    $dtpEnd.Location = New-Object System.Drawing.Point(258, 22)
    $dtpEnd.Size = New-Object System.Drawing.Size(150, 22)
    $grp.Controls.Add($dtpEnd)
    return [PSCustomObject]@{ Start = $dtpStart; End = $dtpEnd }
}
#endregion

#region Main window ----------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Audit'
$form.ClientSize = New-Object System.Drawing.Size(1120, 870)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

# THE WARNING IS FORM-LEVEL, NOT PER TAB. It is true of both searches, and a
# caveat repeated on every tab is a caveat people stop reading.
$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Requires an active Exchange Online connection AND the 'View-Only Audit Logs' or 'Audit Logs' role assigned to your account in Microsoft Purview. Audit entries can take 30+ minutes to appear after an action happens -- a just-performed action may not show up yet. Default retention is 180 days unless Audit (Premium) or a custom retention policy extends it."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(1092, 48)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, 64)
$tabs.Size = New-Object System.Drawing.Size(1096, 750)
$form.Controls.Add($tabs)

$tabMbx = New-Object System.Windows.Forms.TabPage
$tabMbx.Text = 'Audit Mailbox'
$tabMbx.UseVisualStyleBackColor = $true
$tabs.TabPages.Add($tabMbx)

$tabSmb = New-Object System.Windows.Forms.TabPage
$tabSmb.Text = 'Audit SMB'
$tabSmb.UseVisualStyleBackColor = $true
$tabs.TabPages.Add($tabSmb)

$tabSpo = New-Object System.Windows.Forms.TabPage
$tabSpo.Text = 'Audit SharePoint'
$tabSpo.UseVisualStyleBackColor = $true
$tabs.TabPages.Add($tabSpo)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(1008, 826)
$btnClose.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnClose)
#endregion

#region Audit Mailbox tab ----------------------------------------------------------
$rangeMbx = New-AuditRangeGroup -Parent $tabMbx -X 12 -Y 32
$dtpStartMbx = $rangeMbx.Start
$dtpEndMbx   = $rangeMbx.End

$lblSubjectMbx = New-Object System.Windows.Forms.Label
$lblSubjectMbx.Text = "Subject contains (optional, applied after retrieval):"
$lblSubjectMbx.Location = New-Object System.Drawing.Point(424, 40)
$lblSubjectMbx.AutoSize = $true
$tabMbx.Controls.Add($lblSubjectMbx)
$txtSubjectMbx = New-Object System.Windows.Forms.TextBox
$txtSubjectMbx.Location = New-Object System.Drawing.Point(424, 60)
$txtSubjectMbx.Size = New-Object System.Drawing.Size(300, 22)
$txtSubjectMbx.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabMbx.Controls.Add($txtSubjectMbx)

$clbOpsMbx = New-AuditOpsGroup -Parent $tabMbx -X 742 -Y 32 -W 310

$btnAuditMbx = New-Object System.Windows.Forms.Button
$btnAuditMbx.Text = "Audit Now"
$btnAuditMbx.Location = New-Object System.Drawing.Point(424, 94)
$btnAuditMbx.Size = New-Object System.Drawing.Size(150, 32)
$btnAuditMbx.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$tabMbx.Controls.Add($btnAuditMbx)
$btnCancelMbx = New-Object System.Windows.Forms.Button
$btnCancelMbx.Text = "Cancel"
# CANCEL SITS BESIDE 'AUDIT NOW', NOT ON TOP OF IT.
# The two scripts this replaces stacked them at the same coordinates and
# toggled Visible. That works, but the layout checker reads static positions
# and cannot see Visible, so it reports a real overlap on every run -- and a
# gate that cries wolf is a gate that stops being read. Side by side costs a
# gap while one is hidden and keeps the check meaningful.
$btnCancelMbx.Location = New-Object System.Drawing.Point(584, 94)
$btnCancelMbx.Size = New-Object System.Drawing.Size(150, 32)
$btnCancelMbx.Visible = $false
$tabMbx.Controls.Add($btnCancelMbx)

$lblStatusMbx = New-Object System.Windows.Forms.Label
$lblStatusMbx.Text = "Set a date range and click 'Audit Now.'"
$lblStatusMbx.Location = New-Object System.Drawing.Point(12, 302)
$lblStatusMbx.Size = New-Object System.Drawing.Size(1048, 36)
$lblStatusMbx.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$tabMbx.Controls.Add($lblStatusMbx)

$grdMbx = New-Object System.Windows.Forms.DataGridView
$grdMbx.Location = New-Object System.Drawing.Point(12, 342)
$grdMbx.Size = New-Object System.Drawing.Size(1048, 260)
$grdMbx.ReadOnly = $true
$grdMbx.AllowUserToAddRows = $false
$grdMbx.AllowUserToDeleteRows = $false
$grdMbx.AllowUserToResizeRows = $false
$grdMbx.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grdMbx.MultiSelect = $false
$grdMbx.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grdMbx.Font = New-Object System.Drawing.Font("Consolas", 8)
[void]$grdMbx.Columns.Add('TimeLocal', 'Time (Local)')
[void]$grdMbx.Columns.Add('OperationFriendly', 'Activity')
[void]$grdMbx.Columns.Add('Actor', 'Actor')
[void]$grdMbx.Columns.Add('Subject', 'Subject')
[void]$grdMbx.Columns.Add('Folder', 'Folder')
[void]$grdMbx.Columns.Add('ClientInfo', 'Client')
[void]$grdMbx.Columns.Add('ClientIP', 'Client IP')
[void]$grdMbx.Columns.Add('ResultStatus', 'Result')
$grdMbx.Columns['TimeLocal'].FillWeight         = 14
$grdMbx.Columns['OperationFriendly'].FillWeight = 20
$grdMbx.Columns['Actor'].FillWeight             = 16
$grdMbx.Columns['Subject'].FillWeight           = 20
$grdMbx.Columns['Folder'].FillWeight            = 10
$grdMbx.Columns['ClientInfo'].FillWeight        = 12
$grdMbx.Columns['ClientIP'].FillWeight          = 10
$grdMbx.Columns['ResultStatus'].FillWeight      = 8
foreach ($col in $grdMbx.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
$tabMbx.Controls.Add($grdMbx)

$btnExportMbx = New-Object System.Windows.Forms.Button
$btnExportMbx.Text = "Export Results to CSV"
$btnExportMbx.Location = New-Object System.Drawing.Point(12, 612)
$btnExportMbx.Size = New-Object System.Drawing.Size(170, 28)
$tabMbx.Controls.Add($btnExportMbx)
$lblHintMbx = New-Object System.Windows.Forms.Label
$lblHintMbx.Text = "Double-click a row to see its full raw audit record (fields not yet broken into columns are still in there)."
$lblHintMbx.Location = New-Object System.Drawing.Point(196, 618)
$lblHintMbx.Size = New-Object System.Drawing.Size(700, 20)
$lblHintMbx.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
$lblHintMbx.ForeColor = [System.Drawing.Color]::DimGray
$tabMbx.Controls.Add($lblHintMbx)

$btnAuditMbx.Add_Click({
    if (-not "$script:TargetMbx".Trim()) {
        $lblStatusMbx.Text = 'Enter a mailbox address above and press Look Up first.'
        return
    }
    if ($dtpEndMbx.Value -le $dtpStartMbx.Value) {
        $null = Show-ToolDialog("End date/time must be after the start date/time.", 'Invalid Range', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $selectedOps = @($clbOpsMbx.CheckedItems | ForEach-Object { $script:MailboxOperations[$_] })
    $btnAuditMbx.Visible = $false
    $btnCancelMbx.Visible = $true
    $script:AuditCancelRequested = $false
    $grdMbx.Rows.Clear()
    $script:AuditResultsMbx = @()
    $lblStatusMbx.Text = "Searching..."
    $form.Refresh()
    try {
        $startUtc = $dtpStartMbx.Value.ToUniversalTime()
        $endUtc = $dtpEndMbx.Value.ToUniversalTime()
        $progress = {
            param($count)
            $lblStatusMbx.Text = "Fetched $count record(s) so far..."
            [System.Windows.Forms.Application]::DoEvents()
        }
        $searchResult = Invoke-AuditUserSearch -Mailbox $script:TargetMbx -StartUtc $startUtc -EndUtc $endUtc -Operations $selectedOps -ProgressCallback $progress
        $parsed = @($searchResult.Records | ForEach-Object { ConvertTo-AuditUserRecord -RawRecord $_ })
        $subjectFilter = $txtSubjectMbx.Text.Trim()
        if ($subjectFilter -ne '') {
            $parsed = @($parsed | Where-Object { $_.Subject -like "*$subjectFilter*" })
        }
        $script:AuditResultsMbx = $parsed
        foreach ($rec in $parsed) {
            [void]$grdMbx.Rows.Add($rec.TimeLocal.ToString('yyyy-MM-dd HH:mm:ss'), $rec.OperationFriendly, $rec.Actor, $rec.Subject, $rec.Folder, $rec.ClientInfo, $rec.ClientIP, $rec.ResultStatus)
        }
        $statusMsg = "$($parsed.Count) result(s)."
        if ($script:AuditCancelRequested) { $statusMsg += " Search cancelled -- results may be incomplete." }
        if ($searchResult.HitCeiling) { $statusMsg += " Hit the 50,000-record search ceiling -- narrow the date range or Operations filter to be sure you have everything in this window." }
        if ($parsed.Count -eq 0 -and -not $script:AuditCancelRequested) {
            $statusMsg += " Zero results can mean nothing happened, OR that this falls outside the retention window, OR that mailbox auditing isn't capturing this action for this mailbox's role -- it doesn't necessarily mean 'confirmed nothing happened.'"
        }
        $lblStatusMbx.Text = $statusMsg
        Write-CompletedLog -Task "Ran MailboxAudit search $script:TargetMbx ($($dtpStartMbx.Value.ToString('yyyy-MM-dd HH:mm')) to $($dtpEndMbx.Value.ToString('yyyy-MM-dd HH:mm')); $($selectedOps.Count) operation(s) selected; $($parsed.Count) result(s))"
    } catch {
        $lblStatusMbx.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Search-UnifiedAuditLog'
    } finally {
        $btnAuditMbx.Visible = $true
        $btnCancelMbx.Visible = $false
    }
})
$btnCancelMbx.Add_Click({ $script:AuditCancelRequested = $true })

$grdMbx.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.RowIndex -ge $script:AuditResultsMbx.Count) { return }
    $rec = $script:AuditResultsMbx[$e.RowIndex]
    Show-AuditRawDetail -Json $rec.RawJson -Title "Raw Audit Record: $($rec.OperationFriendly) at $($rec.TimeLocal)"
})

$btnExportMbx.Add_Click({
    if ($script:AuditResultsMbx.Count -eq 0) {
        $null = Show-ToolDialog("Nothing to export -- run a search first.", 'Export Results', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    try {
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $sfd.FileName = "$($script:TargetMbx -replace '[^a-zA-Z0-9]','_')_audit_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $script:AuditResultsMbx | Select-Object TimeLocal, OperationFriendly, Operation, Actor, MailboxOwnerUPN, Subject, Folder, ClientInfo, ClientIP, ResultStatus |
            Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
        $lblStatusMbx.Text = "Exported $($script:AuditResultsMbx.Count) result(s) to $($sfd.FileName)."
    } catch {
        $lblStatusMbx.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Export audit results'
    }
})
#endregion

#region Audit SMB tab --------------------------------------------------------------
$grpHolders = New-Object System.Windows.Forms.GroupBox
$grpHolders.Text = "Current Full Access Holders (used to flag results below)"
$grpHolders.Location = New-Object System.Drawing.Point(12, 32)
$grpHolders.Size = New-Object System.Drawing.Size(300, 138)
$tabSmb.Controls.Add($grpHolders)
$lstHolders = New-Object System.Windows.Forms.ListBox
$lstHolders.Location = New-Object System.Drawing.Point(8, 20)
$lstHolders.Size = New-Object System.Drawing.Size(284, 80)
$lstHolders.Font = New-Object System.Drawing.Font("Consolas", 8)
$grpHolders.Controls.Add($lstHolders)
$btnRefreshHolders = New-Object System.Windows.Forms.Button
$btnRefreshHolders.Text = "Refresh Holders"
$btnRefreshHolders.Location = New-Object System.Drawing.Point(8, 104)
$btnRefreshHolders.Size = New-Object System.Drawing.Size(120, 26)
$grpHolders.Controls.Add($btnRefreshHolders)

$rangeSmb = New-AuditRangeGroup -Parent $tabSmb -X 324 -Y 32
$dtpStartSmb = $rangeSmb.Start
$dtpEndSmb   = $rangeSmb.End

$lblSubjectSmb = New-Object System.Windows.Forms.Label
$lblSubjectSmb.Text = "Subject contains (optional, applied after retrieval):"
$lblSubjectSmb.Location = New-Object System.Drawing.Point(324, 100)
$lblSubjectSmb.AutoSize = $true
$tabSmb.Controls.Add($lblSubjectSmb)
$txtSubjectSmb = New-Object System.Windows.Forms.TextBox
$txtSubjectSmb.Location = New-Object System.Drawing.Point(324, 120)
$txtSubjectSmb.Size = New-Object System.Drawing.Size(300, 22)
$txtSubjectSmb.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabSmb.Controls.Add($txtSubjectSmb)

$clbOpsSmb = New-AuditOpsGroup -Parent $tabSmb -X 742 -Y 32 -W 334

$btnAuditSmb = New-Object System.Windows.Forms.Button
$btnAuditSmb.Text = "Audit Now"
$btnAuditSmb.Location = New-Object System.Drawing.Point(324, 150)
$btnAuditSmb.Size = New-Object System.Drawing.Size(150, 32)
$btnAuditSmb.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$tabSmb.Controls.Add($btnAuditSmb)
$btnCancelSmb = New-Object System.Windows.Forms.Button
$btnCancelSmb.Text = "Cancel"
$btnCancelSmb.Location = New-Object System.Drawing.Point(484, 150)
$btnCancelSmb.Size = New-Object System.Drawing.Size(150, 32)
$btnCancelSmb.Visible = $false
$tabSmb.Controls.Add($btnCancelSmb)

$lblStatusSmb = New-Object System.Windows.Forms.Label
$lblStatusSmb.Text = "Loading Full Access holders..."
$lblStatusSmb.Location = New-Object System.Drawing.Point(12, 302)
$lblStatusSmb.Size = New-Object System.Drawing.Size(1048, 36)
$lblStatusSmb.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$tabSmb.Controls.Add($lblStatusSmb)

$grdSmb = New-Object System.Windows.Forms.DataGridView
$grdSmb.Location = New-Object System.Drawing.Point(12, 342)
$grdSmb.Size = New-Object System.Drawing.Size(1048, 300)
$grdSmb.ReadOnly = $true
$grdSmb.AllowUserToAddRows = $false
$grdSmb.AllowUserToDeleteRows = $false
$grdSmb.AllowUserToResizeRows = $false
$grdSmb.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grdSmb.MultiSelect = $false
$grdSmb.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grdSmb.Font = New-Object System.Drawing.Font("Consolas", 8)
[void]$grdSmb.Columns.Add('TimeLocal', 'Time (Local)')
[void]$grdSmb.Columns.Add('OperationFriendly', 'Activity')
[void]$grdSmb.Columns.Add('Actor', 'Actor')
[void]$grdSmb.Columns.Add('HolderStatus', 'vs. Full Access List')
[void]$grdSmb.Columns.Add('Subject', 'Subject')
[void]$grdSmb.Columns.Add('Folder', 'Folder')
[void]$grdSmb.Columns.Add('ClientInfo', 'Client')
[void]$grdSmb.Columns.Add('ResultStatus', 'Result')
$grdSmb.Columns['TimeLocal'].FillWeight         = 13
$grdSmb.Columns['OperationFriendly'].FillWeight = 17
$grdSmb.Columns['Actor'].FillWeight             = 15
$grdSmb.Columns['HolderStatus'].FillWeight      = 16
$grdSmb.Columns['Subject'].FillWeight           = 17
$grdSmb.Columns['Folder'].FillWeight            = 8
$grdSmb.Columns['ClientInfo'].FillWeight        = 10
$grdSmb.Columns['ResultStatus'].FillWeight      = 7
foreach ($col in $grdSmb.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
$tabSmb.Controls.Add($grdSmb)

$btnExportSmb = New-Object System.Windows.Forms.Button
$btnExportSmb.Text = "Export Results to CSV"
$btnExportSmb.Location = New-Object System.Drawing.Point(12, 652)
$btnExportSmb.Size = New-Object System.Drawing.Size(170, 28)
$tabSmb.Controls.Add($btnExportSmb)
$lblHintSmb = New-Object System.Windows.Forms.Label
$lblHintSmb.Text = "Double-click a row to see its full raw audit record. 'NOT a current holder' can mean revoked access, an admin/owner action, or a permission type other than Full Access."
$lblHintSmb.Location = New-Object System.Drawing.Point(196, 658)
$lblHintSmb.Size = New-Object System.Drawing.Size(770, 20)
$lblHintSmb.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
$lblHintSmb.ForeColor = [System.Drawing.Color]::DimGray
$tabSmb.Controls.Add($lblHintSmb)

function Update-HolderList {
    $lstHolders.Items.Clear()
    if (-not "$script:TargetSmb".Trim()) {
        $script:CurrentHolders = @()
        $lblStatusSmb.Text = 'Enter a shared mailbox address above and press Look Up.'
        return
    }
    try {
        $script:CurrentHolders = Get-FullAccessHolders -Smtp $script:TargetSmb
        if ($script:CurrentHolders.Count -eq 0) {
            [void]$lstHolders.Items.Add('(none found)')
        } else {
            foreach ($h in $script:CurrentHolders) { [void]$lstHolders.Items.Add($h) }
        }
        $lblStatusSmb.Text = "$($script:CurrentHolders.Count) Full Access holder(s) loaded. Set a date range and click 'Audit Now.'"
    } catch {
        $script:CurrentHolders = @()
        $lblStatusSmb.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access holders)'
    }
}
$btnRefreshHolders.Add_Click({ Update-HolderList })

$btnAuditSmb.Add_Click({
    if (-not "$script:TargetSmb".Trim()) {
        $lblStatusSmb.Text = 'Enter a shared mailbox address above and press Look Up first.'
        return
    }
    if ($dtpEndSmb.Value -le $dtpStartSmb.Value) {
        $null = Show-ToolDialog("End date/time must be after the start date/time.", 'Invalid Range', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $selectedOps = @($clbOpsSmb.CheckedItems | ForEach-Object { $script:MailboxOperations[$_] })
    $btnAuditSmb.Visible = $false
    $btnCancelSmb.Visible = $true
    $script:AuditCancelRequested = $false
    $grdSmb.Rows.Clear()
    $script:AuditResultsSmb = @()
    $lblStatusSmb.Text = "Searching..."
    $form.Refresh()
    try {
        $startUtc = $dtpStartSmb.Value.ToUniversalTime()
        $endUtc = $dtpEndSmb.Value.ToUniversalTime()
        $progress = {
            param($count)
            $lblStatusSmb.Text = "Fetched $count record(s) so far..."
            [System.Windows.Forms.Application]::DoEvents()
        }
        $searchResult = Invoke-SharedMailboxAuditSearch -Mailbox $script:TargetSmb -StartUtc $startUtc -EndUtc $endUtc -Operations $selectedOps -ProgressCallback $progress
        $parsed = @($searchResult.Records | ForEach-Object { ConvertTo-SharedMailboxAuditRecord -RawRecord $_ })
        $subjectFilter = $txtSubjectSmb.Text.Trim()
        if ($subjectFilter -ne '') {
            $parsed = @($parsed | Where-Object { $_.Subject -like "*$subjectFilter*" })
        }
        foreach ($rec in $parsed) {
            $isHolder = $false
            foreach ($h in $script:CurrentHolders) {
                if ($h -and $rec.Actor -and $h.ToLowerInvariant() -eq $rec.Actor.ToLowerInvariant()) { $isHolder = $true; break }
            }
            $rec.HolderStatus = if ($isHolder) { 'Full Access holder' } else { 'NOT a current holder' }
        }
        $script:AuditResultsSmb = $parsed
        foreach ($rec in $parsed) {
            [void]$grdSmb.Rows.Add($rec.TimeLocal.ToString('yyyy-MM-dd HH:mm:ss'), $rec.OperationFriendly, $rec.Actor, $rec.HolderStatus, $rec.Subject, $rec.Folder, $rec.ClientInfo, $rec.ResultStatus)
        }
        $statusMsg = "$($parsed.Count) result(s)."
        if ($script:AuditCancelRequested) { $statusMsg += " Search cancelled -- results may be incomplete." }
        if ($searchResult.HitCeiling) { $statusMsg += " Hit the 50,000-record search ceiling -- narrow the date range or Operations filter to be sure you have everything in this window." }
        if ($parsed.Count -eq 0 -and -not $script:AuditCancelRequested) {
            $statusMsg += " Zero results can mean nothing happened, OR that this falls outside the retention window, OR that mailbox auditing isn't capturing this action -- it doesn't necessarily mean 'confirmed nothing happened.'"
        }
        $lblStatusSmb.Text = $statusMsg
        Write-CompletedLog -Task "Ran SharedMailboxAudit search $script:TargetSmb ($($dtpStartSmb.Value.ToString('yyyy-MM-dd HH:mm')) to $($dtpEndSmb.Value.ToString('yyyy-MM-dd HH:mm')); $($selectedOps.Count) operation(s) selected; $($script:CurrentHolders.Count) known holder(s); $($parsed.Count) result(s))"
    } catch {
        $lblStatusSmb.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Search-UnifiedAuditLog'
    } finally {
        $btnAuditSmb.Visible = $true
        $btnCancelSmb.Visible = $false
    }
})
$btnCancelSmb.Add_Click({ $script:AuditCancelRequested = $true })

$grdSmb.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.RowIndex -ge $script:AuditResultsSmb.Count) { return }
    $rec = $script:AuditResultsSmb[$e.RowIndex]
    Show-AuditRawDetail -Json $rec.RawJson -Title "Raw Audit Record: $($rec.OperationFriendly) at $($rec.TimeLocal)"
})

$btnExportSmb.Add_Click({
    if ($script:AuditResultsSmb.Count -eq 0) {
        $null = Show-ToolDialog("Nothing to export -- run a search first.", 'Export Results', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    try {
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $sfd.FileName = "$($script:TargetSmb -replace '[^a-zA-Z0-9]','_')_sharedaudit_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $script:AuditResultsSmb | Select-Object TimeLocal, OperationFriendly, Operation, Actor, HolderStatus, MailboxOwnerUPN, Subject, Folder, ClientInfo, ClientIP, ResultStatus |
            Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
        $lblStatusSmb.Text = "Exported $($script:AuditResultsSmb.Count) result(s) to $($sfd.FileName)."
    } catch {
        $lblStatusSmb.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Export audit results'
    }
})
#endregion

#region Audit SharePoint tab -------------------------------------------------------
$rangeSpo = New-AuditRangeGroup -Parent $tabSpo -X 12 -Y 32
$dtpStartSpo = $rangeSpo.Start
$dtpEndSpo   = $rangeSpo.End
# A LONGER DEFAULT than the mailbox tabs. 'Who opened this file' is usually
# asked about weeks, not days, and a quiet file shows nothing in seven.
$dtpStartSpo.Value = (Get-Date).AddDays(-90)

$grpMatchSpo = New-Object System.Windows.Forms.GroupBox
$grpMatchSpo.Text = 'What to match'
$grpMatchSpo.Location = New-Object System.Drawing.Point(424, 32)
$grpMatchSpo.Size = New-Object System.Drawing.Size(300, 66)
$tabSpo.Controls.Add($grpMatchSpo)
$rbExactSpo = New-Object System.Windows.Forms.RadioButton
$rbExactSpo.Text = 'This exact file'
$rbExactSpo.Location = New-Object System.Drawing.Point(10, 18)
$rbExactSpo.Size = New-Object System.Drawing.Size(280, 20)
$rbExactSpo.Checked = $true
$grpMatchSpo.Controls.Add($rbExactSpo)
$rbUnderSpo = New-Object System.Windows.Forms.RadioButton
$rbUnderSpo.Text = 'Anything under this folder or site'
$rbUnderSpo.Location = New-Object System.Drawing.Point(10, 40)
$rbUnderSpo.Size = New-Object System.Drawing.Size(280, 20)
$grpMatchSpo.Controls.Add($rbUnderSpo)

$lblUserSpo = New-Object System.Windows.Forms.Label
$lblUserSpo.Text = 'Only this person (optional) -- first.last is enough:'
$lblUserSpo.Location = New-Object System.Drawing.Point(12, 108)
$lblUserSpo.AutoSize = $true
$tabSpo.Controls.Add($lblUserSpo)
$txtUserSpo = New-Object System.Windows.Forms.TextBox
$txtUserSpo.Location = New-Object System.Drawing.Point(12, 128)
$txtUserSpo.Size = New-Object System.Drawing.Size(400, 22)
$txtUserSpo.Font = New-Object System.Drawing.Font("Consolas", 9)
$tabSpo.Controls.Add($txtUserSpo)

$clbOpsSpo = New-AuditOpsGroup -Parent $tabSpo -X 742 -Y 32 -W 334 -Operations $script:SpoOperations

$btnAuditSpo = New-Object System.Windows.Forms.Button
$btnAuditSpo.Text = "Audit Now"
$btnAuditSpo.Location = New-Object System.Drawing.Point(424, 118)
$btnAuditSpo.Size = New-Object System.Drawing.Size(150, 32)
$btnAuditSpo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$tabSpo.Controls.Add($btnAuditSpo)
$btnCancelSpo = New-Object System.Windows.Forms.Button
$btnCancelSpo.Text = "Cancel"
$btnCancelSpo.Location = New-Object System.Drawing.Point(584, 118)
$btnCancelSpo.Size = New-Object System.Drawing.Size(140, 32)
$btnCancelSpo.Visible = $false
$tabSpo.Controls.Add($btnCancelSpo)

$lblTipSpo = New-Object System.Windows.Forms.Label
$lblTipSpo.Text = "Paste the file, folder or site path from the address bar or the file's Details pane -- not a 'Copy link' sharing link, which does not contain the path. Needs Exchange Online only; SharePoint Online is not used."
$lblTipSpo.Location = New-Object System.Drawing.Point(12, 162)
$lblTipSpo.Size = New-Object System.Drawing.Size(712, 34)
$lblTipSpo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
$lblTipSpo.ForeColor = [System.Drawing.Color]::DimGray
$tabSpo.Controls.Add($lblTipSpo)

$lblStatusSpo = New-Object System.Windows.Forms.Label
$lblStatusSpo.Text = "Paste a SharePoint or OneDrive URL above, set a date range and click 'Audit Now.'"
$lblStatusSpo.Location = New-Object System.Drawing.Point(12, 302)
$lblStatusSpo.Size = New-Object System.Drawing.Size(1048, 36)
$lblStatusSpo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$tabSpo.Controls.Add($lblStatusSpo)

$grdSpo = New-Object System.Windows.Forms.DataGridView
$grdSpo.Location = New-Object System.Drawing.Point(12, 342)
$grdSpo.Size = New-Object System.Drawing.Size(1048, 300)
$grdSpo.ReadOnly = $true
$grdSpo.AllowUserToAddRows = $false
$grdSpo.AllowUserToDeleteRows = $false
$grdSpo.AllowUserToResizeRows = $false
$grdSpo.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grdSpo.MultiSelect = $false
$grdSpo.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grdSpo.Font = New-Object System.Drawing.Font("Consolas", 8)
[void]$grdSpo.Columns.Add('TimeLocal', 'Time (Local)')
[void]$grdSpo.Columns.Add('OperationFriendly', 'Activity')
[void]$grdSpo.Columns.Add('Actor', 'Who')
[void]$grdSpo.Columns.Add('File', 'File')
[void]$grdSpo.Columns.Add('Folder', 'Folder')
[void]$grdSpo.Columns.Add('SharedWith', 'Shared with')
[void]$grdSpo.Columns.Add('ClientIP', 'Client IP')
[void]$grdSpo.Columns.Add('UserAgent', 'App')
$grdSpo.Columns['TimeLocal'].FillWeight         = 12
$grdSpo.Columns['OperationFriendly'].FillWeight = 15
$grdSpo.Columns['Actor'].FillWeight             = 16
$grdSpo.Columns['File'].FillWeight              = 15
$grdSpo.Columns['Folder'].FillWeight            = 14
$grdSpo.Columns['SharedWith'].FillWeight        = 10
$grdSpo.Columns['ClientIP'].FillWeight          = 9
$grdSpo.Columns['UserAgent'].FillWeight         = 12
foreach ($col in $grdSpo.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
$tabSpo.Controls.Add($grdSpo)

$btnExportSpo = New-Object System.Windows.Forms.Button
$btnExportSpo.Text = "Export Results to CSV"
$btnExportSpo.Location = New-Object System.Drawing.Point(12, 652)
$btnExportSpo.Size = New-Object System.Drawing.Size(170, 28)
$tabSpo.Controls.Add($btnExportSpo)
$lblHintSpo = New-Object System.Windows.Forms.Label
$lblHintSpo.Text = "Double-click a row for the full raw record. 'Accessed' is logged at most once per person per file every five minutes, so a count of rows is not a count of opens."
$lblHintSpo.Location = New-Object System.Drawing.Point(196, 658)
$lblHintSpo.Size = New-Object System.Drawing.Size(860, 20)
$lblHintSpo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
$lblHintSpo.ForeColor = [System.Drawing.Color]::DimGray
$tabSpo.Controls.Add($lblHintSpo)

$btnAuditSpo.Add_Click({
    # THE BOX IS READ HERE, not only on Look Up. A pasted URL is the whole
    # input on this tab; making the operator press Look Up first would be a
    # step that exists only to be forgotten.
    $url = ConvertTo-SpoAuditUrl -Text $rowSpo.TextBox.Text
    if (-not $url) { $lblStatusSpo.Text = 'Paste a SharePoint or OneDrive URL above first.'; return }
    if ($url -notmatch '^(?i)https://[^/]+\.sharepoint\.com/') {
        $lblStatusSpo.Text = "That does not look like a SharePoint or OneDrive address (https://<tenant>.sharepoint.com/... or -my.sharepoint.com/...)."
        return
    }
    $script:TargetSpo = $url
    $rowSpo.TextBox.Text = $url
    if ($dtpEndSpo.Value -le $dtpStartSpo.Value) {
        $null = Show-ToolDialog("End date/time must be after the start date/time.", 'Invalid Range', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $who = ''
    if ("$($txtUserSpo.Text)".Trim()) { $who = Get-TargetSmtp -Text $txtUserSpo.Text }
    $under = [bool]$rbUnderSpo.Checked
    $selectedOps = @($clbOpsSpo.CheckedItems | ForEach-Object { $script:SpoOperations[$_] })
    $btnAuditSpo.Visible = $false
    $btnCancelSpo.Visible = $true
    $script:AuditCancelRequested = $false
    $grdSpo.Rows.Clear()
    $script:AuditResultsSpo = @()
    $lblStatusSpo.Text = "Searching..."
    $form.Refresh()
    try {
        $startUtc = $dtpStartSpo.Value.ToUniversalTime()
        $endUtc = $dtpEndSpo.Value.ToUniversalTime()
        $progress = {
            param($count)
            $lblStatusSpo.Text = "Fetched $count record(s) so far..."
            [System.Windows.Forms.Application]::DoEvents()
        }
        $searchResult = Invoke-SpoAuditSearch -Url $url -Under $under -UserId $who -StartUtc $startUtc -EndUtc $endUtc -Operations $selectedOps -ProgressCallback $progress
        $parsed = @($searchResult.Records | ForEach-Object { ConvertTo-SpoAuditRecord -RawRecord $_ })

        # KEEP ONLY WHAT IS REALLY THIS ITEM. A text search matches the path
        # anywhere in a record, so a folder search could also return a file in
        # a sibling folder whose name starts the same way, and the file-name
        # fallback could return a same-named file elsewhere.
        $prefix = $url.ToLowerInvariant()
        $parsed = @($parsed | Where-Object {
            $o = "$($_.ObjectId)".ToLowerInvariant()
            if (-not $o) { $false }
            elseif ($under) { $o -eq $prefix -or $o.StartsWith($prefix + '/') }
            else { $o -eq $prefix }
        } | Sort-Object TimeLocal -Descending)

        $script:AuditResultsSpo = $parsed
        foreach ($rec in $parsed) {
            [void]$grdSpo.Rows.Add($rec.TimeLocal.ToString('yyyy-MM-dd HH:mm:ss'), $rec.OperationFriendly, $rec.Actor, $rec.File, $rec.Folder, $rec.SharedWith, $rec.ClientIP, $rec.UserAgent)
        }
        $people = @($parsed | ForEach-Object { $_.Actor } | Where-Object { $_ } | Sort-Object -Unique).Count
        $statusMsg = "$($parsed.Count) result(s) from $people person/people."
        if ($searchResult.UsedFallback) { $statusMsg += " The exact-path search found nothing, so it was repeated as a search on the file name and matched back to this path." }
        if ($script:AuditCancelRequested) { $statusMsg += " Search cancelled -- results may be incomplete." }
        if ($searchResult.HitCeiling) { $statusMsg += " Hit the 50,000-record search ceiling -- narrow the date range, person or Operations filter." }
        if ($parsed.Count -eq 0 -and -not $script:AuditCancelRequested) {
            $statusMsg += " Zero results can mean nothing happened, OR it is outside the retention window, OR the path does not match how SharePoint records it -- try the folder above with 'Anything under this folder or site'."
        }
        $lblStatusSpo.Text = $statusMsg
        Write-CompletedLog -Task "Ran SharePoint audit search $url ($(if ($under) { 'under' } else { 'exact' }); $($dtpStartSpo.Value.ToString('yyyy-MM-dd HH:mm')) to $($dtpEndSpo.Value.ToString('yyyy-MM-dd HH:mm')); person '$who'; $($selectedOps.Count) operation(s) selected; $($parsed.Count) result(s))"
    } catch {
        $lblStatusSpo.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Search-UnifiedAuditLog (SharePoint)'
    } finally {
        $btnAuditSpo.Visible = $true
        $btnCancelSpo.Visible = $false
    }
})
$btnCancelSpo.Add_Click({ $script:AuditCancelRequested = $true })

$grdSpo.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.RowIndex -ge $script:AuditResultsSpo.Count) { return }
    # SORTING MOVES ROWS, the results array does not move with them -- so the
    # row is found by its time and activity, not by its index.
    $row = $grdSpo.Rows[$e.RowIndex]
    $t = "$($row.Cells['TimeLocal'].Value)"; $op = "$($row.Cells['OperationFriendly'].Value)"; $who = "$($row.Cells['Actor'].Value)"
    $rec = @($script:AuditResultsSpo | Where-Object { $_.TimeLocal.ToString('yyyy-MM-dd HH:mm:ss') -eq $t -and $_.OperationFriendly -eq $op -and $_.Actor -eq $who })[0]
    if (-not $rec) { $rec = $script:AuditResultsSpo[$e.RowIndex] }
    Show-AuditRawDetail -Json $rec.RawJson -Title "Raw Audit Record: $($rec.OperationFriendly) at $($rec.TimeLocal)"
})

$btnExportSpo.Add_Click({
    if ($script:AuditResultsSpo.Count -eq 0) {
        $null = Show-ToolDialog("Nothing to export -- run a search first.", 'Export Results', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    try {
        $leaf = (("$script:TargetSpo" -split '/')[-1]) -replace '[^a-zA-Z0-9]', '_'
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $sfd.FileName = "$($leaf)_spoaudit_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $script:AuditResultsSpo | Select-Object TimeLocal, OperationFriendly, Operation, Actor, File, Folder, Site, SharedWith, ClientIP, UserAgent, Workload, ObjectId |
            Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
        $lblStatusSpo.Text = "Exported $($script:AuditResultsSpo.Count) result(s) to $($sfd.FileName)."
    } catch {
        $lblStatusSpo.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Export audit results'
    }
})
#endregion

#region Target and wiring ----------------------------------------------------------
function Set-AuditTargetMbx {
    # THE SEARCH IS NOT RE-RUN. It can take minutes and hit a 50,000-record
    # ceiling, so Look Up retargets and clears; 'Audit Now' stays the
    # deliberate act it always was. Clearing the grid matters more than it
    # looks: results left on screen under a new mailbox name are results about
    # the wrong person.
    #
    # ONLY THIS TAB IS TOUCHED. The Audit SMB tab keeps its own address and its
    # own results, so retargeting here cannot silently invalidate what is on
    # screen over there.
    $script:TargetMbx = Get-TargetSmtp -Text $rowMbx.TextBox.Text
    $script:AuditResultsMbx = @()
    $grdMbx.Rows.Clear()
    Update-AuditFormTitle
    $lblStatusMbx.Text = if ("$script:TargetMbx".Trim()) {
        "Set a date range and click 'Audit Now'."
    } else {
        'Enter a mailbox address above and press Look Up.'
    }
}

function Set-AuditTargetSmb {
    # As above, plus: THE HOLDER LIST IS RE-READ. It is one Get-MailboxPermission
    # call, it is what says whether an actor in the results was supposed to be
    # there, and a holder list belonging to the previous mailbox would mislabel
    # every row.
    $script:TargetSmb = Get-TargetSmtp -Text $rowSmb.TextBox.Text
    $script:AuditResultsSmb = @()
    $grdSmb.Rows.Clear()
    Update-AuditFormTitle
    $lblStatusSmb.Text = if ("$script:TargetSmb".Trim()) {
        "Set a date range and click 'Audit Now'."
    } else {
        'Enter a shared mailbox address above and press Look Up.'
    }
    Update-HolderList
}

function Update-AuditFormTitle {
    # Two targets, one title bar. Show whichever are set, so the window in the
    # task bar still says who is being looked at.
    $m = "$script:TargetMbx".Trim()
    $b = "$script:TargetSmb".Trim()
    $form.Text = if ($m -and $b -and $m -ne $b) { "Audit: $m | SMB: $b" }
                 elseif ($m) { "Audit: $m" }
                 elseif ($b) { "Audit SMB: $b" }
                 else { 'Audit' }
}

$btnClose.Add_Click({ $form.Close() })

# ONE ADDRESS ROW PER TAB, each directly above the controls it drives.
# AFTER every control is added to the tab: the row shifts what is already
# there down by its own height, and anything added later would be left behind
# at the old offset. The TabControl and the form were sized with that shift
# already accounted for -- Add-TabTargetRow grows nothing itself.
$rowMbx = Add-TabTargetRow -Tab $tabMbx -Label 'Mailbox or address:' -InitialValue $script:TargetMbx
$rowMbx.Button.Add_Click({ Set-AuditTargetMbx })

$rowSmb = Add-TabTargetRow -Tab $tabSmb -Label 'Shared mailbox address:' -InitialValue $script:TargetSmb
$rowSmb.Button.Add_Click({ Set-AuditTargetSmb })

# THE URL BOX IS WIDE: a SharePoint path is routinely 150 characters.
$rowSpo = Add-TabTargetRow -Tab $tabSpo -Label 'File, folder or site URL:' -ButtonText 'Tidy' -LabelWidth 160 -BoxWidth 790
$rowSpo.Button.Add_Click({
    # 'Tidy', not 'Look Up': there is nothing to look up. It shows the path
    # the way the audit log stores it -- decoded, no query string -- so what
    # will be searched is visible before the search runs.
    $rowSpo.TextBox.Text = ConvertTo-SpoAuditUrl -Text $rowSpo.TextBox.Text
    $script:TargetSpo = $rowSpo.TextBox.Text
    $script:AuditResultsSpo = @()
    $grdSpo.Rows.Clear()
    $lblStatusSpo.Text = $(if ($script:TargetSpo) { "Set a date range and click 'Audit Now'." } else { 'Paste a SharePoint or OneDrive URL above.' })
})

Set-AuditTargetMbx
Set-AuditTargetSmb

$form.StartPosition = 'Manual'
Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Audit' | Out-Null

[void]$form.ShowDialog()
#endregion
