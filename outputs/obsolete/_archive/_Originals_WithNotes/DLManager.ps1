#Requires -Version 5.1
<#
.SYNOPSIS
    Distribution List Manager -- scan, browse, and edit on-prem Exchange
    distribution groups, with delegation guardrails.
.DESCRIPTION
    ON-PREM EXCHANGE ONLY. Unlike the rest of this tool suite, this script
    never calls Get-EXOCommand / Invoke-EXOCommand -- it calls
    Get-DistributionGroup, Get-DistributionGroupMember,
    Add-DistributionGroupMember, and Remove-DistributionGroupMember directly
    by name. Those Invoke-EXOCommand wrappers exist elsewhere in this suite
    specifically to force resolution to the CLOUD-sourced copy of a same-
    named cmdlet when an on-prem module might also be loaded (see
    Get-EXOCommand's header comment in Common.ps1) -- here the goal is the
    exact opposite: this tool wants whatever on-prem Exchange Management
    Shell session/snap-in is already loaded, and has no Exchange Online
    dependency at all. Like the original monolith's own assumption
    ("Connect to Exchange Online before running"), this script assumes the
    operator already has an on-prem Exchange Management Shell session open
    (or the Exchange snap-in loaded) before running it.

    Dot-sources Common.ps1 for Format-ErrorDetail, Write-CompletedLog, the
    Franchise DC list / Get-ADServerString, and Get-TargetSmtp -- but does
    NOT use Resolve-ToolTarget, because this tool isn't scoped to one
    mailbox or group the way the other ~10 tool scripts in this suite are.
    It's a standalone utility, launched directly by the operator, in the
    same spirit as ToolManager.ps1 -- not wired into Tools.json / launched
    from AdminConsole.ps1's per-object button flow. (It CAN be added as a
    Tools.json entry later if you want a button for it too; nothing here
    prevents that, it just isn't done by default since it doesn't target a
    single resolved object.)

    The Environment / Domain Controller picker on the main window is used
    ONLY for the two AD-side (not Exchange-side) lookups this tool makes:
    checking whether a member's AD account is disabled (Get-ADUser) before
    letting it be added, and checking whether a DL is synced to Azure AD via
    AAD Connect (Get-ADGroup -Properties mS-DS-ConsistencyGuid). Every
    Exchange cmdlet call (Get-DistributionGroup, Add/Remove-
    DistributionGroupMember, etc.) uses whatever on-prem Exchange session is
    already active and does not depend on this picker at all.

    KNOWN-UNVERIFIED / FIRST-PASS ITEMS (flagged per standards Section 2 --
    "if verification is unavailable, clearly say so before delivery"):
      - "Synced (AAD Connect)" detection uses the presence of a non-empty
        mS-DS-ConsistencyGuid attribute on the underlying AD group as its
        indicator. This is a well-corroborated community convention (AAD
        Connect populates this attribute, by default since sync engine
        v1.5.18.0, as the sourceAnchor for groups in scope of sync) but it
        is NOT a single official Microsoft Learn "this property always means
        synced" reference the way Get-DistributionGroup's own parameters
        are. Treat a "Synced" label as a strong indicator to double-check
        against your AAD Connect sync scope config on first real use, not
        as an unconditional guarantee.
      - "Removing the last owner" is the one guardrail from the original
        request that has no corresponding action in this scaffold yet --
        the membership actions built here are Add / Remove / Bulk Add /
        Bulk Remove / Export MEMBER, not an owner-editing UI (on-prem DL
        owners are the ManagedBy property, a separate field from
        membership). Test-RemovingLastOwner is included and ready to wire
        up, but nothing currently calls it. Add an "Edit Owners" action to
        Show-DLEditorForm and call it there when that's wanted.
      - Get-DistributionGroup -ResultSize Unlimited is fast (it does not
        enumerate members), but a per-DL member COUNT is not a property
        Get-DistributionGroup exposes -- getting it requires
        Get-DistributionGroupMember per group, which is exactly the kind of
        "iterate everything to get one number" scan that already caused a
        documented UI freeze elsewhere in this suite (see
        Get-MailboxDLMemberships's header in Common.ps1/DLMemberships.ps1
        for that history). This script applies the same lesson: Get-
        DLInventory returns the fast metadata only, with MemberCount blank;
        Start-MemberCountScan fills MemberCount in afterward, one DL at a
        time, with a visible status label and a Cancel button -- same
        progress/cancel shape as Show-DLMembershipsForm.
    REVISED 2026-07-11 (design discussion): a full inventory scan (Get-
    DistributionGroup -ResultSize Unlimited, plus one Get-ADGroup call per DL
    for the Source label) is the expensive part of opening this tool, and DL
    membership/metadata genuinely doesn't change often enough to justify
    paying that cost every single time. 'Scan All Distribution Groups' now
    also writes its result to a JSON cache file (-CacheFilePath, default
    '.\DLInventoryCache.json' next to this script) with a timestamp. On
    startup, if that cache file exists and parses, it's loaded automatically
    instead of starting from an empty grid -- no live Exchange call happens
    until 'Scan All Distribution Groups' is clicked, which always does a
    live rescan and overwrites the cache. If member counts had already been
    loaded before a scan/rescan, they're cached too; a scan alone (before
    'Load Member Counts' is clicked) caches with MemberCount still blank,
    same as it displays in the grid.
    IMPORTANT -- this cache is read-only convenience for the BROWSE/SEARCH
    grid, never for editing: Show-DLEditorForm always re-pulls membership and
    re-runs every guardrail check live when it opens, regardless of whether
    the DL row it opened from came from a fresh scan or a week-old cache. A
    stale DisplayName or member count in the grid is a minor inconvenience;
    a stale guardrail check would not be.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-11
#>
param(
    [string]$Environment,
    [string]$DomainController,
    [string]$CacheFilePath = (Join-Path $PSScriptRoot 'DLInventoryCache.json')
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Constants --------------------------------------------------------------
# Standard Exchange RecipientType values (Microsoft Learn / well-established,
# already relied on elsewhere in this suite -- e.g. Resolve-Target's
# GroupCategory check in Common.ps1) used to classify each member row.
$script:GroupRecipientTypes   = @('MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', 'MailNonUniversalGroup')
$script:ContactRecipientTypes = @('MailContact', 'MailUser')
$script:UserRecipientTypes    = @('UserMailbox', 'RemoteUserMailbox', 'SharedMailbox', 'RemoteSharedMailbox')
#endregion

#region Data layer -- on-prem Exchange (bare cmdlet names -- see header) -------
function Get-DLInventory {
    # Fast pass: metadata only, no per-DL member enumeration. MemberCount is
    # deliberately left $null here -- see header notes -- and is filled in
    # afterward by Start-MemberCountScan.
    $groups = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)
    return @($groups | ForEach-Object {
        [PSCustomObject]@{
            Identity                = $_.Identity.ToString()
            DisplayName             = $_.DisplayName
            PrimarySmtpAddress      = $_.PrimarySmtpAddress.ToString()
            Alias                   = $_.Alias
            GroupType               = if ($_.RecipientTypeDetails -like '*Security*') { 'Security' } else { 'Distribution' }
            ModerationEnabled       = [bool]$_.ModerationEnabled
            HiddenFromAddressLists  = [bool]$_.HiddenFromAddressListsEnabled
            ManagedBy               = @($_.ManagedBy | ForEach-Object { "$_" })
            DistinguishedName       = $_.DistinguishedName
            MemberCount             = $null   # filled in by Start-MemberCountScan
            Source                  = $null   # filled in by Update-DLSourceLabels
        }
    })
}

function Get-DLSourceLabel {
    # See header "KNOWN-UNVERIFIED" note -- mS-DS-ConsistencyGuid presence is
    # a strong, community-corroborated indicator of AAD Connect sync scope,
    # not a single-source-of-truth Microsoft Learn property.
    param([string]$DistinguishedName, [string]$ADServer)
    try {
        $adGroup = Get-ADGroup -Identity $DistinguishedName -Server $ADServer -Properties 'mS-DS-ConsistencyGuid' -ErrorAction Stop
        if ($adGroup.'mS-DS-ConsistencyGuid') { return 'Synced (AAD Connect)' }
        return 'On-Prem Only'
    } catch {
        return "Unknown ($($_.Exception.Message))"
    }
}

function Get-DLMemberCount {
    param([string]$Identity)
    return @(Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop).Count
}

function Get-DLMemberDetail {
    # Returns display-ready member rows. IsDisabledAccount is NOT computed
    # here (it needs a live AD lookup per user -- see
    # Test-AddingDisabledAccount) -- computing it for every existing member
    # of every DL up front would repeat the same "scan cost" mistake flagged
    # in the header notes. It's shown as "?" here and only resolved on
    # demand (Refresh Disabled Status button in the editor).
    param([string]$Identity)
    $members = @(Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop)
    return @($members | ForEach-Object {
        $typeLabel = if ($script:GroupRecipientTypes -contains $_.RecipientType) { 'Group' }
                     elseif ($script:ContactRecipientTypes -contains $_.RecipientType) { 'Contact' }
                     else { 'User' }
        [PSCustomObject]@{
            Identity        = $_.Identity.ToString()
            Name            = $_.DisplayName
            Email           = if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress.ToString() } else { '' }
            RecipientType   = "$($_.RecipientType)"
            TypeLabel       = $typeLabel
            IsNestedGroup   = ($script:GroupRecipientTypes -contains $_.RecipientType)
            IsExternalContact = ($_.RecipientType -eq 'MailContact')
            DisabledStatus  = if ($typeLabel -eq 'User') { '?' } else { 'n/a' }
            SamAccountName  = $_.SamAccountName
        }
    })
}
#endregion

#region Inventory cache -- scan-to-file, load-from-file on startup -------------
function Save-DLInventoryCache {
    param([string]$Path, $DLs)
    $payload = [PSCustomObject]@{
        ScanTimestamp = (Get-Date).ToString('o')
        DLs           = @($DLs)
    }
    $payload | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Read-DLInventoryCache {
    # Returns $null (not a throw) on anything short of a fully-readable,
    # non-empty cache -- callers treat "no usable cache" as equivalent to
    # "no cache file at all" and just fall back to an empty grid.
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ("$raw".Trim() -eq '') { return $null }
        $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
        if (-not $parsed.DLs) { return $null }
        # ConvertFrom-Json collapses a single-item JSON array to a bare
        # value -- re-wrap ManagedBy per DL explicitly, same class of gotcha
        # already documented/handled in ToolManager.ps1's Read-ToolsJson.
        # PSCustomObject properties from ConvertFrom-Json are ordinary
        # settable NoteProperties, so assigning back onto $_ here is safe.
        $dls = @($parsed.DLs | ForEach-Object {
            $_.ManagedBy = @($_.ManagedBy)
            $_
        })
        return [PSCustomObject]@{
            ScanTimestamp = "$($parsed.ScanTimestamp)"
            DLs           = $dls
        }
    } catch {
        return $null
    }
}
#endregion

#region Guardrail checks --------------------------------------------------------
function Test-RemovingLastOwner {
    # NOT YET WIRED TO AN ACTION -- see header notes. Included so an "Edit
    # Owners" feature can call it directly when that's built, rather than
    # re-deriving this logic later.
    param($DL, [string]$OwnerBeingRemoved)
    $owners = @($DL.ManagedBy)
    return ($owners.Count -le 1 -and $owners -contains $OwnerBeingRemoved)
}

function Test-RemovingNestedGroup {
    param($Member)
    return [bool]$Member.IsNestedGroup
}

function Test-AddingDisabledAccount {
    # Live AD lookup -- only called at the moment a specific member is being
    # added (or on-demand refresh in the editor), never during the bulk
    # inventory/member scans. Only meaningful for user-type recipients;
    # returns $false (no warning) for anything else without attempting a
    # lookup that wouldn't apply.
    param([string]$SamAccountName, [string]$RecipientType, [string]$ADServer)
    if ($script:UserRecipientTypes -notcontains $RecipientType) { return $false }
    if (-not $SamAccountName) { return $false }
    try {
        $adUser = Get-ADUser -Identity $SamAccountName -Server $ADServer -Properties Enabled -ErrorAction Stop
        return ($adUser.Enabled -eq $false)
    } catch {
        # Can't confirm -- treat as "unknown" rather than silently claiming
        # enabled. Caller decides how to present this (see Show-DLEditorForm).
        return $null
    }
}

function Test-AddingExternalContact {
    param([string]$RecipientType)
    return ($RecipientType -eq 'MailContact')
}

function Get-SyncedDLWarning {
    # Informational only -- editing a synced DL always happens here (this
    # tool IS on-prem, the authoritative source for a synced object), unlike
    # the EXO-era "view-only, edit on-prem instead" framing this replaced.
    param($DL)
    if ($DL.Source -eq 'Synced (AAD Connect)') {
        return "This DL is synced to Azure AD via AAD Connect. Membership changes made here are authoritative, but will not appear in Exchange Online / Azure AD until the next sync cycle runs."
    }
    return $null
}

function Get-ModeratedDLWarning {
    param($DL)
    if ($DL.ModerationEnabled) {
        return "Moderation is enabled on this DL. Messages sent to it (including from members you add) are held for moderator approval until released -- this does not block membership changes themselves, it's a reminder about mail flow."
    }
    return $null
}

# Centralizes the "block on unacknowledged guardrail, otherwise confirm"
# pattern for every write action below, so Add/Remove/Bulk Add/Bulk Remove
# don't each reimplement it. $Guardrails are DANGEROUS conditions -- if any
# are present and Operator Mode is off, the action is blocked outright.
# $Warnings are informational (synced/moderated) and never block, only
# appear in the confirmation text.
function Confirm-GuardrailedAction {
    param(
        [string]$ActionLabel,
        [string[]]$Warnings = @(),
        [string[]]$Guardrails = @(),
        [bool]$OperatorMode
    )
    if ($Guardrails.Count -gt 0 -and -not $OperatorMode) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Blocked by guardrail(s):`n`n$($Guardrails -join "`n")`n`nEnable 'Operator Mode' below to proceed with actions like this.",
            'Blocked -- Operator Mode Required', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    $lines = @()
    if ($Warnings.Count -gt 0) { $lines += $Warnings }
    if ($Guardrails.Count -gt 0) { $lines += ($Guardrails | ForEach-Object { "GUARDRAIL (Operator Mode acknowledged): $_" }) }
    $lines += ''
    $lines += "Proceed with $ActionLabel?"
    $icon = if ($Guardrails.Count -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Question }
    $result = [System.Windows.Forms.MessageBox]::Show(
        ($lines -join "`n"), "Confirm $ActionLabel", [System.Windows.Forms.MessageBoxButtons]::YesNo, $icon)
    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}
#endregion

#region Actions -- writes -------------------------------------------------------
function Add-DLMember {
    param([string]$Identity, [string]$MemberIdentity, [string]$ADServer)
    Add-DistributionGroupMember -Identity $Identity -Member $MemberIdentity -Confirm:$false -DomainController $ADServer -ErrorAction Stop
}
function Remove-DLMember {
    param([string]$Identity, [string]$MemberIdentity, [string]$ADServer)
    Remove-DistributionGroupMember -Identity $Identity -Member $MemberIdentity -Confirm:$false -DomainController $ADServer -ErrorAction Stop
}
function Add-DLMemberBulk {
    # Collects per-item success/failure rather than stopping at the first
    # error -- same "don't let one bad row hide the other nine" convention
    # used throughout this suite (e.g. Get-MailboxHealthSnapshot's Errors
    # array).
    param([string]$Identity, [string[]]$MemberIdentities, [string]$ADServer)
    $results = @()
    foreach ($m in $MemberIdentities) {
        $m = $m.Trim()
        if ($m -eq '') { continue }
        try {
            Add-DLMember -Identity $Identity -MemberIdentity $m -ADServer $ADServer
            $results += [PSCustomObject]@{ Member = $m; Success = $true; Error = $null }
        } catch {
            $results += [PSCustomObject]@{ Member = $m; Success = $false; Error = (Format-ErrorDetail -ErrorRecord $_ -Context "Add-DistributionGroupMember ($m)") }
        }
    }
    return $results
}
function Remove-DLMemberBulk {
    param([string]$Identity, [string[]]$MemberIdentities, [string]$ADServer)
    $results = @()
    foreach ($m in $MemberIdentities) {
        $m = $m.Trim()
        if ($m -eq '') { continue }
        try {
            Remove-DLMember -Identity $Identity -MemberIdentity $m -ADServer $ADServer
            $results += [PSCustomObject]@{ Member = $m; Success = $true; Error = $null }
        } catch {
            $results += [PSCustomObject]@{ Member = $m; Success = $false; Error = (Format-ErrorDetail -ErrorRecord $_ -Context "Remove-DistributionGroupMember ($m)") }
        }
    }
    return $results
}
function Export-DLMembersToCsv {
    param($Members, [string]$Path)
    $Members | Select-Object Name, Email, TypeLabel, IsNestedGroup, IsExternalContact, DisabledStatus |
        Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}
#endregion

#region DL Editor window ----------------------------------------------------------
# Defined BEFORE the main window below on purpose (matching every other
# script in this suite -- e.g. MailboxHealth.ps1 defines its Show-XForm
# before calling it, GroupEditor.ps1 defines Show-GroupEditorForm before its
# trailing call). PowerShell does hoist top-level function definitions
# across an entire script, so calling Show-DLEditorForm from inside the main
# window's button handlers below would technically still resolve even with
# the function defined later in the file -- but this codebase's convention
# is "define before use," and following it here avoids being the one file
# that relies on hoisting while everything else doesn't.
function Show-DLEditorForm {
    param($DL, [string]$ADServer)
    $editForm = New-Object System.Windows.Forms.Form
    $editForm.Text = "DL Editor: $($DL.DisplayName)"
    $editForm.Size = New-Object System.Drawing.Size(900, 760)
    $editForm.StartPosition = "CenterScreen"
    $editForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

    $grpMeta = New-Object System.Windows.Forms.GroupBox
    $grpMeta.Text = "DL Metadata"
    $grpMeta.Location = New-Object System.Drawing.Point(12, 12)
    $grpMeta.Size = New-Object System.Drawing.Size(860, 96)
    $editForm.Controls.Add($grpMeta)
    $lblMeta = New-Object System.Windows.Forms.Label
    $lblMeta.Text = "Display Name : $($DL.DisplayName)`r`nAlias        : $($DL.Alias)`r`nPrimary SMTP : $($DL.PrimarySmtpAddress)`r`nSource       : $($DL.Source)`r`nModeration   : $(if ($DL.ModerationEnabled) {'Enabled'} else {'Disabled'})`r`nHidden From GAL: $($DL.HiddenFromAddressLists)`r`nGroup Type   : $($DL.GroupType)   Owners (ManagedBy): $(($DL.ManagedBy -join ', '))"
    $lblMeta.Location = New-Object System.Drawing.Point(10, 16)
    $lblMeta.Size = New-Object System.Drawing.Size(840, 76)
    $lblMeta.Font = New-Object System.Drawing.Font("Consolas", 8)
    $grpMeta.Controls.Add($lblMeta)

    $lblWarnBanner = New-Object System.Windows.Forms.Label
    $warnings = @((Get-SyncedDLWarning -DL $DL), (Get-ModeratedDLWarning -DL $DL)) | Where-Object { $_ }
    $lblWarnBanner.Text = if ($warnings.Count -gt 0) { ($warnings -join "`r`n`r`n") } else { "No standing warnings for this DL." }
    $lblWarnBanner.Location = New-Object System.Drawing.Point(12, 114)
    $lblWarnBanner.Size = New-Object System.Drawing.Size(860, 50)
    $lblWarnBanner.ForeColor = if ($warnings.Count -gt 0) { [System.Drawing.Color]::DarkOrange } else { [System.Drawing.Color]::DimGray }
    $lblWarnBanner.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
    $editForm.Controls.Add($lblWarnBanner)

    $grdMembers = New-Object System.Windows.Forms.DataGridView
    $grdMembers.Location = New-Object System.Drawing.Point(12, 170)
    $grdMembers.Size = New-Object System.Drawing.Size(860, 300)
    $grdMembers.ReadOnly = $true
    $grdMembers.AllowUserToAddRows = $false
    $grdMembers.AllowUserToDeleteRows = $false
    $grdMembers.AllowUserToResizeRows = $false
    $grdMembers.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdMembers.MultiSelect = $true
    $grdMembers.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdMembers.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdMembers.Columns.Add('Name', 'Name')
    [void]$grdMembers.Columns.Add('Email', 'Email')
    [void]$grdMembers.Columns.Add('TypeLabel', 'Type')
    [void]$grdMembers.Columns.Add('Flags', 'Flags')
    [void]$grdMembers.Columns.Add('Identity', 'Identity')
    $grdMembers.Columns['Name'].FillWeight      = 26
    $grdMembers.Columns['Email'].FillWeight     = 32
    $grdMembers.Columns['TypeLabel'].FillWeight = 14
    $grdMembers.Columns['Flags'].FillWeight     = 28
    $grdMembers.Columns['Identity'].Visible     = $false
    foreach ($col in $grdMembers.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
    $editForm.Controls.Add($grdMembers)

    $script:EditorMembers = @()
    function Update-MembersGrid {
        $grdMembers.Rows.Clear()
        $script:EditorMembers = Get-DLMemberDetail -Identity $DL.Identity
        foreach ($m in $script:EditorMembers) {
            $flags = @()
            if ($m.IsNestedGroup) { $flags += 'NESTED GROUP' }
            if ($m.IsExternalContact) { $flags += 'EXTERNAL CONTACT' }
            if ($m.DisabledStatus -eq $true) { $flags += 'DISABLED ACCOUNT' }
            elseif ($m.DisabledStatus -eq '?') { $flags += '(disabled status unchecked)' }
            [void]$grdMembers.Rows.Add($m.Name, $m.Email, $m.TypeLabel, ($flags -join ', '), $m.Identity)
        }
    }

    $lblOutStatus = New-Object System.Windows.Forms.Label
    $lblOutStatus.Text = ""
    $lblOutStatus.Location = New-Object System.Drawing.Point(12, 476)
    $lblOutStatus.Size = New-Object System.Drawing.Size(860, 40)
    $editForm.Controls.Add($lblOutStatus)

    $grpSingle = New-Object System.Windows.Forms.GroupBox
    $grpSingle.Text = "Add / Remove One Member"
    $grpSingle.Location = New-Object System.Drawing.Point(12, 518)
    $grpSingle.Size = New-Object System.Drawing.Size(860, 66)
    $editForm.Controls.Add($grpSingle)
    $lblGrantee = New-Object System.Windows.Forms.Label
    $lblGrantee.Text = "Member (email or logon):"
    $lblGrantee.Location = New-Object System.Drawing.Point(10, 26)
    $lblGrantee.AutoSize = $true
    $grpSingle.Controls.Add($lblGrantee)
    $txtGrantee = New-Object System.Windows.Forms.TextBox
    $txtGrantee.Location = New-Object System.Drawing.Point(170, 23)
    $txtGrantee.Size = New-Object System.Drawing.Size(260, 22)
    $txtGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grpSingle.Controls.Add($txtGrantee)
    $btnAddOne = New-Object System.Windows.Forms.Button
    $btnAddOne.Text = "Add Member"
    $btnAddOne.Location = New-Object System.Drawing.Point(440, 20)
    $btnAddOne.Size = New-Object System.Drawing.Size(120, 28)
    $grpSingle.Controls.Add($btnAddOne)
    $btnRemoveSelected = New-Object System.Windows.Forms.Button
    $btnRemoveSelected.Text = "Remove Selected"
    $btnRemoveSelected.Location = New-Object System.Drawing.Point(566, 20)
    $btnRemoveSelected.Size = New-Object System.Drawing.Size(140, 28)
    $grpSingle.Controls.Add($btnRemoveSelected)
    $btnRefreshDisabled = New-Object System.Windows.Forms.Button
    $btnRefreshDisabled.Text = "Check Disabled"
    $btnRefreshDisabled.Location = New-Object System.Drawing.Point(712, 20)
    $btnRefreshDisabled.Size = New-Object System.Drawing.Size(130, 28)
    $grpSingle.Controls.Add($btnRefreshDisabled)

    $grpBulk = New-Object System.Windows.Forms.GroupBox
    $grpBulk.Text = "Bulk Add / Bulk Remove (one address per line)"
    $grpBulk.Location = New-Object System.Drawing.Point(12, 590)
    $grpBulk.Size = New-Object System.Drawing.Size(860, 110)
    $editForm.Controls.Add($grpBulk)
    $txtBulk = New-Object System.Windows.Forms.TextBox
    $txtBulk.Multiline = $true
    $txtBulk.ScrollBars = "Vertical"
    $txtBulk.Location = New-Object System.Drawing.Point(10, 20)
    $txtBulk.Size = New-Object System.Drawing.Size(600, 80)
    $txtBulk.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grpBulk.Controls.Add($txtBulk)
    $btnBulkAdd = New-Object System.Windows.Forms.Button
    $btnBulkAdd.Text = "Bulk Add"
    $btnBulkAdd.Location = New-Object System.Drawing.Point(620, 20)
    $btnBulkAdd.Size = New-Object System.Drawing.Size(220, 28)
    $grpBulk.Controls.Add($btnBulkAdd)
    $btnBulkRemove = New-Object System.Windows.Forms.Button
    $btnBulkRemove.Text = "Bulk Remove"
    $btnBulkRemove.Location = New-Object System.Drawing.Point(620, 54)
    $btnBulkRemove.Size = New-Object System.Drawing.Size(220, 28)
    $grpBulk.Controls.Add($btnBulkRemove)

    $chkOperatorMode = New-Object System.Windows.Forms.CheckBox
    $chkOperatorMode.Text = "Operator Mode (required to proceed past a guardrail warning)"
    $chkOperatorMode.Location = New-Object System.Drawing.Point(12, 706)
    $chkOperatorMode.Size = New-Object System.Drawing.Size(500, 22)
    $chkOperatorMode.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $editForm.Controls.Add($chkOperatorMode)

    $btnExportEditor = New-Object System.Windows.Forms.Button
    $btnExportEditor.Text = "Export Members"
    $btnExportEditor.Location = New-Object System.Drawing.Point(560, 704)
    $btnExportEditor.Size = New-Object System.Drawing.Size(140, 26)
    $editForm.Controls.Add($btnExportEditor)
    $btnEditorClose = New-Object System.Windows.Forms.Button
    $btnEditorClose.Text = "Close"
    $btnEditorClose.Location = New-Object System.Drawing.Point(772, 704)
    $btnEditorClose.Size = New-Object System.Drawing.Size(100, 26)
    $editForm.Controls.Add($btnEditorClose)

    Update-MembersGrid

    $btnAddOne.Add_Click({
        $grantee = $txtGrantee.Text.Trim()
        if (-not $grantee) { $lblOutStatus.Text = "Enter a member (email or logon) first."; return }
        $guardrails = @()
        $disabled = Test-AddingDisabledAccount -SamAccountName $grantee -RecipientType 'UserMailbox' -ADServer $ADServer
        # NOTE: RecipientType is assumed 'UserMailbox' here since the operator
        # is typing a raw identity, not selecting from a known-typed row --
        # this only affects whether the disabled-account guardrail check
        # runs at all (Test-AddingDisabledAccount short-circuits for non-user
        # types); it does not affect what actually gets added.
        if ($disabled -eq $true) { $guardrails += "The account '$grantee' appears to be DISABLED in AD." }
        $warnings = @((Get-SyncedDLWarning -DL $DL), (Get-ModeratedDLWarning -DL $DL)) | Where-Object { $_ }
        if (-not (Confirm-GuardrailedAction -ActionLabel "Add Member '$grantee'" -Warnings $warnings -Guardrails $guardrails -OperatorMode $chkOperatorMode.Checked)) {
            $lblOutStatus.Text = "Cancelled."
            return
        }
        try {
            Add-DLMember -Identity $DL.Identity -MemberIdentity $grantee -ADServer $ADServer
            Write-CompletedLog -Task "Added DL Member $grantee to $($DL.PrimarySmtpAddress) (On-Prem DL Manager; $(Get-SelectedEnvLabel))"
            $lblOutStatus.Text = "Added '$grantee'."
            Update-MembersGrid
        } catch {
            $lblOutStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add-DistributionGroupMember'
        }
    })

    $btnRemoveSelected.Add_Click({
        if ($grdMembers.SelectedRows.Count -eq 0) { $lblOutStatus.Text = "Select one or more members first."; return }
        $selected = @($grdMembers.SelectedRows | ForEach-Object {
            $identity = $_.Cells['Identity'].Value
            $script:EditorMembers | Where-Object { $_.Identity -eq $identity } | Select-Object -First 1
        })
        $guardrails = @()
        $nested = @($selected | Where-Object { $_.IsNestedGroup })
        if ($nested.Count -gt 0) { $guardrails += "Removing nested group member(s): $(($nested | ForEach-Object { $_.Name }) -join ', ')" }
        $names = ($selected | ForEach-Object { $_.Name }) -join ', '
        $warnings = @((Get-SyncedDLWarning -DL $DL), (Get-ModeratedDLWarning -DL $DL)) | Where-Object { $_ }
        if (-not (Confirm-GuardrailedAction -ActionLabel "Remove $($selected.Count) member(s): $names" -Warnings $warnings -Guardrails $guardrails -OperatorMode $chkOperatorMode.Checked)) {
            $lblOutStatus.Text = "Cancelled."
            return
        }
        $results = Remove-DLMemberBulk -Identity $DL.Identity -MemberIdentities ($selected | ForEach-Object { $_.Email }) -ADServer $ADServer
        foreach ($r in ($results | Where-Object { $_.Success })) {
            Write-CompletedLog -Task "Removed DL Member $($r.Member) from $($DL.PrimarySmtpAddress) (On-Prem DL Manager; $(Get-SelectedEnvLabel))"
        }
        $failed = @($results | Where-Object { -not $_.Success })
        $lblOutStatus.Text = "Removed $($results.Count - $failed.Count) of $($results.Count). $(if ($failed.Count -gt 0) { "$($failed.Count) failed -- see log/errors." })"
        Update-MembersGrid
    })

    $btnRefreshDisabled.Add_Click({
        foreach ($m in $script:EditorMembers) {
            if ($m.TypeLabel -eq 'User') {
                $m.DisabledStatus = Test-AddingDisabledAccount -SamAccountName $m.SamAccountName -RecipientType $m.RecipientType -ADServer $ADServer
            }
        }
        Update-MembersGrid
        $lblOutStatus.Text = "Disabled-account status refreshed."
    })

    $btnBulkAdd.Add_Click({
        $lines = @($txtBulk.Text -split "`r`n|`n" | Where-Object { $_.Trim() -ne '' })
        if ($lines.Count -eq 0) { $lblOutStatus.Text = "Enter at least one address to add."; return }
        $warnings = @((Get-SyncedDLWarning -DL $DL), (Get-ModeratedDLWarning -DL $DL)) | Where-Object { $_ }
        $guardrails = @("Bulk adding $($lines.Count) member(s) -- disabled-account and external-contact checks are NOT run per-address for bulk operations (would require one AD/Exchange lookup per line before any writes happen). Review the list above carefully.")
        if (-not (Confirm-GuardrailedAction -ActionLabel "Bulk Add $($lines.Count) member(s)" -Warnings $warnings -Guardrails $guardrails -OperatorMode $chkOperatorMode.Checked)) {
            $lblOutStatus.Text = "Cancelled."
            return
        }
        $results = Add-DLMemberBulk -Identity $DL.Identity -MemberIdentities $lines -ADServer $ADServer
        foreach ($r in ($results | Where-Object { $_.Success })) {
            Write-CompletedLog -Task "Bulk-added DL Member $($r.Member) to $($DL.PrimarySmtpAddress) (On-Prem DL Manager; $(Get-SelectedEnvLabel))"
        }
        $failed = @($results | Where-Object { -not $_.Success })
        $lblOutStatus.Text = "Bulk add: $($results.Count - $failed.Count) succeeded, $($failed.Count) failed."
        if ($failed.Count -gt 0) {
            $lblOutStatus.Text += " First failure: $($failed[0].Error)"
        }
        Update-MembersGrid
    })

    $btnBulkRemove.Add_Click({
        $lines = @($txtBulk.Text -split "`r`n|`n" | Where-Object { $_.Trim() -ne '' })
        if ($lines.Count -eq 0) { $lblOutStatus.Text = "Enter at least one address to remove."; return }
        $warnings = @((Get-SyncedDLWarning -DL $DL), (Get-ModeratedDLWarning -DL $DL)) | Where-Object { $_ }
        $guardrails = @("Bulk removing $($lines.Count) member(s) -- nested-group checks are NOT run per-address for bulk operations. Review the list above carefully.")
        if (-not (Confirm-GuardrailedAction -ActionLabel "Bulk Remove $($lines.Count) member(s)" -Warnings $warnings -Guardrails $guardrails -OperatorMode $chkOperatorMode.Checked)) {
            $lblOutStatus.Text = "Cancelled."
            return
        }
        $results = Remove-DLMemberBulk -Identity $DL.Identity -MemberIdentities $lines -ADServer $ADServer
        foreach ($r in ($results | Where-Object { $_.Success })) {
            Write-CompletedLog -Task "Bulk-removed DL Member $($r.Member) from $($DL.PrimarySmtpAddress) (On-Prem DL Manager; $(Get-SelectedEnvLabel))"
        }
        $failed = @($results | Where-Object { -not $_.Success })
        $lblOutStatus.Text = "Bulk remove: $($results.Count - $failed.Count) succeeded, $($failed.Count) failed."
        if ($failed.Count -gt 0) {
            $lblOutStatus.Text += " First failure: $($failed[0].Error)"
        }
        Update-MembersGrid
    })

    $btnExportEditor.Add_Click({
        try {
            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = "CSV files (*.csv)|*.csv"
            $sfd.FileName = "$($DL.Alias -replace '[^a-zA-Z0-9]','_')_members_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
            if ($sfd.ShowDialog($editForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }
            Export-DLMembersToCsv -Members $script:EditorMembers -Path $sfd.FileName
            $lblOutStatus.Text = "Exported $($script:EditorMembers.Count) member(s) to $($sfd.FileName)."
        } catch {
            $lblOutStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Export member list'
        }
    })

    $btnEditorClose.Add_Click({ $editForm.Close() })
    # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
    # Console, which opens top-left. Helper is in Common.ps1.
    Set-FormScreenCorner -Form $editForm -Corner 'BottomLeft'
    [void]$editForm.ShowDialog()
}
#endregion

#region Main window --------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Distribution List Manager (On-Prem Exchange)"
$form.Size = New-Object System.Drawing.Size(1040, 720)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "On-prem Exchange only -- requires an already-open Exchange Management Shell session (or Exchange snap-in loaded). Environment/DC below is used only for AD-side lookups (disabled-account and AAD-sync checks), not for Exchange cmdlets."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(1000, 32)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$grpEnv = New-Object System.Windows.Forms.GroupBox
$grpEnv.Text = "Environment (AD lookups only)"
$grpEnv.Location = New-Object System.Drawing.Point(12, 46)
$grpEnv.Size = New-Object System.Drawing.Size(300, 50)
$form.Controls.Add($grpEnv)
$radFranchise = New-Object System.Windows.Forms.RadioButton
$radFranchise.Text = "Franchise"
$radFranchise.Location = New-Object System.Drawing.Point(10, 20)
$radFranchise.AutoSize = $true
$radFranchise.Checked = $true
$grpEnv.Controls.Add($radFranchise)
$radHQ = New-Object System.Windows.Forms.RadioButton
$radHQ.Text = "HQ"
$radHQ.Location = New-Object System.Drawing.Point(120, 20)
$radHQ.AutoSize = $true
$grpEnv.Controls.Add($radHQ)

$lblDC = New-Object System.Windows.Forms.Label
$lblDC.Text = "Domain Controller:"
$lblDC.Location = New-Object System.Drawing.Point(324, 60)
$lblDC.AutoSize = $true
$form.Controls.Add($lblDC)
$cboDC = New-Object System.Windows.Forms.ComboBox
$cboDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:FranchiseDCs | ForEach-Object { [void]$cboDC.Items.Add($_) }
$cboDC.SelectedIndex = 0
$cboDC.Location = New-Object System.Drawing.Point(440, 57)
$cboDC.Size = New-Object System.Drawing.Size(360, 20)
$cboDC.Font = New-Object System.Drawing.Font("Consolas", 8)
$form.Controls.Add($cboDC)

function Get-SelectedEnvLabel { if ($radHQ.Checked) { return 'HQ' } else { return 'Franchise' } }
function Update-EnvironmentUI { $cboDC.Enabled = -not $radHQ.Checked }
$radFranchise.Add_CheckedChanged({ Update-EnvironmentUI })
$radHQ.Add_CheckedChanged({ Update-EnvironmentUI })
if ($Environment -eq 'HQ') { $radHQ.Checked = $true }
elseif ($DomainController -and $script:FranchiseDCs -contains $DomainController) { $cboDC.SelectedItem = $DomainController }
Update-EnvironmentUI

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan All Distribution Groups"
$btnScan.Location = New-Object System.Drawing.Point(812, 46)
$btnScan.Size = New-Object System.Drawing.Size(200, 28)
$form.Controls.Add($btnScan)
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($btnScan, "Runs a live scan against Exchange and saves the result to a cache file ($CacheFilePath). Next time this tool opens, that cache loads automatically instead of re-scanning.")

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search (Name or Email, partial match):"
$lblSearch.Location = New-Object System.Drawing.Point(12, 106)
$lblSearch.AutoSize = $true
$form.Controls.Add($lblSearch)
$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(300, 103)
$txtSearch.Size = New-Object System.Drawing.Size(400, 20)
$txtSearch.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtSearch)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Not scanned yet. Click 'Scan All Distribution Groups' to begin."
$lblStatus.Location = New-Object System.Drawing.Point(12, 132)
$lblStatus.Size = New-Object System.Drawing.Size(900, 20)
$lblStatus.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblStatus)

$btnLoadCounts = New-Object System.Windows.Forms.Button
$btnLoadCounts.Text = "Load Member Counts"
$btnLoadCounts.Location = New-Object System.Drawing.Point(812, 128)
$btnLoadCounts.Size = New-Object System.Drawing.Size(200, 26)
$btnLoadCounts.Enabled = $false
$form.Controls.Add($btnLoadCounts)
$btnCancelCounts = New-Object System.Windows.Forms.Button
$btnCancelCounts.Text = "Cancel"
$btnCancelCounts.Location = New-Object System.Drawing.Point(812, 128)
$btnCancelCounts.Size = New-Object System.Drawing.Size(200, 26)
$btnCancelCounts.Visible = $false
$form.Controls.Add($btnCancelCounts)

$grdDL = New-Object System.Windows.Forms.DataGridView
$grdDL.Location = New-Object System.Drawing.Point(12, 160)
$grdDL.Size = New-Object System.Drawing.Size(1000, 480)
$grdDL.ReadOnly = $true
$grdDL.AllowUserToAddRows = $false
$grdDL.AllowUserToDeleteRows = $false
$grdDL.AllowUserToResizeRows = $false
$grdDL.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grdDL.MultiSelect = $false
$grdDL.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grdDL.Font = New-Object System.Drawing.Font("Consolas", 9)
[void]$grdDL.Columns.Add('Name', 'Name')
[void]$grdDL.Columns.Add('Email', 'Email')
[void]$grdDL.Columns.Add('Source', 'Source')
[void]$grdDL.Columns.Add('MemberCount', 'Members')
[void]$grdDL.Columns.Add('Moderation', 'Moderation')
[void]$grdDL.Columns.Add('HiddenFromGAL', 'Hidden From GAL')
[void]$grdDL.Columns.Add('Identity', 'Identity')
$grdDL.Columns['Name'].FillWeight          = 22
$grdDL.Columns['Email'].FillWeight         = 26
$grdDL.Columns['Source'].FillWeight        = 16
$grdDL.Columns['MemberCount'].FillWeight   = 10
$grdDL.Columns['Moderation'].FillWeight    = 12
$grdDL.Columns['HiddenFromGAL'].FillWeight = 14
$grdDL.Columns['Identity'].Visible         = $false
foreach ($col in $grdDL.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
$form.Controls.Add($grdDL)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(912, 650)
$btnClose.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnClose)

# Right-click context menu on the DL grid -- matches the double-click
# behavior below plus a couple of read-only conveniences.
$ctxDL = New-Object System.Windows.Forms.ContextMenuStrip
$miOpen = $ctxDL.Items.Add("Open DL Editor")
$miRefreshCount = $ctxDL.Items.Add("Refresh Member Count")
$miExport = $ctxDL.Items.Add("Export Member List...")
[void]$ctxDL.Items.Add("-")
$miCopyEmail = $ctxDL.Items.Add("Copy Email Address")
$miCopyName = $ctxDL.Items.Add("Copy Name")
$grdDL.ContextMenuStrip = $ctxDL

$script:AllDLs = @()

function Get-SelectedDL {
    if ($grdDL.SelectedRows.Count -eq 0) { return $null }
    $identity = $grdDL.SelectedRows[0].Cells['Identity'].Value
    return $script:AllDLs | Where-Object { $_.Identity -eq $identity } | Select-Object -First 1
}

function Update-DLGridRow {
    # Refreshes a single visible row's cells from $script:AllDLs -- used
    # after the member-count scan updates one DL at a time, and after a
    # source-label refresh, without a full grid rebuild.
    param($DL)
    foreach ($row in $grdDL.Rows) {
        if ($row.Cells['Identity'].Value -eq $DL.Identity) {
            $row.Cells['Name'].Value          = $DL.DisplayName
            $row.Cells['Email'].Value         = $DL.PrimarySmtpAddress
            $row.Cells['Source'].Value        = if ($DL.Source) { $DL.Source } else { '(not checked)' }
            $row.Cells['MemberCount'].Value   = if ($null -ne $DL.MemberCount) { "$($DL.MemberCount)" } else { '...' }
            $row.Cells['Moderation'].Value    = if ($DL.ModerationEnabled) { 'Enabled' } else { 'Disabled' }
            $row.Cells['HiddenFromGAL'].Value = "$($DL.HiddenFromAddressLists)"
            return
        }
    }
}

function Update-DLGridFromCurrent {
    $grdDL.Rows.Clear()
    foreach ($dl in $script:AllDLs) {
        [void]$grdDL.Rows.Add(
            $dl.DisplayName, $dl.PrimarySmtpAddress,
            $(if ($dl.Source) { $dl.Source } else { '(not checked)' }),
            $(if ($null -ne $dl.MemberCount) { "$($dl.MemberCount)" } else { '...' }),
            $(if ($dl.ModerationEnabled) { 'Enabled' } else { 'Disabled' }),
            "$($dl.HiddenFromAddressLists)",
            $dl.Identity
        )
    }
    Update-SearchFilter
}

# Real-time "contains" filter across Name and Email -- covers prefix/suffix
# matching as trivial cases of substring matching, so one filter function
# handles all three modes the original request called out separately.
function Update-SearchFilter {
    $needle = $txtSearch.Text.Trim()
    foreach ($row in $grdDL.Rows) {
        if ($needle -eq '') { $row.Visible = $true; continue }
        $name = "$($row.Cells['Name'].Value)"
        $email = "$($row.Cells['Email'].Value)"
        $row.Visible = ($name -like "*$needle*") -or ($email -like "*$needle*")
    }
}
$txtSearch.Add_TextChanged({ Update-SearchFilter })

$btnScan.Add_Click({
    $btnScan.Enabled = $false
    $btnLoadCounts.Enabled = $false
    $lblStatus.Text = "Scanning distribution groups..."
    $form.Refresh()
    try {
        $script:AllDLs = Get-DLInventory
        $adServer = Get-ADServerString -Environment (Get-SelectedEnvLabel) -DomainController $cboDC.SelectedItem
        foreach ($dl in $script:AllDLs) {
            $dl.Source = Get-DLSourceLabel -DistinguishedName $dl.DistinguishedName -ADServer $adServer
        }
        Update-DLGridFromCurrent
        $cacheNote = ''
        try {
            Save-DLInventoryCache -Path $CacheFilePath -DLs $script:AllDLs
            $cacheNote = " Saved to cache ($CacheFilePath)."
        } catch {
            $cacheNote = " (Could not write cache file: $($_.Exception.Message))"
        }
        $lblStatus.Text = "$($script:AllDLs.Count) distribution group(s) found.$cacheNote Member counts not loaded yet -- click 'Load Member Counts'."
        $btnLoadCounts.Enabled = ($script:AllDLs.Count -gt 0)
    } catch {
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroup (DL scan)'
    } finally {
        $btnScan.Enabled = $true
    }
})

$script:CountScanCancelRequested = $false
$btnLoadCounts.Add_Click({
    $btnLoadCounts.Visible = $false
    $btnCancelCounts.Visible = $true
    $script:CountScanCancelRequested = $false
    $i = 0
    $total = $script:AllDLs.Count
    foreach ($dl in $script:AllDLs) {
        $i++
        $lblStatus.Text = "Loading member count $i of $total`: $($dl.DisplayName)..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $dl.MemberCount = Get-DLMemberCount -Identity $dl.Identity
        } catch {
            $dl.MemberCount = -1
        }
        Update-DLGridRow -DL $dl
        if ($script:CountScanCancelRequested) { break }
    }
    $cacheNote = ''
    try {
        Save-DLInventoryCache -Path $CacheFilePath -DLs $script:AllDLs
        $cacheNote = ' Cache updated with member counts.'
    } catch {
        $cacheNote = " (Could not write cache file: $($_.Exception.Message))"
    }
    $lblStatus.Text = if ($script:CountScanCancelRequested) { "Member count scan cancelled after $i of $total.$cacheNote" } else { "Member counts loaded for all $total distribution group(s).$cacheNote" }
    $btnLoadCounts.Visible = $true
    $btnCancelCounts.Visible = $false
})
$btnCancelCounts.Add_Click({ $script:CountScanCancelRequested = $true })

$grdDL.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $dl = Get-SelectedDL
    if (-not $dl) { return }
    $adServer = Get-ADServerString -Environment (Get-SelectedEnvLabel) -DomainController $cboDC.SelectedItem
    Show-DLEditorForm -DL $dl -ADServer $adServer
    Update-DLGridRow -DL $dl
})
$miOpen.Add_Click({
    $dl = Get-SelectedDL
    if (-not $dl) { return }
    $adServer = Get-ADServerString -Environment (Get-SelectedEnvLabel) -DomainController $cboDC.SelectedItem
    Show-DLEditorForm -DL $dl -ADServer $adServer
    Update-DLGridRow -DL $dl
})
$miRefreshCount.Add_Click({
    $dl = Get-SelectedDL
    if (-not $dl) { return }
    try {
        $dl.MemberCount = Get-DLMemberCount -Identity $dl.Identity
        Update-DLGridRow -DL $dl
    } catch {
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Refresh member count'
    }
})
$miExport.Add_Click({
    $dl = Get-SelectedDL
    if (-not $dl) { return }
    try {
        $members = Get-DLMemberDetail -Identity $dl.Identity
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $sfd.FileName = "$($dl.Alias -replace '[^a-zA-Z0-9]','_')_members_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Export-DLMembersToCsv -Members $members -Path $sfd.FileName
        $lblStatus.Text = "Exported $($members.Count) member(s) of '$($dl.DisplayName)' to $($sfd.FileName)."
    } catch {
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Export member list'
    }
})
$miCopyEmail.Add_Click({
    $dl = Get-SelectedDL
    if ($dl) { [System.Windows.Forms.Clipboard]::SetText($dl.PrimarySmtpAddress) }
})
$miCopyName.Add_Click({
    $dl = Get-SelectedDL
    if ($dl) { [System.Windows.Forms.Clipboard]::SetText($dl.DisplayName) }
})

$btnClose.Add_Click({ $form.Close() })

# Auto-load from the inventory cache file if one exists and is readable,
# instead of opening to an empty grid every time. This runs once, at
# startup, after every control/handler above it is already wired -- 'Scan
# All Distribution Groups' is still the only thing that ever talks to
# Exchange; this block only ever reads the cache file. See header REVISED
# 2026-07-11 note above for the full reasoning.
$script:CacheData = Read-DLInventoryCache -Path $CacheFilePath
if ($script:CacheData -and $script:CacheData.DLs.Count -gt 0) {
    $script:AllDLs = $script:CacheData.DLs
    Update-DLGridFromCurrent
    $btnLoadCounts.Enabled = $true
    $lblStatus.Text = "Loaded $($script:AllDLs.Count) distribution group(s) from cache, scanned $($script:CacheData.ScanTimestamp). Click 'Scan All Distribution Groups' for a fresh live scan."
} elseif (Test-Path -LiteralPath $CacheFilePath) {
    $lblStatus.Text = "Cache file at '$CacheFilePath' exists but could not be read -- click 'Scan All Distribution Groups' to replace it with a fresh one."
}
# else: no cache file at all yet -- $lblStatus already shows the default
# "Not scanned yet" text set when the label was created above.

[void]$form.ShowDialog()
#endregion
