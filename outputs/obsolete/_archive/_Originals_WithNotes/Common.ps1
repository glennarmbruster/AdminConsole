﻿#Requires -Version 5.1
<#
.SYNOPSIS
    Shared helper library for the Admin Console launcher and every standalone
    tool script (MailboxHealth.ps1, MessageTrace.ps1, InboxRules.ps1, etc.).
.DESCRIPTION
    Dot-source this file from any tool script:  . "$PSScriptRoot\Common.ps1"

    Every function in here is MOVED, UNCHANGED, from the original monolithic
    Mailbox-Group-Inspector.ps1 -- Get-TargetSmtp, Format-ErrorDetail,
    Get-EXOCommand, Invoke-EXOCommand, Resolve-Target, Write-CompletedLog, and
    Get-InboxRulesForMailbox are copy-pasted verbatim from that script (their
    original line numbers are noted below for traceability). Nothing about
    their behavior was redesigned, simplified, or rewritten.

    Two things in this file ARE new, because the monolith never needed them
    (everything used to share one WinForms $form and its control objects):
      - Get-ADServerString: a parameter-based equivalent of the monolith's
        Get-SelectedADServer, which read $radHQ.Checked / $cboDC.SelectedItem
        directly off the shared form. Tool scripts now receive Environment /
        DomainController as plain strings (from CLI params or their own local
        prompt), so this takes those same two inputs as parameters instead of
        reading a control.
      - Show-TargetPromptDialog / Resolve-ToolTarget: the shared "manual
        entry" prompt every tool script falls back to when launched
        standalone (no CLI parameters), so each of the ~10 tool scripts
        doesn't reimplement its own copy of the input box / Environment
        radios / DC dropdown. This is orchestration/UI glue for the
        parameter-passing model described in the architecture doc, not any
        tool's business logic.

    REVISED 2026-07-23 (operator request -- auto-detect HQ vs Franchise here
    too, matching AdminConsole.ps1's same-day change): account names are
    unique across the HQ and Franchise forests, so there's no reason any tool
    script's manual-entry fallback should still make the operator pick an
    environment by hand.
      - New function Find-TargetEnvironment does the detection: Get-ADObject
        -Filter "mail -eq '<address>'" against HQ first ($script:HQDomain,
        forest-wide, smaller forest), then Franchise only if HQ comes back
        empty (using the new $script:FranchiseAutoDetectDC below). Whichever
        forest answers determines Environment + ADServer. Returns Environment
        = $null if neither forest has a match.
      - Show-TargetPromptDialog dropped its Environment radio group and DC
        dropdown entirely -- it now collects ONLY the mailbox/group address;
        Resolve-ToolTarget calls Find-TargetEnvironment once that address is
        known, instead of trusting whatever the operator picked on the
        now-removed radios.
      - Resolve-ToolTarget's existing fast path (all three of SmtpAddress /
        Environment / DomainController already valid -- the normal case when
        a tool is launched from AdminConsole.ps1, which already ran its own
        auto-detect) is UNCHANGED: it still trusts those values as-is rather
        than re-querying AD a second time for the same address. Auto-detect
        only runs when SmtpAddress is present but Environment/DomainController
        aren't (or a tool is launched fully standalone with nothing at all,
        in which case the trimmed-down Show-TargetPromptDialog collects just
        the address first).
      - $script:FranchiseAutoDetectDC is a NEW, deliberately separately-named
        constant -- NOT reusing the name $script:DefaultDC that AdminConsole.
        ps1 (and SendToDL.ps1/SendAsDL.ps1) each keep as their OWN local
        copy. This isn't the same situation as the 2026-07-21 incident where
        a shared $script:DefaultDC in Common.ps1 silently changed
        disable-intake.ps1's behavior: Find-TargetEnvironment is a brand-new
        function nothing called before this revision, so there's nothing for
        it to unexpectedly override -- every script in this codebase that
        dot-sources Common.ps1 gains a new, inert function/variable pair it
        simply doesn't call unless it already invokes Resolve-ToolTarget.
      - Scope: this only changes the manual-entry FALLBACK path shared by
        tool scripts that already call Resolve-ToolTarget (AllAttributes,
        CalendarPermissions, DLMemberships, GroupEditor, InboxRules,
        MailAttributes, MailboxAudit, MailboxHealth, MailboxPermissions,
        OOFEditor, SharedMailboxAudit -- 11 scripts, unchanged themselves,
        since Resolve-ToolTarget's own signature and contract didn't change).
        SendToDL.ps1, SendAsDL.ps1, DLManager.ps1, and UnifiedGroupMembership.
        ps1 each have their OWN separate, local Environment radio UI (they
        look up Distribution Lists by Name, not by mailbox/group SMTP
        address, so Find-TargetEnvironment's mail-attribute AD filter doesn't
        directly apply) and were NOT touched by this revision -- see the
        2026-07-23 conversation for the follow-up plan on those four.
        MessageTrace.ps1 no longer calls Resolve-ToolTarget at all (see its
        own header) and is unaffected either way. SRG/ANZ Common.ps1 copies
        (SRGCommon.ps1, ANZCommon.ps1) are entirely separate files and were
        not touched -- those suites have no HQ/Franchise concept at all.

    Why Get-InboxRulesForMailbox lives here instead of only in InboxRules.ps1:
    Get-MailboxHealthSnapshot (in MailboxHealth.ps1) calls it too ("Reuses
    Get-InboxRulesForMailbox ... rather than querying it twice" -- see that
    function's own comment). Since every tool script must be fully
    functional on its own, this one small, genuinely shared function is
    promoted to Common.ps1 rather than duplicated or left to create a
    file-to-file dependency between two otherwise-independent tool scripts.
.NOTES
    Author   : Glenn Armbruster
    Created  : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 as part
               of the launcher/tool-script architecture refactor -- see
               REFACTOR-DESIGN.md for the full migration plan).
    Revised  : 2026-07-23 (auto-detect HQ vs Franchise in
               Resolve-ToolTarget/Show-TargetPromptDialog -- see REVISED
               note above)
#>

#region Constants (moved from the WinForms UI region, lines ~3672-3682) ------
$script:MailDomain   = 'expresspros.com'
$script:FranchiseDCs = @(
    'fran-vm-dc01.franchises.expresspersonnel.com'
    'fran-vm-dc02.franchises.expresspersonnel.com'
    'fran-vm-dc03.franchises.expresspersonnel.com'
    'fran-vm-dc04.franchises.expresspersonnel.com'
    'fran-vm-dc05.franchises.expresspersonnel.com'
    'fran-vm-dc06.franchises.expresspersonnel.com'
    'fran-vm-dc07.franchises.expresspersonnel.com'
    'fran-vm-dc08.franchises.expresspersonnel.com'
)
$script:HQDomain = 'expresspersonnel.com'
# Franchise DC used ONLY by Find-TargetEnvironment's auto-detect check below.
# Deliberately a separately-named constant, NOT $script:DefaultDC (the name
# AdminConsole.ps1/SendToDL.ps1/SendAsDL.ps1 each keep as their OWN local
# copy) -- see the REVISED 2026-07-23 header note above for why this is safe
# to centralize here even though a shared $script:DefaultDC caused a real
# problem in disable-intake.ps1 back on 2026-07-21.
$script:FranchiseAutoDetectDC = 'fran-vm-dc01.franchises.expresspersonnel.com'
#endregion

#region Parsing (moved unchanged -- original lines 550-559) ------------------
function Get-TargetSmtp {
    param([string]$Text)
    if (-not $Text) { return $null }
    $t = $Text.Trim()
    if ($t -eq '') { return $null }
    if ($t -like '*@*') { return $t }
    return "$t@$($script:MailDomain)"
}
#endregion

#region Error formatting (moved unchanged -- original lines 561-587) --------
function Format-ErrorDetail {
    # Central formatter so every catch block reports the same level of
    # detail: cmdlet, script line, exception type, message, category, and
    # inner exception if present. Never just "$_.Exception.Message" alone.
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

#region EXO cmdlet resolution (moved unchanged -- original lines 588-625) ---
function Get-EXOCommand {
    # Resolves a cmdlet name to the Exchange Online-sourced command specifically.
    # Field-tested finding (2026-07-08): even with no explicit on-prem EMS session
    # connected, an on-prem Exchange module/snap-in can still be loaded in the host
    # (leftover profile/autoload) and win a plain-name lookup -- the classic symptom
    # is a "couldn't be found on '<DC>'" error from what should be a cloud-only
    # cmdlet, or a script-line number in the thousands (an auto-generated remoting
    # proxy, not this script). Every EXO-flavored call in ANY tool script MUST go
    # through Invoke-EXOCommand below instead of calling the cmdlet name directly.
    param([Parameter(Mandatory)][string]$Name)
    $all = @(Get-Command -Name $Name -All -ErrorAction SilentlyContinue)
    if ($all.Count -eq 0) {
        throw "No command named '$Name' exists in this session. Is Exchange Online connected (Connect-ExchangeOnline)?"
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
    # Splat-friendly wrapper: Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    $cmd = Get-EXOCommand -Name $Name
    & $cmd @Parameters
}
#endregion

#region Object-type resolution (moved unchanged -- original lines 628-668) --
function Resolve-Target {
    # Used by AdminConsole.ps1 to detect Mailbox / UnifiedGroup / OnpremGroup
    # so it knows which tool buttons to render. Individual tool scripts do
    # NOT need to call this themselves -- each one already assumes its own
    # target type (e.g. MailboxHealth.ps1 assumes it was handed a mailbox);
    # if handed the wrong type of object, the underlying Get-Mailbox /
    # Get-ADGroup call in that tool will simply fail with a normal
    # Format-ErrorDetail message, same as the monolith's behavior when
    # someone tried, say, "Inbox Rules" against a group.
    param(
        [string]$Smtp,
        [string]$ADServer
    )
    $result = [ordered]@{
        Type              = 'NotFound'
        Smtp              = $Smtp
        ADGroupDN         = $null
        ResolutionErrors  = @()
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
        $grp = @(Get-ADGroup -Filter "mail -eq '$Smtp'" -Server $ADServer `
                    -Properties mail, ManagedBy, GroupCategory, GroupScope, msExchHideFromAddressLists `
                    -ErrorAction Stop)
        if ($grp.Count -eq 1) {
            $result.Type      = 'OnpremGroup'
            $result.ADGroupDN = $grp[0].DistinguishedName
            return [PSCustomObject]$result
        } elseif ($grp.Count -gt 1) {
            $result.ResolutionErrors += "ERROR [Get-ADGroup (on-prem) resolution attempt] -- $($grp.Count) AD groups matched mail '$Smtp'. Expected 0 or 1. Resolve the duplicate mail attribute before using this tool."
        }
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADGroup (on-prem) resolution attempt')
    }
    return [PSCustomObject]$result
}
#endregion

#region Inbox rule read (moved unchanged -- original lines 798-806) ---------
# Shared by InboxRules.ps1 (its primary reason for existing) and
# MailboxHealth.ps1 (rule count / suspicious-rule scan) -- see header note.
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

#region Completed-tasks logging (moved unchanged -- original lines 3655-3669)
function Write-CompletedLog {
    param([string]$Task)
    $folder = "E:\Scripts\M365\Logs\Completed"
    $file   = Join-Path $folder "CompletedTasks.csv"
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

#region NEW -- parameter-based AD server resolution --------------------------
# Equivalent to the monolith's Get-SelectedADServer, but takes plain string
# parameters instead of reading $radHQ.Checked / $cboDC.SelectedItem off a
# shared form (there is no shared form anymore -- every tool script owns its
# own controls, or receives these values as CLI parameters with no controls
# at all). Same two-branch logic, nothing behavioral changed.
function Get-ADServerString {
    param(
        [Parameter(Mandatory)][ValidateSet('Franchise', 'HQ')][string]$Environment,
        [string]$DomainController
    )
    if ($Environment -eq 'HQ') { return $script:HQDomain }
    return $DomainController
}
#endregion

#region NEW -- auto-detect HQ vs Franchise -----------------------------------
# Added 2026-07-23 -- see the header REVISED note. Auto-detects which AD
# forest a mailbox/group's on-prem object lives in: account names are unique
# across HQ and Franchise (operator-confirmed), so "an AD object with this
# mail address exists here" reliably identifies environment -- no separate
# identity mapping needed. HQ is checked first (forest-wide $script:HQDomain,
# smaller forest, no fixed-DC risk); Franchise is only checked if HQ comes
# back empty, using $script:FranchiseAutoDetectDC above. Returns Environment
# = $null if neither forest has a match -- callers treat that as "not found
# in HQ or Franchise", not a silent fallback to one side. Identical logic to
# AdminConsole.ps1's own copy of this function (kept separate there rather
# than refactored to call this one, to avoid touching an already-tested file
# for a pure deduplication with no behavior change).
function Find-TargetEnvironment {
    param([string]$Smtp)
    $errors = @()
    $hqServer = Get-ADServerString -Environment 'HQ' -DomainController $null
    try {
        $hqHit = @(Get-ADObject -Filter "mail -eq '$Smtp'" -Server $hqServer -ErrorAction Stop)
        if ($hqHit.Count -gt 0) {
            return [PSCustomObject]@{ Environment = 'HQ'; ADServer = $hqServer; Errors = $errors }
        }
    } catch {
        $errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'HQ AD lookup (environment auto-detect)')
    }
    $franServer = Get-ADServerString -Environment 'Franchise' -DomainController $script:FranchiseAutoDetectDC
    try {
        $franHit = @(Get-ADObject -Filter "mail -eq '$Smtp'" -Server $franServer -ErrorAction Stop)
        if ($franHit.Count -gt 0) {
            return [PSCustomObject]@{ Environment = 'Franchise'; ADServer = $franServer; Errors = $errors }
        }
    } catch {
        $errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Franchise AD lookup (environment auto-detect)')
    }
    return [PSCustomObject]@{ Environment = $null; ADServer = $null; Errors = $errors }
}
#endregion

#region NEW -- auto-detect HQ vs Franchise BY sAMAccountName ------------------
# ADDED 2026-07-24 for newuser.ps1 (Enable Remote Mailbox).
#
# WHY THIS EXISTS RATHER THAN REUSING Find-TargetEnvironment ABOVE -- this is
# the important bit, do not "simplify" by merging them:
#
#   Find-TargetEnvironment locates an object by its MAIL ATTRIBUTE
#   (Get-ADObject -Filter "mail -eq '<smtp>'"). That is correct for every
#   other tool in this suite, because those tools act on mailboxes and groups
#   that are ALREADY mail-enabled.
#
#   The Enable-RemoteMailbox workflow is the exact opposite case: the accounts
#   it targets are NOT mail-enabled yet -- that is the entire reason the tool
#   is being run. Their mail attribute is empty, so a mail-based lookup finds
#   nothing and would report "not found in HQ or Franchise" for every single
#   valid new account.
#
# So this one resolves by sAMAccountName, which exists from the moment the
# account is created. Same HQ-first-then-Franchise order and same
# "Environment = $null means genuinely not found" contract as its sibling.
function Find-UserEnvironmentBySam {
    <#
    .SYNOPSIS
        Finds which forest an account lives in, by sAMAccountName.
    .OUTPUTS
        PSCustomObject:
          Environment [string] 'HQ' | 'Franchise' | $null when not found
          ADServer    [string] the server that answered ($null when not found)
          DCParam     [string] '' for HQ (forest-wide), the DC FQDN for
                               Franchise -- ready to pass to
                               Set-ADServerSettings / -DomainController
          User        [object] the matched AD user, for display
          Errors      [string[]]
    #>
    param([Parameter(Mandatory)][string]$SamAccountName)

    $errors = @()
    $sam = $SamAccountName.Trim()

    # HQ first: smaller forest, and forest-wide so no DC is pinned.
    $hqServer = Get-ADServerString -Environment 'HQ' -DomainController $null
    try {
        $hqHit = Get-ADUser -Identity $sam -Server $hqServer -ErrorAction Stop
        if ($hqHit) {
            return [PSCustomObject]@{
                Environment = 'HQ'; ADServer = $hqServer; DCParam = ''
                User = $hqHit; Errors = $errors
            }
        }
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        # Genuinely absent from HQ -- expected, fall through to Franchise
        # rather than recording it as an error.
    } catch {
        $errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'HQ AD lookup (sAMAccountName auto-detect)')
    }

    $franServer = Get-ADServerString -Environment 'Franchise' -DomainController $script:FranchiseAutoDetectDC
    try {
        $franHit = Get-ADUser -Identity $sam -Server $franServer -ErrorAction Stop
        if ($franHit) {
            return [PSCustomObject]@{
                Environment = 'Franchise'; ADServer = $franServer
                DCParam = $script:FranchiseAutoDetectDC
                User = $franHit; Errors = $errors
            }
        }
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        # Absent from Franchise too -- the caller reports "does not exist yet".
    } catch {
        $errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Franchise AD lookup (sAMAccountName auto-detect)')
    }

    return [PSCustomObject]@{
        Environment = $null; ADServer = $null; DCParam = $null
        User = $null; Errors = $errors
    }
}
#endregion

#region NEW -- consistent window placement -----------------------------------
# ADDED 2026-07-24 (operator request): the console should always open at the
# top-left of the screen, and every tool it launches at the bottom-left, so
# the two never cover each other and both land somewhere predictable.
#
# Uses WorkingArea rather than Bounds so the taskbar is respected -- a form
# placed by Bounds can end up underneath it.
#
# Clamped to the working area: if a form is taller than the screen, the
# bottom-left calculation would otherwise push its title bar off the top,
# leaving a window that can't be dragged. Better to sit flush at the top than
# to be unreachable.
function Set-FormScreenCorner {
    <#
    .SYNOPSIS
        Places a WinForms form in a screen corner. Call BEFORE ShowDialog().
    .PARAMETER Corner
        TopLeft | BottomLeft
    #>
    param(
        [Parameter(Mandatory)]$Form,
        [ValidateSet('TopLeft', 'BottomLeft')][string]$Corner = 'TopLeft'
    )
    try {
        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual

        $x = $wa.Left
        if ($Corner -eq 'TopLeft') {
            $y = $wa.Top
        } else {
            $y = $wa.Bottom - $Form.Height
            if ($y -lt $wa.Top) { $y = $wa.Top }   # taller than the screen
        }
        $Form.Location = New-Object System.Drawing.Point($x, $y)
    } catch {
        # Placement is cosmetic -- never let it stop a tool from opening.
        $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    }
}
#endregion

#region NEW -- connection state probes (for AdminConsole's Connect buttons) --
# ADDED 2026-07-24 -- so AdminConsole.ps1 can colour its Connect Scripts
# buttons green when a service is connected and red when it isn't.
#
# SAFE ADDITION, same reasoning as Find-TargetEnvironment above: these are
# brand-new FUNCTIONS that nothing previously called, so no existing script's
# behaviour can change because of them. This is deliberately NOT a new shared
# CONSTANT -- adding a shared $script:DefaultDC here on 2026-07-21 is what
# broke disable-intake.ps1's own DC handling, and that lesson stands.
#
# DELIBERATE DUPLICATION with M365Logic.psm1's Test-M365ExoConnection /
# Test-M365SpoConnection: the four suites in this codebase (this one,
# SRGCommon, ANZCommon, M365Logic) are intentionally independent, sharing no
# runtime state, so a change in one can never surprise another. Copying ~20
# lines is the price of that isolation and is the established pattern here --
# do NOT "fix" this by having Common.ps1 import M365Logic.
function Test-ExoConnectionState {
    <#
    .SYNOPSIS
        Is Exchange Online usable in THIS process? Returns an object, never
        throws.
    .DESCRIPTION
        Cheap: reads in-process state only, no service call, so it's safe to
        call on a form's startup path.

        Two stages. First, are EXO cmdlets even present (the same Get-Command
        probe Get-EXOCommand uses)? If not, nothing is connected. Second, if
        the module offers Get-ConnectionInformation (EXO v3+), read live
        session state -- that's what distinguishes "module loaded but the
        session expired" from "genuinely connected", which matters for a
        console left open all day.

        Where state can't be determined it reports CONNECTED rather than
        disconnected, on purpose: a false green is corrected the moment a
        cmdlet actually fails with its own clear error, whereas a false red
        would send the operator chasing a connection problem that isn't there.
    #>
    param()
    $r = [ordered]@{ Connected = $false; Detail = ''; Organization = $null }

    $probe = @(Get-Command -Name 'Get-Mailbox' -All -ErrorAction SilentlyContinue |
        Where-Object { $_.ModuleName -like 'tmpEXO_*' -or $_.Source -like 'tmpEXO_*' -or
                       $_.ModuleName -eq 'ExchangeOnlineManagement' -or $_.Source -eq 'ExchangeOnlineManagement' })
    if ($probe.Count -eq 0) {
        $r.Detail = 'Not connected -- no Exchange Online cmdlets in this PowerShell process. Connect-ExchangeOnline is per-process, so connecting in another window does not count.'
        return [PSCustomObject]$r
    }

    $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
    if (-not $gci) {
        $r.Connected = $true
        $r.Detail = 'Exchange Online cmdlets are loaded. This module version has no Get-ConnectionInformation, so live session state could not be verified.'
        return [PSCustomObject]$r
    }
    try {
        $conns = @(& $gci -ErrorAction Stop)
        if ($conns.Count -eq 0) {
            $r.Detail = 'Not connected -- cmdlets are loaded but there is no active session (it may have expired). Run Connect-ExchangeOnline again.'
            return [PSCustomObject]$r
        }
        $healthy = $conns | Where-Object {
            (-not $_.PSObject.Properties['State']       -or "$($_.State)"       -eq 'Connected') -and
            (-not $_.PSObject.Properties['TokenStatus'] -or "$($_.TokenStatus)" -ne 'Expired')
        }
        $chosen = $conns | Select-Object -First 1
        if ($chosen.PSObject.Properties['Organization']) { $r.Organization = "$($chosen.Organization)" }
        if (@($healthy).Count -gt 0) {
            $r.Connected = $true
            $r.Detail = "Connected to Exchange Online$(if ($r.Organization) { " ($($r.Organization))" })."
        } else {
            $r.Detail = 'Not connected -- the existing session reports itself expired or disconnected. Run Connect-ExchangeOnline again.'
        }
    } catch {
        $r.Connected = $true
        $r.Detail = "Exchange Online cmdlets are loaded; Get-ConnectionInformation could not be read ($($_.Exception.Message))."
    }
    return [PSCustomObject]$r
}

function Get-ConnectedTenantDomain {
    <#
    .SYNOPSIS
        The CONNECTED tenant's default accepted domain, or $null. Never throws.
    .DESCRIPTION
        ADDED 2026-07-24 -- so the console can show WHICH tenant it's pointed
        at, the way M365_Host_Web.ps1's top bar does. "Connected" on its own
        doesn't tell you connected to what, which matters when several tenants
        get worked in a day.

        NOTE this is the LIVE tenant, deliberately distinct from
        $script:MailDomain above. That constant is what this suite APPENDS to a
        bare name typed in the box ('expresspros.com'); this function reports
        what Exchange Online is actually connected to. They should normally
        agree -- and when they don't, that's exactly the mistake worth seeing,
        because the box would be resolving names into a tenant you aren't
        connected to.

        Cached after the first success, so repeated calls (startup, Refresh,
        each Look Up) cost nothing. Returns $null rather than throwing when
        nothing is connected, since callers use it for a label.
    #>
    param()
    if ($script:ConnectedTenantDomainCache) { return $script:ConnectedTenantDomainCache }
    try {
        $accepted = @(Invoke-EXOCommand -Name 'Get-AcceptedDomain' -Parameters @{ ErrorAction = 'Stop' })
        if ($accepted.Count -eq 0) { return $null }

        # REVISED 2026-07-24 -- prefer the VANITY domain, not the default one.
        #
        # The first version returned whichever domain was flagged Default, and
        # in this tenant that is expresspros.onmicrosoft.com. That is completely
        # normal (the onmicrosoft domain is very often the default even when
        # every user's primary SMTP is the vanity domain) -- but it is not what
        # an admin means by "which tenant am I in", and it made the label read
        # as a mismatch against $script:MailDomain when nothing was wrong.
        #
        # Preference order:
        #   1. an accepted domain equal to $script:MailDomain -- the exact
        #      domain this console appends to bare names, so the most
        #      meaningful thing to show
        #   2. any non-*.onmicrosoft.com domain -- the vanity domain
        #   3. the Default domain
        #   4. whatever came first
        $names = @($accepted | ForEach-Object { "$($_.DomainName)" })

        $pick = $names | Where-Object { $_ -ieq $script:MailDomain } | Select-Object -First 1
        if (-not $pick) {
            $pick = $names | Where-Object { $_ -notlike '*.onmicrosoft.com' } | Select-Object -First 1
        }
        if (-not $pick) {
            $d = $accepted | Where-Object { $_.Default -eq $true } | Select-Object -First 1
            if ($d) { $pick = "$($d.DomainName)" }
        }
        if (-not $pick) { $pick = $names[0] }

        if ($pick) {
            $script:ConnectedTenantDomainCache = $pick
            return $pick
        }
    } catch {
        # Not connected, or the lookup failed -- not fatal, the caller shows
        # "unknown". Deliberately swallowed.
    }
    return $null
}

function Clear-ConnectedTenantDomainCache {
    # Needed if the operator disconnects and reconnects to a different tenant
    # inside one console session.
    $script:ConnectedTenantDomainCache = $null
}

function Test-SpoConnectionState {
    <#
    .SYNOPSIS
        Is SharePoint Online connected? Returns an object, never throws.
    .DESCRIPTION
        COSTS A REAL SERVICE CALL, and callers on a WinForms thread should know
        that: the SPO module exposes no local connection-state cmdlet (there is
        no SPO equivalent of Get-ConnectionInformation), so the only reliable
        probe is a cheap real call. Get-SPOTenant is used because it is
        tenant-scoped, needs no arguments, and returns quickly.

        Because it blocks, AdminConsole.ps1 calls this on startup and after a
        Connect click -- NOT on a timer, which would freeze the form at random
        moments during real work.
    #>
    param()
    $r = [ordered]@{ Connected = $false; Detail = '' }
    if (-not (Get-Command -Name 'Get-SPOTenant' -ErrorAction SilentlyContinue)) {
        $r.Detail = 'Not connected -- SharePoint Online cmdlets are not loaded in this process. Import Microsoft.Online.SharePoint.PowerShell and run Connect-SPOService.'
        return [PSCustomObject]$r
    }
    try {
        $null = Get-SPOTenant -ErrorAction Stop
        $r.Connected = $true
        $r.Detail = 'Connected to SharePoint Online.'
    } catch {
        $r.Detail = "Not connected -- SPO cmdlets are loaded but the session is missing or expired: $($_.Exception.Message)"
    }
    return [PSCustomObject]$r
}
#endregion

#region NEW -- shared "enter a target" prompt for standalone tool launches --
# Mirrors the monolith's own top-of-form controls (original lines ~3688-3726)
# as a small modal dialog, so every tool script gets the exact same "manual
# entry" experience without each of the ~14 tool scripts re-implementing its
# own copy of that UI. This function contains NO business logic -- it only
# collects an address and returns it.
#
# REVISED 2026-07-23 -- the Environment radios and DC dropdown this dialog
# used to show are gone; Resolve-ToolTarget below now auto-detects that via
# Find-TargetEnvironment once an address is known, so there's nothing left
# for the operator to pick here except the address itself. See the header
# REVISED 2026-07-23 note for the full rationale.
function Show-TargetPromptDialog {
    param(
        [string]$InitialSmtp,
        [string]$Title = 'Enter Target'
    )
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(480, 168)
    $dlg.StartPosition = 'CenterScreen'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = 'Mailbox or Group (email address):'
    $lblUser.Location = New-Object System.Drawing.Point(12, 12)
    $lblUser.AutoSize = $true
    $dlg.Controls.Add($lblUser)
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(12, 30)
    $txtUser.Size = New-Object System.Drawing.Size(440, 22)
    $txtUser.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtUser.Text = $InitialSmtp
    $dlg.Controls.Add($txtUser)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = 'HQ and Franchise are checked automatically -- no environment selection needed.'
    $lblHint.Location = New-Object System.Drawing.Point(12, 58)
    $lblHint.Size = New-Object System.Drawing.Size(440, 18)
    $lblHint.ForeColor = [System.Drawing.Color]::DimGray
    $lblHint.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Italic)
    $dlg.Controls.Add($lblHint)

    $lblErr = New-Object System.Windows.Forms.Label
    $lblErr.Location = New-Object System.Drawing.Point(12, 78)
    $lblErr.Size = New-Object System.Drawing.Size(440, 20)
    $lblErr.ForeColor = [System.Drawing.Color]::Red
    $dlg.Controls.Add($lblErr)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(280, 100)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Location = New-Object System.Drawing.Point(368, 100)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk

    $btnOk.Add_Click({
        $smtp = Get-TargetSmtp -Text $txtUser.Text
        if (-not $smtp) { $lblErr.Text = 'Enter a mailbox or group email address.'; return }
        $dlg.Tag = $smtp
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    })

    $result = $dlg.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $dlg.Tag
}

# Every tool script's param block gives it -SmtpAddress / -Environment /
# -DomainController. When AdminConsole launches the tool, all three arrive
# pre-populated (AdminConsole already ran its own Find-TargetEnvironment) and
# this returns immediately with no UI/no re-query, same fast path as before.
#
# REVISED 2026-07-23 -- when that fast path doesn't apply (a tool launched
# standalone, or with only a partial/invalid set of the three), this no
# longer pops the old Environment-radio/DC-dropdown prompt. Instead: if
# there's no usable address yet, Show-TargetPromptDialog collects ONLY the
# address (see its own REVISED note); either way, once an address is known,
# Find-TargetEnvironment auto-detects HQ vs Franchise the same way
# AdminConsole.ps1 does. If neither forest has a matching object, the
# operator sees "not found in HQ or Franchise" and this returns $null, same
# as a cancelled prompt.
function Resolve-ToolTarget {
    param(
        [string]$SmtpAddress,
        [string]$Environment,
        [string]$DomainController,
        [string]$ToolTitle = 'Enter Target'
    )
    $smtp = Get-TargetSmtp -Text $SmtpAddress
    $envOk = $Environment -in @('Franchise', 'HQ')
    if ($smtp -and $envOk -and ($Environment -ne 'Franchise' -or $DomainController)) {
        return [PSCustomObject]@{
            Smtp        = $smtp
            Environment = $Environment
            ADServer    = Get-ADServerString -Environment $Environment -DomainController $DomainController
        }
    }
    if (-not $smtp) {
        $smtp = Show-TargetPromptDialog -InitialSmtp $SmtpAddress -Title $ToolTitle
        if (-not $smtp) { return $null }   # operator cancelled
    }
    $envResult = Find-TargetEnvironment -Smtp $smtp
    if (-not $envResult.Environment) {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [void][System.Windows.Forms.MessageBox]::Show(
            "'$smtp' was not found in HQ or Franchise AD.`n`nDouble-check the address and try again.",
            $ToolTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $null
    }
    return [PSCustomObject]@{
        Smtp        = $smtp
        Environment = $envResult.Environment
        ADServer    = $envResult.ADServer
    }
}
#endregion
