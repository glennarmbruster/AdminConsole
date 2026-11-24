#Requires -Version 5.1
<#
.SYNOPSIS
    Bulk add/remove users to/from Microsoft 365 (Unified) Group membership --
    many-to-many, no target lookup.
.DESCRIPTION
    Two independent paste lists -- Users and Unified Groups, one entry per
    line (commas/semicolons also tolerated as separators). Every user listed
    is added to (or removed from) every group listed: a full Users x Groups
    cross-product, per the standards' batch/cross-product pattern (Section
    21). Confirmed with the operator before building that this tool would
    NOT use the single-object lookup flow every other tool in this suite
    uses (Section 1 triage) -- there is no single "target" here, there are
    two independent lists, each de-duplicated and normalized on its own.

    The two lists use DIFFERENT lookup keys, deliberately (Section 5 is
    written around a single-object tool; this is a two-list batch tool, so
    each list gets the key that matches what the operator actually has in
    hand for that list):
      - Users: PrimarySmtpAddress (Section 5 exception -- users are
        mail-enabled objects). Bare local-parts (no '@') get
        '@expresspros.com' silently appended via Common.ps1's existing
        Get-TargetSmtp; a full address is passed through unchanged.
      - Unified Groups: Name (the group's display name), passed straight
        through to -Identity with no transformation and no domain
        appended. Confirmed against the Microsoft Learn reference (Section
        2) that Add-/Remove-/Get-UnifiedGroupLinks' -Identity parameter
        resolves on Name, Alias, DN, Canonical DN, email address, or GUID --
        Name is a fully supported, documented value, not a guess. An email
        address still works too if that's easier to paste, since it's one
        of the same accepted values -- it just isn't required.
    Nothing in either list is required to already exist in the operator's
    EXO session before it's typed; unresolvable entries surface as errors
    at Add/Remove time, not as a validation step beforehand.

    Only the Members link type is touched. Owners and Subscribers are never
    read or modified by this tool -- use Group Editor (GroupEditor.ps1) for
    single-group owner delegation.

    Deliberate deviation from the standard Parse/Run/Close button triad
    (Section 13): the operator explicitly asked for exactly three buttons --
    Add, Remove, Cancel -- so there is no separate Parse button. In its
    place, the parsed-preview box (the Section 21 "show the scale before
    Run" requirement) refreshes live on every keystroke in either list box,
    so the true Users x Groups operation count is always visible before Add
    or Remove is ever clicked. Because there's no separate Parse step, there
    is also no persisted $script:Parsed state to go stale (Section 7) --
    Add/Remove each re-read and re-normalize both textboxes directly at
    click time.

    Verify mode (Section 19) is deliberately omitted, not an oversight:
    unlike on-prem AD writes elsewhere in this suite, a Get-UnifiedGroupLinks
    read run immediately after an Add-/Remove-UnifiedGroupLinks write against
    Exchange Online is not subject to the DC-replication-lag concern Verify
    mode exists to work around. The inline AFTER snapshot captured per group
    immediately after its writes (Section 17) is already authoritative.

    Already-a-member (on Add) / not-a-member (on Remove) operations are
    reported as SKIP, not errors (Section 21) -- determined from each
    group's own live BEFORE membership snapshot, never inferred from a
    cmdlet failure. BEFORE/AFTER is captured once per target GROUP, not once
    per user (Section 17/21); one CompletedTasks.csv row is written per
    group that had at least one real (non-skipped, non-failed) change,
    joining the successful members with '+' (Section 21 batch log grammar),
    never one row per user.

    Environment required: Exchange Online only. No on-prem AD/Exchange calls
    are made, so there is no DC selector and no Franchise/HQ split --
    Microsoft 365 Groups have no on-prem equivalent (Section 22 does not
    apply to this tool).

    -SmtpAddress/-Environment/-DomainController are declared, exactly as in
    PowerShellConsole.ps1, only so a future Start-ToolScript launch would
    not throw a parameter-binding error -- this tool ignores all three. It
    is not currently wired into Tools.json: ToolManager.ps1 validates every
    entry's ObjectTypes against a single resolved object (Mailbox /
    UnifiedGroup / OnpremGroup), which this many-to-many, lookup-free tool
    does not fit, the same reasoning DLManager.ps1's header gives for why it
    is launched standalone instead of from AdminConsole's per-object button
    row.

    Cmdlets used -- Add-UnifiedGroupLinks, Remove-UnifiedGroupLinks,
    Get-UnifiedGroupLinks -- are EXO-only and were verified against
    Microsoft Learn before delivery (Section 2):
      https://learn.microsoft.com/powershell/module/exchangepowershell/add-unifiedgrouplinks
      https://learn.microsoft.com/powershell/module/exchangepowershell/remove-unifiedgrouplinks
    Add-/Remove-UnifiedGroupLinks are also already in confirmed working use
    elsewhere in this suite, with the identical Identity/LinkType='Members'/
    Links/ErrorAction='Stop' shape, via GroupEditor.ps1's
    Add-UnifiedMember/Remove-UnifiedMember.

    KNOWN LIMITATION -- no PowerShell host was available to run the Section
    15 AST parse check before delivery. The script has been reviewed by eye
    for balanced braces/parens/quotes and structural consistency with the
    other tool scripts in this suite, but that is not a substitute for
    Parser::ParseFile -- run the Section 15 snippet against this file before
    first real use and treat any reported errors as blocking.
.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-07-21
    Revised         : 2026-07-21 1357
    Run from        : standard PowerShell with Exchange Online connected
                       (Connect-ExchangeOnline)
    Required modules: ExchangeOnlineManagement
#>
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Parsing --------------------------------------------------------------
# Normalizes one paste box into a de-duplicated array. De-dup is
# case-insensitive but preserves first-seen casing and order.
#   -AsEmail (Users box): tolerates newline, comma, and semicolon as
#     separators (Section 21) since pasted address lists often arrive
#     comma-delimited rather than one-per-line; runs every entry through
#     Common.ps1's Get-TargetSmtp so a bare local-part gets
#     '@expresspros.com' appended and a full address passes through as-is.
#   Groups box (no -AsEmail): entries are Unified Group DISPLAY NAMES, not
#     addresses, passed straight through with no domain-append and no other
#     transformation. Only newline and semicolon are treated as separators
#     here -- NOT comma -- because a group display name can legitimately
#     contain a comma (e.g. "IT, Facilities, and Ops"); splitting on comma
#     would silently corrupt a name like that into two bogus entries.
function ConvertTo-NormalizedList {
    param([string]$Text, [switch]$AsEmail)
    if (-not $Text) { return @() }
    $splitPattern = if ($AsEmail) { '[\r\n,;]+' } else { '[\r\n;]+' }
    $raw = @($Text -split $splitPattern | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    $items = if ($AsEmail) { @($raw | ForEach-Object { Get-TargetSmtp -Text $_ } | Where-Object { $_ }) } else { $raw }
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $result = @()
    foreach ($s in $items) {
        if ($seen.Add($s.ToLowerInvariant())) { $result += $s }
    }
    return $result
}
#endregion

#region Actions -- EXO, resolved via Invoke-EXOCommand (see Common.ps1) ------
function Get-UnifiedGroupMemberSmtps {
    param([string]$GroupIdentity)
    $links = @(Invoke-EXOCommand -Name 'Get-UnifiedGroupLinks' -Parameters @{
        Identity = $GroupIdentity; LinkType = 'Members'; ErrorAction = 'Stop'
    })
    return @($links | ForEach-Object { "$($_.PrimarySmtpAddress)" } | Where-Object { $_ -ne '' })
}

function Format-UnifiedGroupMembersState {
    param([string]$Heading, [string]$GroupIdentity, [string[]]$Members)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("  $Heading -- $GroupIdentity ($($Members.Count) member(s)):")
    if ($Members.Count -eq 0) {
        [void]$sb.AppendLine("    (none)")
    } else {
        $Members | Sort-Object | ForEach-Object { [void]$sb.AppendLine("    $_") }
    }
    return $sb.ToString()
}

function Add-UnifiedGroupMemberLink {
    param([string]$GroupIdentity, [string]$MemberSmtp)
    Invoke-EXOCommand -Name 'Add-UnifiedGroupLinks' -Parameters @{
        Identity = $GroupIdentity; LinkType = 'Members'; Links = $MemberSmtp; ErrorAction = 'Stop'
    } | Out-Null
}

function Remove-UnifiedGroupMemberLink {
    param([string]$GroupIdentity, [string]$MemberSmtp)
    Invoke-EXOCommand -Name 'Remove-UnifiedGroupLinks' -Parameters @{
        Identity = $GroupIdentity; LinkType = 'Members'; Links = $MemberSmtp; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}

# Runs the full Users x Groups cross-product for one direction (Add or
# Remove). One BEFORE/AFTER snapshot per target GROUP (Section 17/21), not
# per operation -- each user add/remove within that group is logged as a
# single line between the two snapshots. Already-satisfied operations
# (member already present on Add; member already absent on Remove) are SKIP,
# not errors, determined from the live BEFORE snapshot -- never from cmdlet
# failure (Section 21). A group that fails to resolve at all (bad address,
# not actually a Unified Group, etc.) is reported and skipped in its
# entirety without aborting the rest of the batch -- same "one bad row
# doesn't hide the other nine" convention used throughout this suite (e.g.
# DLManager.ps1's Add-DLMemberBulk). One CompletedTasks.csv row is written
# per group that had at least one real change (Section 21 batch grammar),
# never one row per user, and never for a group where every operation was a
# SKIP or a failure.
function Invoke-UnifiedGroupBatch {
    param(
        [string[]]$Users,
        [string[]]$Groups,
        [ValidateSet('Add', 'Remove')][string]$Verb
    )
    $out = New-Object System.Text.StringBuilder
    $totalDone = 0; $totalSkipped = 0; $totalFailed = 0
    $verbPast = if ($Verb -eq 'Add') { 'Added' } else { 'Removed' }

    foreach ($groupId in $Groups) {
        [void]$out.AppendLine("=== $groupId ===")
        try {
            $before = Get-UnifiedGroupMemberSmtps -GroupIdentity $groupId
        } catch {
            [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Get-UnifiedGroupLinks ($groupId) -- skipping this group entirely"))
            [void]$out.AppendLine("")
            $totalFailed += $Users.Count
            continue
        }
        [void]$out.Append((Format-UnifiedGroupMembersState -Heading "BEFORE" -GroupIdentity $groupId -Members $before))

        $beforeSet = New-Object System.Collections.Generic.HashSet[string]
        foreach ($m in $before) { [void]$beforeSet.Add($m.ToLowerInvariant()) }

        $successMembers = @()
        foreach ($userSmtp in $Users) {
            $isMember = $beforeSet.Contains($userSmtp.ToLowerInvariant())
            if ($Verb -eq 'Add') {
                if ($isMember) {
                    [void]$out.AppendLine("  SKIP (already a member): $userSmtp")
                    $totalSkipped++
                    continue
                }
                try {
                    Add-UnifiedGroupMemberLink -GroupIdentity $groupId -MemberSmtp $userSmtp
                    [void]$out.AppendLine("  ADDED: $userSmtp")
                    $successMembers += $userSmtp
                    $totalDone++
                } catch {
                    [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Add-UnifiedGroupLinks ($userSmtp -> $groupId)"))
                    $totalFailed++
                }
            } else {
                if (-not $isMember) {
                    [void]$out.AppendLine("  SKIP (not a member): $userSmtp")
                    $totalSkipped++
                    continue
                }
                try {
                    Remove-UnifiedGroupMemberLink -GroupIdentity $groupId -MemberSmtp $userSmtp
                    [void]$out.AppendLine("  REMOVED: $userSmtp")
                    $successMembers += $userSmtp
                    $totalDone++
                } catch {
                    [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Remove-UnifiedGroupLinks ($userSmtp -> $groupId)"))
                    $totalFailed++
                }
            }
        }

        try {
            $after = Get-UnifiedGroupMemberSmtps -GroupIdentity $groupId
            [void]$out.Append((Format-UnifiedGroupMembersState -Heading "AFTER" -GroupIdentity $groupId -Members $after))
        } catch {
            [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Get-UnifiedGroupLinks ($groupId) -- AFTER snapshot"))
        }
        [void]$out.AppendLine("")

        if ($successMembers.Count -gt 0) {
            Write-CompletedLog -Task "$verbPast Members $groupId (EXO; members: $($successMembers -join '+'))"
        }
    }

    $totalOps = $Users.Count * $Groups.Count
    [void]$out.AppendLine("--- Summary: $verbPast $totalDone, Skipped $totalSkipped, Failed $totalFailed (of $totalOps operation(s) across $($Groups.Count) group(s)) ---")
    return $out.ToString()
}
#endregion

#region WinForms UI -----------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Unified Group Membership -- Bulk Add / Remove (Exchange Online)"
$form.Size = New-Object System.Drawing.Size(900, 760)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Make sure you are connected to Exchange Online before running. This tool is EXO-only -- no on-prem AD/Exchange calls are made."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(860, 34)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$grpUsers = New-Object System.Windows.Forms.GroupBox
$grpUsers.Text = "Users (one per line -- email or logon; @expresspros.com optional)"
$grpUsers.Location = New-Object System.Drawing.Point(12, 50)
$grpUsers.Size = New-Object System.Drawing.Size(430, 230)
$form.Controls.Add($grpUsers)
$txtUsers = New-Object System.Windows.Forms.TextBox
$txtUsers.Multiline = $true
$txtUsers.ScrollBars = "Vertical"
$txtUsers.Location = New-Object System.Drawing.Point(10, 22)
$txtUsers.Size = New-Object System.Drawing.Size(410, 196)
$txtUsers.Font = New-Object System.Drawing.Font("Consolas", 9)
$grpUsers.Controls.Add($txtUsers)

$grpGroups = New-Object System.Windows.Forms.GroupBox
$grpGroups.Text = "Unified Groups (one per line -- group display Name; email also works but isn't required)"
$grpGroups.Location = New-Object System.Drawing.Point(452, 50)
$grpGroups.Size = New-Object System.Drawing.Size(430, 230)
$form.Controls.Add($grpGroups)
$txtGroups = New-Object System.Windows.Forms.TextBox
$txtGroups.Multiline = $true
$txtGroups.ScrollBars = "Vertical"
$txtGroups.Location = New-Object System.Drawing.Point(10, 22)
$txtGroups.Size = New-Object System.Drawing.Size(410, 196)
$txtGroups.Font = New-Object System.Drawing.Font("Consolas", 9)
$grpGroups.Controls.Add($txtGroups)

$lblParsed = New-Object System.Windows.Forms.Label
$lblParsed.Text = "Parsed preview (updates live as you type/paste -- no Parse button needed):"
$lblParsed.Location = New-Object System.Drawing.Point(12, 288)
$lblParsed.AutoSize = $true
$form.Controls.Add($lblParsed)
$txtParsed = New-Object System.Windows.Forms.TextBox
$txtParsed.Multiline = $true
$txtParsed.ReadOnly = $true
$txtParsed.ScrollBars = "Vertical"
$txtParsed.BackColor = [System.Drawing.Color]::WhiteSmoke
$txtParsed.Location = New-Object System.Drawing.Point(12, 308)
$txtParsed.Size = New-Object System.Drawing.Size(860, 100)
$txtParsed.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtParsed)

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add"
$btnAdd.Location = New-Object System.Drawing.Point(12, 418)
$btnAdd.Size = New-Object System.Drawing.Size(120, 30)
$btnAdd.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnAdd)
$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remove"
$btnRemove.Location = New-Object System.Drawing.Point(140, 418)
$btnRemove.Size = New-Object System.Drawing.Size(120, 30)
$btnRemove.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRemove)
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(268, 418)
$btnCancel.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnCancel)

$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Result:"
$lblOutput.Location = New-Object System.Drawing.Point(12, 458)
$lblOutput.AutoSize = $true
$form.Controls.Add($lblOutput)
$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Multiline = $true
$txtOutput.ReadOnly = $true
$txtOutput.ScrollBars = "Both"
$txtOutput.WordWrap = $false
$txtOutput.Location = New-Object System.Drawing.Point(12, 478)
$txtOutput.Size = New-Object System.Drawing.Size(860, 232)
$txtOutput.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOutput)

function Update-ParsedPreview {
    $users = ConvertTo-NormalizedList -Text $txtUsers.Text -AsEmail
    $groups = ConvertTo-NormalizedList -Text $txtGroups.Text
    $lines = @("Users ($($users.Count)):")
    if ($users.Count -eq 0) { $lines += "  (none)" } else { $users | ForEach-Object { $lines += "  $_" } }
    $lines += ""
    $lines += "Unified Groups ($($groups.Count)):"
    if ($groups.Count -eq 0) { $lines += "  (none)" } else { $groups | ForEach-Object { $lines += "  $_" } }
    $lines += ""
    $lines += "Total operations: $($users.Count * $groups.Count)"
    $txtParsed.Lines = $lines
}
$txtUsers.Add_TextChanged({ Update-ParsedPreview })
$txtGroups.Add_TextChanged({ Update-ParsedPreview })
Update-ParsedPreview

$btnAdd.Add_Click({
    $btnAdd.Enabled = $false
    $btnRemove.Enabled = $false
    try {
        $users = ConvertTo-NormalizedList -Text $txtUsers.Text -AsEmail
        $groups = ConvertTo-NormalizedList -Text $txtGroups.Text
        if ($users.Count -eq 0 -or $groups.Count -eq 0) {
            $txtOutput.Text = "Enter at least one user and at least one Unified Group first."
            return
        }
        $ops = $users.Count * $groups.Count
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Add $($users.Count) member(s) to $($groups.Count) Unified Group(s) = $ops operation(s).",
            "Confirm Add",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $txtOutput.Text = "Cancelled."
            return
        }
        $txtOutput.Text = Invoke-UnifiedGroupBatch -Users $users -Groups $groups -Verb 'Add'
    } catch {
        $txtOutput.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add Members (Unified Group batch)'
    } finally {
        $btnAdd.Enabled = $true
        $btnRemove.Enabled = $true
    }
})

$btnRemove.Add_Click({
    $btnAdd.Enabled = $false
    $btnRemove.Enabled = $false
    try {
        $users = ConvertTo-NormalizedList -Text $txtUsers.Text -AsEmail
        $groups = ConvertTo-NormalizedList -Text $txtGroups.Text
        if ($users.Count -eq 0 -or $groups.Count -eq 0) {
            $txtOutput.Text = "Enter at least one user and at least one Unified Group first."
            return
        }
        $ops = $users.Count * $groups.Count
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($users.Count) member(s) from $($groups.Count) Unified Group(s) = $ops operation(s).",
            "Confirm Remove",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $txtOutput.Text = "Cancelled."
            return
        }
        $txtOutput.Text = Invoke-UnifiedGroupBatch -Users $users -Groups $groups -Verb 'Remove'
    } catch {
        $txtOutput.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove Members (Unified Group batch)'
    } finally {
        $btnAdd.Enabled = $true
        $btnRemove.Enabled = $true
    }
})

$btnCancel.Add_Click({ $form.Close() })

# ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin

# Console, which opens top-left. Helper is in Common.ps1.

Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
#endregion
