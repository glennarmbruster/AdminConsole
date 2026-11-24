#Requires -Version 5.1
<#
    UnifiedGroup.ps1
    View and edit Microsoft 365 group membership.

    Author   : Glenn Armbruster
    Revised  : 2026-07-28
#>
# TOOLKIT-HELP-BEGIN
# View and edit the membership of a Microsoft 365 group.
#
# Microsoft 365 groups are not distribution lists, even though they receive mail
# like one. They also carry a Teams team, a SharePoint site and a shared mailbox,
# so adding somebody here gives them more than an email feed.
#
# Owners and members are separate lists and both matter -- see Group Editor if
# you need to change ownership.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Parsing --------------------------------------------------------------
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

$form.StartPosition = 'Manual'
Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'UnifiedGroup Editor' | Out-Null

[void]$form.ShowDialog()
#endregion
