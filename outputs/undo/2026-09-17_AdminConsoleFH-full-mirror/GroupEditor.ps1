#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
# CLOUD VERSION, AND NOW THE ONLY VERSION.
#
# AD-GroupEditor.ps1 was the on-premises twin. It has been PARKED to
# DisabledTools\, so there is no longer an AD- variant beside this file and
# nothing here needs a domain controller.
#
# The header used to say the AD- variant was "beside it", which stopped being
# true the moment the file moved. A comment that names a sibling is a
# reference like any other -- it goes stale silently, and the next person
# reads it as current.
# -------------------------------------------------------------------------
<#
    GroupEditor.ps1
    Standalone Group Editor (owner / member delegation for M365 Groups and
    on-prem synced groups) tool -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# View and change who owns and who belongs to a group -- both Microsoft 365
# groups and on-premises groups synced to the cloud.
#
# Look the group up, then add or remove owners and members. Owners matter more
# than they look: a group with no owner cannot be managed by anyone except an
# admin, and nobody gets told when the last one leaves.
#
# Changes take effect immediately and are written to the audit log.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress
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


# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. Resolve-ToolTarget is not used any more:
# its job was to decide which DIRECTORY an address belongs to, and this is the
# CLOUD variant -- every read and write below goes to Exchange Online.
#
# An empty address is fine. The window opens and asks.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

# WHICH KIND OF GROUP, RESOLVED PER LOOKUP RATHER THAN ONCE AT STARTUP.
#
# A Microsoft 365 group and a distribution group are edited by different
# cmdlets, so the type is not decoration -- it decides which code runs. It used
# to be resolved here, before the form existed, which was correct while the
# address could not change. It can now, so it is resolved inside
# Update-StateDisplay and left EMPTY when the address is not a group.
#
# Empty is a real state, not a missing value: the window stays open showing why,
# and every handler refuses while it lasts. The old behaviour -- a modal error
# and exit -- is wrong once the operator can simply type a different address.
$script:GroupType = $null

#region Actions - group -------------------------------------------------------
function ConvertTo-MemberList {
    <#
        Splits a text box into a de-duplicated list of names.

        Accepts newlines, commas and semicolons, because a list pasted out of a
        ticket arrives in whichever of those the writer happened to use.

        DUPLICATES ARE DROPPED CASE-INSENSITIVELY. The same person twice is not
        two operations -- the second would fail with 'already a member' and be
        reported as an error the operator has to read and dismiss.
    #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $seen = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    $kept = New-Object System.Collections.ArrayList
    foreach ($line in ($Text -split '[\r\n,;]+')) {
        $v = "$line".Trim()
        if (-not $v) { continue }
        if ($seen.Add($v)) { [void]$kept.Add($v) }
    }
    return @($kept.ToArray())
}

function Get-GroupState {
    <#
        Everything the editor shows about one group, either kind.

        Microsoft 365 groups and distribution groups are read by different
        cmdlets that return different property names, so they are normalised
        here into one shape -- otherwise every caller has to branch, and the
        display code drifts apart between the two.

        Errors are collected rather than thrown: a group whose members cannot
        be listed is still worth showing with its properties, and an editor
        that refuses to open teaches nobody anything.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][ValidateSet('UnifiedGroup', 'DistributionGroup')][string]$Type
    )

    $state = [PSCustomObject]@{
        DisplayName  = ''
        PrimarySmtp  = $Smtp
        GroupType    = $Type
        Hidden       = $false
        Owners       = @()
        Members      = @()
        Errors       = @()
    }

    if ($Type -eq 'UnifiedGroup') {
        try {
            $g = Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{
                Identity = $Smtp; ErrorAction = 'Stop' }
            $state.DisplayName = "$($g.DisplayName)"
            $state.PrimarySmtp = "$($g.PrimarySmtpAddress)"
            $state.GroupType   = "$($g.RecipientTypeDetails)"
            $state.Hidden      = [bool]$g.HiddenFromAddressListsEnabled
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroup')
            return $state
        }
        foreach ($link in @('Owners', 'Members')) {
            try {
                $vals = @(Invoke-EXOCommand -Name 'Get-UnifiedGroupLinks' -Parameters @{
                    Identity = $Smtp; LinkType = $link; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
                    ForEach-Object { "$($_.DisplayName) <$($_.PrimarySmtpAddress)>" })
                if ($link -eq 'Owners') { $state.Owners = $vals } else { $state.Members = $vals }
            } catch {
                $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context "Get-UnifiedGroupLinks ($link)")
            }
        }
        return $state
    }

    try {
        $g = Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop' }
        $state.DisplayName = "$($g.DisplayName)"
        $state.PrimarySmtp = "$($g.PrimarySmtpAddress)"
        $state.GroupType   = "$($g.RecipientTypeDetails)"
        $state.Hidden      = [bool]$g.HiddenFromAddressListsEnabled

        # A directory-synced group is mastered somewhere else. Saying so up
        # front is the difference between "this tool is broken" and "this
        # change has to be made at the source".
        if ($g.IsDirSynced) {
            $state.Errors += 'This group is directory-synced (IsDirSynced True). Owner and membership changes must be made where it is mastered; Exchange Online will refuse them here.'
        }
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroup')
        return $state
    }

    foreach ($o in @($g.ManagedBy)) {
        try {
            $r = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
                Identity = "$o"; ErrorAction = 'Stop' }
            $state.Owners += "$($r.DisplayName) <$($r.PrimarySmtpAddress)>"
        } catch {
            $state.Owners += "$o (could not re-resolve)"
        }
    }
    try {
        $state.Members = @(Invoke-EXOCommand -Name 'Get-DistributionGroupMember' -Parameters @{
            Identity = $Smtp; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
            ForEach-Object { "$($_.DisplayName) <$($_.PrimarySmtpAddress)>  [$($_.RecipientTypeDetails)]" })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroupMember')
    }
    return $state
}
function Format-GroupState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading (Group) ---")
    [void]$sb.AppendLine("  Display Name        : $($State.DisplayName)")
    [void]$sb.AppendLine("  Primary SMTP        : $($State.PrimarySmtp)")
    [void]$sb.AppendLine("  Group Type          : $($State.GroupType)")
    [void]$sb.AppendLine("  Hidden From GAL     : $($State.Hidden)")
    [void]$sb.AppendLine("  Owner(s)            :")
    if ($State.Owners.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.Owners | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    [void]$sb.AppendLine("  Members ($($State.Members.Count)) :")
    if ($State.Members.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.Members | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    if ($State.Errors.Count -gt 0) {
        [void]$sb.AppendLine("  --- Field errors (partial data above is still valid) ---")
        $State.Errors | ForEach-Object { [void]$sb.AppendLine("  $_") }
    }
    return $sb.ToString()
}
function Add-UnifiedOwner {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Owners'; Links = $Grantee; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-UnifiedOwner {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Owners'; Links = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Add-UnifiedMember {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Members'; Links = $Grantee; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-UnifiedMember {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Members'; Links = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Set-GroupOwner {
    param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string]$Owner)
    Invoke-EXOCommand -Name 'Set-DistributionGroup' -Parameters @{
        Identity = $Identity; ManagedBy = $Owner; ErrorAction = 'Stop' } | Out-Null
}
function Add-GroupMemberCloud {
    param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string]$Member)
    Invoke-EXOCommand -Name 'Add-DistributionGroupMember' -Parameters @{
        Identity = $Identity; Member = $Member; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
}
function Remove-GroupMemberCloud {
    param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string]$Member)
    Invoke-EXOCommand -Name 'Remove-DistributionGroupMember' -Parameters @{
        Identity = $Identity; Member = $Member; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
}
#endregion

#region NEW - wrapper window for group delegation (see .DESCRIPTION above)
function Show-GroupEditorForm {
    # NO PARAMETERS. Both the address and the group type can change while the
    # window is open, so they live in $script:Target and $script:GroupType
    # rather than in parameters fixed the moment the form was built.
    $editForm = New-Object System.Windows.Forms.Form
    $editForm.Text = 'Group Editor'
    $editForm.ClientSize = New-Object System.Drawing.Size(884, 712)
    $editForm.StartPosition = "CenterScreen"
    $editForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline = $true
    $txtOut.ReadOnly = $true
    $txtOut.ScrollBars = "Vertical"
    $txtOut.Location = New-Object System.Drawing.Point(12, 12)
    $txtOut.Size = New-Object System.Drawing.Size(860, 290)
    $txtOut.Anchor = 'Top,Left,Right,Bottom'
    $txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
    $editForm.Controls.Add($txtOut)

    # TWO BOXES, BECAUSE THESE ARE TWO DIFFERENT OPERATIONS.
    #
    # One 'Grantee' field drove four buttons -- two owner actions and two
    # membership actions -- which made ownership and membership look like the
    # same job done twice. They are not: an owner is normally one person and is
    # changed rarely, while membership moves in handfuls and moves both ways at
    # once.
    #
    # Sharing the field also forced membership work to be done one name at a
    # time, with a confirmation prompt for each. Adding four people and removing
    # three meant seven round trips through the same box.
    $grpOwner = New-Object System.Windows.Forms.GroupBox
    $grpOwner.Text = 'Owner'
    $grpOwner.Location = New-Object System.Drawing.Point(12, 312)
    $grpOwner.Size = New-Object System.Drawing.Size(860, 92)
    $grpOwner.Anchor = 'Left,Right,Bottom'
    $editForm.Controls.Add($grpOwner)

    $lblGrpGrantee = New-Object System.Windows.Forms.Label
    $lblGrpGrantee.Text = "Grantee (email or logon):"
    $lblGrpGrantee.Location = New-Object System.Drawing.Point(10, 24)
    $lblGrpGrantee.AutoSize = $true
    $grpOwner.Controls.Add($lblGrpGrantee)

    # ONE name, deliberately. An owner change is a single-subject operation and
    # the field should say so rather than accepting a list it will ignore.
    $txtGrpGrantee = New-Object System.Windows.Forms.TextBox
    $txtGrpGrantee.Location = New-Object System.Drawing.Point(180, 21)
    $txtGrpGrantee.Size = New-Object System.Drawing.Size(330, 22)
    $txtGrpGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grpOwner.Controls.Add($txtGrpGrantee)

    $btnOwnerSet = New-Object System.Windows.Forms.Button
    $btnOwnerSet.Text = "Set/Add Owner"
    $btnOwnerSet.Location = New-Object System.Drawing.Point(10, 54)
    $btnOwnerSet.Size = New-Object System.Drawing.Size(160, 28)
    $grpOwner.Controls.Add($btnOwnerSet)

    $btnOwnerRemove = New-Object System.Windows.Forms.Button
    $btnOwnerRemove.Text = "Remove Owner"
    $btnOwnerRemove.Location = New-Object System.Drawing.Point(180, 54)
    $btnOwnerRemove.Size = New-Object System.Drawing.Size(160, 28)
    $grpOwner.Controls.Add($btnOwnerRemove)

    $grpMembers = New-Object System.Windows.Forms.GroupBox
    $grpMembers.Text = "Membership"
    $grpMembers.Location = New-Object System.Drawing.Point(12, 412)
    $grpMembers.Size = New-Object System.Drawing.Size(860, 250)
    $grpMembers.Anchor = 'Left,Right,Bottom'
    $editForm.Controls.Add($grpMembers)

    $lblAdd = New-Object System.Windows.Forms.Label
    $lblAdd.Text = "Members to add  (one per line):"
    $lblAdd.Location = New-Object System.Drawing.Point(10, 22)
    $lblAdd.AutoSize = $true
    $grpMembers.Controls.Add($lblAdd)

    $txtAdd = New-Object System.Windows.Forms.TextBox
    $txtAdd.Multiline = $true
    $txtAdd.ScrollBars = "Vertical"
    $txtAdd.Location = New-Object System.Drawing.Point(10, 42)
    $txtAdd.Size = New-Object System.Drawing.Size(410, 150)
    $txtAdd.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grpMembers.Controls.Add($txtAdd)

    $lblRemove = New-Object System.Windows.Forms.Label
    $lblRemove.Text = "Members to remove  (one per line):"
    $lblRemove.Location = New-Object System.Drawing.Point(440, 22)
    $lblRemove.AutoSize = $true
    $grpMembers.Controls.Add($lblRemove)

    $txtRemove = New-Object System.Windows.Forms.TextBox
    $txtRemove.Multiline = $true
    $txtRemove.ScrollBars = "Vertical"
    $txtRemove.Location = New-Object System.Drawing.Point(440, 42)
    $txtRemove.Size = New-Object System.Drawing.Size(410, 150)
    $txtRemove.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grpMembers.Controls.Add($txtRemove)

    # ONE button for both boxes. Adds run before removals -- see
    # Invoke-MembershipChange for why that order is not arbitrary.
    $btnMembership = New-Object System.Windows.Forms.Button
    $btnMembership.Text = "Process Membership Change"
    $btnMembership.Location = New-Object System.Drawing.Point(10, 204)
    $btnMembership.Size = New-Object System.Drawing.Size(240, 32)
    $grpMembers.Controls.Add($btnMembership)

    $lblMemHint = New-Object System.Windows.Forms.Label
    $lblMemHint.Text = "Either box may be left empty. Nothing is written until you confirm."
    $lblMemHint.Location = New-Object System.Drawing.Point(262, 212)
    $lblMemHint.AutoSize = $true
    $lblMemHint.ForeColor = [System.Drawing.Color]::DimGray
    $grpMembers.Controls.Add($lblMemHint)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 672)
    $btnRefresh.Anchor = 'Left,Bottom'
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $editForm.Controls.Add($btnRefresh)
    $btnEditClose = New-Object System.Windows.Forms.Button
    $btnEditClose.Text = "Close"
    $btnEditClose.Location = New-Object System.Drawing.Point(772, 672)
    $btnEditClose.Anchor = 'Right,Bottom'
    $btnEditClose.Size = New-Object System.Drawing.Size(100, 28)
    $editForm.Controls.Add($btnEditClose)

    function Set-EditingEnabled {
        # The two boxes are the only things that WRITE. Greying them is how the
        # window says "I do not know what this is yet" without a dialog.
        param([bool]$On)
        $grpOwner.Enabled   = $On
        $grpMembers.Enabled = $On
        $grpOwner.Text = if ($On -and $script:GroupType) { "Owner  ($script:GroupType)" } else { 'Owner' }
    }

    function Update-StateDisplay {
        $script:GroupType = $null

        if (-not "$script:Target".Trim()) {
            $editForm.Text = 'Group Editor'
            $txtOut.Text   = 'Enter a group address above and press Look Up.'
            Set-EditingEnabled $false
            return
        }

        $editForm.Text = "Group Editor: $script:Target"
        Set-EditingEnabled $false

        try {
            $info = Resolve-Target -Smtp $script:Target
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Resolve-Target'
            return
        }

        if ($info.Type -ne 'UnifiedGroup' -and $info.Type -ne 'DistributionGroup') {
            # REPORTED IN THE WINDOW, NOT IN A MESSAGE BOX. The operator's next
            # move is to correct the address in the box that is already in front
            # of them, and a modal they have to dismiss first is in the way.
            $detail = if ($info.ResolutionErrors.Count -gt 0) {
                ($info.ResolutionErrors -join "`n`n")
            } else {
                "No Microsoft 365 group or distribution group was found for '$script:Target'."
            }
            $txtOut.Text = "'$script:Target' did not resolve to a group (Type: $($info.Type)).`r`n`r`n$detail"
            return
        }

        $script:GroupType = $info.Type

        # ADOPT THE ADDRESS RESOLVE-TARGET ACTUALLY FOUND.
        #
        # Typing a bare name gets a domain appended by Get-TargetSmtp, and for
        # an object whose only address is @tenant.onmicrosoft.com that guess is
        # wrong. Resolve-Target now retries with the bare name and returns the
        # REAL primary address when it finds one that way.
        #
        # Without this line the lookup would succeed and every edit after it
        # would still be aimed at the invented address -- a worse failure than
        # the original, because the window would look like it had worked.
        if ("$($info.Smtp)".Trim() -and "$($info.Smtp)" -ne "$script:Target") {
            $script:Target  = "$($info.Smtp)"
            $txtAddr.Text   = $script:Target
            $form.Text      = "Group Editor: $script:Target"
        }

        try {
            $state = Get-GroupState -Smtp $script:Target -Type $script:GroupType
            $txtOut.Text = Format-GroupState -Heading "CURRENT STATE" -State $state
            Set-EditingEnabled $true
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Group Editor'
        }
    }

    function Invoke-GroupEdit {
        param([string]$ActionLabel, [string]$Verb, [scriptblock]$Action)
        if (-not $script:GroupType) { $txtOut.Text = 'Look up a group first.'; return }
        $grantee = $txtGrpGrantee.Text.Trim()
        if (-not $grantee) { $txtOut.Text = "Enter a grantee (email or logon) first."; return }
        $confirm = Show-ToolDialog(
            "$ActionLabel for '$grantee' on group '$script:Target'?",
            "Confirm $ActionLabel",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $grpOwner.Enabled = $false
        try {
            $before = Get-GroupState -Smtp $script:Target -Type $script:GroupType
            $out = New-Object System.Text.StringBuilder
            [void]$out.Append((Format-GroupState -Heading "BEFORE" -State $before))
            & $Action
            [void]$out.AppendLine("$ActionLabel : SUCCESS ($grantee)")
            $after = Get-GroupState -Smtp $script:Target -Type $script:GroupType
            [void]$out.Append((Format-GroupState -Heading "AFTER" -State $after))
            $txtOut.Text = $out.ToString()
            Write-CompletedLog -Task "$Verb $script:Target (grantee: $grantee)"
        }
        catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context $ActionLabel
        }
        finally {
            $grpOwner.Enabled = $true
        }
    }

    Update-StateDisplay
    $btnRefresh.Add_Click({ Update-StateDisplay })
    $btnOwnerSet.Add_Click({
        Invoke-GroupEdit -ActionLabel "Set/Add Owner" -Verb "Set Owner" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($script:GroupType -eq 'UnifiedGroup') { Add-UnifiedOwner -Smtp $script:Target -Grantee $grantee }
            else { Set-GroupOwner -Identity $script:Target -Owner $grantee }
        }
    })
    $btnOwnerRemove.Add_Click({
        Invoke-GroupEdit -ActionLabel "Remove Owner" -Verb "Removed Owner" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($script:GroupType -eq 'UnifiedGroup') { Remove-UnifiedOwner -Smtp $script:Target -Grantee $grantee }
            else { throw "Set-DistributionGroup -ManagedBy replaces the whole owner list rather than removing one entry. Use 'Set/Add Owner' with the owners you want to keep." }
        }
    })
    $btnMembership.Add_Click({
        # ADDS FIRST, THEN REMOVALS, AND THE ORDER IS NOT ARBITRARY.
        #
        # A list that both adds and removes people is usually a handover. Doing
        # the removals first means that between the two passes the group is
        # smaller than it was at either end -- and if the run fails in the
        # middle, it fails having taken access away and not yet granted it.
        # Adding first means a mid-run failure leaves the group with too many
        # members rather than too few, which is the safer way to be wrong.
        #
        # EACH NAME IS INDEPENDENT. One that will not resolve is reported and
        # skipped, not fatal -- a typo on line four should not silently abandon
        # lines five onwards after four have already been written.
        if (-not $script:GroupType) { $txtOut.Text = 'Look up a group first.'; return }

        $toAdd    = @(ConvertTo-MemberList -Text $txtAdd.Text)
        $toRemove = @(ConvertTo-MemberList -Text $txtRemove.Text)

        if ($toAdd.Count -eq 0 -and $toRemove.Count -eq 0) {
            $txtOut.Text = "Nothing to do -- put at least one name in Members to add or Members to remove."
            return
        }

        # The count is in the prompt because the boxes scroll. Twelve names and
        # three names look identical once the box is full.
        $lines = @()
        if ($toAdd.Count)    { $lines += "ADD ($($toAdd.Count)):";    $lines += ($toAdd    | ForEach-Object { "   $_" }) }
        if ($toRemove.Count) { $lines += "REMOVE ($($toRemove.Count)):"; $lines += ($toRemove | ForEach-Object { "   $_" }) }
        $confirm = Show-ToolDialog(
            ("Apply these membership changes to '$script:Target'?`r`n`r`n" + ($lines -join "`r`n")),
            'Confirm membership change',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = 'Cancelled. Nothing was changed.'; return }

        $grpMembers.Enabled = $false
        $grpOwner.Enabled   = $false
        try {
            $out = New-Object System.Text.StringBuilder
            [void]$out.Append((Format-GroupState -Heading 'BEFORE' -State (Get-GroupState -Smtp $script:Target -Type $script:GroupType)))
            [void]$out.AppendLine('')

            $ok = 0; $bad = 0
            foreach ($m in $toAdd) {
                try {
                    if ($script:GroupType -eq 'UnifiedGroup') { Add-UnifiedMember -Smtp $script:Target -Grantee $m }
                    else { Add-GroupMemberCloud -Identity $script:Target -Member $m }
                    [void]$out.AppendLine("  ADD    $m  -- OK")
                    Write-CompletedLog -Task "Added Member $script:Target (grantee: $m)"
                    $ok++
                } catch {
                    [void]$out.AppendLine("  ADD    $m  -- FAILED: $($_.Exception.Message)")
                    Write-FailedLog -Action 'AddGroupMember' -Target $script:Target -ErrorRecord $_ -Detail "member: $m"
                    $bad++
                }
            }
            foreach ($m in $toRemove) {
                try {
                    if ($script:GroupType -eq 'UnifiedGroup') { Remove-UnifiedMember -Smtp $script:Target -Grantee $m }
                    else { Remove-GroupMemberCloud -Identity $script:Target -Member $m }
                    [void]$out.AppendLine("  REMOVE $m  -- OK")
                    Write-CompletedLog -Task "Removed Member $script:Target (grantee: $m)"
                    $ok++
                } catch {
                    [void]$out.AppendLine("  REMOVE $m  -- FAILED: $($_.Exception.Message)")
                    Write-FailedLog -Action 'RemoveGroupMember' -Target $script:Target -ErrorRecord $_ -Detail "member: $m"
                    $bad++
                }
            }

            [void]$out.AppendLine('')
            [void]$out.AppendLine("$ok succeeded, $bad failed.")
            [void]$out.AppendLine('')
            [void]$out.Append((Format-GroupState -Heading 'AFTER' -State (Get-GroupState -Smtp $script:Target -Type $script:GroupType)))
            $txtOut.Text = $out.ToString()

            # CLEARED ONLY WHEN EVERYTHING WORKED. A box emptied after a partial
            # run destroys the record of which names still need attention.
            if ($bad -eq 0) { $txtAdd.Clear(); $txtRemove.Clear() }
        }
        catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Membership change'
        }
        finally {
            $grpMembers.Enabled = $true
            $grpOwner.Enabled   = $true
        }
    })
    $btnEditClose.Add_Click({ $editForm.Close() })

    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down and
    # grows the form, and the help button places itself from ClientSize at run
    # time. Before Set-FormScreenCorner too, which positions from the size.
    $bar = Add-ToolTargetBar -Form $editForm -Label 'Group address:' -InitialValue $script:Target

    # WIRED HERE, NOT PASSED INTO THE HELPER. Update-StateDisplay is nested in
    # this function, so the handler has to be built in this scope to see it.
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-StateDisplay
    })

    Set-ConsoleTheme -Form $editForm
    Add-ToolHelpButton -Form $editForm -ScriptPath $PSCommandPath -Title 'Group Editor' | Out-Null
    Set-FormScreenCorner -Form $editForm -Corner 'BottomLeft'
    [void]$editForm.ShowDialog()
}
#endregion

Show-GroupEditorForm
