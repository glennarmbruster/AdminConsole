#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
# CLOUD ONLY. New-DistributionGroup exists in BOTH Exchanges and creates the
# group in whichever one the session is pointed at. This tool means EXCHANGE
# ONLINE, and every call goes through Invoke-EXOCommand to say so.
#
# ManageMailObject.ps1 is the ON-PREMISES sibling: it creates groups in AD,
# which then sync up. Use that one when the group has to exist on-premises --
# for a hybrid user population, or because something on-premises needs to
# secure against it. Use THIS one for a group that only ever lives in the
# cloud.
# ---------------------------------------------------------------------------
<#
    NewCloudGroup.ps1
    Create a cloud distribution group or mail-enabled security group.

    Author   : Glenn Armbruster
    Revised  : 2026-09-11
#>
# TOOLKIT-HELP-RUNS: Preview only reads. Create writes -- it asks first and shows exactly what it will do.
# TOOLKIT-HELP-BEGIN
# Creates a group in Exchange Online. Nothing on-premises is touched, and the
# group does not sync down -- it exists only in the cloud.
#
# TWO TYPES, AND THE DIFFERENCE MATTERS:
#
#   Distribution group            Carries mail. Cannot be granted permissions
#                                 on anything. The right choice for a plain
#                                 "everyone in Accounts" list.
#   Mail-enabled security group   Carries mail AND can be granted permissions
#                                 -- mailbox Full Access, Send As, SharePoint,
#                                 and so on. Choose this if the group will ever
#                                 be used to grant access to something.
#
# YOU CANNOT CONVERT ONE INTO THE OTHER LATER. Exchange has no cmdlet for it;
# the fix is to create the other kind and move the members. If you are unsure,
# a mail-enabled security group does everything a distribution group does.
#
# THE OWNER IS REQUIRED. A group with nobody responsible for it is a group
# nobody prunes. By default the owner is also added as a member, which is what
# the Exchange admin centre does and what people expect -- untick the box if
# the owner should manage the group without receiving its mail.
#
# EXTERNAL MAIL IS BLOCKED by default: only authenticated senders inside the
# organisation can send to the group. Tick the box for a group that has to
# receive mail from outside (info@, careers@, a supplier alias).
#
# PRESS PREVIEW FIRST. It resolves the owner and every member, works out the
# addresses, and checks that nothing already holds them -- all without
# creating anything. Create is not enabled until Preview has passed.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

#region Helpers ---------------------------------------------------------------
function exo { param([string]$n, [hashtable]$p) Invoke-EXOCommand -Name $n -Parameters $p }

function Test-CloudReady { try { return [bool](Test-ExoInProcess).Present } catch { return $false } }

function Get-CloudGroupAlias {
    <#
        An alias Exchange will accept, derived from the group's name.

        THE RULES ARE MICROSOFT'S, NOT MINE. Learn (New-DistributionGroup,
        -Alias): letters, numbers and ! # % * + - / = ? ^ _ ~ are allowed;
        "periods must be surrounded by other valid characters"; the maximum
        length is 64. Spaces are not in that list, which is why "Accounts
        Payable" cannot be its own alias.

        Everything outside the set is dropped rather than replaced with a
        hyphen: 'A/P (Shared)' becoming 'A-P--Shared-' is a worse name than
        'APShared', and the operator can always type a better one over it.
    #>
    param([string]$Name)
    $s = ("$Name" -replace '[^A-Za-z0-9!#%*+\-/=?^_~.]', '')
    $s = $s -replace '^\.+', '' -replace '\.+$', ''   # no leading or trailing period
    $s = $s -replace '\.{2,}', '.'                    # and never two in a row
    if ($s.Length -gt 64) { $s = $s.Substring(0, 64) }
    return $s
}

function Resolve-CloudRecipient {
    <#
        One address or name to a cloud recipient, or $null.

        Tried three ways because operators type all three: the display name,
        the primary address, and an alias that is only a secondary proxy.
    #>
    param([string]$Text)
    $t = "$Text".Trim()
    if (-not $t) { return $null }
    $r = $null
    try { $r = exo 'Get-Recipient' @{ Identity = $t; ErrorAction = 'Stop' } } catch { }
    if (-not $r) { try { $r = exo 'Get-Recipient' @{ Filter = "PrimarySmtpAddress -eq '$t'"; ErrorAction = 'SilentlyContinue' } } catch { } }
    if (-not $r) { try { $r = exo 'Get-Recipient' @{ Filter = "EmailAddresses -eq 'smtp:$t'"; ErrorAction = 'SilentlyContinue' } } catch { } }
    if ($r) { return @($r)[0] }
    return $null
}

function Test-CloudAddressFree {
    <#
        Is this address or alias unheld in the tenant?

        CHECKED BEFORE CREATING, AND THAT IS NOT BELT AND BRACES. Microsoft's
        documentation for New-DistributionGroup says nothing about what happens
        when the alias or primary address is already taken -- no error text, no
        described behaviour. An undocumented failure mode halfway through a
        create is exactly the thing to find out about beforehand rather than
        afterwards, so the tool asks first and refuses with a sentence naming
        whatever is in the way.
    #>
    param([string]$Address)
    $a = "$Address".Trim()
    if (-not $a) { return @{ Free = $true; HeldBy = '' } }
    $hit = $null
    try { $hit = exo 'Get-Recipient' @{ Filter = "EmailAddresses -eq 'smtp:$a'"; ErrorAction = 'SilentlyContinue' } } catch { }
    if (-not $hit) { try { $hit = exo 'Get-Recipient' @{ Identity = $a; ErrorAction = 'Stop' } } catch { } }
    if ($hit) {
        $h = @($hit)[0]
        return @{ Free = $false; HeldBy = "$($h.DisplayName) <$($h.PrimarySmtpAddress)> ($($h.RecipientTypeDetails))" }
    }
    return @{ Free = $true; HeldBy = '' }
}

function ConvertFrom-AddressList {
    # One address per line, or separated by ; or , -- because people paste from
    # all three and none of them is wrong.
    param([string]$Text)
    if (-not "$Text".Trim()) { return @() }
    return @("$Text" -split '[;,\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
}
#endregion

#region Form ------------------------------------------------------------------
function Show-NewCloudGroupForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Create Cloud Group'
    $form.ClientSize = New-Object System.Drawing.Size(780, 700)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblConn = New-Object System.Windows.Forms.Label
    $lblConn.Location = New-Object System.Drawing.Point(12, $y); $lblConn.Size = New-Object System.Drawing.Size(756, 18); $lblConn.Tag = 'NoTheme'
    $form.Controls.Add($lblConn)
    $y += 26

    # ---- type -------------------------------------------------------------
    $grpType = New-Object System.Windows.Forms.GroupBox
    $grpType.Text = 'Group type'
    $grpType.Location = New-Object System.Drawing.Point(12, $y); $grpType.Size = New-Object System.Drawing.Size(756, 84)
    $form.Controls.Add($grpType)

    $rdoDist = New-Object System.Windows.Forms.RadioButton
    $rdoDist.Text = 'Distribution group  --  carries mail only'
    $rdoDist.Location = New-Object System.Drawing.Point(14, 24); $rdoDist.Size = New-Object System.Drawing.Size(340, 22)
    $rdoDist.Checked = $true
    $grpType.Controls.Add($rdoDist)

    $rdoSec = New-Object System.Windows.Forms.RadioButton
    $rdoSec.Text = 'Mail-enabled security group  --  can also be granted permissions'
    $rdoSec.Location = New-Object System.Drawing.Point(14, 50); $rdoSec.Size = New-Object System.Drawing.Size(420, 22)
    $grpType.Controls.Add($rdoSec)

    $lblTypeNote = New-Object System.Windows.Forms.Label
    $lblTypeNote.Location = New-Object System.Drawing.Point(440, 24); $lblTypeNote.Size = New-Object System.Drawing.Size(306, 50)
    $lblTypeNote.ForeColor = [System.Drawing.Color]::DimGray
    $lblTypeNote.Text = 'This cannot be changed later -- Exchange has no cmdlet to convert one into the other. If in doubt, a security group does everything a distribution group does.'
    $grpType.Controls.Add($lblTypeNote)
    $y += 92

    # ---- identity ---------------------------------------------------------
    $grpId = New-Object System.Windows.Forms.GroupBox
    $grpId.Text = 'Name and addresses'
    $grpId.Location = New-Object System.Drawing.Point(12, $y); $grpId.Size = New-Object System.Drawing.Size(756, 146)
    $form.Controls.Add($grpId)

    $mk = {
        param([string]$Caption, [int]$Top, [int]$Width, [string]$Hint)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $Caption; $l.Location = New-Object System.Drawing.Point(14, ($Top + 4)); $l.Size = New-Object System.Drawing.Size(120, 18)
        $grpId.Controls.Add($l)
        $t = New-Object System.Windows.Forms.TextBox
        $t.Location = New-Object System.Drawing.Point(140, $Top); $t.Size = New-Object System.Drawing.Size($Width, 24)
        $t.Font = New-Object System.Drawing.Font('Consolas', 9)
        $grpId.Controls.Add($t)
        if ($Hint) {
            $h = New-Object System.Windows.Forms.Label
            $h.Text = $Hint; $h.Location = New-Object System.Drawing.Point((146 + $Width), ($Top + 4)); $h.Size = New-Object System.Drawing.Size((600 - $Width), 18)
            $h.ForeColor = [System.Drawing.Color]::DimGray
            $grpId.Controls.Add($h)
        }
        return $t
    }
    $txtName    = & $mk 'Group name:'   22 300 'max 64 characters'
    $txtDisplay = & $mk 'Display name:' 52 300 'as it appears in the address book'
    $txtAlias   = & $mk 'Alias:'        82 300 'no spaces; filled in from the name'
    $txtSmtp    = & $mk 'Primary SMTP:' 112 420 ''
    $y += 154

    # ---- owner and members -------------------------------------------------
    $grpPeople = New-Object System.Windows.Forms.GroupBox
    $grpPeople.Text = 'Owner and members'
    $grpPeople.Location = New-Object System.Drawing.Point(12, $y); $grpPeople.Size = New-Object System.Drawing.Size(756, 176)
    $form.Controls.Add($grpPeople)

    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Owner (required):'; $l.Location = New-Object System.Drawing.Point(14, 26); $l.Size = New-Object System.Drawing.Size(120, 18)
    $grpPeople.Controls.Add($l)
    $txtOwner = New-Object System.Windows.Forms.TextBox
    $txtOwner.Location = New-Object System.Drawing.Point(140, 22); $txtOwner.Size = New-Object System.Drawing.Size(300, 24)
    $txtOwner.Font = New-Object System.Drawing.Font('Consolas', 9)
    $grpPeople.Controls.Add($txtOwner)

    $chkOwnerMember = New-Object System.Windows.Forms.CheckBox
    $chkOwnerMember.Text = 'Owner is also a member'
    $chkOwnerMember.Location = New-Object System.Drawing.Point(452, 24); $chkOwnerMember.Size = New-Object System.Drawing.Size(290, 22)
    $chkOwnerMember.Checked = $true
    $grpPeople.Controls.Add($chkOwnerMember)

    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Members:'; $l.Location = New-Object System.Drawing.Point(14, 56); $l.Size = New-Object System.Drawing.Size(120, 18)
    $grpPeople.Controls.Add($l)
    $txtMembers = New-Object System.Windows.Forms.TextBox
    $txtMembers.Location = New-Object System.Drawing.Point(140, 52); $txtMembers.Size = New-Object System.Drawing.Size(602, 110)
    $txtMembers.Multiline = $true; $txtMembers.ScrollBars = 'Vertical'
    $txtMembers.Font = New-Object System.Drawing.Font('Consolas', 9)
    $grpPeople.Controls.Add($txtMembers)
    $y += 184

    # ---- options -----------------------------------------------------------
    $grpOpt = New-Object System.Windows.Forms.GroupBox
    $grpOpt.Text = 'Options'
    $grpOpt.Location = New-Object System.Drawing.Point(12, $y); $grpOpt.Size = New-Object System.Drawing.Size(756, 82)
    $form.Controls.Add($grpOpt)

    $chkExternal = New-Object System.Windows.Forms.CheckBox
    $chkExternal.Text = 'Accept mail from outside the organisation'
    $chkExternal.Location = New-Object System.Drawing.Point(14, 24); $chkExternal.Size = New-Object System.Drawing.Size(310, 22)
    $grpOpt.Controls.Add($chkExternal)

    $chkHidden = New-Object System.Windows.Forms.CheckBox
    $chkHidden.Text = 'Hide from the address book'
    $chkHidden.Location = New-Object System.Drawing.Point(14, 50); $chkHidden.Size = New-Object System.Drawing.Size(310, 22)
    $grpOpt.Controls.Add($chkHidden)

    # IgnoreNamingPolicy is offered because a tenant CAN carry a group naming
    # policy (Set-OrganizationConfig -DistributionGroupNamingPolicy) that
    # decorates every new group's display name with a prefix or suffix. When
    # one is in force and nobody remembers it, the group appears in the address
    # book under a name the operator did not type -- and this checkbox is the
    # documented way out.
    $chkIgnorePolicy = New-Object System.Windows.Forms.CheckBox
    $chkIgnorePolicy.Text = 'Ignore the tenant group naming policy'
    $chkIgnorePolicy.Location = New-Object System.Drawing.Point(340, 24); $chkIgnorePolicy.Size = New-Object System.Drawing.Size(400, 22)
    $grpOpt.Controls.Add($chkIgnorePolicy)

    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Notes:'; $l.Location = New-Object System.Drawing.Point(340, 54); $l.Size = New-Object System.Drawing.Size(50, 18)
    $grpOpt.Controls.Add($l)
    $txtNotes = New-Object System.Windows.Forms.TextBox
    $txtNotes.Location = New-Object System.Drawing.Point(392, 50); $txtNotes.Size = New-Object System.Drawing.Size(350, 24)
    $grpOpt.Controls.Add($txtNotes)
    $y += 90

    # ---- output ------------------------------------------------------------
    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline = $true; $txtOut.ReadOnly = $true; $txtOut.ScrollBars = 'Vertical'
    $txtOut.Location = New-Object System.Drawing.Point(12, $y); $txtOut.Size = New-Object System.Drawing.Size(756, 150)
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)
    $y += 158

    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = 'Preview'; $btnPreview.Location = New-Object System.Drawing.Point(12, $y); $btnPreview.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnPreview)
    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = 'Create group'; $btnCreate.Location = New-Object System.Drawing.Point(150, $y); $btnCreate.Size = New-Object System.Drawing.Size(150, 30)
    $btnCreate.Enabled = $false
    $form.Controls.Add($btnCreate)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(668, $y); $btnClose.Size = New-Object System.Drawing.Size(100, 30); $btnClose.Tag = 'Plain'
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    # STATE ON THE FORM, NOT IN VARIABLES BESIDE IT. A '$plan = ...' inside one
    # handler creates that variable in the handler's own scope and is gone the
    # moment it returns -- the next handler reads the stale outer copy and
    # quietly does nothing. A property assignment on the form mutates the thing
    # every handler can see.
    $form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'Plan' -NotePropertyValue $null  -Force

    $log = { param([string]$s) $txtOut.AppendText("$s`r`n") }

    $refreshConn = {
        $ok = Test-CloudReady
        $lblConn.Text = "Exchange Online: $(if ($ok) { 'connected' } else { 'NOT connected -- connect it from the console' })"
        $lblConn.ForeColor = $(if ($ok) { [System.Drawing.Color]::FromArgb(90, 220, 140) } else { [System.Drawing.Color]::FromArgb(240, 175, 60) })
    }

    # Any edit invalidates a preview. Without this, changing the name after a
    # good preview and pressing Create would build the PREVIOUSLY previewed
    # group -- the plan and the form would disagree and the form would win the
    # argument in the operator's head while the plan won it in the tenant.
    $invalidate = {
        if ($form.Plan) { & $log 'Something changed -- press Preview again before creating.' }
        $form.Plan = $null
        $btnCreate.Enabled = $false
    }
    foreach ($c in @($txtName, $txtDisplay, $txtAlias, $txtSmtp, $txtOwner, $txtMembers, $txtNotes)) { $c.Add_TextChanged($invalidate) }
    foreach ($c in @($rdoDist, $rdoSec, $chkExternal, $chkHidden, $chkIgnorePolicy, $chkOwnerMember)) { $c.Add_CheckedChanged($invalidate) }

    # Name drives display name, alias and address until the operator overrides
    # one. Tracked with a flag per box rather than by comparing values, because
    # a comparison cannot tell "typed the same thing" from "never touched it".
    $form | Add-Member -NotePropertyName 'TouchedDisplay' -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'TouchedAlias'   -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'TouchedSmtp'    -NotePropertyValue $false -Force
    $txtDisplay.Add_KeyDown({ $form.TouchedDisplay = $true })
    $txtAlias.Add_KeyDown({   $form.TouchedAlias   = $true })
    $txtSmtp.Add_KeyDown({    $form.TouchedSmtp    = $true })

    $txtName.Add_TextChanged({
        $n = "$($txtName.Text)".Trim()
        if (-not $form.TouchedDisplay) { $txtDisplay.Text = $n }
        $a = Get-CloudGroupAlias -Name $n
        if (-not $form.TouchedAlias) { $txtAlias.Text = $a }
        if (-not $form.TouchedSmtp) {
            $dom = ''
            try { $dom = "$(Get-MailDomain)" } catch { }
            $txtSmtp.Text = $(if ($a -and $dom) { "$a@$dom" } else { '' })
        }
    })

    # ---- preview -----------------------------------------------------------
    $btnPreview.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $txtOut.Clear()
            $form.Plan = $null
            $btnCreate.Enabled = $false
            if (-not (Test-CloudReady)) { & $log 'Exchange Online is not connected in this window.'; return }

            $name    = "$($txtName.Text)".Trim()
            $display = "$($txtDisplay.Text)".Trim()
            $alias   = "$($txtAlias.Text)".Trim()
            $smtp    = "$($txtSmtp.Text)".Trim()
            $ownerIn = "$($txtOwner.Text)".Trim()
            $type    = $(if ($rdoSec.Checked) { 'Security' } else { 'Distribution' })

            # Length caps are Exchange's (Learn: Name 64, DisplayName 256,
            # Alias 64). Checked here so the message names the limit instead of
            # leaving the cmdlet to say something vaguer.
            if (-not $name)    { & $log 'Group name is required.'; $txtName.Focus(); return }
            if ($name.Length -gt 64) { & $log "Group name is limited to 64 characters; that is $($name.Length)."; return }
            if (-not $display) { $display = $name }
            if ($display.Length -gt 256) { & $log "Display name is limited to 256 characters; that is $($display.Length)."; return }
            if (-not $alias)   { & $log 'Alias is required -- it could not be derived from the name. Type one.'; $txtAlias.Focus(); return }
            if ($alias.Length -gt 64) { & $log "Alias is limited to 64 characters; that is $($alias.Length)."; return }
            if (-not $smtp)    { & $log 'Primary SMTP address is required.'; $txtSmtp.Focus(); return }
            if ($smtp -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { & $log "'$smtp' is not a valid email address."; return }
            if (-not $ownerIn) { & $log 'OWNER IS REQUIRED. A group with nobody responsible for it is a group nobody prunes.'; $txtOwner.Focus(); return }

            Write-ToolStep 'Resolving the owner ...'
            $owner = Resolve-CloudRecipient -Text $ownerIn
            if (-not $owner) { & $log "Owner '$ownerIn' was not found in the cloud."; return }
            & $log "Owner   : $($owner.DisplayName) <$($owner.PrimarySmtpAddress)>  ($($owner.RecipientTypeDetails))"

            Write-ToolStep 'Checking the addresses are free ...'
            foreach ($probe in @(@{ L = 'Primary SMTP'; V = $smtp }, @{ L = 'Alias'; V = $alias }, @{ L = 'Name'; V = $name })) {
                $r = Test-CloudAddressFree -Address $probe.V
                if (-not $r.Free) {
                    & $log ''
                    & $log "CONFLICT: $($probe.L) '$($probe.V)' is already held by:"
                    & $log "  $($r.HeldBy)"
                    & $log 'Nothing was created. Change it and preview again.'
                    return
                }
            }
            & $log "Addresses: '$smtp' and alias '$alias' are free."

            Write-ToolStep 'Resolving members ...'
            $wanted = @(ConvertFrom-AddressList -Text $txtMembers.Text)
            $members = @(); $badMembers = @()
            foreach ($m in $wanted) {
                $r = Resolve-CloudRecipient -Text $m
                if ($r) { $members += "$($r.PrimarySmtpAddress)" } else { $badMembers += $m }
            }
            if ($chkOwnerMember.Checked) {
                # ADDED EXPLICITLY, NOT LEFT TO THE CMDLET. Learn is blunt about
                # it: "Owners that you specify with this parameter are not
                # automatically added as group members." There is a
                # CopyOwnerToMember switch, but naming the owner in the members
                # list does the same job with a parameter this tool has already
                # verified, and leaves one less thing to be wrong about.
                $oa = "$($owner.PrimarySmtpAddress)"
                if ($members -notcontains $oa) { $members += $oa }
            }
            $members = @($members | Select-Object -Unique)
            foreach ($b in $badMembers) { & $log "  member '$b' -- NOT FOUND, it will be skipped" }
            & $log "Members : $($members.Count) resolved$(if ($badMembers.Count) { ", $($badMembers.Count) not found" })"

            $form.Plan = [PSCustomObject]@{
                Name = $name; Display = $display; Alias = $alias; Smtp = $smtp
                Type = $type; Owner = $owner; Members = $members
                External = [bool]$chkExternal.Checked
                Hidden = [bool]$chkHidden.Checked
                IgnorePolicy = [bool]$chkIgnorePolicy.Checked
                Notes = "$($txtNotes.Text)".Trim()
            }
            & $log ''
            & $log '---- will create ----'
            & $log "  Type        : $(if ($type -eq 'Security') { 'mail-enabled SECURITY group (can hold permissions)' } else { 'distribution group (mail only)' })"
            & $log "  Name        : $name"
            & $log "  Display     : $display"
            & $log "  Alias       : $alias"
            & $log "  Primary     : $smtp"
            & $log "  Owner       : $($owner.PrimarySmtpAddress)$(if ($chkOwnerMember.Checked) { '  (also a member)' } else { '  (not a member)' })"
            & $log "  Members     : $($members.Count)"
            & $log "  External    : $(if ($chkExternal.Checked) { 'ACCEPTED from outside the organisation' } else { 'blocked -- internal authenticated senders only' })"
            & $log "  Address book: $(if ($chkHidden.Checked) { 'hidden' } else { 'visible' })"
            if ($chkIgnorePolicy.Checked) { & $log '  Naming policy: IGNORED' }
            if ("$($txtNotes.Text)".Trim()) { & $log "  Notes       : $("$($txtNotes.Text)".Trim())" }
            & $log ''
            & $log 'Preview passed. Create is unlocked.'
            $btnCreate.Enabled = $true
        } catch {
            & $log "Preview failed: $($_.Exception.Message)"
            Show-HandlerError -ErrorRecord $_ -Context 'Preview the new cloud group'
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    # ---- create ------------------------------------------------------------
    $btnCreate.Add_Click({
        if ($form.Busy) { return }
        $p = $form.Plan
        if (-not $p) { & $log 'Press Preview first.'; return }
        $form.Busy = $true
        try {
            if (-not (Test-CloudReady)) { & $log 'Exchange Online is not connected in this window.'; return }
            $a = Show-ToolDialog(
                ("Create this group in Exchange Online?" +
                 "`r`n`r`n  $($p.Display)  <$($p.Smtp)>" +
                 "`r`n  $(if ($p.Type -eq 'Security') { 'Mail-enabled SECURITY group' } else { 'Distribution group' })" +
                 "`r`n  Owner: $($p.Owner.PrimarySmtpAddress)" +
                 "`r`n  Members: $($p.Members.Count)" +
                 "`r`n  External mail: $(if ($p.External) { 'ACCEPTED' } else { 'blocked' })" +
                 "`r`n`r`nThe type cannot be changed afterwards."),
                'Confirm create', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { & $log 'Cancelled.'; return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            Write-ToolStep "Creating $($p.Name) ..."
            # Learn (New-DistributionGroup): -Name is mandatory; -Type accepts
            # Distribution or Security, where Security creates a MAIL-ENABLED
            # SECURITY GROUP and Distribution creates a group that "can't have
            # permissions assigned"; -ManagedBy and -Members both take multiple
            # values; -RequireSenderAuthenticationEnabled $true means "Messages
            # are accepted only from authenticated (internal) senders".
            #
            # RequireSenderAuthenticationEnabled IS SENT EXPLICITLY even though
            # Learn says its default is already $true. A default is a fact about
            # today's Exchange, not a decision this tool made -- and the one
            # security-relevant setting here should be readable in the command
            # rather than inferred from documentation.
            $np = @{
                Name                              = $p.Name
                DisplayName                       = $p.Display
                Alias                             = $p.Alias
                PrimarySmtpAddress                = $p.Smtp
                Type                              = $p.Type
                ManagedBy                         = "$($p.Owner.PrimarySmtpAddress)"
                RequireSenderAuthenticationEnabled = (-not $p.External)
                ErrorAction                       = 'Stop'
            }
            if ($p.Notes) { $np.Notes = $p.Notes }
            # Learn (New-DistributionGroup): -IgnoreNamingPolicy "specifies
            # whether to prevent this group from being affected by your
            # organization's group naming policy".
            if ($p.IgnorePolicy) { $np.IgnoreNamingPolicy = $true }
            exo 'New-DistributionGroup' $np | Out-Null
            & $log "Created: $($p.Display) <$($p.Smtp)>  ($($p.Type))"
            Write-CompletedLog -Action 'NewCloudGroup' -Target (Get-LogTarget $p.Smtp $p.Name) -Result 'Success' `
                -Detail "$($p.Type) group; owner $($p.Owner.PrimarySmtpAddress); external mail $(if ($p.External) { 'accepted' } else { 'blocked' })"

            if ($p.Hidden) {
                try {
                    # Learn (Set-DistributionGroup): -HiddenFromAddressListsEnabled.
                    # Set AFTER creation rather than at create: New- does not
                    # document the parameter, and sending one it does not take
                    # would fail the whole create over a cosmetic setting.
                    exo 'Set-DistributionGroup' @{ Identity = $p.Smtp; HiddenFromAddressListsEnabled = $true; ErrorAction = 'Stop' } | Out-Null
                    & $log 'Hidden from the address book.'
                } catch { & $log "Could not hide it from the address book: $($_.Exception.Message)" }
            }

            $ok = 0; $fail = 0
            foreach ($m in @($p.Members)) {
                Write-ToolStep "Adding $m ..."
                try {
                    # Learn (Add-DistributionGroupMember): -Identity, -Member
                    # (singular -- one call per member),
                    # -BypassSecurityGroupManagerCheck, which is required for
                    # "a user who isn't an owner of the group".
                    exo 'Add-DistributionGroupMember' @{ Identity = $p.Smtp; Member = $m; BypassSecurityGroupManagerCheck = $true; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                    & $log "  added   $m"
                    $ok++
                } catch { & $log "  FAILED  $m -- $($_.Exception.Message)"; $fail++ }
            }
            if (@($p.Members).Count) {
                & $log "Members: $ok added, $fail failed."
                Write-CompletedLog -Action 'NewCloudGroupMembers' -Target (Get-LogTarget $p.Smtp $p.Name) -Result $(if ($fail -and $ok) { 'Partial' } elseif ($fail) { 'Failed' } else { 'Success' }) `
                    -Detail "$ok added, $fail failed"
            }

            & $log ''
            & $log 'Done. The group is usable immediately; the address book can take a few minutes to show it.'
            $btnCreate.Enabled = $false
            $form.Plan = $null
        } catch {
            & $log "Create failed: $($_.Exception.Message)"
            Show-HandlerError -ErrorRecord $_ -Context 'New-DistributionGroup (cloud)'
            Write-FailedLog -Action 'NewCloudGroup' -Target (Get-LogTarget "$($p.Smtp)" "$($p.Name)") -ErrorRecord $_
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    $form.Add_Shown({
        & $refreshConn
        # The console passes the looked-up address in. Nine times out of ten
        # the person on screen is the one who should own the new group, so it
        # is offered rather than assumed -- the box is editable and Preview
        # resolves whatever is finally in it.
        if ("$SmtpAddress".Trim() -and -not "$($txtOwner.Text)".Trim()) {
            $txtOwner.Text = (Get-TargetSmtp -Text $SmtpAddress)
        }
        $txtName.Focus()
    })

    Set-ConsoleTheme -Form $form
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Create Cloud Group' | Out-Null
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    [void]$form.ShowDialog()
}
#endregion

Show-NewCloudGroupForm
