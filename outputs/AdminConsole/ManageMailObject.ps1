#Requires -Version 5.1
# TOOLKIT-VARIANT: OnPrem
# TOOLKIT-REQUIRES: OwnProcess  -- New-/Remove-DistributionGroup exist in BOTH
#   Exchanges and create or delete a different object in each. New-RemoteMailbox
#   is on-premises only. Sharing a process with a connected Exchange Online
#   session is how a create lands in the wrong directory with no error.
# ---------------------------------------------------------------------------
<#
    ManageMailObject.ps1
    Create or remove a shared mailbox or a distribution group.

    Named ManageMailObject rather than New-Object: New-Object is a built-in
    PowerShell cmdlet, and this tool removes as well as creates.

    Author   : Glenn Armbruster
    Revised  : 2026-08-10
#>
# TOOLKIT-HELP-BEGIN
# Creates and removes shared mailboxes and distribution groups on-premises.
#
# Type the address -- with or without the domain -- pick the object type, pick
# Add or Remove, then Preview and Apply.
#
# ON ADD it needs a directory, because that decides which OU the object lands
# in. Directories that have an OU configured in Tenant Manager use it without
# asking; ones that do not will ask you for the OU.
#
# Shared mailboxes are created as REMOTE mailboxes -- the AD account is created
# on-premises and the mailbox itself lives in the cloud. That is what makes them
# appear after a sync cycle rather than immediately.
#
# ON REMOVE it needs nothing but the address, but it will not act on what it has
# not identified. It looks the object up, shows you what it found, and refuses
# if the object is not the type you selected. Typing a person's address into
# Remove Shared Mailbox would otherwise delete their mailbox.
#
#   Remove   deletes the AD account. The cloud mailbox is soft-deleted and
#            recoverable for a limited window.
#   Disable  keeps the AD account and strips its mail attributes.
# Disable is the reversible one and is offered as a checkbox on Remove.
#
# Owner is optional. On a shared mailbox it grants FullAccess and SendAs; on a
# distribution group it sets ManagedBy, which also lets that person edit the
# membership from Outlook.
#
# Distribution groups reject external senders unless you tick the box.
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

if (-not (Get-Command -Name 'Get-TenantDirectoryOu' -ErrorAction SilentlyContinue)) {
    $null = Show-ToolDialog(
        ("This tool needs a newer Common.ps1 than the one in this folder." + "`r`n`r`n" +
         "Missing: Get-TenantDirectoryOu" + "`r`n`r`n" +
         "Copy Common.ps1 from the same release into" + "`r`n$PSScriptRoot"),
        'Common.ps1 is out of date', 'OK', 'Warning')
    return
}

#region Identity ---------------------------------------------------------------

function Resolve-ObjectAddress {
    <#
        'helpdesk' and 'helpdesk@contoso.com' are the same request.

        The domain comes from Tenant.json rather than being typed, so a copy
        pointed at another tenant builds the right address without anyone
        editing anything.
    #>
    param([string]$Entry)
    $e = "$Entry".Trim()
    if (-not $e) { return '' }
    if ($e -match '@') { return $e }
    $dom = Get-MailDomain
    if (-not $dom) { return '' }
    return "$e@$dom"
}

function Get-ObjectAlias {
    param([string]$Smtp)
    return ("$Smtp" -split '@')[0]
}

function Get-DefaultAccountName {
    <#
        A legal sAMAccountName suggested from the alias.

        WHY THIS EXISTS AT ALL: sAMAccountName IS CAPPED AT 20 CHARACTERS and
        the alias is not. Left to itself Exchange derives the account name from
        Name, so an address whose local part is longer than 20 characters --
        'HoustonFacilitiesRequests' is 25 -- fails the create outright, or gets
        silently mangled into something nobody would guess. That is the failure
        this field was added to remove.

        A SUGGESTION, NOT A DECISION. The box stays editable, because truncating
        at 20 produces collisions ('AccountsPayableTexas' and
        'AccountsPayableTennessee' both truncate to the same 20 characters) and
        only a person can pick which one wins.

        Learn 2026-08-26 (New-RemoteMailbox / New-DistributionGroup,
        -SamAccountName): letters, numbers, spaces, periods and
        ! # $ % ^ & - _ { } ~ ; maximum 20; the last character cannot be a
        period.
    #>
    param([string]$Alias)

    $s = "$Alias".Trim()
    if (-not $s) { return '' }
    $s = ($s -replace '[^A-Za-z0-9 .!#$%^&\-_{}~]', '')
    if ($s.Length -gt 20) { $s = $s.Substring(0, 20) }
    # Trailing periods are rejected by AD, and truncation is how you get one.
    $s = $s.TrimEnd('.')
    return $s
}

function Test-AccountName {
    <#
        Is this a sAMAccountName AD will actually accept?

        Checked here rather than left to the cmdlet, because the error Exchange
        returns for a bad account name names the attribute and not the field the
        operator typed into -- and this tool has two text boxes that could
        plausibly be at fault.
    #>
    param([string]$Name)

    $n = "$Name".Trim()
    if (-not $n)            { return @{ Ok = $false; Msg = 'An account name is required.' } }
    if ($n.Length -gt 20)   { return @{ Ok = $false; Msg = "Account name is $($n.Length) characters. The maximum is 20." } }
    if ($n.EndsWith('.'))   { return @{ Ok = $false; Msg = 'An account name cannot end with a period.' } }
    if ($n -match '[^A-Za-z0-9 .!#$%^&\-_{}~]') {
        return @{ Ok = $false; Msg = "Account name contains a character AD will not accept: '$n'." }
    }
    return @{ Ok = $true; Msg = '' }
}

function Find-AdAccount {
    <#
        Is there already an AD object using this account name, and is it
        mail-enabled?

        USES Get-User RATHER THAN THE AD MODULE. Get-User returns on-premises AD
        users whether or not they have a mailbox, which is exactly the question,
        and it keeps this tool inside the Exchange session it already has rather
        than adding a dependency on RSAT.

        Get-User exists in BOTH Exchange editions (C1), so it goes through
        Invoke-OnPremCommand. An Exchange Online answer here would be about a
        different directory entirely.

        THREE OUTCOMES, AND THEY LEAD TO THREE DIFFERENT ACTIONS:
          not found          create the account   (New-RemoteMailbox)
          found, no mailbox  mail-enable it       (Enable-RemoteMailbox)
          found, mail-enabled  refuse -- it belongs to something else
    #>
    param([string]$SamAccountName)

    $out = [PSCustomObject]@{ Found = $false; MailEnabled = $false; Kind = ''; Display = ''; Upn = '' }
    try {
        $u = Invoke-OnPremCommand -Name 'Get-User' -Parameters @{
            Identity = $SamAccountName; ErrorAction = 'Stop'
        }
        if (-not $u) { return $out }
        $out.Found       = $true
        $out.Kind        = "$($u.RecipientTypeDetails)"
        $out.Display     = "$($u.DisplayName)"
        $out.Upn         = "$($u.UserPrincipalName)"
        # 'User' is the plain AD account with nothing mail-related on it. Every
        # other RecipientTypeDetails means Exchange already owns this object.
        $out.MailEnabled = ($out.Kind -and $out.Kind -ne 'User')
    } catch {
        # Not found is the normal path on a create and is not an error.
        return $out
    }
    return $out
}

function Find-MailObject {
    <#
        What is actually at this address, whatever type it is.

        DELIBERATELY NOT FILTERED BY THE SELECTED TYPE. The caller compares what
        was found against what was asked for, so the tool can say "that is a
        UserMailbox, not a shared mailbox" instead of "not found" -- the second
        is true and would let someone try again harder at deleting a person.
    #>
    param([string]$Smtp)

    $out = [PSCustomObject]@{ Found = $false; Kind = ''; Detail = $null; Display = ''; Extra = '' }
    try {
        $r = Invoke-OnPremCommand -Name 'Get-Recipient' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop'
        }
        if ($r) {
            $out.Found   = $true
            $out.Kind    = "$($r.RecipientTypeDetails)"
            $out.Display = "$($r.DisplayName)"
            $out.Detail  = $r
        }
    } catch { return $out }

    # Member count is the fact that makes a deletion feel real. A group with 40
    # people in it should not be removed on the same shrug as an empty one.
    if ($out.Kind -match '(?i)DistributionGroup|MailUniversal') {
        try {
            $m = @(Invoke-OnPremCommand -Name 'Get-DistributionGroupMember' -Parameters @{
                Identity = $Smtp; ResultSize = 'Unlimited'; ErrorAction = 'Stop'
            })
            $out.Extra = "$($m.Count) member(s)"
        } catch { $out.Extra = 'members could not be read' }
    }
    return $out
}

function Test-KindMatches {
    param([string]$Kind, [string]$Wanted)
    if ($Wanted -eq 'SharedMailbox') {
        return ($Kind -match '(?i)SharedMailbox|RemoteSharedMailbox')
    }
    # Plain distribution groups only -- mail-enabled security groups are a
    # different object and this tool does not create or remove them.
    return ($Kind -match '(?i)^MailUniversalDistributionGroup$')
}

#endregion

#region UI ---------------------------------------------------------------------

function Show-MailObjectForm {
    param([string]$Seed)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Mail Object -- create or remove, on-premises'
    $form.Size = New-Object System.Drawing.Size(880, 700)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblAddr = New-Object System.Windows.Forms.Label
    $lblAddr.Text = 'Address (with or without the domain):'
    $lblAddr.Location = New-Object System.Drawing.Point(12, $y)
    $lblAddr.Size = New-Object System.Drawing.Size(400, 18)
    $form.Controls.Add($lblAddr); $y += 20

    $txtAddr = New-Object System.Windows.Forms.TextBox
    $txtAddr.Location = New-Object System.Drawing.Point(12, $y)
    $txtAddr.Size = New-Object System.Drawing.Size(420, 24)
    $txtAddr.Text = "$Seed"
    $form.Controls.Add($txtAddr)

    $lblResolved = New-Object System.Windows.Forms.Label
    $lblResolved.Location = New-Object System.Drawing.Point(444, ($y + 3))
    $lblResolved.Size = New-Object System.Drawing.Size(400, 18)
    $lblResolved.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblResolved); $y += 34

    # --- object type ---
    $grpType = New-Object System.Windows.Forms.GroupBox
    $grpType.Text = 'Object'
    $grpType.Location = New-Object System.Drawing.Point(12, $y)
    $grpType.Size = New-Object System.Drawing.Size(280, 52)
    $form.Controls.Add($grpType)

    $rbShared = New-Object System.Windows.Forms.RadioButton
    $rbShared.Text = 'Shared mailbox'; $rbShared.Checked = $true
    $rbShared.Location = New-Object System.Drawing.Point(12, 20)
    $rbShared.Size = New-Object System.Drawing.Size(120, 22)
    $grpType.Controls.Add($rbShared)

    $rbDl = New-Object System.Windows.Forms.RadioButton
    $rbDl.Text = 'Distribution group'
    $rbDl.Location = New-Object System.Drawing.Point(140, 20)
    $rbDl.Size = New-Object System.Drawing.Size(130, 22)
    $grpType.Controls.Add($rbDl)

    # --- action ---
    $grpAct = New-Object System.Windows.Forms.GroupBox
    $grpAct.Text = 'Action'
    $grpAct.Location = New-Object System.Drawing.Point(302, $y)
    $grpAct.Size = New-Object System.Drawing.Size(200, 52)
    $form.Controls.Add($grpAct)

    $rbAdd = New-Object System.Windows.Forms.RadioButton
    $rbAdd.Text = 'Add'; $rbAdd.Checked = $true
    $rbAdd.Location = New-Object System.Drawing.Point(12, 20)
    $rbAdd.Size = New-Object System.Drawing.Size(70, 22)
    $grpAct.Controls.Add($rbAdd)

    $rbRemove = New-Object System.Windows.Forms.RadioButton
    $rbRemove.Text = 'Remove'
    $rbRemove.Location = New-Object System.Drawing.Point(90, 20)
    $rbRemove.Size = New-Object System.Drawing.Size(90, 22)
    $grpAct.Controls.Add($rbRemove)
    $y += 60

    # --- add-only panel ---
    $pnlAdd = New-Object System.Windows.Forms.Panel
    $pnlAdd.Location = New-Object System.Drawing.Point(12, $y)
    # 110, not 96: the Account name row was added at y=78 inside this panel.
    $pnlAdd.Size = New-Object System.Drawing.Size(830, 110)
    $form.Controls.Add($pnlAdd)

    $lblDir = New-Object System.Windows.Forms.Label
    $lblDir.Text = 'Directory:'
    $lblDir.Location = New-Object System.Drawing.Point(0, 4)
    $lblDir.Size = New-Object System.Drawing.Size(70, 18)
    $pnlAdd.Controls.Add($lblDir)

    $cboDir = New-Object System.Windows.Forms.ComboBox
    $cboDir.DropDownStyle = 'DropDownList'
    $cboDir.Location = New-Object System.Drawing.Point(72, 0)
    $cboDir.Size = New-Object System.Drawing.Size(160, 24)
    foreach ($d in @(Get-TenantDirectory)) { [void]$cboDir.Items.Add("$($d.Id)") }
    if ($cboDir.Items.Count -gt 0) {
        $def = Get-DefaultTenantDirectory
        $cboDir.SelectedItem = $(if ($def) { "$($def.Id)" } else { $cboDir.Items[0] })
    }
    $pnlAdd.Controls.Add($cboDir)

    $lblOu = New-Object System.Windows.Forms.Label
    $lblOu.Text = 'OU:'
    $lblOu.Location = New-Object System.Drawing.Point(244, 4)
    $lblOu.Size = New-Object System.Drawing.Size(30, 18)
    $pnlAdd.Controls.Add($lblOu)

    $txtOu = New-Object System.Windows.Forms.TextBox
    $txtOu.Location = New-Object System.Drawing.Point(276, 0)
    $txtOu.Size = New-Object System.Drawing.Size(554, 24)
    $txtOu.Font = New-Object System.Drawing.Font('Consolas', 9)
    $pnlAdd.Controls.Add($txtOu)

    $lblOuNote = New-Object System.Windows.Forms.Label
    $lblOuNote.Location = New-Object System.Drawing.Point(0, 26)
    $lblOuNote.Size = New-Object System.Drawing.Size(830, 18)
    $lblOuNote.ForeColor = [System.Drawing.Color]::DimGray
    $pnlAdd.Controls.Add($lblOuNote)

    $lblOwner = New-Object System.Windows.Forms.Label
    $lblOwner.Text = 'Owner (optional):'
    $lblOwner.Location = New-Object System.Drawing.Point(0, 52)
    $lblOwner.Size = New-Object System.Drawing.Size(110, 18)
    $pnlAdd.Controls.Add($lblOwner)

    $txtOwner = New-Object System.Windows.Forms.TextBox
    $txtOwner.Location = New-Object System.Drawing.Point(112, 48)
    $txtOwner.Size = New-Object System.Drawing.Size(300, 24)
    $pnlAdd.Controls.Add($txtOwner)

    $chkExternal = New-Object System.Windows.Forms.CheckBox
    $chkExternal.Text = 'Allow email from external senders'
    $chkExternal.Location = New-Object System.Drawing.Point(430, 50)
    $chkExternal.Size = New-Object System.Drawing.Size(300, 22)
    $pnlAdd.Controls.Add($chkExternal)

    $lblAccount = New-Object System.Windows.Forms.Label
    $lblAccount.Text = 'Account name:'
    $lblAccount.Location = New-Object System.Drawing.Point(0, 82)
    $lblAccount.Size = New-Object System.Drawing.Size(110, 18)
    $pnlAdd.Controls.Add($lblAccount)

    $txtAccount = New-Object System.Windows.Forms.TextBox
    $txtAccount.Location = New-Object System.Drawing.Point(112, 78)
    $txtAccount.Size = New-Object System.Drawing.Size(200, 24)
    $txtAccount.Font = New-Object System.Drawing.Font('Consolas', 9)
    # MaxLength IS THE STOP, not a validator.
    #
    # The box simply refuses the 21st character -- no beep, no message, no
    # red text appearing after the fact. sAMAccountName is capped at 20 by
    # Active Directory itself, so there is no such thing as a valid 21st
    # character to warn about, and a control that accepts input it will later
    # reject is a control that wasted the operator's typing.
    $txtAccount.MaxLength = 20
    $pnlAdd.Controls.Add($txtAccount)

    $lblAccountNote = New-Object System.Windows.Forms.Label
    $lblAccountNote.Location = New-Object System.Drawing.Point(320, 82)
    $lblAccountNote.Size = New-Object System.Drawing.Size(510, 18)
    $lblAccountNote.ForeColor = [System.Drawing.Color]::DimGray
    $lblAccountNote.Text = 'Pre-Windows 2000 name (sAMAccountName). 20 characters maximum -- filled in from the address, edit if you need to.'
    $pnlAdd.Controls.Add($lblAccountNote)

    # --- remove-only panel ---
    $pnlRemove = New-Object System.Windows.Forms.Panel
    $pnlRemove.Location = New-Object System.Drawing.Point(12, $y)
    $pnlRemove.Size = New-Object System.Drawing.Size(830, 96)
    $pnlRemove.Visible = $false
    $form.Controls.Add($pnlRemove)

    $chkDisable = New-Object System.Windows.Forms.CheckBox
    $chkDisable.Text = 'Disable instead of remove  (keeps the AD account, strips mail attributes)'
    $chkDisable.Location = New-Object System.Drawing.Point(0, 4)
    $chkDisable.Size = New-Object System.Drawing.Size(560, 22)
    $pnlRemove.Controls.Add($chkDisable)

    $lblRemoveNote = New-Object System.Windows.Forms.Label
    $lblRemoveNote.Location = New-Object System.Drawing.Point(0, 30)
    $lblRemoveNote.Size = New-Object System.Drawing.Size(820, 50)
    $lblRemoveNote.ForeColor = [System.Drawing.Color]::DimGray
    $lblRemoveNote.Text = ('Remove deletes the AD account; the cloud mailbox is soft-deleted and ' +
                           'recoverable for a limited window. Disable keeps the account.' + "`r`n" +
                           'Nothing is removed until the object has been identified and shown to you.')
    $pnlRemove.Controls.Add($lblRemoveNote)
    # 118, not 104: pnlAdd grew to 110 for the Account name row, and this
    # advance has to clear the TALLER of the two panels -- they occupy the same
    # y and only one is visible at a time.
    $y += 118

    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = 'Preview'
    $btnPreview.Location = New-Object System.Drawing.Point(12, $y)
    $btnPreview.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnPreview)

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = 'Apply'
    $btnApply.Location = New-Object System.Drawing.Point(150, $y)
    $btnApply.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnApply)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Location = New-Object System.Drawing.Point(712, $y)
    $btnClose.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })
    $y += 40

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(12, $y)
    $txtOut.Size = New-Object System.Drawing.Size(830, 280)
    $txtOut.Multiline = $true
    $txtOut.ScrollBars = 'Vertical'
    $txtOut.ReadOnly = $true
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)

    # --- reactive bits -------------------------------------------------------

    $wantedKind = { if ($rbShared.Checked) { 'SharedMailbox' } else { 'DistributionGroup' } }

    $syncOu = {
        # The configured OU fills the box and the box stays editable -- config
        # is a default, not a cage. A directory with nothing configured leaves
        # it empty and says so, which is the prompt.
        $d = "$($cboDir.SelectedItem)"
        $ou = Get-TenantDirectoryOu -DirectoryId $d -Kind (& $wantedKind)
        if ($ou) {
            $txtOu.Text = $ou
            $lblOuNote.Text = "From Tenant.json for '$d'. Edit if this one is different."
        } else {
            $txtOu.Text = ''
            $lblOuNote.Text = "No OU configured -- type the canonical OU, e.g. contoso.com/Accounts/Resource Mailbox"
        }
    }

    $syncMode = {
        $add = $rbAdd.Checked
        $pnlAdd.Visible = $add
        $pnlRemove.Visible = -not $add
        $chkExternal.Enabled = $rbDl.Checked
        if ($add) { & $syncOu }
    }

    $syncAddr = {
        $s = Resolve-ObjectAddress -Entry $txtAddr.Text
        $lblResolved.Text = $(if ($s) { "-> $s" } else { '' })

        # THE SUGGESTION STOPS THE MOMENT THE OPERATOR TYPES THEIR OWN.
        #
        # TextBox.Modified is exactly this flag and is maintained by WinForms:
        # true once a PERSON has edited the box, and reset to false whenever the
        # Text property is assigned in code. So the assignment below both fills
        # the box and leaves it still counting as untouched -- which is what
        # lets the suggestion keep tracking the address until it shouldn't.
        #
        # A hand-rolled flag would have to distinguish a programmatic
        # TextChanged from a typed one, which is the bug this avoids rather than
        # solves.
        if (-not $txtAccount.Modified) {
            $txtAccount.Text = Get-DefaultAccountName -Alias (Get-ObjectAlias -Smtp $s)
        }
    }

    $txtAddr.Add_TextChanged($syncAddr)
    $rbAdd.Add_CheckedChanged($syncMode)
    $rbRemove.Add_CheckedChanged($syncMode)
    $rbShared.Add_CheckedChanged($syncMode)
    $rbDl.Add_CheckedChanged($syncMode)
    $cboDir.Add_SelectedIndexChanged($syncOu)

    # --- gather --------------------------------------------------------------

    $gather = {
        $smtp = Resolve-ObjectAddress -Entry $txtAddr.Text
        if (-not $smtp) {
            return @{ Ok = $false; Msg = 'An address is required (and Tenant.json must have a primary domain).' }
        }
        if ($smtp -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            return @{ Ok = $false; Msg = "'$smtp' does not look like an address." }
        }
        $g = @{
            Ok = $true; Smtp = $smtp; Alias = (Get-ObjectAlias -Smtp $smtp)
            Kind = (& $wantedKind); Add = $rbAdd.Checked
            Owner = "$($txtOwner.Text)".Trim()
            External = $chkExternal.Checked
            DisableOnly = $chkDisable.Checked
            Dir = "$($cboDir.SelectedItem)"
            Ou = "$($txtOu.Text)".Trim()
            Account = "$($txtAccount.Text)".Trim()
        }
        if ($g.Add -and -not $g.Ou) {
            return @{ Ok = $false; Msg = "No OU. Set one for '$($g.Dir)' in Tenant Manager, or type it above." }
        }
        # Only Add uses the account name. Remove works from the address and must
        # not start refusing because a field it never reads is empty.
        if ($g.Add) {
            $chk = Test-AccountName -Name $g.Account
            if (-not $chk.Ok) { return @{ Ok = $false; Msg = $chk.Msg } }
        }
        return $g
    }

    $resolveOwner = {
        param([string]$Owner)
        if (-not $Owner) { return $null }
        $o = Resolve-ObjectAddress -Entry $Owner
        try {
            $r = Invoke-OnPremCommand -Name 'Get-Recipient' -Parameters @{ Identity = $o; ErrorAction = 'Stop' }
            return $r
        } catch { return $null }
    }

    # --- preview -------------------------------------------------------------

    $doPreview = {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $g = & $gather
            if (-not $g.Ok) { $txtOut.Text = $g.Msg; return }

            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('PREVIEW -- nothing has been changed')
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("  Address : $($g.Smtp)")
            [void]$sb.AppendLine("  Object  : $($g.Kind)")
            [void]$sb.AppendLine('')

            $found = Find-MailObject -Smtp $g.Smtp

            if ($g.Add) {
                if ($found.Found) {
                    [void]$sb.AppendLine("  ALREADY EXISTS as $($found.Kind) -- '$($found.Display)'")
                    [void]$sb.AppendLine('  Nothing would be created.')
                    $txtOut.Text = $sb.ToString(); return
                }

                # SAY WHICH OF THE TWO THINGS APPLY WOULD DO.
                #
                # A preview that says "would create" and then mail-enables
                # somebody else's existing account is worse than no preview --
                # the operator approved a different action from the one that
                # ran. So the account-name probe happens here too, and the
                # answer is stated in the words Apply will use.
                if ($g.Kind -eq 'SharedMailbox') {
                    $ad = Find-AdAccount -SamAccountName $g.Account
                    [void]$sb.AppendLine("  Account : $($g.Account)")
                    if (-not $ad.Found) {
                        [void]$sb.AppendLine('            no AD account with this name -- it would be CREATED')
                    } elseif ($ad.MailEnabled) {
                        [void]$sb.AppendLine("            TAKEN by a mail-enabled object: '$($ad.Display)' ($($ad.Kind))")
                        [void]$sb.AppendLine('')
                        [void]$sb.AppendLine('  Apply would REFUSE. Choose a different account name.')
                        $txtOut.Text = $sb.ToString(); return
                    } else {
                        [void]$sb.AppendLine("            EXISTS as a plain AD account -- '$($ad.Display)'")
                        [void]$sb.AppendLine('            it would be MAIL-ENABLED in place, not created')
                        [void]$sb.AppendLine('            the OU below is not used on that path')
                    }
                    [void]$sb.AppendLine('')
                }
                $ou = $null
                try {
                    $ou = @(Invoke-OnPremCommand -Name 'Get-OrganizationalUnit' -Parameters @{
                        Identity = $g.Ou; ErrorAction = 'Stop' })
                } catch { $ou = @() }
                if (@($ou).Count -eq 0) {
                    [void]$sb.AppendLine("  OU NOT FOUND: $($g.Ou)")
                    [void]$sb.AppendLine('  Canonical form, the shape ADUC shows:')
                    [void]$sb.AppendLine('    contoso.com/Accounts/Resource Mailbox')
                    $txtOut.Text = $sb.ToString(); return
                }
                [void]$sb.AppendLine("  Directory : $($g.Dir)")
                [void]$sb.AppendLine("  OU        : $($g.Ou)")
                [void]$sb.AppendLine("  Alias     : $($g.Alias)")
                if ($g.Kind -eq 'SharedMailbox') {
                    [void]$sb.AppendLine('  Creates a REMOTE shared mailbox -- AD account on-premises,')
                    [void]$sb.AppendLine('  mailbox in the cloud after a sync cycle.')
                } else {
                    [void]$sb.AppendLine('  Creates a plain distribution group.')
                    [void]$sb.AppendLine("  External senders: $(if ($g.External) { 'ALLOWED' } else { 'rejected' })")
                }
                if ($g.Owner) {
                    $o = & $resolveOwner $g.Owner
                    if ($o) {
                        $what = $(if ($g.Kind -eq 'SharedMailbox') { 'FullAccess + SendAs' } else { 'ManagedBy' })
                        [void]$sb.AppendLine("  Owner     : $($o.DisplayName) <$($o.PrimarySmtpAddress)>  ($what)")
                    } else {
                        [void]$sb.AppendLine("  Owner     : '$($g.Owner)' DOES NOT RESOLVE -- fix or clear it.")
                    }
                }
            } else {
                if (-not $found.Found) {
                    [void]$sb.AppendLine('  NOT FOUND. Nothing to remove.')
                    $txtOut.Text = $sb.ToString(); return
                }
                [void]$sb.AppendLine("  Found   : $($found.Display)")
                [void]$sb.AppendLine("  Type    : $($found.Kind)")
                if ($found.Extra) { [void]$sb.AppendLine("  Members : $($found.Extra)") }
                [void]$sb.AppendLine('')
                if (-not (Test-KindMatches -Kind $found.Kind -Wanted $g.Kind)) {
                    [void]$sb.AppendLine("  REFUSED -- that is a $($found.Kind), not a $($g.Kind).")
                    [void]$sb.AppendLine('  Apply will not act on it. Change the object type, or check the address.')
                } elseif ($g.Kind -eq 'SharedMailbox') {
                    [void]$sb.AppendLine($(if ($g.DisableOnly) {
                        '  Would DISABLE -- AD account kept, mail attributes stripped.'
                    } else {
                        '  Would REMOVE -- AD account deleted, cloud mailbox soft-deleted.'
                    }))
                } else {
                    [void]$sb.AppendLine('  Would REMOVE the distribution group. Membership is not recoverable.')
                }
            }
            $txtOut.Text = $sb.ToString()
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Preview'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    # --- apply ---------------------------------------------------------------

    $doApply = {
        $g = & $gather
        if (-not $g.Ok) { $txtOut.Text = $g.Msg; return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $found = Find-MailObject -Smtp $g.Smtp

            if ($g.Add) {
                if ($found.Found) {
                    $txtOut.Text = "'$($g.Smtp)' already exists as $($found.Kind). Nothing created."
                    return
                }
                # WHAT IS ALREADY IN AD UNDER THIS ACCOUNT NAME?
                #
                # Find-MailObject above answered "is anything mail-enabled at
                # this ADDRESS". That is a different question from "is this
                # ACCOUNT NAME taken", and the gap between them is where creates
                # fail with a raw AD collision error naming an attribute rather
                # than a field.
                $existing = $null
                if ($g.Kind -eq 'SharedMailbox') {
                    $existing = Find-AdAccount -SamAccountName $g.Account
                    if ($existing.Found -and $existing.MailEnabled) {
                        $txtOut.Text = ("REFUSED." + [Environment]::NewLine + [Environment]::NewLine +
                            "  Account name '$($g.Account)' already belongs to a mail-enabled object:" + [Environment]::NewLine +
                            "    $($existing.Display)  ($($existing.Kind))" + [Environment]::NewLine +
                            "    $($existing.Upn)" + [Environment]::NewLine + [Environment]::NewLine +
                            "  It is not at the address you typed, so this would be a different object." +
                            [Environment]::NewLine + "  Choose another account name.")
                        return
                    }
                }

                # The confirmation says which of the two things is about to
                # happen. 'Create' and 'mail-enable an account somebody else
                # already made' are not the same action and should not share a
                # prompt.
                $reuse = ($existing -and $existing.Found)
                $verbTxt = if ($reuse) { "Mail-enable the EXISTING account '$($g.Account)'" }
                           else        { "Create this $($g.Kind) on-premises" }
                $ouTxt = if ($reuse) { "`r`n`r`nThe account stays where it is -- the OU box is not used when`r`nmail-enabling an existing account." }
                         else        { "`r`n`r`nIn OU:`r`n  $($g.Ou)" }
                $a = Show-ToolDialog(
                    ($verbTxt + '?' + "`r`n`r`n" +
                     "  $($g.Smtp)" + "`r`n  account name: $($g.Account)" + $ouTxt),
                    'Confirm', 'OKCancel', 'Question')
                if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

                if ($g.Kind -eq 'SharedMailbox') {
                    if ($reuse) {
                        # Learn 2026-08-26 (Enable-RemoteMailbox): -Identity is
                        # mandatory and positional; -Shared, -Alias and
                        # -PrimarySmtpAddress are all valid in the Shared set.
                        #
                        # THERE IS NO -OnPremisesOrganizationalUnit ON THIS
                        # CMDLET. The account already exists somewhere, and
                        # mail-enabling it does not move it. Passing the OU here
                        # would be a parameter that does not exist; silently
                        # ignoring the operator's OU box would be worse, so the
                        # confirmation above says so out loud.
                        Invoke-OnPremCommand -Name 'Enable-RemoteMailbox' -Parameters @{
                            Identity           = $g.Account
                            Alias              = $g.Alias
                            PrimarySmtpAddress = $g.Smtp
                            Shared             = $true
                            ErrorAction        = 'Stop'
                        } | Out-Null
                    } else {
                        # Learn 2026-08-26 (New-RemoteMailbox): -SamAccountName
                        # is valid in every parameter set, maximum 20 chars. In
                        # the Shared set -Password and -UserPrincipalName are
                        # NOT mandatory, which is why neither is passed here.
                        Invoke-OnPremCommand -Name 'New-RemoteMailbox' -Parameters @{
                            Name                          = $g.Alias
                            Alias                         = $g.Alias
                            SamAccountName                = $g.Account
                            PrimarySmtpAddress            = $g.Smtp
                            Shared                        = $true
                            OnPremisesOrganizationalUnit  = $g.Ou
                            ErrorAction                   = 'Stop'
                        } | Out-Null
                    }
                } else {
                    # Learn 2026-08-26 (New-DistributionGroup): -SamAccountName
                    # is on-premises only, maximum 20 characters, same character
                    # set as the user cmdlet. A group name over 20 characters
                    # fails here for exactly the same reason a mailbox one does.
                    Invoke-OnPremCommand -Name 'New-DistributionGroup' -Parameters @{
                        Name               = $g.Alias
                        Alias              = $g.Alias
                        SamAccountName     = $g.Account
                        PrimarySmtpAddress = $g.Smtp
                        Type               = 'Distribution'
                        OrganizationalUnit = $g.Ou
                        ErrorAction        = 'Stop'
                    } | Out-Null

                    # Unticked means REJECT external senders, so the parameter is
                    # the inverse of the checkbox. Set explicitly either way
                    # rather than relying on whatever the default happens to be.
                    Invoke-OnPremCommand -Name 'Set-DistributionGroup' -Parameters @{
                        Identity                            = $g.Smtp
                        RequireSenderAuthenticationEnabled  = (-not $g.External)
                        ErrorAction                         = 'Stop'
                    } | Out-Null
                }

                $ownerNote = ''
                if ($g.Owner) {
                    $o = & $resolveOwner $g.Owner
                    if (-not $o) {
                        $ownerNote = "  Owner '$($g.Owner)' did not resolve -- created without an owner."
                    } elseif ($g.Kind -eq 'SharedMailbox') {
                        Invoke-OnPremCommand -Name 'Add-MailboxPermission' -Parameters @{
                            Identity = $g.Smtp; User = "$($o.PrimarySmtpAddress)"
                            AccessRights = 'FullAccess'; AutoMapping = $false
                            Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        Invoke-OnPremCommand -Name 'Add-ADPermission' -Parameters @{
                            Identity = $g.Smtp; User = "$($o.PrimarySmtpAddress)"
                            ExtendedRights = 'Send As'; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        $ownerNote = "  Owner: $($o.DisplayName) -- FullAccess + SendAs."
                    } else {
                        Invoke-OnPremCommand -Name 'Set-DistributionGroup' -Parameters @{
                            Identity = $g.Smtp; ManagedBy = "$($o.PrimarySmtpAddress)"
                            BypassSecurityGroupManagerCheck = $true; ErrorAction = 'Stop' } | Out-Null
                        $ownerNote = "  Owner: $($o.DisplayName) -- ManagedBy."
                    }
                }

                # The audit distinguishes the two paths. 'Create' on a row where
                # the account already existed would misdescribe what happened,
                # and the account name is recorded because it is the only handle
                # that finds the object in AD if the address is later changed.
                Write-CompletedLog `
                    -Action $(if ($reuse) { 'MailEnable' } else { 'Create' }) `
                    -Target $g.Smtp -Result 'Success' -Environment $g.Dir `
                    -Detail $(
                        if ($reuse) {
                            "Mail-enabled existing on-premises account '$($g.Account)' as $($g.Kind)$(if ($g.Owner) { "; owner $($g.Owner)" })"
                        } else {
                            "Created on-premises $($g.Kind) in $($g.Ou); account name '$($g.Account)'$(if ($g.Owner) { "; owner $($g.Owner)" })"
                        })

                $txtOut.Text = ($(if ($reuse) { 'MAIL-ENABLED (existing account)' } else { 'CREATED' }) +
                    [Environment]::NewLine + [Environment]::NewLine +
                    "  $($g.Smtp)  ($($g.Kind))" + [Environment]::NewLine +
                    "  account name: $($g.Account)" + [Environment]::NewLine +
                    $(if ($reuse) { "  OU: unchanged -- the account already existed" }
                      else        { "  OU: $($g.Ou)" }) + [Environment]::NewLine +
                    $(if ($ownerNote) { $ownerNote + [Environment]::NewLine } else { '' }) +
                    [Environment]::NewLine + "  Allow a sync cycle before it appears in the cloud.")
                return
            }

            # ---- remove ----
            if (-not $found.Found) { $txtOut.Text = "'$($g.Smtp)' not found. Nothing removed."; return }

            if (-not (Test-KindMatches -Kind $found.Kind -Wanted $g.Kind)) {
                # The guard that matters. Without it, a mistyped address plus the
                # wrong radio button deletes a person's mailbox.
                $txtOut.Text = ("REFUSED. '$($g.Smtp)' is a $($found.Kind), not a $($g.Kind)." +
                    [Environment]::NewLine + 'Nothing has been removed.')
                return
            }

            $verb = $(if ($g.Kind -eq 'SharedMailbox' -and $g.DisableOnly) { 'Disable' } else { 'Remove' })
            $a = Show-ToolDialog(
                ("$verb this $($found.Kind)?" + "`r`n`r`n" +
                 "  $($found.Display)" + "`r`n  $($g.Smtp)" +
                 $(if ($found.Extra) { "`r`n  $($found.Extra)" } else { '' }) + "`r`n`r`n" +
                 $(if ($verb -eq 'Disable') {
                     'The AD account is kept and its mail attributes are stripped.'
                   } elseif ($g.Kind -eq 'SharedMailbox') {
                     'The AD account is deleted. The cloud mailbox is soft-deleted and recoverable for a limited window.'
                   } else {
                     'The group and its membership are deleted. Membership is not recoverable.'
                   })),
                "Confirm $verb", 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            if ($g.Kind -eq 'SharedMailbox') {
                $cmd = $(if ($g.DisableOnly) { 'Disable-RemoteMailbox' } else { 'Remove-RemoteMailbox' })
                Invoke-OnPremCommand -Name $cmd -Parameters @{
                    Identity = $g.Smtp; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
            } else {
                Invoke-OnPremCommand -Name 'Remove-DistributionGroup' -Parameters @{
                    Identity = $g.Smtp; BypassSecurityGroupManagerCheck = $true
                    Confirm = $false; ErrorAction = 'Stop' } | Out-Null
            }

            Write-CompletedLog -Action $(if ($verb -eq 'Disable') { 'Disable' } else { 'Delete' }) `
                -Target $g.Smtp -Result 'Success' `
                -Detail "$verb on-premises $($found.Kind) '$($found.Display)'$(if ($found.Extra) { " ($($found.Extra))" })"

            $txtOut.Text = ("$($verb.ToUpper())D" + [Environment]::NewLine + [Environment]::NewLine +
                "  $($found.Display)" + [Environment]::NewLine + "  $($g.Smtp)" + [Environment]::NewLine +
                [Environment]::NewLine + '  Allow a sync cycle for the cloud to catch up.')
        } catch {
            Write-FailedLog -Action $(if ($g.Add) { 'Create' } else { 'Delete' }) -Target $g.Smtp -ErrorRecord $_ `
                -Detail "$(if ($g.Add) { 'Create' } else { 'Remove' }) on-premises $($g.Kind)"
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "$(if ($g.Add) { 'Creating' } else { 'Removing' }) the object"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnPreview.Add_Click($doPreview)
    $btnApply.Add_Click($doApply)

    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Mail Object' | Out-Null
    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    & $syncMode
    & $syncAddr
    $form.Add_Shown({ $txtAddr.Focus() })
    [void]$form.ShowDialog()
}

#endregion

Show-MailObjectForm -Seed $SmtpAddress
