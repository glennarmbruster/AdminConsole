#Requires -Version 5.1
<#
    dl_send_to-as.ps1
    Send-As and Send-To on a distribution list, for one person at a time.

    THE TWO PERMISSIONS ARE NOT ALIKE, and mixing them up is the reason this
    tool shows both at once.

    SEND-AS is a cloud ACL on the group object: Add-RecipientPermission with
    -AccessRights SendAs. Mail sent that way appears to come FROM the group.

    SEND-TO is not really a permission at all. It is the group's delivery
    restriction, AcceptMessagesOnlyFromSendersOrMembers, and it has two
    behaviours that catch people:

    1. AN EMPTY LIST MEANS EVERYONE CAN SEND. It does not mean the restriction
       is broken or unset. So adding the first entry does not GRANT anything --
       it takes a group the whole company could mail and restricts it to one
       person. That is the single most destructive thing this window can do and
       it is confirmed separately, in words, before it happens.

    2. THE LIST IS INERT UNLESS RequireSenderAuthenticationEnabled IS ON.
       With authentication off the entries are stored and ignored and anybody
       can mail the group -- so a carefully built list can be completely
       useless while looking perfect in the portal. The state of that flag is
       shown next to the list rather than left to be discovered.

    THE LIST IS REPLACED WHOLE by Set-DistributionGroup, never appended to.
    Every write here reads the current value, changes one entry and writes the
    lot back. Passing only the new name does not error -- it silently removes
    everybody else. Same trap as ManagedBy in GroupEditor.ps1.
#>
# TOOLKIT-HELP-RUNS: Reads and WRITES Send-As and delivery restrictions on a distribution list in Exchange Online.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Give one person, or take away from them, the right to send AS a distribution
# list or to send TO one.
#
# TYPE THE LIST, TYPE THE PERSON, PICK ONE ACTION, PRESS APPLY. Either address
# box accepts a full address or just the first part -- the tenant domain is
# added for you.
#
# SEND-AS means mail appears to come from the list. SEND-TO means being allowed
# to mail the list at all, which only matters when the list is restricted.
#
# THE ONE THING TO UNDERSTAND BEFORE USING "ADD SEND-TO":
#
# A list with an EMPTY allowed-senders list accepts mail from everyone. Adding
# the first person to that list does not grant them anything -- it locks
# everybody else OUT. The window checks for this and asks again, in plain
# words, because the operator's intention ("let Bob mail this list") and the
# effect ("only Bob may mail this list") are opposites.
#
# Removing the last entry does the reverse: the list goes back to accepting
# mail from anyone. That is also confirmed.
#
# AND THE OTHER THING: the allowed-senders list does nothing at all unless
# "Require sender authentication" is on. The window shows that flag beside the
# list. If it says OFF, adding a sender changes nothing anybody will notice.
#
# A LIST SYNCED FROM ON-PREMISES cannot have its delivery restriction changed
# here -- it is mastered in Active Directory and the next sync would put it
# back. The window refuses and says so. Send-As is a cloud ACL and is allowed.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

#region Load Common -------------------------------------------------------------
$root = $PSScriptRoot
while ($root -and -not (Test-Path -LiteralPath (Join-Path $root 'Common.ps1'))) {
    $parent = Split-Path -Parent $root
    if (-not $parent -or $parent -eq $root) { $root = ''; break }
    $root = $parent
}
if (-not $root) { throw "Common.ps1 not found above '$PSScriptRoot'." }
. (Join-Path $root 'Common.ps1')
#endregion

if (-not (Test-ConsoleLaunch)) { return }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:Dl   = $null   # the resolved distribution list
$script:User = $null   # the resolved person
$script:SendTo = @()   # AcceptMessagesOnlyFromSendersOrMembers as it stands
$script:AuthReq = $false

#region Resolve -----------------------------------------------------------------
function Resolve-TheList {
    <#
        The group, and whether we may write to it.

        Get-Recipient FIRST, because the type decides everything after it.
        Opening with Get-DistributionGroup means a Microsoft 365 Group answers
        with "The current operation is not supported on GroupMailbox", which
        reads like a broken tool rather than a category mistake.

        THE GUID IS KEPT, NOT THE NAME. C4: a group's Identity renders as its
        AD common name, and passing that back is a fresh forest-wide lookup
        that can land on a different object.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $r = [PSCustomObject]@{
        Ok = $false; Reason = ''
        Guid = ''; Smtp = ''; Display = ''; Rtd = ''; IsDirSynced = $false
    }

    $rec = $null
    try {
        $rec = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $Text; ErrorAction = 'Stop' }
    } catch { $rec = $null }
    if (-not $rec) { $r.Reason = "Nothing in the directory matched '$Text'."; return $r }

    $rtd = "$($rec.RecipientTypeDetails)"
    $r.Rtd         = $rtd
    $r.Guid        = "$($rec.Guid)"
    $r.Smtp        = "$($rec.PrimarySmtpAddress)"
    $r.Display     = "$($rec.DisplayName)"
    $r.IsDirSynced = [bool]$rec.IsDirSynced

    # C12: EXACT matches only. 'MailUniversalDistributionGroup' and
    # 'MailUniversalSecurityGroup' are the two this tool handles; a dynamic
    # list has no stored membership and a Microsoft 365 Group is a different
    # object with different cmdlets.
    switch -Exact ($rtd) {
        'MailUniversalDistributionGroup' { $r.Ok = $true; return $r }
        'MailUniversalSecurityGroup'     { $r.Ok = $true; return $r }
        'MailNonUniversalGroup'          { $r.Ok = $true; return $r }
        'DynamicDistributionGroup' {
            $r.Reason = "'$($rec.DisplayName)' is a DYNAMIC distribution list. Send-As can be set on it, but its membership is a filter evaluated at send time -- and this tool does not handle the different cmdlets that go with it."
            return $r
        }
        'GroupMailbox' {
            $r.Reason = "'$($rec.DisplayName)' is a Microsoft 365 Group, not a distribution list. Use the M365 Group tool -- it takes different cmdlets and this one would report failures that are really category mistakes."
            return $r
        }
    }
    $r.Reason = "'$($rec.DisplayName)' is a $rtd, which is not a distribution list."
    return $r
}

function Resolve-ThePerson {
    <#
        The trustee.

        RESOLVED BEFORE ANY WRITE, because both cmdlets accept an address that
        does not exist and fail in ways that read like a permissions problem.
        Knowing the display name also lets every message and audit row name a
        person rather than a string somebody typed.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Smtp = ''; Display = ''; Rtd = '' }
    $rec = $null
    try {
        $rec = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $Text; ErrorAction = 'Stop' }
    } catch { $rec = $null }
    if (-not $rec) { $r.Reason = "Nothing in the directory matched '$Text'."; return $r }

    $r.Smtp    = "$($rec.PrimarySmtpAddress)"
    $r.Display = "$($rec.DisplayName)"
    $r.Rtd     = "$($rec.RecipientTypeDetails)"
    if (-not $r.Smtp) {
        $r.Reason = "'$($rec.DisplayName)' has no email address, so it cannot be given either of these rights."
        return $r
    }
    $r.Ok = $true
    return $r
}
#endregion

#region Read --------------------------------------------------------------------
function Get-DlSendAs {
    <# Who can already send AS this list. #>
    $out = [PSCustomObject]@{ Ok = $false; Reason = ''; Rows = @() }
    try {
        $rows = @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{
            Identity = $script:Dl.Guid; ErrorAction = 'Stop' })
        # NT AUTHORITY\SELF is on every object and is not a delegation anybody
        # granted, so listing it would pad the answer with noise.
        $out.Rows = @($rows |
            Where-Object { "$($_.AccessRights)" -match 'SendAs' -and "$($_.Trustee)" -ne 'NT AUTHORITY\SELF' } |
            ForEach-Object { "$($_.Trustee)" })
        $out.Ok = $true
    } catch {
        $out.Reason = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission'
    }
    return $out
}

function Get-DlSendTo {
    <#
        The allowed-sender list, and the flag that decides whether it means
        anything. Read together because they are read together to be understood.
    #>
    $out = [PSCustomObject]@{ Ok = $false; Reason = ''; Rows = @(); AuthRequired = $false }
    try {
        $g = Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
            Identity = $script:Dl.Guid; ErrorAction = 'Stop' }
        $out.Rows         = @($g.AcceptMessagesOnlyFromSendersOrMembers | ForEach-Object { "$_" } | Where-Object { "$_".Trim() })
        $out.AuthRequired = [bool]$g.RequireSenderAuthenticationEnabled
        $out.Ok = $true
    } catch {
        $out.Reason = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroup'
    }
    return $out
}

function Test-IsAllowedSender {
    <#
        Is this person already in the list?

        MATCHED LOOSELY ON PURPOSE. Exchange returns these entries as display
        names or canonical AD paths depending on the object -- almost never as
        the address the operator typed -- so an exact test would answer "no"
        for entries that are plainly there.
    #>
    param([Parameter(Mandatory)][string]$Smtp, [Parameter(Mandatory)][string]$Display)

    $needles = @("$Smtp".ToLowerInvariant(), "$Display".ToLowerInvariant()) | Where-Object { $_ }
    foreach ($e in @($script:SendTo)) {
        $v    = "$e".ToLowerInvariant()
        $leaf = ("$e" -split '/')[-1].ToLowerInvariant()
        if ($needles -contains $v -or $needles -contains $leaf) { return $true }
    }
    return $false
}
#endregion

#region Write -------------------------------------------------------------------
function Test-CanWriteSendTo {
    <#
        C18, ANSWERED BEFORE THE WRITE.

        AcceptMessagesOnlyFromSendersOrMembers on a directory-synced list is
        mastered in on-premises AD. Exchange Online will either refuse the
        change or accept it and have the next sync overwrite it -- and the
        second case is worse, because the tool reported success, the audit row
        says success, and the restriction is gone by morning.
    #>
    $r = [PSCustomObject]@{ Allowed = $false; Reason = '' }
    if (-not $script:Dl) { $r.Reason = 'No list is loaded.'; return $r }
    if ($script:Dl.IsDirSynced) {
        $r.Reason = "'$($script:Dl.Display)' is SYNCED FROM ON-PREMISES (IsDirSynced is true). Its delivery restriction is mastered in Active Directory, so this change has to be made there -- Exchange Online will either refuse it or accept it and have the next directory sync undo it. Use the on-premises tools, or Group Manager, which opens the on-premises editor in its own process." +
            [Environment]::NewLine + [Environment]::NewLine +
            'Send-As is different: it is a cloud ACL and can be set from here.'
        return $r
    }
    $r.Allowed = $true
    return $r
}

function Set-DlSendAs {
    param([Parameter(Mandatory)][ValidateSet('Add', 'Remove')][string]$Mode)

    $name = if ($Mode -eq 'Add') { 'Add-RecipientPermission' } else { 'Remove-RecipientPermission' }
    [void](Invoke-EXOCommand -Name $name -Parameters @{
        Identity     = $script:Dl.Guid
        Trustee      = $script:User.Smtp
        AccessRights = 'SendAs'
        Confirm      = $false
        ErrorAction  = 'Stop'
    })
}

function Set-DlSendTo {
    <#
        Add or remove ONE allowed sender by rebuilding the whole list.

        THE VALUE IS REPLACED, NOT APPENDED TO. Writing just the new entry does
        not error; it silently removes everybody else. So the current list is
        read, one entry changes, and the whole thing goes back.

        BypassSecurityGroupManagerCheck because an administrator who is not
        listed in the group's ManagedBy is otherwise refused -- with an error
        about ownership that sounds like a bug in the tool.
    #>
    param([Parameter(Mandatory)][ValidateSet('Add', 'Remove')][string]$Mode)

    $list = New-Object System.Collections.ArrayList
    foreach ($e in @($script:SendTo)) { [void]$list.Add("$e") }

    if ($Mode -eq 'Add') {
        [void]$list.Add($script:User.Smtp)
    } else {
        $needles = @("$($script:User.Smtp)".ToLowerInvariant(), "$($script:User.Display)".ToLowerInvariant()) | Where-Object { $_ }
        $keep = @($list | Where-Object {
            $v    = "$_".ToLowerInvariant()
            $leaf = ("$_" -split '/')[-1].ToLowerInvariant()
            -not ($needles -contains $v -or $needles -contains $leaf)
        })
        $list = New-Object System.Collections.ArrayList
        foreach ($e in $keep) { [void]$list.Add("$e") }
    }

    [void](Invoke-EXOCommand -Name 'Set-DistributionGroup' -Parameters @{
        Identity = $script:Dl.Guid
        AcceptMessagesOnlyFromSendersOrMembers = @($list.ToArray())
        BypassSecurityGroupManagerCheck = $true
        Confirm     = $false
        ErrorAction = 'Stop'
    })
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Distribution list -- Send-As and Send-To'
$form.ClientSize      = New-Object System.Drawing.Size(900, 704)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$lblHead          = New-Object System.Windows.Forms.Label
$lblHead.Text     = 'Send-As and Send-To on a distribution list'
$lblHead.Location = New-Object System.Drawing.Point(14, 12)
$lblHead.Size     = New-Object System.Drawing.Size(600, 22)
$lblHead.Font     = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($lblHead)

$grpWho          = New-Object System.Windows.Forms.GroupBox
$grpWho.Text     = 'Which list, and who'
$grpWho.Location = New-Object System.Drawing.Point(14, 40)
$grpWho.Size     = New-Object System.Drawing.Size(872, 122)
[void]$form.Controls.Add($grpWho)

$lblDl          = New-Object System.Windows.Forms.Label
$lblDl.Text     = 'Distribution list'
$lblDl.Location = New-Object System.Drawing.Point(16, 26)
$lblDl.Size     = New-Object System.Drawing.Size(100, 20)
[void]$grpWho.Controls.Add($lblDl)

$txtDl          = New-Object System.Windows.Forms.TextBox
$txtDl.Location = New-Object System.Drawing.Point(122, 23)
$txtDl.Size     = New-Object System.Drawing.Size(290, 22)
$txtDl.Text     = "$(Get-TargetSmtp -Text $SmtpAddress)"
[void]$grpWho.Controls.Add($txtDl)

$btnLook           = New-Object System.Windows.Forms.Button
$btnLook.Text      = 'Look up'
$btnLook.Location  = New-Object System.Drawing.Point(420, 22)
$btnLook.Size      = New-Object System.Drawing.Size(96, 26)
$btnLook.FlatStyle = 'Flat'
[void]$grpWho.Controls.Add($btnLook)

$lblDlWho          = New-Object System.Windows.Forms.Label
$lblDlWho.Location = New-Object System.Drawing.Point(524, 26)
$lblDlWho.Size     = New-Object System.Drawing.Size(332, 20)
[void]$grpWho.Controls.Add($lblDlWho)

$lblUser          = New-Object System.Windows.Forms.Label
$lblUser.Text     = 'User'
$lblUser.Location = New-Object System.Drawing.Point(16, 58)
$lblUser.Size     = New-Object System.Drawing.Size(100, 20)
[void]$grpWho.Controls.Add($lblUser)

$txtUser          = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(122, 55)
$txtUser.Size     = New-Object System.Drawing.Size(290, 22)
[void]$grpWho.Controls.Add($txtUser)

$lblUserWho          = New-Object System.Windows.Forms.Label
$lblUserWho.Location = New-Object System.Drawing.Point(524, 58)
$lblUserWho.Size     = New-Object System.Drawing.Size(332, 20)
[void]$grpWho.Controls.Add($lblUserWho)

$lblAddrHint          = New-Object System.Windows.Forms.Label
$lblAddrHint.Text     = 'Both boxes take a full address or just the first part -- the tenant domain is added for you. Look up reads the list and shows what is set on it now; nothing is changed until Apply.'
$lblAddrHint.Location = New-Object System.Drawing.Point(122, 84)
$lblAddrHint.Size     = New-Object System.Drawing.Size(734, 32)
[void]$grpWho.Controls.Add($lblAddrHint)

# RADIO BUTTONS, NOT TICK BOXES -- and that is a change from what was asked for.
#
# The six actions are mutually exclusive: "Add both" already is Send-As plus
# Send-To, so tick boxes would allow "Add Send-As" and "Remove both" together
# and the tool would have to guess. A radio group makes the impossible
# combinations unrepresentable rather than validated afterwards.
$grpWhat          = New-Object System.Windows.Forms.GroupBox
$grpWhat.Text     = 'What to do -- one of these'
$grpWhat.Location = New-Object System.Drawing.Point(14, 170)
$grpWhat.Size     = New-Object System.Drawing.Size(872, 112)
[void]$form.Controls.Add($grpWhat)

$rbAddAs          = New-Object System.Windows.Forms.RadioButton
$rbAddAs.Text     = 'Add Send-As'
$rbAddAs.Location = New-Object System.Drawing.Point(16, 24)
$rbAddAs.Size     = New-Object System.Drawing.Size(180, 22)
$rbAddAs.Checked  = $true
[void]$grpWhat.Controls.Add($rbAddAs)

$rbAddTo          = New-Object System.Windows.Forms.RadioButton
$rbAddTo.Text     = 'Add Send-To'
$rbAddTo.Location = New-Object System.Drawing.Point(16, 50)
$rbAddTo.Size     = New-Object System.Drawing.Size(180, 22)
[void]$grpWhat.Controls.Add($rbAddTo)

$rbAddBoth          = New-Object System.Windows.Forms.RadioButton
$rbAddBoth.Text     = 'Add both'
$rbAddBoth.Location = New-Object System.Drawing.Point(16, 76)
$rbAddBoth.Size     = New-Object System.Drawing.Size(180, 22)
[void]$grpWhat.Controls.Add($rbAddBoth)

$rbRemAs          = New-Object System.Windows.Forms.RadioButton
$rbRemAs.Text     = 'Remove Send-As'
$rbRemAs.Location = New-Object System.Drawing.Point(210, 24)
$rbRemAs.Size     = New-Object System.Drawing.Size(180, 22)
[void]$grpWhat.Controls.Add($rbRemAs)

$rbRemTo          = New-Object System.Windows.Forms.RadioButton
$rbRemTo.Text     = 'Remove Send-To'
$rbRemTo.Location = New-Object System.Drawing.Point(210, 50)
$rbRemTo.Size     = New-Object System.Drawing.Size(180, 22)
[void]$grpWhat.Controls.Add($rbRemTo)

$rbRemBoth          = New-Object System.Windows.Forms.RadioButton
$rbRemBoth.Text     = 'Remove both'
$rbRemBoth.Location = New-Object System.Drawing.Point(210, 76)
$rbRemBoth.Size     = New-Object System.Drawing.Size(180, 22)
[void]$grpWhat.Controls.Add($rbRemBoth)

$lblWhatHint          = New-Object System.Windows.Forms.Label
$lblWhatHint.Text     = "SEND-AS  mail appears to come from the list." + [Environment]::NewLine +
                        "SEND-TO  allowed to mail the list at all -- only meaningful when the list is restricted." + [Environment]::NewLine + [Environment]::NewLine +
                        'Adding the FIRST allowed sender restricts the list to that person alone. The window asks again before doing it.'
$lblWhatHint.Location = New-Object System.Drawing.Point(404, 24)
$lblWhatHint.Size     = New-Object System.Drawing.Size(452, 76)
[void]$grpWhat.Controls.Add($lblWhatHint)

$btnApply           = New-Object System.Windows.Forms.Button
$btnApply.Text      = 'Apply'
$btnApply.Location  = New-Object System.Drawing.Point(14, 292)
$btnApply.Size      = New-Object System.Drawing.Size(180, 30)
$btnApply.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnApply)

$btnRefresh           = New-Object System.Windows.Forms.Button
$btnRefresh.Text      = 'Refresh'
$btnRefresh.Location  = New-Object System.Drawing.Point(204, 292)
$btnRefresh.Size      = New-Object System.Drawing.Size(140, 30)
$btnRefresh.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnRefresh)

$lblAuth          = New-Object System.Windows.Forms.Label
$lblAuth.Location = New-Object System.Drawing.Point(356, 298)
$lblAuth.Size     = New-Object System.Drawing.Size(530, 20)
$lblAuth.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($lblAuth)

$lblCurCap          = New-Object System.Windows.Forms.Label
$lblCurCap.Text     = 'What is set on the list now'
$lblCurCap.Location = New-Object System.Drawing.Point(14, 330)
$lblCurCap.Size     = New-Object System.Drawing.Size(300, 18)
$lblCurCap.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($lblCurCap)

$txtCurrent            = New-Object System.Windows.Forms.TextBox
$txtCurrent.Location   = New-Object System.Drawing.Point(14, 350)
$txtCurrent.Size       = New-Object System.Drawing.Size(872, 146)
$txtCurrent.Multiline  = $true
$txtCurrent.ReadOnly   = $true
$txtCurrent.ScrollBars = 'Vertical'
$txtCurrent.WordWrap   = $true
$txtCurrent.Font       = New-Object System.Drawing.Font('Consolas', 9)
[void]$form.Controls.Add($txtCurrent)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(14, 506)
$txtOut.Size       = New-Object System.Drawing.Size(872, 184)
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
# B6.
$txtOut.WordWrap   = $true
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
[void]$form.Controls.Add($txtOut)
#endregion

#region View --------------------------------------------------------------------
function Show-Out {
    param([string]$Text, [switch]$Keep)
    if ($Keep) { $txtOut.Text = "$Text$([Environment]::NewLine)$([Environment]::NewLine)$($txtOut.Text)" }
    else { $txtOut.Text = "$Text" }
}

function Get-ChosenAction {
    if ($rbAddTo.Checked)   { return 'AddTo' }
    if ($rbAddBoth.Checked) { return 'AddBoth' }
    if ($rbRemAs.Checked)   { return 'RemoveAs' }
    if ($rbRemTo.Checked)   { return 'RemoveTo' }
    if ($rbRemBoth.Checked) { return 'RemoveBoth' }
    return 'AddAs'
}

function Update-ListView {
    <# Reads the list and paints everything known about it. #>
    $txtCurrent.Text = ''
    $lblAuth.Text    = ''
    $script:SendTo   = @()

    $typed = "$($txtDl.Text)".Trim()
    if (-not $typed) { Show-Out 'Type a distribution list address, or just the first part of one.'; return }

    $addr = Get-TargetSmtp -Text $typed
    if ("$addr" -ne $typed) { $txtDl.Text = "$addr" }

    $old = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $g = Resolve-TheList -Text "$addr"
        $script:Dl = $null
        if (-not $g.Ok) {
            $lblDlWho.Text = 'Not a distribution list'
            Show-Out $g.Reason
            return
        }
        $script:Dl = $g
        $lblDlWho.Text = "$($g.Display)  [$($g.Rtd)]"
        $form.Text = "Send-As and Send-To: $($g.Smtp)"

        $sa = Get-DlSendAs
        $st = Get-DlSendTo
        if ($st.Ok) { $script:SendTo = @($st.Rows); $script:AuthReq = $st.AuthRequired }

        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("$($g.Display)   <$($g.Smtp)>")
        [void]$L.Add("  $($g.Rtd)   GUID $($g.Guid)")
        if ($g.IsDirSynced) {
            [void]$L.Add('  SYNCED FROM ON-PREMISES -- the allowed-sender list cannot be changed here.')
        }
        [void]$L.Add('')

        [void]$L.Add("SEND-AS  ($(@($sa.Rows).Count))")
        if (-not $sa.Ok) { [void]$L.Add("  could not be read: $($sa.Reason)") }
        elseif (@($sa.Rows).Count -eq 0) { [void]$L.Add('  nobody') }
        else { foreach ($t in $sa.Rows) { [void]$L.Add("  $t") } }
        [void]$L.Add('')

        [void]$L.Add("SEND-TO -- allowed senders  ($(@($script:SendTo).Count))")
        if (-not $st.Ok) { [void]$L.Add("  could not be read: $($st.Reason)") }
        elseif (@($script:SendTo).Count -eq 0) {
            [void]$L.Add('  EMPTY -- which means ANYONE may mail this list. It is not unset or broken.')
        } else {
            foreach ($e in $script:SendTo) { [void]$L.Add("  $e") }
        }
        $txtCurrent.Text = ($L -join [Environment]::NewLine)

        # THE FLAG THAT DECIDES WHETHER THE LIST ABOVE MEANS ANYTHING.
        if ($st.Ok -and -not $st.AuthRequired) {
            $lblAuth.Text = 'Require sender authentication: OFF -- the allowed-sender list is IGNORED.'
            $lblAuth.ForeColor = [System.Drawing.Color]::FromArgb(230, 170, 60)
        } elseif ($st.Ok) {
            $lblAuth.Text = 'Require sender authentication: on -- the allowed-sender list is enforced.'
            $lblAuth.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
        }
        # B3: a semantic colour has to survive theming.
        $lblAuth.Tag = 'SemanticColour'

        Show-Out "Loaded $($g.Smtp). Nothing has been changed."
    } catch {
        Show-Out (Format-ErrorDetail -ErrorRecord $_ -Context "Looking up '$addr'")
    } finally {
        $form.Cursor = $old
    }
}
#endregion

#region Apply -------------------------------------------------------------------
function Confirm-Sentence {
    param([Parameter(Mandatory)][string]$Sentence, [switch]$DefaultCancel)
    $btn = [System.Windows.Forms.MessageBoxDefaultButton]::Button1
    if ($DefaultCancel) { $btn = [System.Windows.Forms.MessageBoxDefaultButton]::Button2 }
    $r = Show-ToolDialog($Sentence, 'Confirm',
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Warning, $btn)
    return ($r -eq [System.Windows.Forms.DialogResult]::OK)
}

function Invoke-Apply {
    if (-not $script:Dl) { Show-Out 'Look a distribution list up first.'; return }

    $typedUser = "$($txtUser.Text)".Trim()
    if (-not $typedUser) { Show-Out 'Type the user to add or remove.'; return }
    $userAddr = Get-TargetSmtp -Text $typedUser
    if ("$userAddr" -ne $typedUser) { $txtUser.Text = "$userAddr" }

    $u = Resolve-ThePerson -Text "$userAddr"
    if (-not $u.Ok) { Show-Out $u.Reason; return }
    $script:User = $u
    $lblUserWho.Text = "$($u.Display)  [$($u.Rtd)]"

    $action  = Get-ChosenAction
    $doAs    = $action -in @('AddAs', 'AddBoth', 'RemoveAs', 'RemoveBoth')
    $doTo    = $action -in @('AddTo', 'AddBoth', 'RemoveTo', 'RemoveBoth')
    $adding  = $action -like 'Add*'
    $mode    = if ($adding) { 'Add' } else { 'Remove' }

    $dlName = "$($script:Dl.Display) <$($script:Dl.Smtp)>"
    $who    = "$($u.Display) <$($u.Smtp)>"

    # --- the refusals, before anything is written --------------------------
    if ($doTo) {
        $can = Test-CanWriteSendTo
        if (-not $can.Allowed) {
            Show-Out $can.Reason
            Write-FailedLog -Action $mode -Target (Get-LogTarget $script:Dl.Smtp $script:Dl.Guid 'DL Send-As/Send-To') `
                -Detail "REFUSED before writing: $($can.Reason)" -Environment $Environment
            return
        }
    }

    $already = Test-IsAllowedSender -Smtp $u.Smtp -Display $u.Display
    if ($doTo -and $adding -and $already) {
        Show-Out "$who is already an allowed sender on $dlName, so Send-To needs no change."
        $doTo = $false
        if (-not $doAs) { return }
    }
    if ($doTo -and -not $adding -and -not $already) {
        Show-Out "$who is not in the allowed-sender list on $dlName, so there is no Send-To to remove."
        $doTo = $false
        if (-not $doAs) { return }
    }

    # --- the sentence, and the two dangerous cases spelled out -------------
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("List : $dlName")
    [void]$lines.Add("User : $who")
    [void]$lines.Add('')
    if ($doAs) { [void]$lines.Add("$mode Send-As") }
    if ($doTo) { [void]$lines.Add("$mode Send-To (allowed senders)") }

    $defaultCancel = $false

    # THE ONE THAT REVERSES THE OPERATOR'S INTENTION.
    if ($doTo -and $adding -and @($script:SendTo).Count -eq 0) {
        [void]$lines.Add('')
        [void]$lines.Add('READ THIS. The allowed-sender list on this list is currently EMPTY, which means ANYONE can mail it.')
        [void]$lines.Add('')
        [void]$lines.Add("Adding $($u.Display) does NOT give them something they lack. It makes them the ONLY person who may mail this list. Everybody else -- the whole company -- stops being able to.")
        [void]$lines.Add('')
        [void]$lines.Add('If what you wanted was "let this person mail the list", they can already. Cancel.')
        $defaultCancel = $true
    }

    # THE REVERSE, WHICH IS EASIER TO DO BY ACCIDENT THAN IT LOOKS.
    if ($doTo -and -not $adding -and @($script:SendTo).Count -eq 1) {
        [void]$lines.Add('')
        [void]$lines.Add("This is the LAST entry in the allowed-sender list. Removing it leaves the list empty, which means the restriction is gone and ANYONE will be able to mail $($script:Dl.Display).")
        $defaultCancel = $true
    }

    if ($doTo -and -not $script:AuthReq) {
        [void]$lines.Add('')
        [void]$lines.Add('Note: "Require sender authentication" is OFF on this list, so the allowed-sender list is not enforced. This change will be stored and will not alter who can actually mail the list.')
    }

    [void]$lines.Add('')
    [void]$lines.Add('Proceed?')

    if (-not (Confirm-Sentence -Sentence ($lines -join [Environment]::NewLine) -DefaultCancel:$defaultCancel)) {
        Show-Out 'Nothing was changed.'
        return
    }

    # --- do it, one at a time, and report each -----------------------------
    $done   = New-Object System.Collections.ArrayList
    $failed = New-Object System.Collections.ArrayList

    $old = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnApply.Enabled = $false
    try {
        if ($doAs) {
            try { Set-DlSendAs -Mode $mode; [void]$done.Add("$mode Send-As") }
            catch { [void]$failed.Add("Send-As: $(Format-ErrorDetail -ErrorRecord $_ -Context "$mode Send-As")") }
        }
        if ($doTo) {
            try { Set-DlSendTo -Mode $mode; [void]$done.Add("$mode Send-To") }
            catch { [void]$failed.Add("Send-To: $(Format-ErrorDetail -ErrorRecord $_ -Context "$mode Send-To")") }
        }

        # READ BACK. An absent exception is not evidence the change landed, and
        # on a synced object it positively is not.
        Update-ListView

        $L = New-Object System.Collections.ArrayList
        if (@($done).Count -gt 0)   { [void]$L.Add("DONE: $(@($done) -join '; ') for $who on $dlName.") }
        if (@($failed).Count -gt 0) { [void]$L.Add(''); foreach ($x in $failed) { [void]$L.Add("FAILED: $x") } }
        [void]$L.Add('')
        [void]$L.Add('The panel above is a fresh read, not what was asked for -- if a change is not visible there, it did not land.')
        Show-Out ($L -join [Environment]::NewLine) -Keep

        $detail = "$($mode): $(@($done) -join '; ') for $($u.Smtp) on $($script:Dl.Smtp)."
        if (@($failed).Count -eq 0) {
            Write-CompletedLog -Action $mode -Target (Get-LogTarget $script:Dl.Smtp $script:Dl.Guid 'DL Send-As/Send-To') `
                -Result 'Success' -Detail $detail -Environment $Environment
        } elseif (@($done).Count -gt 0) {
            Write-CompletedLog -Action $mode -Target (Get-LogTarget $script:Dl.Smtp $script:Dl.Guid 'DL Send-As/Send-To') `
                -Result 'Partial' -Detail "$detail Failed: $(@($failed) -join ' | ')" -Environment $Environment
        } else {
            Write-FailedLog -Action $mode -Target (Get-LogTarget $script:Dl.Smtp $script:Dl.Guid 'DL Send-As/Send-To') `
                -Detail "Nothing applied for $($u.Smtp) on $($script:Dl.Smtp). $(@($failed) -join ' | ')" -Environment $Environment
        }
    } catch {
        Show-Out (Format-ErrorDetail -ErrorRecord $_ -Context 'Applying')
        Write-FailedLog -Action $mode -Target (Get-LogTarget $script:Dl.Smtp $script:Dl.Guid 'DL Send-As/Send-To') `
            -Detail 'Apply failed.' -ErrorRecord $_ -Environment $Environment
    } finally {
        $btnApply.Enabled = $true
        $form.Cursor = $old
    }
}
#endregion

#region Handlers ----------------------------------------------------------------
# B1: every handler wrapped, or an escaping exception is caught by .NET and the
# operator gets a CLR dialog with no script name on it.
$btnLook.Add_Click({ try { Update-ListView } catch { Show-HandlerError -ErrorRecord $_ -Context 'Look up' } })
$btnRefresh.Add_Click({ try { Update-ListView } catch { Show-HandlerError -ErrorRecord $_ -Context 'Refresh' } })
$btnApply.Add_Click({ try { Invoke-Apply } catch { Show-HandlerError -ErrorRecord $_ -Context 'Apply' } })

$txtDl.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $_.SuppressKeyPress = $true
        try { Update-ListView } catch { Show-HandlerError -ErrorRecord $_ -Context 'Look up' }
    }
})
#endregion

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL Send-As and Send-To' | Out-Null

$form.Add_Shown({
    try {
        if ("$($txtDl.Text)".Trim()) { Update-ListView }
        else { Show-Out 'Type a distribution list above and press Look up, then the user, then pick an action.' }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening' }
})

[void]$form.ShowDialog()
