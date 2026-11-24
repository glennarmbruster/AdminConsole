#Requires -Version 5.1
# TOOLKIT-VARIANT: Dual
# TOOLKIT-REQUIRES: OwnProcess
#
# ONE WINDOW FOR EVERY KIND OF GROUP, ON EITHER SIDE.
#
# It needs its OWN PROCESS because it can hold an on-premises Exchange session
# and an Exchange Online one in the same window, and both export
# Get-DistributionGroup. Every call here is qualified -- Invoke-OnPremCommand
# or Invoke-EXOCommand -- so which directory answers is decided by this code
# and never by import order.
<#
    Group-Editor.ps1
    One group, every setting, whatever kind of group it is and wherever it
    lives.

    Author   : Glenn Armbruster
    Created  : 2026-09-18

    WHY THIS EXISTS. Three tools looked like group editors and were really a
    router plus two half-editors: Group-Manager.ps1 decided which to open,
    Group-Manager-Editor.ps1 did distribution lists properly on both sides,
    and GroupEditor.ps1 did Microsoft 365 groups because the other one cannot
    -- a Microsoft 365 group answers *-DistributionGroup with "The current
    operation is not supported on GroupMailbox". Add the membership tools and
    the reports and there were eight windows for one object.

    THIS IS THE ROUTER AND BOTH EDITORS IN ONE WINDOW. The routing did not go
    away -- it cannot, the cmdlet families really are different -- it moved
    inside, where it decides which tabs are live rather than which program to
    start.

    NOTHING HERE CREATES OR DELETES A GROUP. An editor that can delete its own
    subject is a different and more dangerous kind of tool; creation lives in
    NewCloudGroup.ps1 and ManageMailObject.ps1, and conversion in
    DLtoSharedMB.ps1.
#>
# TOOLKIT-HELP-RUNS: Reads one group across several tabs, on-premises or in Exchange Online. Each section's Apply changes only that section, and the risky ones ask first.
# TOOLKIT-HELP-BEGIN
# Everything about one group, in tabs. Type an address, press Load.
#
# IT WORKS OUT WHAT KIND OF GROUP IT IS, and that decides what you can edit:
#
#   Distribution list, mail-enabled security group
#                   -- owners, members, addresses, mail flow, Send As.
#                      On-premises or cloud, decided by where it is mastered.
#   Microsoft 365 group
#                   -- owners and members through a different cmdlet family,
#                      plus the settings Exchange keeps for it.
#   Dynamic distribution group
#                   -- membership is a FILTER, not a list. It is shown and
#                      cannot be edited, because there is nothing to edit.
#
# A SYNCED GROUP IS EDITED ON-PREMISES. The window says so, and writes go to
# Active Directory, because a cloud write to a synced group appears to succeed
# and is undone by the next directory sync.
#
# EACH BLOCK HAS ITS OWN APPLY. There is no global Save.
#
# WHAT IT WILL NOT DO: create a group, delete a group, or convert one type
# into another. Those are separate tools on purpose.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    # Forced on-premises, the way Group-Manager used to launch its editor.
    [switch]$OnPrem
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

#region State and routing -------------------------------------------------------
$script:Grp     = $null    # the resolved group
$script:Mode    = 'Cloud'  # what the Editing switch says
$script:OnPremWhy = ''     # why the switch is unavailable, if it is
$script:Loaded  = @{}
$script:Busy    = $false
$script:TabReadFailed = $false

# THE DOMAIN CONTROLLER, PINNED ONCE AND USED FOR EVERY ON-PREMISES CALL.
#
# -DomainController wants a REAL CONTROLLER, never a domain name.
# Get-TenantDirectoryServer deliberately returns the DOMAIN when a directory
# has several controllers, and Exchange rejects that with "Active Directory
# error 0x52 ... suitability of server", which reads as a broken server and is
# not one. A wrong pin fails as "out of the current user's write scope", which
# reads as a permissions problem and is not one either.
#
# NO PIN IS A VALID ANSWER. An empty splat lets Exchange choose, which works.
# What must not happen is the READ landing on one controller and the WRITE on
# another, so whatever is chosen is used for everything.
$script:DCParam = @{}

function Invoke-GrpCmd {
    <#
        THE ONLY PLACE THAT KNOWS THE DIFFERENCE between the two directories.

        Every read and every write in this file goes through here. Written the
        other way -- an if/else at each call site -- there are forty chances to
        get it wrong and one of them will be a write.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    # The caller's hashtable is COPIED before the controller is added to it.
    # Adding in place mutates a table the caller may reuse, and the second call
    # then carries a parameter the caller never asked for.
    $p = @{} + $Parameters
    if ($script:Grp -and $script:Grp.OnPrem) {
        foreach ($k in $script:DCParam.Keys) { $p[$k] = $script:DCParam[$k] }
        return Invoke-OnPremCommand -Name $Name -Parameters $p
    }
    return Invoke-EXOCommand -Name $Name -Parameters $p
}

function Set-GrpDomainController {
    <#
        Pin a real controller for the on-premises side, or pin nothing.

        The first directory that resolves to a server wins, and that server is
        used for every later call so the read and the write cannot land in
        different forests.
    #>
    $script:DCParam = @{}
    try {
        foreach ($d in @(Get-TenantDirectory)) {
            $srv = ''
            try { $srv = Resolve-DirectoryDomainController -DirectoryId "$($d.Id)" } catch { }
            if ("$srv".Trim()) { $script:DCParam = @{ DomainController = "$srv" }; return }
        }
    } catch { }
}

function Resolve-GroupTarget {
    <#
        What kind of group is this, and which directory masters it.

        GET-RECIPIENT IS ASKED FIRST, and that ordering is load-bearing. A
        Microsoft 365 group answers Get-DistributionGroup with "The current
        operation is not supported on GroupMailbox", so a tool that asks that
        first reports three real groups as unroutable.

        IsDirSynced IS NEVER CACHED between loads. It is correct until the
        weekend the group converts, and then it is confidently wrong.
    #>
    param([Parameter(Mandatory)][string]$Identity, [bool]$ForceOnPrem = $false)

    $r = [PSCustomObject]@{
        Ok = $false; Smtp = ''; Display = ''; Name = ''; Alias = ''
        Kind = ''            # Distribution | Security | Unified | Dynamic
        Type = ''            # the raw RecipientTypeDetails
        OnPrem = $false; IsSynced = $false
        EmsHere = $false     # was an on-premises Exchange session available
        FoundIn = ''         # 'Cloud' or 'OnPrem' -- which directory answered
        Raw = $null; Dn = ''; Reason = ''
    }

    # THE BARE LOCAL PART IS TRIED TOO. Get-TargetSmtp completes a bare word
    # with the tenant's primary domain, and a group whose only address is
    # @tenant.onmicrosoft.com does not have one -- which is most groups nobody
    # gave a vanity address to.
    $tries = @("$Identity")
    if ("$Identity" -match '^([^@]+)@') { $tries += $Matches[1] }

    # BOTH DIRECTORIES ARE ASKED, AND THE CLOUD IS NOT ASKED FIRST WHEN THE
    # CALLER ALREADY SAID ON-PREMISES.
    #
    # THE BUG THIS FIXES, because it made the window useless for a whole class
    # of group: the resolve was Exchange Online only. A group that lives in a
    # child domain, or that has not synchronised, or that is simply not
    # mail-enabled in the cloud, was reported as "not found in Exchange
    # Online" -- and every on-premises group in a directory the cloud does not
    # carry showed nothing at all.
    #
    # THE FOREST SCOPE IS SET BEFORE THE ON-PREMISES READ. A child-domain
    # object is invisible to an Exchange session scoped to its own domain, and
    # invisible looks exactly like absent. Set-ADSessionScope puts the session
    # forest-wide first and only then tries to pin a controller.
    $rec = $null
    $onPremFirst = $ForceOnPrem
    $order = $(if ($onPremFirst) { @('OnPrem', 'Cloud') } else { @('Cloud', 'OnPrem') })
    $emsHere = $false
    try { $emsHere = [bool](Test-EmsAvailable).Available } catch { }

    foreach ($where in $order) {
        if ($rec) { break }
        if ($where -eq 'OnPrem') {
            if (-not $emsHere) { continue }
            try { $null = Set-ADSessionScope } catch { }
        }
        foreach ($t in $tries) {
            try {
                $hit = @(& $(if ($where -eq 'OnPrem') { 'Invoke-OnPremCommand' } else { 'Invoke-EXOCommand' }) `
                            -Name 'Get-Recipient' -Parameters @{
                                Identity = $t; ResultSize = 1; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' })
                if (@($hit).Count -ge 1) { $rec = $hit[0]; $r.FoundIn = $where; break }
            } catch { }
        }
    }

    if (-not $rec) {
        $r.Reason = $(if ($emsHere) {
            'was not found in Exchange Online or on-premises. Check the spelling, and that the group is in a directory this console is configured for.'
        } else {
            'was not found in Exchange Online, and on-premises Exchange is not connected in this window so the on-premises directory was not searched.'
        })
        return $r
    }

    $r.Smtp    = "$($rec.PrimarySmtpAddress)"
    $r.Display = "$($rec.DisplayName)"
    $r.Name    = "$($rec.Name)"
    $r.Alias   = "$($rec.Alias)"
    $r.Type    = "$($rec.RecipientTypeDetails)"
    $r.Dn      = "$($rec.DistinguishedName)"

    switch -Regex ($r.Type) {
        'GroupMailbox'                      { $r.Kind = 'Unified' }
        'DynamicDistributionGroup'          { $r.Kind = 'Dynamic' }
        'MailUniversalSecurityGroup'        { $r.Kind = 'Security' }
        'MailUniversalDistributionGroup'    { $r.Kind = 'Distribution' }
        'RoomList'                          { $r.Kind = 'Distribution' }
        default                             { $r.Kind = '' }
    }
    if (-not $r.Kind) {
        $r.Reason = "is a $($r.Type), which is not a group this window edits."
        return $r
    }

    # WHICH SIDE IS DECIDED FIRST, FROM THE CHEAP CLOUD LOOKUP, because the
    # detail has to be read from the directory that is going to be WRITTEN.
    #
    # Reading it from Exchange Online for an on-premises group looks harmless
    # and is not: an on-premises write is not visible in the cloud until the
    # next sync, so every re-read after an Apply repaints the tab from the
    # pre-change copy. The checkboxes snap back, the owner grid shows the old
    # list, and the window contradicts what it has just done.
    # A GROUP THE CLOUD HAS NEVER HEARD OF IS AN ON-PREMISES GROUP, whatever
    # IsDirSynced says -- that property is meaningless on an object the cloud
    # did not return.
    # ON-PREMISES ONLY IF THERE IS AN ON-PREMISES SESSION TO USE. Without one,
    # claiming the on-prem side sends the detail read to a module that is not
    # loaded and the group appears unreadable -- when the cloud copy was
    # sitting right there. The banner then says the cloud copy is being shown
    # and why.
    $r.IsSynced = [bool]$rec.IsDirSynced
    $r.OnPrem   = ($emsHere -and ($ForceOnPrem -or $r.IsSynced -or ($r.FoundIn -eq 'OnPrem')))
    $r.EmsHere  = $emsHere
    # A Microsoft 365 group has no on-premises half to write to.
    if ($r.Kind -eq 'Unified') { $r.OnPrem = $false }

    # $script:Grp is what Invoke-GrpCmd routes on, so it has to be set before
    # the detail read goes through it.
    $script:Grp = $r

    try {
        switch ($r.Kind) {
            'Unified' {
                # ALWAYS CLOUD. There is no on-premises Get-UnifiedGroup.
                $r.Raw = Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{ Identity = $r.Smtp; ErrorAction = 'Stop' }
            }
            'Dynamic' {
                $r.Raw = Invoke-GrpCmd -Name 'Get-DynamicDistributionGroup' -Parameters @{ Identity = $r.Smtp; ErrorAction = 'Stop' }
            }
            default {
                $r.Raw = Invoke-GrpCmd -Name 'Get-DistributionGroup' -Parameters @{ Identity = $r.Smtp; ErrorAction = 'Stop' }
            }
        }
    } catch {
        $r.Reason = "was found as a $($r.Type) but its detail could not be read$(if ($r.OnPrem) { ' on-premises' }): $($_.Exception.Message)"
        return $r
    }

    $r.Ok = $true
    return $r
}
#endregion

#region Form shell --------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Group Editor'
$form.ClientSize    = New-Object System.Drawing.Size(1120, 760)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$y = 12
$lblWho          = New-Object System.Windows.Forms.Label
$lblWho.Text     = 'Group:'
$lblWho.Location = New-Object System.Drawing.Point(12, ($y + 3))
$lblWho.Size     = New-Object System.Drawing.Size(48, 18)
$form.Controls.Add($lblWho)

$txtGrp          = New-Object System.Windows.Forms.TextBox
$txtGrp.Location = New-Object System.Drawing.Point(62, $y)
$txtGrp.Size     = New-Object System.Drawing.Size(340, 24)
$txtGrp.Font     = New-Object System.Drawing.Font('Consolas', 9)
$txtGrp.Text     = "$(Get-TargetSmtp -Text $SmtpAddress)"
$form.Controls.Add($txtGrp)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load'
$btnLoad.Location = New-Object System.Drawing.Point(412, ($y - 2))
$btnLoad.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnLoad)

# THE EDITING SWITCH, from Common.ps1 -- the identical control the Mailbox and
# Room editors get, in the same place and with the same wording.
$cmbMode = New-EditorModeSwitch -Form $form -X 718 -Y $y -LabelX 660

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(952, ($y - 2))
$btnClose.Size     = New-Object System.Drawing.Size(90, 28)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)
$y += 34

$lblIdent           = New-Object System.Windows.Forms.Label
$lblIdent.Tag       = 'NoTheme'
$lblIdent.Location  = New-Object System.Drawing.Point(12, $y)
$lblIdent.Size      = New-Object System.Drawing.Size(1096, 18)
$lblIdent.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblIdent.ForeColor = [System.Drawing.Color]::DimGray
$lblIdent.Text      = 'No group loaded.'
$form.Controls.Add($lblIdent)
$y += 20

$lblSide           = New-Object System.Windows.Forms.Label
$lblSide.Tag       = 'NoTheme'
$lblSide.Location  = New-Object System.Drawing.Point(12, $y)
$lblSide.Size      = New-Object System.Drawing.Size(1096, 18)
$lblSide.ForeColor = [System.Drawing.Color]::DimGray
$lblSide.Text      = ''
$form.Controls.Add($lblSide)
$y += 22

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, $y)
$tabs.Size     = New-Object System.Drawing.Size(1096, 640)
$form.Controls.Add($tabs)

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(12, ($y + 646))
$lblStatus.Size      = New-Object System.Drawing.Size(1096, 18)
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$lblStatus.Text      = 'Enter a group address and press Load.'
$form.Controls.Add($lblStatus)

$form.ClientSize = New-Object System.Drawing.Size(1120, ($lblStatus.Bottom + 12))

function New-GrpTab {
    param([string]$Key, [string]$Text)
    $p = New-Object System.Windows.Forms.TabPage
    $p.Text       = $Text
    $p.Name       = $Key
    $p.Padding    = New-Object System.Windows.Forms.Padding(8)
    $p.AutoScroll = $true
    [void]$tabs.TabPages.Add($p)
    return $p
}

$tabOverview = New-GrpTab -Key 'Overview' -Text 'Overview'
$tabOwners   = New-GrpTab -Key 'Owners'   -Text 'Owners'
$tabMembers  = New-GrpTab -Key 'Members'  -Text 'Members'
$tabMailFlow = New-GrpTab -Key 'MailFlow' -Text 'Mail Flow'
$tabAccess   = New-GrpTab -Key 'Access'   -Text 'Send As'
#endregion

#region Shell helpers -----------------------------------------------------------
function Set-GrpStatus {
    param([string]$Text, [switch]$Bad)
    $lblStatus.Text = $Text
    $lblStatus.ForeColor = $(if ($Bad) { [System.Drawing.Color]::FromArgb(180, 0, 0) } else { [System.Drawing.Color]::DimGray })
    $lblStatus.Refresh()
}

function Invoke-GrpAction {
    param([scriptblock]$Do, [string]$Context)
    if ($script:Busy) { return }
    $script:Busy = $true
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        & $Do
    } catch {
        Set-GrpStatus -Text "$Context failed: $($_.Exception.Message)" -Bad
        Write-GrpAudit -Action 'GroupEditorAction' -Result 'Failed' -Detail "$Context failed: $($_.Exception.Message)"
        Show-HandlerError -ErrorRecord $_ -Context $Context
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $script:Busy = $false
    }
}

function Test-GrpLoaded {
    if (-not $script:Grp -or -not $script:Grp.Ok) {
        Set-GrpStatus -Text 'Load a group first.' -Bad
        return $false
    }
    return $true
}

function Write-GrpAudit {
    param([string]$Action, [string]$Detail, [string]$Result = 'Success')
    try {
        # THE ENVIRONMENT IS WHICH DIRECTORY WAS WRITTEN, not a constant. A
        # membership change made on-premises and one made in the cloud are
        # different acts, and the log has to tell them apart.
        $env = $(if ($script:Grp -and $script:Grp.OnPrem) { 'OnPrem' } else { 'Cloud' })
        Write-CompletedLog -Action $Action -Target (Get-LogTarget $script:Grp.Smtp) `
            -Result $Result -Detail $Detail -Environment $env
    } catch { }
}

function Test-GrpKind {
    <#
        Refuse an action this kind of group cannot do, and say why.

        A DYNAMIC GROUP HAS NO MEMBER LIST AT ALL -- its membership is a
        filter evaluated at send time. Add-DistributionGroupMember against one
        fails with a message about the identity, which reads as the group not
        existing.
    #>
    param([string[]]$Allow, [string]$What = 'This')
    if (-not $script:Grp) { return $false }
    if ($Allow -contains $script:Grp.Kind) { return $true }
    $why = switch ($script:Grp.Kind) {
        'Dynamic' { "$What does not apply to a dynamic distribution group -- its membership is a filter, not a list. The filter is on the Members tab." }
        'Unified' { "$What is not available for a Microsoft 365 group." }
        default   { "$What does not apply to a $($script:Grp.Type)." }
    }
    Set-GrpStatus -Text $why -Bad
    $null = Show-ToolDialog($why, 'Not available for this kind of group', 'OK', 'Information')
    return $false
}

function Update-GrpIdentity {
    $g = $script:Grp
    if (-not $g -or -not $g.Ok) { $lblIdent.Text = 'No group loaded.'; $lblSide.Text = ''; return }
    $kindText = switch ($g.Kind) {
        'Unified'      { 'Microsoft 365 group' }
        'Dynamic'      { 'dynamic distribution group' }
        'Security'     { 'mail-enabled security group' }
        'Distribution' { 'distribution list' }
        default        { $g.Type }
    }
    $lblIdent.Text = "$($g.Display)   <$($g.Smtp)>   $kindText"
    $lblIdent.ForeColor = [System.Drawing.Color]::FromArgb(0, 110, 60)

    if ($g.OnPrem) {
        $dc = $(if ($script:DCParam.ContainsKey('DomainController')) { " via $($script:DCParam['DomainController'])" } else { ' (no controller pinned -- Exchange chooses)' })
        $lblSide.Text = "EDITED ON-PREMISES$dc. This group is mastered in Active Directory; a cloud write would be accepted and undone by the next directory sync."
        $lblSide.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)
    } elseif ($g.IsSynced -and -not $g.EmsHere) {
        $lblSide.Text = 'SYNCED FROM ACTIVE DIRECTORY, and this window has no on-premises Exchange session -- so you are seeing the CLOUD copy. Writes here will be undone by the next directory sync. Open this from a console with on-premises Exchange connected.'
        $lblSide.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
    } elseif ($g.IsSynced) {
        $lblSide.Text = 'SYNCED FROM ACTIVE DIRECTORY but edited in the cloud -- a Microsoft 365 group has no on-premises half.'
        $lblSide.ForeColor = [System.Drawing.Color]::FromArgb(190, 110, 0)
    } else {
        $lblSide.Text = 'Cloud group -- edited in Exchange Online.'
        $lblSide.ForeColor = [System.Drawing.Color]::DimGray
    }
}

function Resolve-GrpRecipient {
    <#
        One name, turned into the things the write actually needs.

        PrimarySmtpAddress IS THE KEY, NEVER Identity. Identity for a group
        comes back as its Active Directory CN, and feeding that back to a
        cmdlet is a fresh forest-wide NAME lookup rather than a bind -- which
        can land on a different object entirely.

        Sam is carried because the on-premises Send-As write needs it: the
        *-ADPermission cmdlets take a trustee that AD recognises, and an SMTP
        address is not one.
    #>
    param([string]$Entry)
    $r = [PSCustomObject]@{ Ok = $false; Raw = "$Entry"; Name = "$Entry"; Address = ''; Sam = ''; Dn = ''; Reason = '' }
    if (-not "$Entry".Trim()) { $r.Reason = 'empty'; return $r }
    # RESOLVED IN THE DIRECTORY THAT WILL BE WRITTEN, which is two separate
    # problems solved at once. Exchange Online's Get-Recipient does not
    # populate SamAccountName, and the on-premises Send As write needs one; and
    # a trustee that exists only on-premises is simply not in Exchange Online
    # to be found, so the handler would abort with 'nothing was changed' for a
    # write that would have worked.
    try {
        $hit = @(Invoke-GrpCmd -Name 'Get-Recipient' -Parameters @{
                    Identity = $Entry; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' })
        if (@($hit).Count -gt 1) { $r.Reason = 'more than one recipient matches'; return $r }
        if (@($hit).Count -eq 1) {
            $r.Ok      = $true
            $r.Name    = "$($hit[0].DisplayName)"
            $r.Address = "$($hit[0].PrimarySmtpAddress)"
            $r.Sam     = "$($hit[0].SamAccountName)"
            $r.Dn      = "$($hit[0].DistinguishedName)"
            return $r
        }
    } catch { $r.Reason = $_.Exception.Message; return $r }
    $r.Reason = 'not found'
    return $r
}
#endregion

#region Overview tab ------------------------------------------------------------
$ovY = 8
[void](New-TabLabel $tabOverview 'IDENTITY' 8 $ovY -Bold)
$ovY += 22
[void](New-TabLabel $tabOverview 'Display name:' 8 $ovY)
$ovName = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Kind:' 560 $ovY 135)
$ovKind = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Address:' 8 $ovY)
$ovSmtp = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Alias:' 560 $ovY 135)
$ovAlias = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Name:' 8 $ovY)
$ovRawName = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Mastered in:' 560 $ovY 135)
$ovSource = New-TabValue $tabOverview 700 $ovY 320
$ovY += 28

# HIDE-FROM-GAL ON THE FIRST TAB, in all three editors. 'Can people find this'
# belongs beside 'what is this', and the Room Editor had it right first.
[void](New-TabLabel $tabOverview 'Hide from the address book:' 8 ($ovY + 3) 180)
$cmbOvHide               = New-Object System.Windows.Forms.ComboBox
$cmbOvHide.DropDownStyle = 'DropDownList'
$cmbOvHide.Location      = New-Object System.Drawing.Point(196, $ovY)
$cmbOvHide.Size          = New-Object System.Drawing.Size(110, 24)
[void]$cmbOvHide.Items.Add('True'); [void]$cmbOvHide.Items.Add('False')
$tabOverview.Controls.Add($cmbOvHide)
$btnOvHide = New-TabButton $tabOverview 'Apply' 320 ($ovY - 2) 100
$ovHideNote = New-TabLabel $tabOverview '' 430 ($ovY + 3) 600 -Muted
$ovY += 32

[void](New-TabLabel $tabOverview 'ADDRESSES' 8 $ovY -Bold)
$ovY += 22
$gridAddr = New-TabGrid $tabOverview @(@('Address', 460), @('Type', 90), @('Primary', 80)) 8 $ovY 700 190
$ovY += 200
$ovAddrNote = New-TabLabel $tabOverview '' 8 $ovY 1040 -Muted
$ovAddrNote.Size = New-Object System.Drawing.Size(1040, 32)
$ovAddrNote.Text = ('Read-only here on purpose. Changing a group''s addresses is a rename, and a rename has an X500 to keep and a LegacyExchangeDN to think about -- that belongs in a tool that does the whole job.' + "`r`n" +
                    'X500 entries are what let replies to older mail still resolve, and are never removed automatically.')
$ovY += 40

[void](New-TabLabel $tabOverview 'ANYTHING WRONG?' 8 $ovY -Bold)
$ovY += 22
$ovVerdict            = New-Object System.Windows.Forms.TextBox
$ovVerdict.Multiline  = $true
$ovVerdict.ReadOnly   = $true
$ovVerdict.ScrollBars = 'Vertical'
$ovVerdict.WordWrap   = $true
$ovVerdict.Location   = New-Object System.Drawing.Point(8, $ovY)
$ovVerdict.Size       = New-Object System.Drawing.Size(1040, 150)
$ovVerdict.Font       = New-Object System.Drawing.Font('Consolas', 9)
$ovVerdict.Tag        = 'NoTheme'
$ovVerdict.BackColor  = [System.Drawing.Color]::WhiteSmoke
$ovVerdict.ForeColor  = [System.Drawing.Color]::Black
$tabOverview.Controls.Add($ovVerdict)

function Update-OverviewTab {
    $g  = $script:Grp
    $gr = $g.Raw
    $ovName.Text    = "$($gr.DisplayName)"
    $ovKind.Text    = "$($g.Type)"
    $ovSmtp.Text    = "$($g.Smtp)"
    $ovAlias.Text   = "$($gr.Alias)"
    $ovRawName.Text = "$($gr.Name)"
    $ovSource.Text  = $(if ($g.OnPrem) { 'Active Directory (on-premises)' } elseif ($g.IsSynced) { 'Active Directory, edited in the cloud' } else { 'Exchange Online' })

    # A MICROSOFT 365 GROUP HIDES DIFFERENTLY. HiddenFromAddressListsEnabled
    # is a distribution-group property; the Unified cmdlets have their own,
    # and Set-DistributionGroup refuses the object outright -- so the control
    # says so rather than offering a write that cannot land.
    $cmbOvHide.SelectedItem = $(if ($gr.HiddenFromAddressListsEnabled) { 'True' } else { 'False' })
    $canHide = ($g.Kind -ne 'Unified')
    $cmbOvHide.Enabled = $canHide; $btnOvHide.Enabled = $canHide
    $ovHideNote.Text = $(if (-not $canHide) {
        'Not available for a Microsoft 365 group here -- it is not a distribution-group property.'
    } elseif ($g.OnPrem) {
        "Written on-premises. $($script:OnPremSyncNotice)"
    } else {
        'Removes the group from the address book. Mail sent straight to its address still delivers.'
    })
    $ovHideNote.ForeColor = $(if (-not $canHide) { [System.Drawing.Color]::FromArgb(150, 80, 0) }
                              elseif ($g.OnPrem) { [System.Drawing.Color]::FromArgb(150, 80, 0) }
                              else { [System.Drawing.Color]::DimGray })

    $gridAddr.Rows.Clear()
    foreach ($a in @($gr.EmailAddresses)) {
        $str = "$a"
        if (-not $str) { continue }
        $prefix = ''; $addr = $str
        if ($str -match '^([A-Za-z0-9]+):(.*)$') { $prefix = $Matches[1]; $addr = $Matches[2] }
        $isPrimary = ($prefix -ceq 'SMTP')
        $i = $gridAddr.Rows.Add($addr, $(if ($prefix -match '(?i)^smtp$') { 'SMTP' } else { $prefix }), $(if ($isPrimary) { 'PRIMARY' } else { '' }))
        if ($isPrimary) {
            $gridAddr.Rows[$i].DefaultCellStyle.Font = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
        } elseif ($prefix -notmatch '(?i)^smtp$') {
            $gridAddr.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DimGray
        }
    }

    $f = New-Object System.Collections.ArrayList
    # OWNERLESS IS THE FINDING PEOPLE ACTUALLY CHASE. A group with no owner is
    # a real state and a bad one: nobody can be asked to approve a member,
    # nobody is told when it goes stale, and for a Microsoft 365 group the
    # expiry policy has nobody to renew it.
    # $null MEANS 'COULD NOT READ', EMPTY MEANS 'THERE ARE NONE'. Collapsing
    # the two asserts that a group has no owner when the owner read merely
    # failed -- which is the finding people act on.
    $owners = $null
    try { $owners = @(Get-GrpOwner) } catch { $owners = $null }
    if ($null -ne $owners -and @($owners).Count -eq 0) {
        [void]$f.Add('THIS GROUP HAS NO OWNER. Nobody can approve a membership request, nobody is asked before it is deleted, and nothing tells anyone when it goes stale. See Owners.')
    }
    if ($g.Kind -ne 'Dynamic') {
        $mc = $null
        try { $mc = @(Get-GrpMember).Count } catch { }
        if ($null -ne $mc) {
            if ($mc -eq 0)    { [void]$f.Add('The group has NO members. Mail sent to it goes nowhere. See Members.') }
            elseif ($mc -eq 1) { [void]$f.Add('The group has ONE member. That is often a list somebody outgrew -- worth asking whether it should still exist. See Members.') }
        }
    }
    if ($g.Kind -ne 'Unified') {
        $req = [bool]$gr.RequireSenderAuthenticationEnabled
        $accept = @($gr.AcceptMessagesOnlyFromSendersOrMembers)
        if (-not $req -and $accept.Count -gt 0) {
            [void]$f.Add("Sender authentication is OFF while $($accept.Count) allowed sender(s) are listed. The list is NOT being enforced -- the group is open while still looking restricted. See Mail Flow.")
        }
        if (-not $req) {
            [void]$f.Add('Anyone outside the organisation can send to this group. That is sometimes deliberate and is the usual route for spam to reach everybody at once. See Mail Flow.')
        }
    }
    if ($gr.HiddenFromAddressListsEnabled) {
        [void]$f.Add('Hidden from the address book. It still receives mail sent to its address.')
    }
    if ($g.IsSynced -and -not $g.OnPrem -and $g.Kind -ne 'Unified') {
        [void]$f.Add('This group is synced from Active Directory but this window is in cloud mode. Writes will be reverted by the next sync.')
    }

    $sb = New-Object System.Text.StringBuilder
    if ($f.Count -eq 0) {
        [void]$sb.AppendLine('Nothing obviously wrong. It has an owner, it has members, and its mail-flow restrictions agree with each other.')
    } else {
        foreach ($x in $f) { [void]$sb.AppendLine("* $x") }
    }
    $ovVerdict.Text = $sb.ToString()
}
$btnOvHide.Add_Click({
    if (-not (Test-GrpLoaded)) { return }
    Invoke-GrpAction {
        if (-not (Test-GrpKind -Allow @('Distribution','Security','Dynamic') -What 'Hiding from the address book')) { return }
        $want = ("$($cmbOvHide.SelectedItem)" -eq 'True')
        if ($want -eq [bool]$script:Grp.Raw.HiddenFromAddressListsEnabled) { Set-GrpStatus -Text 'Nothing changed.'; return }
        $cmdlet = $(if ($script:Grp.Kind -eq 'Dynamic') { 'Set-DynamicDistributionGroup' } else { 'Set-DistributionGroup' })
        $p = @{ Identity = $script:Grp.Smtp; HiddenFromAddressListsEnabled = $want; Confirm = $false; ErrorAction = 'Stop' }
        # -BypassSecurityGroupManagerCheck is a distribution-group parameter;
        # Set-DynamicDistributionGroup does not declare it.
        if ($script:Grp.Kind -ne 'Dynamic') { $p['BypassSecurityGroupManagerCheck'] = $true }
        Invoke-GrpCmd -Name $cmdlet -Parameters $p | Out-Null
        Write-GrpAudit -Action 'GroupHideFromGal' -Detail "Hidden from address list = $want."
        $script:Grp = Resolve-GroupTarget -Identity $script:Grp.Smtp -ForceOnPrem:($script:Mode -eq 'OnPrem')
        Update-OverviewTab
        $script:Loaded['MailFlow'] = $false
        Set-GrpStatus -Text ($(if ($want) { 'Hidden from the address book.' } else { 'Visible in the address book.' }) +
                             $(if ($script:Grp.OnPrem) { " $($script:OnPremSyncNotice)" }))
    } 'Apply address book visibility'
})
#endregion

#region Owners tab --------------------------------------------------------------
<#
    WHO OWNS IT -- AND THE TRAP UNDERNEATH.

    ManagedBy IS REPLACED WHOLE. Set-DistributionGroup -ManagedBy does not add
    to the list, it overwrites it, and getting that wrong does not error -- it
    silently removes everybody else. So every Apply on this tab writes the
    ENTIRE owner list back, built from the grid, and adding one person means
    writing the other four again.

    A MICROSOFT 365 GROUP IS THE EXCEPTION. Its owners are links, added and
    removed one at a time with Add-/Remove-UnifiedGroupLinks, and
    *-DistributionGroup refuses it outright with "not supported on
    GroupMailbox".

    AN OWNER THAT WILL NOT RESOLVE IS SHOWN, NOT DROPPED. A deleted account
    still occupies a slot in ManagedBy, and a tool that quietly left it out of
    the rewrite would be removing it without being asked.
#>
$owY = 8
[void](New-TabLabel $tabOwners 'OWNERS' 8 $owY -Bold)
$owY += 22
$owNote = New-TabLabel $tabOwners '' 8 $owY 1040 -Muted
$owNote.Size = New-Object System.Drawing.Size(1040, 32)
$owY += 40
$gridOwners = New-TabGrid $tabOwners @(@('Name', 380), @('Address', 300), @('State', 180)) 8 $owY 700 320
Add-GridPickColumn -Grid $gridOwners
$owBtnY = $owY
$btnOwAdd    = New-TabButton $tabOwners 'Add owners...'  720 $owBtnY 170
$owBtnY += 34
$btnOwRemove = New-TabButton $tabOwners 'Remove ticked'  720 $owBtnY 170
$owBtnY += 34
$btnOwAll    = New-TabButton $tabOwners 'Tick all'       720 $owBtnY 82
$btnOwNone   = New-TabButton $tabOwners 'None'           808 $owBtnY 82
$owY += 332
$owHint = New-TabLabel $tabOwners '' 8 $owY 1040 -Muted
$owHint.Size = New-Object System.Drawing.Size(1040, 46)
$owHint.Text = ('For a distribution list the whole owner list is rewritten on every change -- that is how Exchange stores it, and it is why removing one owner is not a one-line operation.' + "`r`n" +
                'A group with no owner is allowed and is almost always a mistake.')

function Get-GrpOwner {
    # The owners, whichever family this group belongs to.
    $g = $script:Grp
    if ($g.Kind -eq 'Unified') {
        return @(Invoke-EXOCommand -Name 'Get-UnifiedGroupLinks' -Parameters @{
                    Identity = $g.Smtp; LinkType = 'Owners'; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
    }
    # ManagedBy holds directory names, not addresses; each is re-resolved so
    # the grid can show a person rather than a CN.
    $out = @()
    foreach ($o in @($g.Raw.ManagedBy)) { $out += "$o" }
    return $out
}

function Update-OwnersTab {
    $gridOwners.Rows.Clear()
    $g = $script:Grp
    $raw = @()
    try { $raw = @(Get-GrpOwner) } catch {
        $script:TabReadFailed = $true
        $owNote.Text = "The owners could not be read: $($_.Exception.Message)"
        $owNote.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
        return
    }

    foreach ($o in $raw) {
        if ($g.Kind -eq 'Unified') {
            $i = $gridOwners.Rows.Add($false, "$($o.DisplayName)", "$($o.PrimarySmtpAddress)", 'ok')
            $gridOwners.Rows[$i].Tag = [PSCustomObject]@{ Ok = $true; Name = "$($o.DisplayName)"; Address = "$($o.PrimarySmtpAddress)"; Raw = "$($o.PrimarySmtpAddress)" }
        } else {
            $res = Resolve-GrpRecipient -Entry "$o"
            $state = $(if ($res.Ok) { 'ok' } else { "could not be re-resolved -- may be deleted" })
            $i = $gridOwners.Rows.Add($false, $(if ($res.Ok) { $res.Name } else { "$o" }), $res.Address, $state)
            $gridOwners.Rows[$i].Tag = [PSCustomObject]@{
                Ok = $res.Ok; Name = $res.Name; Address = $res.Address; Raw = "$o"
            }
            if (-not $res.Ok) { $gridOwners.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(190, 110, 0) }
        }
    }

    if ($gridOwners.Rows.Count -eq 0) {
        $owNote.Text = 'THIS GROUP HAS NO OWNER. Nobody can approve a membership request and nothing tells anyone when it goes stale.'
        $owNote.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
    } elseif ($g.Kind -eq 'Unified') {
        $owNote.Text = "$($gridOwners.Rows.Count) owner(s). Added and removed one at a time -- a Microsoft 365 group stores them as links."
        $owNote.ForeColor = [System.Drawing.Color]::DimGray
    } else {
        $owNote.Text = "$($gridOwners.Rows.Count) owner(s). The whole list is rewritten on every change, so an owner shown in amber is carried through unchanged rather than dropped."
        $owNote.ForeColor = [System.Drawing.Color]::DimGray
    }
}

function Set-GrpOwnerList {
    # THE WHOLE LIST, ALWAYS. -BypassSecurityGroupManagerCheck because
    # Exchange checks whether YOU are in the group's ManagedBy regardless of
    # what your RBAC role says -- an administrator who is not an owner is
    # refused without it, with an error that reads like a permissions problem
    # and is one only in the narrowest sense.
    param([string[]]$Owners)
    Invoke-GrpCmd -Name 'Set-DistributionGroup' -Parameters @{
        Identity = $script:Grp.Smtp
        ManagedBy = @($Owners)
        BypassSecurityGroupManagerCheck = $true
        Confirm = $false
        ErrorAction = 'Stop'
    } | Out-Null
}

$btnOwAll.Add_Click({  Set-GridPicked -Grid $gridOwners -On $true })
$btnOwNone.Add_Click({ Set-GridPicked -Grid $gridOwners -On $false })

$btnOwAdd.Add_Click({
    if (-not (Test-GrpLoaded)) { return }
    Invoke-GrpAction {
        if (-not (Test-GrpKind -Allow @('Distribution','Security','Unified') -What 'Editing owners')) { return }
        $who = Show-InputDialog -Title 'Add owners' -Width 620 `
                -Prompt 'Who should own this group? Addresses or aliases, separated by semicolons.'
        $list = ConvertFrom-DialogList -Text "$who"
        if ($list.Count -eq 0) { return }

        $resolved = @()
        foreach ($w in $list) {
            $r = Resolve-GrpRecipient -Entry $w
            if (-not $r.Ok) { Set-GrpStatus -Text "'$w' $($r.Reason). Nothing was changed." -Bad; return }
            $resolved += $r
        }

        if ($script:Grp.Kind -eq 'Unified') {
            $lines = @(); $ok = 0; $bad = 0
            foreach ($r in $resolved) {
                try {
                    Invoke-EXOCommand -Name 'Add-UnifiedGroupLinks' -Parameters @{
                        Identity = $script:Grp.Smtp; LinkType = 'Owners'; Links = $r.Address
                        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                    $lines += "OK      $($r.Name)"; $ok++
                } catch { $lines += "FAILED  $($r.Name): $($_.Exception.Message)"; $bad++ }
            }
            Write-GrpAudit -Action 'GroupOwnerAdd' -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
                -Detail "added owners: $(($resolved | ForEach-Object { $_.Address }) -join '; ')"
            if ($bad) { Show-ResultListDialog -Title "Add owners -- $ok done, $bad failed" -Lines $lines }
        } else {
            # EVERY EXISTING OWNER IS WRITTEN BACK WITH THE NEW ONES.
            $all = @()
            foreach ($row in $gridOwners.Rows) { if ($row.Tag) { $all += "$(if ($row.Tag.Address) { $row.Tag.Address } else { $row.Tag.Raw })" } }
            foreach ($r in $resolved) { if ($all -notcontains $r.Address) { $all += $r.Address } }
            $a = Show-ToolDialog(
                ("Set the owners of $($script:Grp.Display) to:" + "`r`n`r`n  " + ($all -join "`r`n  ") + "`r`n`r`n" +
                 'The whole owner list is rewritten -- that is how Exchange stores it.'),
                'Set owners', 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
            Set-GrpOwnerList -Owners $all
            Write-GrpAudit -Action 'GroupOwnerSet' -Detail "ManagedBy = $($all -join '; ')"
        }
        $script:Grp = Resolve-GroupTarget -Identity $script:Grp.Smtp -ForceOnPrem:($script:Mode -eq 'OnPrem')
        Update-OwnersTab
        $script:Loaded['Overview'] = $false
        Set-GrpStatus -Text 'Owners applied.'
    } 'Add owners'
})

$btnOwRemove.Add_Click({
    if (-not (Test-GrpLoaded)) { return }
    Invoke-GrpAction {
        if (-not (Test-GrpKind -Allow @('Distribution','Security','Unified') -What 'Editing owners')) { return }
        $picked = Get-GridPicked -Grid $gridOwners
        if ($picked.Count -eq 0) { Set-GrpStatus -Text 'Tick one or more owners to remove.' -Bad; return }

        $keep = @()
        foreach ($row in $gridOwners.Rows) {
            if ([bool]$row.Cells['Pick'].Value) { continue }
            if ($row.Tag) { $keep += "$(if ($row.Tag.Address) { $row.Tag.Address } else { $row.Tag.Raw })" }
        }
        $warn = ''
        if ($keep.Count -eq 0) {
            $warn = "`r`n`r`nTHIS REMOVES THE LAST OWNER. A group with no owner is a real state and a bad one: nobody can approve a membership request, and nothing tells anyone when it goes stale."
        }
        $a = Show-ToolDialog(
            ("Remove $($picked.Count) owner(s) from $($script:Grp.Display)?" + "`r`n`r`n  " +
             (@($picked | ForEach-Object { $_.Name }) -join "`r`n  ") + $warn),
            'Remove owners', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        if ($script:Grp.Kind -eq 'Unified') {
            $lines = @(); $ok = 0; $bad = 0
            foreach ($e in $picked) {
                try {
                    Invoke-EXOCommand -Name 'Remove-UnifiedGroupLinks' -Parameters @{
                        Identity = $script:Grp.Smtp; LinkType = 'Owners'; Links = $e.Address
                        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                    $lines += "OK      $($e.Name)"; $ok++
                } catch { $lines += "FAILED  $($e.Name): $($_.Exception.Message)"; $bad++ }
            }
            Write-GrpAudit -Action 'GroupOwnerRemove' -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
                -Detail "removed owners: $(($picked | ForEach-Object { $_.Address }) -join '; ')"
            if ($bad) { Show-ResultListDialog -Title "Remove owners -- $ok done, $bad failed" -Lines $lines }
        } else {
            Set-GrpOwnerList -Owners $keep
            Write-GrpAudit -Action 'GroupOwnerSet' -Detail "ManagedBy = $(if ($keep) { $keep -join '; ' } else { '(none)' })"
        }
        $script:Grp = Resolve-GroupTarget -Identity $script:Grp.Smtp -ForceOnPrem:($script:Mode -eq 'OnPrem')
        Update-OwnersTab
        $script:Loaded['Overview'] = $false
        Set-GrpStatus -Text "Removed $($picked.Count) owner(s)."
    } 'Remove owners'
})
#endregion

#region Members tab -------------------------------------------------------------
<#
    WHO IS IN IT.

    -ResultSize Unlimited IS LOAD-BEARING. The default is 1000 and it
    truncates SILENTLY -- a group with 1,400 members shows 1,000, and a tool
    that then "removes everyone not in the grid" removes four hundred people
    it never saw.

    ADDS BEFORE REMOVALS, always. A run that fails halfway leaves the group
    with too many members rather than too few, which is the safer way to be
    wrong: nobody loses mail while it is put right.

    A DYNAMIC GROUP HAS NO MEMBER LIST. Its membership is a filter evaluated
    when mail arrives. The filter is shown here, and the preview is what it
    resolves to right now -- not a list anybody can edit.
#>
$mbY = 8
[void](New-TabLabel $tabMembers 'MEMBERS' 8 $mbY -Bold)
$mbY += 22
$mbNote = New-TabLabel $tabMembers '' 8 $mbY 1040 -Muted
$mbNote.Size = New-Object System.Drawing.Size(1040, 32)
$mbY += 40
$gridMembers = New-TabGrid $tabMembers @(@('Name', 360), @('Address', 300), @('Type', 160)) 8 $mbY 700 330
Add-GridPickColumn -Grid $gridMembers
$mbBtnY = $mbY
$btnMbAdd    = New-TabButton $tabMembers 'Add members...' 720 $mbBtnY 170
$mbBtnY += 34
$btnMbRemove = New-TabButton $tabMembers 'Remove ticked'  720 $mbBtnY 170
$mbBtnY += 34
$btnMbAll    = New-TabButton $tabMembers 'Tick all'       720 $mbBtnY 82
$btnMbNone   = New-TabButton $tabMembers 'None'           808 $mbBtnY 82
$mbBtnY += 42
$btnMbRefresh = New-TabButton $tabMembers 'Refresh'       720 $mbBtnY 170
$mbY += 342
$mbFilter = New-TabLabel $tabMembers '' 8 $mbY 1040 -Muted
$mbFilter.Size = New-Object System.Drawing.Size(1040, 64)

function Get-GrpMember {
    $g = $script:Grp
    switch ($g.Kind) {
        'Unified' {
            return @(Invoke-EXOCommand -Name 'Get-UnifiedGroupLinks' -Parameters @{
                        Identity = $g.Smtp; LinkType = 'Members'; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
        }
        'Dynamic' {
            # THE PREVIEW NEEDS THE CONTAINER AS WELL AS THE FILTER. Without
            # -OrganizationalUnit the preview can include recipients the group
            # would never actually deliver to.
            $p = @{ RecipientPreviewFilter = "$($g.Raw.RecipientFilter)"; ResultSize = 'Unlimited'; ErrorAction = 'Stop' }
            if ("$($g.Raw.RecipientContainer)".Trim()) { $p['OrganizationalUnit'] = "$($g.Raw.RecipientContainer)" }
            return @(Invoke-GrpCmd -Name 'Get-Recipient' -Parameters $p)
        }
        default {
            return @(Invoke-GrpCmd -Name 'Get-DistributionGroupMember' -Parameters @{
                        Identity = $g.Smtp; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
        }
    }
}

function Update-MembersTab {
    $gridMembers.Rows.Clear()
    $g = $script:Grp
    $mem = @()
    try { $mem = @(Get-GrpMember) } catch {
        $script:TabReadFailed = $true
        $mbNote.Text = "The members could not be read: $($_.Exception.Message)"
        $mbNote.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
        return
    }
    foreach ($m in $mem) {
        $i = $gridMembers.Rows.Add($false, "$($m.DisplayName)", "$($m.PrimarySmtpAddress)", "$($m.RecipientTypeDetails)")
        $gridMembers.Rows[$i].Tag = [PSCustomObject]@{
            Name = "$($m.DisplayName)"; Address = "$($m.PrimarySmtpAddress)"; Type = "$($m.RecipientTypeDetails)"
        }
        if ("$($m.RecipientTypeDetails)" -match '(?i)Group') {
            # A nested group is worth seeing at a glance: it is where "who
            # actually gets this mail" stops being answerable from one screen.
            $gridMembers.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 90, 150)
        }
    }

    if ($g.Kind -eq 'Dynamic') {
        $btnMbAdd.Enabled = $false; $btnMbRemove.Enabled = $false
        $mbNote.Text = "$($gridMembers.Rows.Count) recipient(s) match the filter RIGHT NOW. This is a preview, not a list -- it is recalculated every time mail arrives, so it can be different tomorrow without anybody changing anything."
        $mbNote.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)
        $mbFilter.Text = ("Filter: $($g.Raw.RecipientFilter)" + "`r`n" +
                          "Container: $(if ("$($g.Raw.RecipientContainer)") { $g.Raw.RecipientContainer } else { '(the whole directory)' })" + "`r`n" +
                          'To change who is in this group, change the filter -- which is not something this window does, because a bad filter silently empties a distribution list.')
    } else {
        $btnMbAdd.Enabled = $true; $btnMbRemove.Enabled = $true
        $nested = @($gridMembers.Rows | Where-Object { $_.Tag -and "$($_.Tag.Type)" -match '(?i)Group' }).Count
        $mbNote.Text = "$($gridMembers.Rows.Count) member(s)$(if ($nested) { ", $nested of them nested groups (shown in blue)" })."
        $mbNote.ForeColor = [System.Drawing.Color]::DimGray
        $mbFilter.Text = 'Every member is read, never the first thousand -- the default page size truncates silently and a partial list is how people get removed from groups they were never seen in.'
    }
}

$btnMbAll.Add_Click({  Set-GridPicked -Grid $gridMembers -On $true })
$btnMbNone.Add_Click({ Set-GridPicked -Grid $gridMembers -On $false })
$btnMbRefresh.Add_Click({
    if (-not (Test-GrpLoaded)) { return }
    Invoke-GrpAction { Update-MembersTab; Set-GrpStatus -Text 'Members re-read.' } 'Refresh members'
})

$btnMbAdd.Add_Click({
    if (-not (Test-GrpLoaded)) { return }
    Invoke-GrpAction {
        if (-not (Test-GrpKind -Allow @('Distribution','Security','Unified') -What 'Adding members')) { return }
        $who = Show-InputDialog -Title 'Add members' -Width 620 `
                -Prompt 'Who should be in this group? Addresses or aliases, separated by semicolons.'
        $list = ConvertFrom-DialogList -Text "$who"
        if ($list.Count -eq 0) { return }

        $resolved = @()
        foreach ($w in $list) {
            $r = Resolve-GrpRecipient -Entry $w
            if (-not $r.Ok) { Set-GrpStatus -Text "'$w' $($r.Reason). Nothing was changed." -Bad; return }
            $resolved += $r
        }
        $a = Show-ToolDialog(("Add $($resolved.Count) member(s) to $($script:Grp.Display)?" + "`r`n`r`n  " +
                              (@($resolved | ForEach-Object { $_.Name }) -join "`r`n  ")),
                             'Add members', 'OKCancel', 'Question')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $lines = @(); $ok = 0; $bad = 0
        foreach ($r in $resolved) {
            try {
                if ($script:Grp.Kind -eq 'Unified') {
                    Invoke-EXOCommand -Name 'Add-UnifiedGroupLinks' -Parameters @{
                        Identity = $script:Grp.Smtp; LinkType = 'Members'; Links = $r.Address
                        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                } else {
                    Invoke-GrpCmd -Name 'Add-DistributionGroupMember' -Parameters @{
                        Identity = $script:Grp.Smtp; Member = $r.Address
                        BypassSecurityGroupManagerCheck = $true; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                }
                $lines += "OK      $($r.Name)"; $ok++
            } catch {
                # ALREADY A MEMBER IS NOT A FAILURE. It is the answer to the
                # question, and reporting it as an error makes people stop
                # reading the result list.
                if ($_.Exception.Message -match '(?i)already a member') { $lines += "SKIP    $($r.Name) is already a member"; $ok++ }
                else { $lines += "FAILED  $($r.Name): $($_.Exception.Message)"; $bad++ }
            }
        }
        Write-GrpAudit -Action 'GroupMemberAdd' -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "added: $(($resolved | ForEach-Object { $_.Address }) -join '; ')"
        Show-ResultListDialog -Title "Add members -- $ok done, $bad failed" -Lines $lines
        Update-MembersTab
        $script:Loaded['Overview'] = $false
        Set-GrpStatus -Text "Added $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Add members'
})

$btnMbRemove.Add_Click({
    if (-not (Test-GrpLoaded)) { return }
    Invoke-GrpAction {
        if (-not (Test-GrpKind -Allow @('Distribution','Security','Unified') -What 'Removing members')) { return }
        $picked = Get-GridPicked -Grid $gridMembers
        if ($picked.Count -eq 0) { Set-GrpStatus -Text 'Tick one or more members to remove.' -Bad; return }
        $a = Show-ToolDialog(("Remove $($picked.Count) member(s) from $($script:Grp.Display)?" + "`r`n`r`n  " +
                              (@($picked | ForEach-Object { $_.Name }) -join "`r`n  ") + "`r`n`r`n" +
                              'They stop receiving anything sent to this group.'),
                             'Remove members', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $lines = @(); $ok = 0; $bad = 0
        foreach ($e in $picked) {
            try {
                if ($script:Grp.Kind -eq 'Unified') {
                    Invoke-EXOCommand -Name 'Remove-UnifiedGroupLinks' -Parameters @{
                        Identity = $script:Grp.Smtp; LinkType = 'Members'; Links = $e.Address
                        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                } else {
                    Invoke-GrpCmd -Name 'Remove-DistributionGroupMember' -Parameters @{
                        Identity = $script:Grp.Smtp; Member = $e.Address
                        BypassSecurityGroupManagerCheck = $true; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                }
                $lines += "OK      $($e.Name)"; $ok++
            } catch {
                if ($_.Exception.Message -match '(?i)not a member') { $lines += "SKIP    $($e.Name) was not a member"; $ok++ }
                else { $lines += "FAILED  $($e.Name): $($_.Exception.Message)"; $bad++ }
            }
        }
        Write-GrpAudit -Action 'GroupMemberRemove' -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "removed: $(($picked | ForEach-Object { $_.Address }) -join '; ')"
        Show-ResultListDialog -Title "Remove members -- $ok done, $bad failed" -Lines $lines
        Update-MembersTab
        $script:Loaded['Overview'] = $false
        Set-GrpStatus -Text "Removed $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Remove members'
})
#endregion

#region Mail Flow tab -----------------------------------------------------------
<#
    WHO MAY SEND TO IT, AND WHETHER THE RULE IS BEING ENFORCED.

    THE TWO SETTINGS ARE ON ONE TAB AND SHARE ONE CONFIRMATION, deliberately.
    Turning sender authentication OFF does not clear the allowed-senders list
    -- it stops the list being ENFORCED. The group goes from restricted to
    open while still LOOKING restricted in the portal, and the two settings
    being edited in different places is how that goes unnoticed.

    AcceptMessagesOnlyFromSendersOrMembers IS REPLACED WHOLE, like ManagedBy.
    Adding one allowed sender means writing all the others back, and getting
    that wrong does not error.
#>
$mfY = 8
[void](New-TabLabel $tabMailFlow 'WHO MAY SEND TO THIS GROUP' 8 $mfY -Bold)
$mfY += 24
$chkRequireAuth = New-Object System.Windows.Forms.CheckBox
$chkRequireAuth.Location = New-Object System.Drawing.Point(8, $mfY)
$chkRequireAuth.Size     = New-Object System.Drawing.Size(700, 20)
$chkRequireAuth.Text     = 'Only authenticated senders (untick to let EXTERNAL senders post to this group)'
$tabMailFlow.Controls.Add($chkRequireAuth)
$mfY += 24
# READ-ONLY ECHO. The editable control is on the Overview tab, where it is in
# the other two editors; this stays so somebody looking at mail flow does not
# find it simply missing.
$chkHidden = New-Object System.Windows.Forms.CheckBox
$chkHidden.Location = New-Object System.Drawing.Point(8, $mfY)
$chkHidden.Size     = New-Object System.Drawing.Size(700, 20)
$chkHidden.Text     = 'Hidden from the address book  (change this on the Overview tab)'
$chkHidden.Enabled  = $false
$tabMailFlow.Controls.Add($chkHidden)
$mfY += 26
$btnMfApply = New-TabButton $tabMailFlow 'Apply these two' 8 $mfY 170
$mfY += 34
$mfWarn = New-TabLabel $tabMailFlow '' 8 $mfY 1040 -Muted
$mfWarn.Size = New-Object System.Drawing.Size(1040, 48)
$mfY += 56

[void](New-TabLabel $tabMailFlow 'ALLOWED SENDERS' 8 $mfY -Bold)
$mfY += 22
$gridAllowed = New-TabGrid $tabMailFlow @(@('Name', 360), @('Address', 300), @('State', 160)) 8 $mfY 700 250
Add-GridPickColumn -Grid $gridAllowed
$mfBtnY = $mfY
$btnAlAdd    = New-TabButton $tabMailFlow 'Add senders...' 720 $mfBtnY 170
$mfBtnY += 34
$btnAlRemove = New-TabButton $tabMailFlow 'Remove ticked'  720 $mfBtnY 170
$mfBtnY += 34
$btnAlAll    = New-TabButton $tabMailFlow 'Tick all'       720 $mfBtnY 82
$btnAlNone   = New-TabButton $tabMailFlow 'None'           808 $mfBtnY 82
$mfY += 262
$mfNote = New-TabLabel $tabMailFlow '' 8 $mfY 1040 -Muted
$mfNote.Size = New-Object System.Drawing.Size(1040, 48)
$mfNote.Text = ('An EMPTY list means anybody may send, subject to the authentication setting above. A populated list means only those senders -- and only while authentication is required.' + "`r`n" +
                'The whole list is rewritten on every change, so a sender shown in amber is carried through unchanged rather than dropped.')

function Update-MailFlowTab {
    $g = $script:Grp
    if ($g.Kind -eq 'Unified') {
        # A Microsoft 365 group keeps these elsewhere and *-DistributionGroup
        # refuses it, so the tab says so rather than showing empty controls.
        $chkRequireAuth.Enabled = $false
        $btnMfApply.Enabled = $false; $btnAlAdd.Enabled = $false; $btnAlRemove.Enabled = $false
        $mfWarn.Text = 'Mail-flow restrictions for a Microsoft 365 group are not edited with the distribution-group cmdlets and are not available here.'
        $mfWarn.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)
        $gridAllowed.Rows.Clear()
        return
    }
    $chkRequireAuth.Enabled = $true
    $btnMfApply.Enabled = $true; $btnAlAdd.Enabled = $true; $btnAlRemove.Enabled = $true

    $gr = $g.Raw
    $chkRequireAuth.Checked = [bool]$gr.RequireSenderAuthenticationEnabled
    $chkHidden.Checked      = [bool]$gr.HiddenFromAddressListsEnabled

    $gridAllowed.Rows.Clear()
    foreach ($a in @($gr.AcceptMessagesOnlyFromSendersOrMembers)) {
        $res = Resolve-GrpRecipient -Entry "$a"
        $i = $gridAllowed.Rows.Add($false, $(if ($res.Ok) { $res.Name } else { "$a" }), $res.Address,
                                   $(if ($res.Ok) { 'ok' } else { 'could not be re-resolved' }))
        $gridAllowed.Rows[$i].Tag = [PSCustomObject]@{ Ok = $res.Ok; Name = $res.Name; Address = $res.Address; Raw = "$a" }
        if (-not $res.Ok) { $gridAllowed.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(190, 110, 0) }
    }

    # THE PAIRED WARNING, EVERY TIME THE TAB IS DRAWN.
    if (-not $chkRequireAuth.Checked -and $gridAllowed.Rows.Count -gt 0) {
        $mfWarn.Text = "THE $($gridAllowed.Rows.Count) ALLOWED SENDER(S) BELOW ARE NOT BEING ENFORCED. Sender authentication is off, so anyone -- inside or outside -- can post to this group, while it still looks restricted."
        $mfWarn.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
    } elseif (-not $chkRequireAuth.Checked) {
        $mfWarn.Text = 'Sender authentication is OFF: anyone outside the organisation can send to this group. Sometimes deliberate, and the usual route for spam to reach everybody at once.'
        $mfWarn.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)
    } else {
        $mfWarn.Text = 'Only authenticated senders may post. The allowed-senders list below narrows that further when it is not empty.'
        $mfWarn.ForeColor = [System.Drawing.Color]::DimGray
    }
}

function Set-GrpAllowedSenders {
    param([string[]]$Senders)
    Invoke-GrpCmd -Name 'Set-DistributionGroup' -Parameters @{
        Identity = $script:Grp.Smtp
        AcceptMessagesOnlyFromSendersOrMembers = @($Senders)
        BypassSecurityGroupManagerCheck = $true
        Confirm = $false
        ErrorAction = 'Stop'
    } | Out-Null
}

$btnAlAll.Add_Click({  Set-GridPicked -Grid $gridAllowed -On $true })
$btnAlNone.Add_Click({ Set-GridPicked -Grid $gridAllowed -On $false })

$btnMfApply.Add_Click({
    if (-not (Test-GrpLoaded)) { return }
    Invoke-GrpAction {
        if (-not (Test-GrpKind -Allow @('Distribution','Security','Dynamic') -What 'Mail-flow settings')) { return }
        $gr = $script:Grp.Raw
        $p = @{ Identity = $script:Grp.Smtp; Confirm = $false; ErrorAction = 'Stop' }
        # -BypassSecurityGroupManagerCheck IS A DISTRIBUTION-GROUP PARAMETER.
        # Set-DynamicDistributionGroup does not declare it, and sending it
        # fails at parameter binding -- so every mail-flow Apply on a dynamic
        # group would have failed before reaching Exchange.
        if ($script:Grp.Kind -ne 'Dynamic') { $p['BypassSecurityGroupManagerCheck'] = $true }
        $changed = @()
        if ($chkRequireAuth.Checked -ne [bool]$gr.RequireSenderAuthenticationEnabled) {
            $p['RequireSenderAuthenticationEnabled'] = [bool]$chkRequireAuth.Checked
            $changed += "RequireSenderAuthentication=$($chkRequireAuth.Checked)"
        }
        if ($changed.Count -eq 0) { Set-GrpStatus -Text 'Nothing changed.'; return }

        $extra = ''
        if (-not $chkRequireAuth.Checked -and $gridAllowed.Rows.Count -gt 0) {
            $extra = "`r`n`r`nTURNING AUTHENTICATION OFF DOES NOT CLEAR THE ALLOWED-SENDERS LIST -- it stops the list being enforced. The group becomes open to everyone while still showing $($gridAllowed.Rows.Count) allowed sender(s)."
        } elseif (-not $chkRequireAuth.Checked) {
            $extra = "`r`n`r`nAnyone outside the organisation will be able to send to this group."
        }
        $cmdlet = $(if ($script:Grp.Kind -eq 'Dynamic') { 'Set-DynamicDistributionGroup' } else { 'Set-DistributionGroup' })
        $a = Show-ToolDialog(("On $($script:Grp.Display):" + "`r`n`r`n  " + ($changed -join "`r`n  ") + $extra),
                             'Apply mail-flow settings', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-GrpCmd -Name $cmdlet -Parameters $p | Out-Null
        Write-GrpAudit -Action 'GroupMailFlow' -Detail ($changed -join '; ')
        $script:Grp = Resolve-GroupTarget -Identity $script:Grp.Smtp -ForceOnPrem:($script:Mode -eq 'OnPrem')
        Update-MailFlowTab
        $script:Loaded['Overview'] = $false
        Set-GrpStatus -Text "Applied: $($changed -join '; ')"
    } 'Apply mail-flow settings'
})

$btnAlAdd.Add_Click({
    if (-not (Test-GrpLoaded)) { return }
    Invoke-GrpAction {
        if (-not (Test-GrpKind -Allow @('Distribution','Security') -What 'Allowed senders')) { return }
        $who = Show-InputDialog -Title 'Add allowed senders' -Width 620 `
                -Prompt 'Who may send to this group? Addresses or aliases, separated by semicolons. A group is allowed.'
        $list = ConvertFrom-DialogList -Text "$who"
        if ($list.Count -eq 0) { return }
        $all = @()
        foreach ($row in $gridAllowed.Rows) { if ($row.Tag) { $all += "$(if ($row.Tag.Address) { $row.Tag.Address } else { $row.Tag.Raw })" } }
        foreach ($w in $list) {
            $r = Resolve-GrpRecipient -Entry $w
            if (-not $r.Ok) { Set-GrpStatus -Text "'$w' $($r.Reason). Nothing was changed." -Bad; return }
            if ($all -notcontains $r.Address) { $all += $r.Address }
        }
        $extra = ''
        if (-not [bool]$script:Grp.Raw.RequireSenderAuthenticationEnabled) {
            $extra = "`r`n`r`nNOTE: sender authentication is currently OFF for this group, so this list will be stored and NOT enforced."
        }
        $a = Show-ToolDialog(("Set the allowed senders on $($script:Grp.Display) to:" + "`r`n`r`n  " + ($all -join "`r`n  ") + $extra),
                             'Allowed senders', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Set-GrpAllowedSenders -Senders $all
        Write-GrpAudit -Action 'GroupAllowedSenders' -Detail "AcceptMessagesOnlyFromSendersOrMembers = $($all -join '; ')"
        $script:Grp = Resolve-GroupTarget -Identity $script:Grp.Smtp -ForceOnPrem:($script:Mode -eq 'OnPrem')
        Update-MailFlowTab
        $script:Loaded['Overview'] = $false
        Set-GrpStatus -Text 'Allowed senders applied.'
    } 'Add allowed senders'
})

$btnAlRemove.Add_Click({
    if (-not (Test-GrpLoaded)) { return }
    Invoke-GrpAction {
        if (-not (Test-GrpKind -Allow @('Distribution','Security') -What 'Allowed senders')) { return }
        $picked = Get-GridPicked -Grid $gridAllowed
        if ($picked.Count -eq 0) { Set-GrpStatus -Text 'Tick one or more senders to remove.' -Bad; return }
        $keep = @()
        foreach ($row in $gridAllowed.Rows) {
            if ([bool]$row.Cells['Pick'].Value) { continue }
            if ($row.Tag) { $keep += "$(if ($row.Tag.Address) { $row.Tag.Address } else { $row.Tag.Raw })" }
        }
        $extra = ''
        if ($keep.Count -eq 0) {
            $extra = "`r`n`r`nTHE LIST WILL BE EMPTY. An empty list means anybody may send, subject only to the authentication setting -- it does not mean nobody may send."
        }
        $a = Show-ToolDialog(("Remove $($picked.Count) allowed sender(s) from $($script:Grp.Display)?" + "`r`n`r`n  " +
                              (@($picked | ForEach-Object { $_.Name }) -join "`r`n  ") + $extra),
                             'Remove allowed senders', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Set-GrpAllowedSenders -Senders $keep
        Write-GrpAudit -Action 'GroupAllowedSenders' -Detail "AcceptMessagesOnlyFromSendersOrMembers = $(if ($keep) { $keep -join '; ' } else { '(empty)' })"
        $script:Grp = Resolve-GroupTarget -Identity $script:Grp.Smtp -ForceOnPrem:($script:Mode -eq 'OnPrem')
        Update-MailFlowTab
        $script:Loaded['Overview'] = $false
        Set-GrpStatus -Text "Removed $($picked.Count) allowed sender(s)."
    } 'Remove allowed senders'
})
#endregion

#region Send As tab -------------------------------------------------------------
<#
    WHO MAY SEND AS THIS GROUP -- AND THE TWO SIDES ARE GENUINELY DIFFERENT.

    ON-PREMISES it is an Active Directory permission:
        Add-ADPermission -Identity <the group's DISTINGUISHED NAME>
                         -User <the trustee's sAMAccountName>
                         -ExtendedRights 'Send-As'
      The *-ADPermission cmdlets are AD OBJECT cmdlets. Their -Identity takes
      a name, canonical name, distinguished name or GUID and NOT an SMTP
      address -- given one it fails with "object 'x@y' couldn't be found",
      which reads as a missing group and is nothing of the kind.

    IN THE CLOUD it is a recipient permission:
        Add-RecipientPermission -Identity <smtp> -Trustee <smtp>
                                -AccessRights 'SendAs'

    NOTE THE TWO SPELLINGS. 'Send-As' with a hyphen on-premises, 'SendAs'
    without one in the cloud. They are not interchangeable and neither errors
    usefully when given the other's spelling.

    A DENY IS KEPT AND FLAGGED. A Deny outranks an Allow, so a list that
    quietly dropped them would explain nothing about why somebody still
    cannot send.
#>
$saY = 8
[void](New-TabLabel $tabAccess 'SEND AS' 8 $saY -Bold)
$saY += 22
$saNote = New-TabLabel $tabAccess '' 8 $saY 1040 -Muted
$saNote.Size = New-Object System.Drawing.Size(1040, 32)
$saY += 40
$gridSendAs = New-TabGrid $tabAccess @(@('Name', 360), @('Address', 300), @('Right', 160)) 8 $saY 700 320
Add-GridPickColumn -Grid $gridSendAs
$saBtnY = $saY
$btnSaAdd    = New-TabButton $tabAccess 'Grant Send As...' 720 $saBtnY 170
$saBtnY += 34
$btnSaRemove = New-TabButton $tabAccess 'Revoke ticked'    720 $saBtnY 170
$saBtnY += 34
$btnSaAll    = New-TabButton $tabAccess 'Tick all'         720 $saBtnY 82
$btnSaNone   = New-TabButton $tabAccess 'None'             808 $saBtnY 82
$saY += 332
$saHint = New-TabLabel $tabAccess '' 8 $saY 1040 -Muted
$saHint.Size = New-Object System.Drawing.Size(1040, 46)
$saHint.Text = ('Send As lets somebody send mail that appears to come FROM the group, with no sign of who actually sent it.' + "`r`n" +
                'A row marked DENY overrides any Allow -- if somebody cannot send as the group despite being listed, look for a Deny first.')

function Update-SendAsTab {
    $gridSendAs.Rows.Clear()
    $g = $script:Grp
    if ($g.Kind -eq 'Unified' -or $g.Kind -eq 'Dynamic') {
        $btnSaAdd.Enabled = $false; $btnSaRemove.Enabled = $false
        $saNote.Text = "Send As is not something this window edits for a $(if ($g.Kind -eq 'Unified') { 'Microsoft 365 group' } else { 'dynamic distribution group' })."
        $saNote.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)
        return
    }
    $btnSaAdd.Enabled = $true; $btnSaRemove.Enabled = $true

    $rows = @()
    try {
        if ($g.OnPrem) {
            # THE DISTINGUISHED NAME, NOT THE ADDRESS.
            foreach ($p in @(Invoke-GrpCmd -Name 'Get-ADPermission' -Parameters @{ Identity = $g.Dn; ErrorAction = 'Stop' })) {
                if ($p.IsInherited) { continue }
                if ("$($p.User)" -match '(?i)NT AUTHORITY\\SELF|\\SELF$|^S-1-5') { continue }
                if (-not (@($p.ExtendedRights) -match 'Send-As')) { continue }
                $rows += [PSCustomObject]@{ Who = "$($p.User)"; Right = $(if ($p.Deny) { 'DENY' } else { 'Allow' }) }
            }
        } else {
            foreach ($p in @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $g.Smtp; ErrorAction = 'Stop' })) {
                if ("$($p.Trustee)" -match '(?i)NT AUTHORITY\\SELF|\\SELF$') { continue }
                if (-not (@($p.AccessRights) -match 'SendAs')) { continue }
                $rows += [PSCustomObject]@{ Who = "$($p.Trustee)"; Right = $(if ("$($p.AccessControlType)" -eq 'Deny') { 'DENY' } else { 'Allow' }) }
            }
        }
    } catch {
        $script:TabReadFailed = $true
        $saNote.Text = "Send As could not be read: $($_.Exception.Message)"
        $saNote.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
        return
    }

    foreach ($r in $rows) {
        $res = Resolve-GrpRecipient -Entry $r.Who
        $i = $gridSendAs.Rows.Add($false, $(if ($res.Ok) { $res.Name } else { $r.Who }), $res.Address, $r.Right)
        $gridSendAs.Rows[$i].Tag = [PSCustomObject]@{ Who = $r.Who; Name = $res.Name; Address = $res.Address; Sam = $res.Sam; Right = $r.Right }
        if ($r.Right -eq 'DENY') { $gridSendAs.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0) }
    }
    $saNote.Text = "$($gridSendAs.Rows.Count) Send As entry/entries, read $(if ($g.OnPrem) { 'from Active Directory' } else { 'from Exchange Online' }). Inherited entries and the group's own SELF entry are not shown."
    $saNote.ForeColor = [System.Drawing.Color]::DimGray
}

$btnSaAll.Add_Click({  Set-GridPicked -Grid $gridSendAs -On $true })
$btnSaNone.Add_Click({ Set-GridPicked -Grid $gridSendAs -On $false })

$btnSaAdd.Add_Click({
    if (-not (Test-GrpLoaded)) { return }
    Invoke-GrpAction {
        if (-not (Test-GrpKind -Allow @('Distribution','Security') -What 'Send As')) { return }
        $who = Show-InputDialog -Title 'Grant Send As' -Width 620 `
                -Prompt 'Who should be able to send as this group? Addresses or aliases, separated by semicolons.'
        $list = ConvertFrom-DialogList -Text "$who"
        if ($list.Count -eq 0) { return }
        $resolved = @()
        foreach ($w in $list) {
            $r = Resolve-GrpRecipient -Entry $w
            if (-not $r.Ok) { Set-GrpStatus -Text "'$w' $($r.Reason). Nothing was changed." -Bad; return }
            if ($script:Grp.OnPrem -and -not "$($r.Sam)".Trim()) {
                Set-GrpStatus -Text "'$w' has no sAMAccountName, and the on-premises Send As permission needs one." -Bad; return
            }
            $resolved += $r
        }
        $a = Show-ToolDialog(
            ("Let these people send AS $($script:Grp.Display)?" + "`r`n`r`n  " +
             (@($resolved | ForEach-Object { $_.Name }) -join "`r`n  ") + "`r`n`r`n" +
             'Mail they send will appear to come from the group, with no sign of who actually sent it.'),
            'Grant Send As', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $lines = @(); $ok = 0; $bad = 0
        foreach ($r in $resolved) {
            try {
                if ($script:Grp.OnPrem) {
                    Invoke-GrpCmd -Name 'Add-ADPermission' -Parameters @{
                        Identity = $script:Grp.Dn; User = $r.Sam; ExtendedRights = 'Send-As'
                        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                } else {
                    Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
                        Identity = $script:Grp.Smtp; Trustee = $r.Address; AccessRights = 'SendAs'
                        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                }
                $lines += "OK      $($r.Name)"; $ok++
            } catch { $lines += "FAILED  $($r.Name): $($_.Exception.Message)"; $bad++ }
        }
        Write-GrpAudit -Action 'GroupSendAsGrant' -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "Send As granted to $(($resolved | ForEach-Object { $_.Address }) -join '; ')"
        Show-ResultListDialog -Title "Grant Send As -- $ok done, $bad failed" -Lines $lines
        Update-SendAsTab
        Set-GrpStatus -Text "Granted $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Grant Send As'
})

$btnSaRemove.Add_Click({
    if (-not (Test-GrpLoaded)) { return }
    Invoke-GrpAction {
        if (-not (Test-GrpKind -Allow @('Distribution','Security') -What 'Send As')) { return }
        $picked = Get-GridPicked -Grid $gridSendAs
        if ($picked.Count -eq 0) { Set-GrpStatus -Text 'Tick one or more entries to revoke.' -Bad; return }
        $deny = @($picked | Where-Object { $_.Right -eq 'DENY' })
        $extra = $(if ($deny.Count) { "`r`n`r`n$($deny.Count) of these are DENY entries. Removing a Deny GRANTS access to whatever Allow was underneath it." } else { '' })
        $a = Show-ToolDialog(("Revoke Send As for $($picked.Count) entry/entries on $($script:Grp.Display)?" + "`r`n`r`n  " +
                              (@($picked | ForEach-Object { "$($_.Name)  ($($_.Right))" }) -join "`r`n  ") + $extra),
                             'Revoke Send As', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $lines = @(); $ok = 0; $bad = 0
        foreach ($e in $picked) {
            try {
                if ($script:Grp.OnPrem) {
                    # -Deny IS NOT OPTIONAL FOR A DENY ROW. Without it the
                    # cmdlet targets the ALLOW entry, reports success, and
                    # leaves the Deny in place -- so the person still cannot
                    # send and the window says it fixed it.
                    $rp = @{
                        Identity = $script:Grp.Dn; User = $(if ($e.Sam) { $e.Sam } else { $e.Who })
                        ExtendedRights = 'Send-As'; Confirm = $false; ErrorAction = 'Stop'
                    }
                    if ($e.Right -eq 'DENY') { $rp['Deny'] = $true }
                    Invoke-GrpCmd -Name 'Remove-ADPermission' -Parameters $rp | Out-Null
                } else {
                    Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{
                        Identity = $script:Grp.Smtp; Trustee = $(if ($e.Address) { $e.Address } else { $e.Who }); AccessRights = 'SendAs'
                        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                }
                $lines += "OK      $($e.Name)"; $ok++
            } catch { $lines += "FAILED  $($e.Name): $($_.Exception.Message)"; $bad++ }
        }
        Write-GrpAudit -Action 'GroupSendAsRevoke' -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "Send As revoked for $(($picked | ForEach-Object { $_.Who }) -join '; ')"
        Show-ResultListDialog -Title "Revoke Send As -- $ok done, $bad failed" -Lines $lines
        Update-SendAsTab
        Set-GrpStatus -Text "Revoked $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Revoke Send As'
})
#endregion

#region Load and tab switching --------------------------------------------------
function Invoke-GrpLoad {
    $t = "$($txtGrp.Text)".Trim()
    if (-not $t) { Set-GrpStatus -Text 'Enter a group address.' -Bad; return }
    Set-GrpStatus -Text "Resolving $t..."
    $g = Resolve-GroupTarget -Identity $t -ForceOnPrem:($OnPrem -or $script:Mode -eq 'OnPrem')
    if (-not $g.Ok) {
        $script:Grp = $null
        $lblIdent.Text = ''; $lblSide.Text = ''
        Set-GrpStatus -Text "'$t' $($g.Reason)" -Bad
        return
    }
    $script:Grp    = $g
    $script:Loaded = @{}
    # THE CONTROLLER IS PINNED ONCE PER LOAD, and only when this group is
    # going to be written on-premises.
    if ($g.OnPrem) { Set-GrpDomainController } else { $script:DCParam = @{} }

    # THE SWITCH FOLLOWS THE OBJECT. A synced group loads in on-premises mode
    # because that is where its writes have to go; the operator can still
    # override it, which is the whole reason the switch exists.
    $script:Mode = $(if ($g.OnPrem) { 'OnPrem' } else { 'Cloud' })
    Sync-EditorModeCombo -Combo $cmbMode -Mode $script:Mode

    Update-GrpIdentity
    $script:TabReadFailed = $false
    Update-OverviewTab
    if (-not $script:TabReadFailed) { $script:Loaded['Overview'] = $true }
    Invoke-GrpTabLoad -Key "$($tabs.SelectedTab.Name)"
    Set-GrpStatus -Text "Loaded $($g.Smtp) $(if ($g.FoundIn -eq 'OnPrem') { 'from Active Directory' } else { 'from Exchange Online' })."

}

function Invoke-GrpTabLoad {
    param([string]$Key)
    if (-not $script:Grp -or -not $script:Grp.Ok) { return }
    if ($script:Loaded[$Key]) { return }
    $script:TabReadFailed = $false
    switch ($Key) {
        'Overview' { Update-OverviewTab }
        'Owners'   { Update-OwnersTab }
        'Members'  { Update-MembersTab }
        'MailFlow' { Update-MailFlowTab }
        'Access'   { Update-SendAsTab }
    }
    if (-not $script:TabReadFailed) { $script:Loaded[$Key] = $true }
}

function Set-GrpMode {
    param([Parameter(Mandatory)][ValidateSet('Cloud', 'OnPrem')][string]$Mode)
    $prev     = $script:Grp
    $prevMode = $script:Mode
    $script:Mode   = $Mode
    $script:Loaded = @{}
    if ($prev -and $prev.Ok) {
        # RE-RESOLVED, NOT RE-FLAGGED. Which cmdlet family answers and which
        # controller is pinned both follow from the side, so the object has to
        # be read again rather than relabelled.
        #
        # A FAILED RE-RESOLVE PUTS EVERYTHING BACK. Resolve-GroupTarget assigns
        # $script:Grp internally before its detail read so the read can be
        # routed -- so a failure leaves that variable pointing at the broken
        # record, and simply not assigning it here was not enough.
        $g = Resolve-GroupTarget -Identity $prev.Smtp -ForceOnPrem:($Mode -eq 'OnPrem')
        if (-not $g.Ok) {
            $script:Grp  = $prev
            $script:Mode = $prevMode
            Sync-EditorModeCombo -Combo $cmbMode -Mode $script:Mode
            Update-GrpIdentity
            Set-GrpStatus -Text "Cannot edit this group $(if ($Mode -eq 'OnPrem') { 'on-premises' } else { 'in the cloud' }): $($g.Reason)" -Bad
            return
        }
        $script:Grp = $g
        if ($g.OnPrem) { Set-GrpDomainController } else { $script:DCParam = @{} }
        # THE OBJECT DECIDES, NOT THE BOX. A Microsoft 365 group forced to
        # on-premises comes back cloud-side whatever was asked for, and a
        # drop-down left saying otherwise is a drop-down that lies about where
        # the next write lands.
        $script:Mode = $(if ($g.OnPrem) { 'OnPrem' } else { 'Cloud' })
        Sync-EditorModeCombo -Combo $cmbMode -Mode $script:Mode
    }
    Update-GrpIdentity
    if ($script:Grp -and $script:Grp.Ok) {
        $script:TabReadFailed = $false
        Update-OverviewTab
        if (-not $script:TabReadFailed) { $script:Loaded['Overview'] = $true }
        Invoke-GrpTabLoad -Key "$($tabs.SelectedTab.Name)"
    }
    Set-GrpStatus -Text $(if ($script:Mode -eq 'OnPrem') { "Editing on-premises. $($script:OnPremSyncNotice)" }
                          elseif ($Mode -eq 'OnPrem') { 'That group has no on-premises half -- still editing in Exchange Online.' }
                          else { 'Editing in Exchange Online.' })
}

$cmbMode.Add_SelectedIndexChanged({
    $want = $(if ($cmbMode.SelectedIndex -eq 1) { 'OnPrem' } else { 'Cloud' })
    if ($script:Busy) { Sync-EditorModeCombo -Combo $cmbMode -Mode $script:Mode; return }
    if ($want -eq $script:Mode) { return }
    Invoke-GrpAction { Set-GrpMode -Mode $want } 'Switch editing mode'
})

$btnLoad.Add_Click({ Invoke-GrpAction { Invoke-GrpLoad } 'Load group' })
$txtGrp.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $btnLoad.PerformClick(); $_.SuppressKeyPress = $true } })

$tabs.Add_SelectedIndexChanged({
    if ($script:Busy) { return }
    $k = "$($tabs.SelectedTab.Name)"
    Invoke-GrpAction { Invoke-GrpTabLoad -Key $k } "Load the $k tab"
})
#endregion

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Group Editor' | Out-Null

$form.Add_Shown({
    # THE REASON GOES ON THE SIDE BANNER, not the status line -- the load on
    # the next line overwrites the status line before anyone can read it.
    $script:OnPremWhy = Set-EditorModeAvailability -Combo $cmbMode
    if ($script:OnPremWhy) {
        $lblSide.Text = "On-premises editing is unavailable: $($script:OnPremWhy)"
        $lblSide.ForeColor = [System.Drawing.Color]::DimGray
    }
    # -OnPrem IS A REQUEST, NOT A GUARANTEE. Honouring it when the probe has
    # just said there is no on-premises session sends every read to a module
    # that is not loaded.
    if ($OnPrem -and $cmbMode.Enabled) { $script:Mode = 'OnPrem'; Sync-EditorModeCombo -Combo $cmbMode -Mode 'OnPrem' }
    if ("$($txtGrp.Text)".Trim()) { $btnLoad.PerformClick() }
})

[void]$form.ShowDialog()
$form.Dispose()
