#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    MailboxPermDupe.ps1
    Copies the access on one mailbox onto another. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    2026-09-16  Folder rows that REPLACE an existing trustee's rights are listed
                apart in the preview with the current rights beside the new
                ones, counted in the Apply dialog, and no longer described as
                'nothing is removed' -- Set-MailboxFolderPermission replaces.
#>
# TOOLKIT-HELP-RUNS: reads the source mailbox, then GRANTS the same access on the target. Preview first; nothing is written until Apply.
# TOOLKIT-HELP-BEGIN
# "Give the new person the same access the old one had."
#
# Reads five kinds of access off a SOURCE mailbox and grants them on a TARGET:
#
#   Full Access        Get-/Add-MailboxPermission
#   Send As            Get-/Add-RecipientPermission
#   Send on Behalf     the GrantSendOnBehalfTo property on the mailbox
#   Inbox              folder permissions on the Inbox
#   Calendar           folder permissions on the Calendar
#
# IT ADDS. IT REMOVES NOTHING OUTRIGHT -- BUT IT CAN REPLACE FOLDER RIGHTS.
#
# Full Access, Send As and Send on Behalf are add-only: a trustee the target
# already has is left exactly as they were. Inbox and Calendar are not, and
# cannot be: a folder permission is one entry per trustee, so giving somebody
# the source's rights REPLACES whatever they had on the target. Copying
# Reviewer onto somebody who had Editor takes access away.
#
# So those rows are listed separately in the preview, under WOULD REPLACE
# EXISTING RIGHTS, with the current rights beside the new ones, and the Apply
# confirmation counts them. The tool does not decide for the operator -- it
# refuses to pretend the write is harmless.
#
# If the target has a delegate the source does not, that delegate is left alone. This is a copy, not a sync -- there is
# no mode that makes the target match the source exactly, deliberately, because
# "make these identical" on a live mailbox removes access somebody is using and
# there is no undo button in this window.
#
# WHAT IT WILL NOT COPY, AND WHY:
#
#   NT AUTHORITY\SELF   every mailbox has it; it means "the owner". Copying it
#                       would try to grant the SOURCE owner access to the
#                       target, which is not what anybody means.
#   Inherited entries   they come from the container, not from a decision made
#                       about this mailbox. The target inherits its own.
#   Deny entries        a deny copied out of context is a support ticket.
#   The target itself   a mailbox does not need Full Access to itself.
#
# FULL ACCESS IS GRANTED WITH AUTOMAPPING OFF. Exchange does not report the
# source mailbox's AutoMapping setting through Get-MailboxPermission, so it
# cannot be copied faithfully. Off is the choice that cannot surprise anyone --
# with it on, a copy silently mounts the mailbox in the Outlook profile of every
# delegate. Add it by hand afterwards if somebody needs it.
#
# PREVIEW SHOWS EVERY GRANT BEFORE ANY OF THEM HAPPEN, and marks the ones that
# are already present on the target so you can see what is actually new.
#
# ONE FAILED GRANT DOES NOT STOP THE REST. Each is applied on its own and
# reported on its own; the run finishes and tells you what did not work.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

#region Reading ---------------------------------------------------------------

function Get-FolderIdentity {
    <#
        The folder to ask about, as Exchange wants it: 'smtp:\<folder name>'.

        MATCHED ON FolderType, NEVER ON NAME. Folder names are LOCALISED -- a
        mailbox created in a French locale has 'Boite de reception' where an
        English one has 'Inbox' -- and a user can rename most of them anyway.
        FolderType is the invariant, and getting this wrong produces "the folder
        does not exist" on a mailbox that plainly has an inbox.

        FolderScope narrows the call first so this is not reading every folder
        in the mailbox to find one.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][ValidateSet('Inbox', 'Calendar')][string]$Kind
    )

    $folders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{
        Identity = $Smtp; FolderScope = $Kind; ErrorAction = 'Stop'
    })
    if ($folders.Count -eq 0) { throw "No $Kind folder found on $Smtp." }

    $primary = $folders | Where-Object { "$($_.FolderType)" -eq $Kind } | Select-Object -First 1
    if (-not $primary) {
        # Fall back to the shallowest folder in scope rather than the first one
        # returned. A sub-calendar sorts alongside the real one and picking it
        # would silently read permissions off the wrong folder.
        $primary = $folders |
            Sort-Object { ("$($_.FolderPath)".Split('/') | Where-Object { $_ }).Count } |
            Select-Object -First 1
    }
    if (-not $primary) { throw "No $Kind folder found on $Smtp." }
    return ($Smtp + ':\' + $primary.Name)
}

function Get-CopyableFolderPermission {
    <#
        Folder permissions worth copying, as plain rows.

        Default and Anonymous are KEPT. They are real settings -- "everyone in
        the company can see free/busy" lives on Default -- and a copy that
        dropped them would leave the target quietly more private than the source
        with nothing on screen to say so. They are applied with Set- rather than
        Add- later; see Set-FolderPermission.
    #>
    param([Parameter(Mandatory)][string]$Smtp,
          [Parameter(Mandatory)][string]$Kind)

    $id = Get-FolderIdentity -Smtp $Smtp -Kind $Kind
    $rows = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{
        Identity = $id; ErrorAction = 'Stop'
    })

    $out = New-Object System.Collections.Generic.List[object]
    foreach ($p in $rows) {
        $who = "$($p.User)"
        if (-not $who) { continue }
        # 'None' is the absence of a permission written down. Copying it grants
        # nothing and produces a row in the report that looks like it did.
        $rights = @($p.AccessRights) -join ', '
        if (-not $rights -or $rights -eq 'None') { continue }
        $out.Add([PSCustomObject]@{
            Kind   = $Kind
            User   = $who
            Rights = $rights
            IsWellKnown = ($who -in @('Default', 'Anonymous'))
        })
    }
    return $out.ToArray()
}

function Get-SourcePermissions {
    <#
        Everything this tool knows how to copy, off one mailbox.

        ERRORS ARE COLLECTED, NOT THROWN. A mailbox whose calendar cannot be
        read should still yield its Full Access list -- a partial answer is
        worth having, and the errors are printed rather than swallowed so a
        blank section can never be mistaken for an empty one (D6).
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        # Optional so the function stays usable without a form attached. Called
        # as & $Progress 'message' <percent>.
        [scriptblock]$Progress,
        [string]$Who = 'mailbox'
    )

    $r = [ordered]@{
        FullAccess   = @()
        SendAs       = @()
        SendOnBehalf = @()
        Inbox        = @()
        Calendar     = @()
        Errors       = New-Object System.Collections.Generic.List[string]
    }

    if ($Progress) { & $Progress "Reading Full Access on the $Who..." 5 }
    try {
        $r.FullAccess = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{
                Identity = $Smtp; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
            Where-Object {
                $_.User -notlike 'NT AUTHORITY\SELF' -and
                $_.IsInherited -eq $false -and
                $_.Deny -eq $false -and
                (@($_.AccessRights) -join ',') -match 'FullAccess'
            } | ForEach-Object { "$($_.User)" })
    } catch { $r.Errors.Add((Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access)')) }

    if ($Progress) { & $Progress "Reading Send As on the $Who..." 20 }
    try {
        # SEND AS IN THE CLOUD IS Get-RecipientPermission, NOT Get-ADPermission.
        # Add-ADPermission exists in both Exchange editions and is on-premises
        # machinery; using it here would fail or, worse, act somewhere else.
        $r.SendAs = @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{
                Identity = $Smtp; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
            Where-Object {
                $_.Trustee -notlike 'NT AUTHORITY\SELF' -and
                $_.AccessControlType -eq 'Allow' -and
                (@($_.AccessRights) -join ',') -match 'SendAs'
            } | ForEach-Object { "$($_.Trustee)" })
    } catch { $r.Errors.Add((Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As)')) }

    if ($Progress) { & $Progress "Reading Send on Behalf on the $Who..." 35 }
    try {
        $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop' }
        $r.SendOnBehalf = @(@($mb.GrantSendOnBehalfTo) | ForEach-Object { "$_" } | Where-Object { $_ })
    } catch { $r.Errors.Add((Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (Send on Behalf)')) }

    # FOLDER READS ARE THE SLOW HALF and get their own steps.
    #
    # Each one is two round trips -- Get-MailboxFolderStatistics to find the
    # folder, then Get-MailboxFolderPermission to read it -- so a mailbox with a
    # large folder tree can sit here for a while. Naming the folder is what
    # turns that wait from "frozen" into "reading the calendar".
    $pct = 50
    foreach ($k in @('Inbox', 'Calendar')) {
        if ($Progress) { & $Progress "Reading $k permissions on the $Who..." $pct }
        try {
            $r[$k] = @(Get-CopyableFolderPermission -Smtp $Smtp -Kind $k)
        } catch { $r.Errors.Add((Format-ErrorDetail -ErrorRecord $_ -Context "Get-MailboxFolderPermission ($k)")) }
        $pct += 25
    }

    return $r
}

function Resolve-MailboxOrNull {
    param([string]$Smtp)
    if (-not "$Smtp".Trim()) { return $null }
    try {
        return Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
            Identity = "$Smtp".Trim(); ErrorAction = 'Stop' }
    } catch { return $null }
}

function Get-TargetState {
    <#
        Is the target usable, and if not, WHY not?

        "Not found" IS THREE DIFFERENT SITUATIONS AND THEY NEED DIFFERENT
        ANSWERS. The first version of this tool collapsed them into one line of
        text, which told the operator the tool had failed without telling them
        what to do about it:

          nothing at that address        a typo, or the object was never made
          a recipient but not a mailbox  a mail user, contact or group -- there
                                         is nothing to grant mailbox rights on
          a mailbox that has not synced  created on-premises minutes ago and not
                                         yet through a directory sync cycle

        The third is the one that matters here. ManageMailObject creates a
        shared mailbox as an on-premises REMOTE mailbox, and the cloud mailbox
        does not exist until a sync cycle has run. Copying permissions onto it
        straight afterwards fails, and "not found" reads as a broken tool rather
        than as "wait for the sync".

        Get-Recipient is the second look, because it finds the objects
        Get-Mailbox will not.
    #>
    param([string]$Smtp)

    $out = [PSCustomObject]@{
        Ok = $false; Mailbox = $null; Reason = ''; Hint = ''
    }
    $t = "$Smtp".Trim()
    if (-not $t) { $out.Reason = 'No target mailbox entered.'; return $out }

    $mb = Resolve-MailboxOrNull -Smtp $t
    if ($mb) { $out.Ok = $true; $out.Mailbox = $mb; return $out }

    $rcp = $null
    try {
        $rcp = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
            Identity = $t; ErrorAction = 'Stop' }
    } catch { $rcp = $null }

    if ($rcp) {
        $out.Reason = "'$t' exists as $($rcp.RecipientTypeDetails), which is not a mailbox."
        $out.Hint   = if ("$($rcp.RecipientTypeDetails)" -match 'MailUser|RemoteUserMailbox|RemoteSharedMailbox') {
                          'It is an on-premises object whose cloud mailbox has not been created yet. Allow a directory sync cycle and try again.'
                      } else {
                          'Mailbox permissions can only be granted on a mailbox. Check the address.'
                      }
    } else {
        $out.Reason = "Nothing was found at '$t'."
        $out.Hint   = 'Check the address. If the mailbox was just created on-premises, it does not reach the cloud until a directory sync cycle has run.'
    }
    return $out
}
#endregion

#region Writing ---------------------------------------------------------------

function Set-FolderPermission {
    <#
        Grant one folder permission, whether or not the trustee already has one.

        ADD- FAILS WHEN AN ENTRY ALREADY EXISTS, and Default and Anonymous
        ALWAYS exist -- every folder is created with both. So a copy that only
        ever calls Add-MailboxFolderPermission fails on the two entries that are
        guaranteed to be there, on every single mailbox.

        Try Add, fall back to Set. Deciding in advance by reading the target's
        existing permissions would be one more call per folder and would still
        race anything else touching the mailbox; letting the cmdlet be the
        authority on whether the entry exists is both cheaper and correct.
    #>
    param(
        [Parameter(Mandatory)][string]$FolderId,
        [Parameter(Mandatory)][string]$User,
        [Parameter(Mandatory)][string]$Rights
    )

    $rightsArray = @("$Rights" -split '\s*,\s*' | Where-Object { $_ })
    try {
        Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' -Parameters @{
            Identity = $FolderId; User = $User; AccessRights = $rightsArray
            ErrorAction = 'Stop' } | Out-Null
        return 'added'
    } catch {
        Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{
            Identity = $FolderId; User = $User; AccessRights = $rightsArray
            ErrorAction = 'Stop' } | Out-Null
        return 'updated'
    }
}

function Add-SendOnBehalf {
    <#
        Add names to GrantSendOnBehalfTo WITHOUT losing what is already there.

        A BARE ASSIGNMENT REPLACES THE WHOLE COLLECTION. Set-Mailbox
        -GrantSendOnBehalfTo <list> is not "add these", it is "this is now the
        list", and anyone already on it is silently dropped. This is the same
        shape as the proxy-address trap that costs an X500 address, and it is
        the reason this reads the current value and writes the union rather than
        writing the source's list straight across.

        Returns the names actually added.
    #>
    param(
        [Parameter(Mandatory)][string]$Target,
        [Parameter(Mandatory)][string[]]$Add
    )

    $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
        Identity = $Target; ErrorAction = 'Stop' }
    $current = @(@($mb.GrantSendOnBehalfTo) | ForEach-Object { "$_" } | Where-Object { $_ })

    $new = @()
    foreach ($a in $Add) {
        if (-not ($current | Where-Object { $_ -eq $a })) { $new += $a }
    }
    if ($new.Count -eq 0) { return @() }

    Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
        Identity = $Target
        GrantSendOnBehalfTo = @($current + $new)
        ErrorAction = 'Stop' } | Out-Null
    return $new
}
#endregion

#region Plan ------------------------------------------------------------------

function New-CopyPlan {
    <#
        Every grant that would be made, as rows, before any of them is made.

        WHY A PLAN OBJECT AND NOT JUST DOING IT: Preview and Apply must not be
        two descriptions of the same intention that can drift apart. They run
        this function and then either print it or execute it, so what was shown
        is what runs.

        Rows carry Skip and Why so the preview can show what is being left out
        and for what reason. A row that silently vanishes between the source
        mailbox and the report is indistinguishable from one the tool failed to
        read (D6).
    #>
    param(
        [Parameter(Mandatory)]$Source,
        [Parameter(Mandatory)][string]$SourceSmtp,
        [Parameter(Mandatory)][string]$TargetSmtp,
        [Parameter(Mandatory)]$TargetPerms,
        [Parameter(Mandatory)][hashtable]$Want
    )

    $plan = New-Object System.Collections.Generic.List[object]
    $t = "$TargetSmtp".ToLower()
    $s = "$SourceSmtp".ToLower()

    $add = {
        param($Kind, $User, $Rights, $Existing, $Current)
        $skip = $false; $why = ''
        $u = "$User".ToLower()
        if ($u -eq $t) { $skip = $true; $why = 'this is the target mailbox itself' }
        elseif ($u -eq $s) { $skip = $true; $why = 'this is the source mailbox itself' }
        elseif ($Existing) { $skip = $true; $why = 'already present on the target' }
        # WHAT THE TARGET HAS RIGHT NOW, carried so the preview can show it.
        # A folder row whose trustee already has DIFFERENT rights is not a
        # grant -- Set-MailboxFolderPermission REPLACES that trustee's rights,
        # so copying Reviewer over an existing Editor takes access away. The
        # operator cannot decide about a difference they were never shown.
        [void]$plan.Add([PSCustomObject]@{
            Kind = $Kind; User = "$User"; Rights = "$Rights"
            Current = "$Current"
            Changes = [bool]("$Current" -and -not $skip)
            Skip = $skip; Why = $why; Result = ''; Error = ''
        })
    }

    if ($Want.FullAccess) {
        $have = @($TargetPerms.FullAccess | ForEach-Object { "$_".ToLower() })
        foreach ($u in @($Source.FullAccess)) {
            & $add 'FullAccess' $u 'FullAccess' ($have -contains "$u".ToLower())
        }
    }
    if ($Want.SendAs) {
        $have = @($TargetPerms.SendAs | ForEach-Object { "$_".ToLower() })
        foreach ($u in @($Source.SendAs)) {
            & $add 'SendAs' $u 'SendAs' ($have -contains "$u".ToLower())
        }
    }
    if ($Want.SendOnBehalf) {
        $have = @($TargetPerms.SendOnBehalf | ForEach-Object { "$_".ToLower() })
        foreach ($u in @($Source.SendOnBehalf)) {
            & $add 'SendOnBehalf' $u 'SendOnBehalf' ($have -contains "$u".ToLower())
        }
    }
    foreach ($k in @('Inbox', 'Calendar')) {
        if (-not $Want[$k]) { continue }
        $have = @{}
        foreach ($p in @($TargetPerms[$k])) { $have["$($p.User)".ToLower()] = "$($p.Rights)" }
        foreach ($p in @($Source[$k])) {
            $key = "$($p.User)".ToLower()
            # An existing entry with DIFFERENT rights is not "already present" --
            # it is a difference the operator should see and decide about, so it
            # stays in the plan and Set- will update it.
            $same = $have.ContainsKey($key) -and $have[$key] -eq "$($p.Rights)"
            $cur  = $(if ($have.ContainsKey($key)) { $have[$key] } else { '' })
            & $add $k $p.User $p.Rights $same $cur
        }
    }
    return $plan.ToArray()
}
#endregion

#region UI --------------------------------------------------------------------

function Show-PermDupeForm {
    param([string]$Seed)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Mailbox Permission Copy -- source to target'
    $form.ClientSize = New-Object System.Drawing.Size(880, 640)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12

    $lblSrc = New-Object System.Windows.Forms.Label
    $lblSrc.Text = 'Source mailbox (copy access FROM):'
    $lblSrc.Location = New-Object System.Drawing.Point(12, $y)
    $lblSrc.Size = New-Object System.Drawing.Size(260, 18)
    $form.Controls.Add($lblSrc)

    $txtSrc = New-Object System.Windows.Forms.TextBox
    $txtSrc.Location = New-Object System.Drawing.Point(280, ($y - 3))
    $txtSrc.Size = New-Object System.Drawing.Size(320, 24)
    $txtSrc.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtSrc.Text = "$Seed"
    $form.Controls.Add($txtSrc)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Load source'
    $btnLoad.Location = New-Object System.Drawing.Point(614, ($y - 5))
    $btnLoad.Size = New-Object System.Drawing.Size(120, 28)
    $form.Controls.Add($btnLoad)
    $y += 34

    $lblTgt = New-Object System.Windows.Forms.Label
    $lblTgt.Text = 'Target mailbox (grant access ON):'
    $lblTgt.Location = New-Object System.Drawing.Point(12, $y)
    $lblTgt.Size = New-Object System.Drawing.Size(260, 18)
    $form.Controls.Add($lblTgt)

    $txtTgt = New-Object System.Windows.Forms.TextBox
    $txtTgt.Location = New-Object System.Drawing.Point(280, ($y - 3))
    $txtTgt.Size = New-Object System.Drawing.Size(320, 24)
    $txtTgt.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtTgt)
    $y += 34

    $grp = New-Object System.Windows.Forms.GroupBox
    $grp.Text = 'Copy which permissions'
    $grp.Location = New-Object System.Drawing.Point(12, $y)
    $grp.Size = New-Object System.Drawing.Size(856, 52)
    $form.Controls.Add($grp)

    $chk = [ordered]@{}
    $labels = [ordered]@{
        FullAccess   = 'Full Access'
        SendAs       = 'Send As'
        SendOnBehalf = 'Send on Behalf'
        Inbox        = 'Inbox'
        Calendar     = 'Calendar'
    }
    $cx = 12
    foreach ($k in $labels.Keys) {
        $c = New-Object System.Windows.Forms.CheckBox
        $c.Text = $labels[$k]
        $c.Checked = $true
        $c.Location = New-Object System.Drawing.Point($cx, 20)
        $c.Size = New-Object System.Drawing.Size(150, 22)
        $grp.Controls.Add($c)
        $chk[$k] = $c
        $cx += 160
    }
    $y += 62

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, $y)
    $lblStatus.Size = New-Object System.Drawing.Size(856, 18)
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.Text = 'Load the source, then Preview. Nothing is written until Apply.'
    $form.Controls.Add($lblStatus)
    $y += 22

    # WHAT THE BAR IS ACTUALLY FOR HERE.
    #
    # Every call this tool makes runs on the UI thread, so the window stops
    # redrawing while Exchange is thinking. That is indistinguishable from a
    # hang, and "I think it froze, but I'm not sure" is the worst state to leave
    # an operator in -- they either wait too long or kill a tool mid-write.
    #
    # It cannot make a slow call fast, and it does NOT repaint DURING one. What
    # it does is name the step that is in progress, so a stall is attributable:
    # "Reading Send As" sitting still for a minute is information, a frozen grey
    # rectangle is not.
    $pb = Add-ReportProgressBar -Form $form -X 12 -Y $y -Width 856 -Height 14
    $y += 26

    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = 'Preview'
    $btnPreview.Location = New-Object System.Drawing.Point(12, $y)
    $btnPreview.Size = New-Object System.Drawing.Size(130, 30)
    $btnPreview.Enabled = $false
    $form.Controls.Add($btnPreview)

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = 'Apply'
    $btnApply.Location = New-Object System.Drawing.Point(150, $y)
    $btnApply.Size = New-Object System.Drawing.Size(130, 30)
    $btnApply.Enabled = $false
    $form.Controls.Add($btnApply)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Location = New-Object System.Drawing.Point(738, $y)
    $btnClose.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })
    $y += 40

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(12, $y)
    $txtOut.Size = New-Object System.Drawing.Size(856, 380)
    $txtOut.Multiline = $true
    $txtOut.ScrollBars = 'Vertical'
    $txtOut.WordWrap = $true
    $txtOut.ReadOnly = $true
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)

    # State lives on the form, not in locals: these are read by handlers that
    # run long after this function has returned.
    $form | Add-Member -NotePropertyName 'SourcePerms' -NotePropertyValue $null -Force
    $form | Add-Member -NotePropertyName 'SourceSmtp'  -NotePropertyValue ''   -Force
    # The plan Preview built, plus a fingerprint of the inputs it was built
    # from. See $buildPlan for why this is worth carrying.
    $form | Add-Member -NotePropertyName 'Plan'        -NotePropertyValue $null -Force
    $form | Add-Member -NotePropertyName 'PlanKey'     -NotePropertyValue ''   -Force

    <#
        RE-ENTRANCY GUARD. DISABLING THE BUTTONS IS NOT ENOUGH.

        Clicking a window that is mid-call is the normal thing to do when you
        cannot tell whether it is working -- Windows paints "(Not Responding)"
        over it, which invites exactly that. The click is QUEUED, and it is
        dispatched the next time anything pumps the message queue.

        Two things in this tool pump it. MessageBox.Show runs its own message
        loop, so the confirmation dialog on Apply will deliver a click that was
        made seconds earlier. And ShowDialog is pumping the whole time anyway.

        The buttons are disabled during a run, but a click landing in the gap
        before that takes effect still arrives, and re-entering a handler that is
        halfway through applying permissions is how a run ends up partially
        done. This flag closes the gap: every handler refuses outright if
        another one is already inside.
    #>
    $form | Add-Member -NotePropertyName 'Busy'        -NotePropertyValue $false -Force

    $wanted = {
        $h = @{}
        foreach ($k in @($chk.Keys)) { $h[$k] = $chk[$k].Checked }
        return $h
    }

    $describe = {
        param($Src)
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('SOURCE PERMISSIONS')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('  Full Access    : ' + $(if (@($Src.FullAccess).Count) { @($Src.FullAccess) -join ', ' } else { '(none)' }))
        [void]$sb.AppendLine('  Send As        : ' + $(if (@($Src.SendAs).Count) { @($Src.SendAs) -join ', ' } else { '(none)' }))
        [void]$sb.AppendLine('  Send on Behalf : ' + $(if (@($Src.SendOnBehalf).Count) { @($Src.SendOnBehalf) -join ', ' } else { '(none)' }))
        foreach ($k in @('Inbox', 'Calendar')) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("  $k")
            if (@($Src[$k]).Count -eq 0) {
                [void]$sb.AppendLine('    (none)')
            } else {
                foreach ($p in @($Src[$k])) {
                    [void]$sb.AppendLine('    ' + $p.User.PadRight(40) + $p.Rights)
                }
            }
        }
        if (@($Src.Errors).Count) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('  COULD NOT BE READ -- these sections are unknown, not empty:')
            foreach ($e in @($Src.Errors)) { [void]$sb.AppendLine('    ' + $e) }
        }
        return $sb.ToString()
    }

    $btnLoad.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnPreview.Enabled = $false; $btnApply.Enabled = $false
            $form.SourcePerms = $null

            $s = "$($txtSrc.Text)".Trim()
            if (-not $s) { $lblStatus.Text = 'Enter a source mailbox.'; return }
            $mb = Resolve-MailboxOrNull -Smtp $s
            if (-not $mb) { $txtOut.Text = "Source mailbox '$s' was not found."; $lblStatus.Text = 'Source not found.'; return }

            $resolved = "$($mb.PrimarySmtpAddress)"
            $form.SourceSmtp = $resolved

            # NO -AllowInput. New-ProgressReporter defaults to repainting the
            # label and bar rather than pumping the message queue, and that
            # default is load-bearing: Application::DoEvents dispatches queued
            # INPUT, not just paints, so a stray click during a read can
            # re-enter this handler or dispose the form mid-loop. That is the
            # crash Disable-User had. There is no Cancel button here, so there
            # is nothing that needs input pumped.
            $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Reading source'
            $src = Get-SourcePermissions -Smtp $resolved -Progress $p -Who 'source'
            & $p 'Source read.' 100
            $form.SourcePerms = $src

            $txtOut.Text = "Source: $resolved  ($($mb.DisplayName))" + [Environment]::NewLine +
                           [Environment]::NewLine + (& $describe $src)
            $lblStatus.Text = 'Source loaded. Enter a target and press Preview.'

            # BOTH BUTTONS GO LIVE TOGETHER, matching ManageMailObject.
            #
            # Apply used to stay disabled until a Preview found work, which
            # meant that when the target was wrong the operator pressed a dead
            # button and got NOTHING -- no dialog, no text, no clue. A disabled
            # control cannot explain itself. Apply now runs the same checks as
            # Preview and reports what it found, which is the only way the
            # reason reaches the person who needs it.
            $btnPreview.Enabled = $true
            $btnApply.Enabled = $true
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Loading the source mailbox'
            $lblStatus.Text = 'Source could not be read.'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    $buildPlan = {
        <#
            PREVIEW THEN APPLY USED TO READ THE TARGET TWICE.

            Both handlers called this, and each call is roughly seven round
            trips to Exchange -- three for the mailbox-level rights and two per
            folder, because a folder read is Get-MailboxFolderStatistics to find
            it and then Get-MailboxFolderPermission to read it. Doing the
            obvious thing (Preview, look, Apply) therefore paid for the slow
            half twice, on the UI thread, which is most of what "it froze" was.

            So Preview's plan is kept, along with a fingerprint of what it was
            built from. Apply reuses it when nothing has changed and rebuilds
            when anything has.

            THE FINGERPRINT MUST COVER EVERY INPUT. If it missed one -- the
            checkboxes, say -- Apply would silently write the plan for a
            DIFFERENT set of permissions than the one on screen. Stale-cache
            bugs on a write path are worse than the slowness they fix, which is
            why this rebuilds on any doubt rather than trying to be clever about
            what changed.
        #>
        param([switch]$UseCache)

        $s = $form.SourcePerms
        if (-not $s) { return $null }
        # EVERY REFUSAL WRITES TO THE PREVIEW PANE, NOT JUST THE STATUS LINE.
        #
        # The empty-target case used to set only $lblStatus -- an 18px grey
        # label above a large output box that still held the last successful
        # result. The tool looked like it had ignored the click. If the answer
        # is "I am not going to do this", it belongs where the operator is
        # already looking.
        $t = "$($txtTgt.Text)".Trim()
        $state = Get-TargetState -Smtp $t
        if (-not $state.Ok) {
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('TARGET NOT USABLE -- nothing has been changed')
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('  ' + $state.Reason)
            if ($state.Hint) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine('  ' + $state.Hint)
            }
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('  The source is still loaded. Fix the target and press Preview again.')
            $txtOut.Text = $sb.ToString()
            $lblStatus.Text = 'Target not usable -- see the pane below.'
            return $null
        }
        $tmb = $state.Mailbox
        $tSmtp = "$($tmb.PrimarySmtpAddress)"

        # Source, target and every checkbox. Anything that changes the plan has
        # to change the key.
        $w = & $wanted
        $key = ($form.SourceSmtp + '|' + $tSmtp + '|' +
                (@($w.Keys | Sort-Object | ForEach-Object { "$_=$($w[$_])" }) -join ','))
        if ($UseCache -and $form.Plan -and $form.PlanKey -eq $key) {
            $lblStatus.Text = 'Using the plan from Preview -- the target has not been re-read.'
            return $form.Plan
        }

        $planProgress = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Reading target'
        if ($tSmtp -eq $form.SourceSmtp) {
            $txtOut.Text = 'Source and target are the same mailbox. Nothing to do.'
            $lblStatus.Text = 'Source and target are the same.'
            return $null
        }
        $tp = Get-SourcePermissions -Smtp $tSmtp -Progress $planProgress -Who 'target'
        & $planProgress 'Comparing source against target...' 90
        $plan = [PSCustomObject]@{
            TargetSmtp = $tSmtp
            TargetName = "$($tmb.DisplayName)"
            TargetPerms = $tp
            Rows = @(New-CopyPlan -Source $s -SourceSmtp $form.SourceSmtp `
                        -TargetSmtp $tSmtp -TargetPerms $tp -Want $w)
        }
        $form.Plan = $plan
        $form.PlanKey = $key
        return $plan
    }

    $btnPreview.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $p = & $buildPlan
            if (-not $p) { return }

            $todo = @($p.Rows | Where-Object { -not $_.Skip })
            $skip = @($p.Rows | Where-Object { $_.Skip })

            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('PREVIEW -- nothing has been changed')
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("  From : $($form.SourceSmtp)")
            [void]$sb.AppendLine("  To   : $($p.TargetSmtp)  ($($p.TargetName))")
            [void]$sb.AppendLine('')
            $grants  = @($todo | Where-Object { -not $_.Changes })
            $changes = @($todo | Where-Object { $_.Changes })
            [void]$sb.AppendLine("WOULD GRANT  ($($grants.Count))")
            if ($grants.Count -eq 0) {
                [void]$sb.AppendLine('  nothing -- the target already has everything selected')
            } else {
                foreach ($r in $grants) {
                    [void]$sb.AppendLine('  ' + $r.Kind.PadRight(14) + $r.User.PadRight(40) + $r.Rights)
                }
            }
            # LISTED APART, BECAUSE THESE ARE NOT GRANTS. Setting a folder
            # permission replaces the trustee's rights outright, so a row here
            # can REDUCE what somebody already has on the target. Showing the
            # current rights beside the new ones is the whole point.
            if ($changes.Count) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine("WOULD REPLACE EXISTING RIGHTS  ($($changes.Count))  -- this can REMOVE access")
                foreach ($r in $changes) {
                    [void]$sb.AppendLine('  ' + $r.Kind.PadRight(14) + $r.User.PadRight(40) +
                                         "now: $($r.Current)".PadRight(28) + "-> $($r.Rights)")
                }
            }
            if ($skip.Count) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine("SKIPPED  ($($skip.Count))")
                foreach ($r in $skip) {
                    [void]$sb.AppendLine('  ' + $r.Kind.PadRight(14) + $r.User.PadRight(40) + $r.Why)
                }
            }
            if (@($p.TargetPerms.Errors).Count) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine('  TARGET COULD NOT BE FULLY READ -- "already present" may be understated:')
                foreach ($e in @($p.TargetPerms.Errors)) { [void]$sb.AppendLine('    ' + $e) }
            }
            $txtOut.Text = $sb.ToString()
            $lblStatus.Text = $(if ($changes.Count) {
                "$($grants.Count) grant(s) and $($changes.Count) replacement(s) would be made. Apply writes them."
            } else {
                "$($todo.Count) grant(s) would be made. Apply writes them."
            })
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Building the preview'
            $lblStatus.Text = 'Preview failed.'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    $btnApply.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            # -UseCache: reuse the plan Preview built if source, target and the
            # checkboxes are all unchanged. In the normal Preview-then-Apply
            # flow this removes an entire re-read of the target -- about seven
            # round trips, on the UI thread, which is where a lot of the
            # apparent freeze came from.
            $p = & $buildPlan -UseCache
            if (-not $p) { return }
            $todo = @($p.Rows | Where-Object { -not $_.Skip })
            if ($todo.Count -eq 0) {
                # Says it in the pane, not just the status strip. "Nothing to
                # do" and "the button is broken" look identical when the only
                # difference is 18px of grey text above an unchanged output box.
                $txtOut.Text = ('NOTHING TO APPLY' + [Environment]::NewLine + [Environment]::NewLine +
                    "  Target: $($p.TargetSmtp)" + [Environment]::NewLine + [Environment]::NewLine +
                    '  Every permission selected is already present on the target,' + [Environment]::NewLine +
                    '  or was skipped. Press Preview to see the reason for each one.')
                $lblStatus.Text = 'Nothing to apply -- see the pane below.'
                return
            }

            $chg = @($todo | Where-Object { $_.Changes })
            $tail = $(if ($chg.Count) {
                "$($chg.Count) of them REPLACE rights the target already has, which can reduce access. See WOULD REPLACE EXISTING RIGHTS in the pane below."
            } else {
                'Nothing on the target is removed.'
            })
            $answer = Show-ToolDialog(
                ("Write $($todo.Count) permission(s) on:" + "`r`n`r`n" +
                 "  $($p.TargetSmtp)" + "`r`n`r`n" +
                 "copied from:" + "`r`n  $($form.SourceSmtp)" + "`r`n`r`n" +
                 $tail),
                'Confirm', 'OKCancel', $(if ($chg.Count) { 'Warning' } else { 'Question' }))
            if ($answer -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnApply.Enabled = $false; $btnPreview.Enabled = $false; $btnLoad.Enabled = $false

            $ok = 0; $failed = 0
            $folderIds = @{}

            # THE ONLY PLACE IN THIS TOOL WITH A REAL PERCENTAGE.
            #
            # The reads are a fixed handful of calls whose durations are
            # unknowable, so their numbers are milestones rather than
            # measurements. Here the denominator is genuine -- n grants, one at a
            # time -- and the number means what it looks like it means.
            $applyProgress = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Applying'
            $i = 0
            $total = $todo.Count

            foreach ($r in $todo) {
                $i++
                & $applyProgress "Granting $($r.Kind) to $($r.User)  ($i of $total)" `
                    ([int](100 * $i / [Math]::Max(1, $total)))
                # EVERY GRANT IN ITS OWN TRY.
                #
                # A batch that stops at the first failure leaves the target
                # half-done with no record of where it got to -- which is the
                # inbox-rule bug from Disable-User, and it is worse here because
                # the half that ran is a permission change.
                try {
                    switch ($r.Kind) {
                        'FullAccess' {
                            # InheritanceType 'All' matches Grant-FullAccessPermission
                            # in MailboxPermissions.ps1 -- the same grant made by
                            # two tools should be the same grant.
                            #
                            # AutoMapping IS OFF, AND THAT IS A CHOICE.
                            # Get-MailboxPermission does not report the source's
                            # AutoMapping setting -- it lives on
                            # msExchDelegateListLink, not on the ACE -- so it
                            # cannot be copied faithfully whatever this does.
                            # Off is the answer that cannot surprise anybody: a
                            # bulk copy with it on silently mounts the mailbox in
                            # the Outlook profile of every delegate.
                            Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
                                Identity = $p.TargetSmtp; User = $r.User
                                AccessRights = 'FullAccess'; InheritanceType = 'All'
                                AutoMapping = $false
                                Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                            $r.Result = 'granted'
                        }
                        'SendAs' {
                            Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
                                Identity = $p.TargetSmtp; Trustee = $r.User
                                AccessRights = 'SendAs'
                                Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                            $r.Result = 'granted'
                        }
                        'SendOnBehalf' {
                            # Applied ONE AT A TIME through the union helper.
                            # Batching them would be fewer calls and would make a
                            # single bad name fail the whole set.
                            $added = Add-SendOnBehalf -Target $p.TargetSmtp -Add @($r.User)
                            $r.Result = $(if (@($added).Count) { 'granted' } else { 'already present' })
                        }
                        default {
                            if (-not $folderIds.ContainsKey($r.Kind)) {
                                $folderIds[$r.Kind] = Get-FolderIdentity -Smtp $p.TargetSmtp -Kind $r.Kind
                            }
                            $r.Result = Set-FolderPermission -FolderId $folderIds[$r.Kind] `
                                            -User $r.User -Rights $r.Rights
                        }
                    }
                    $ok++
                } catch {
                    $failed++
                    $r.Result = 'FAILED'
                    $r.Error = "$($_.Exception.Message)"
                }
            }

            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine($(if ($failed) { 'APPLIED WITH FAILURES' } else { 'APPLIED' }))
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("  From : $($form.SourceSmtp)")
            [void]$sb.AppendLine("  To   : $($p.TargetSmtp)")
            [void]$sb.AppendLine('')
            foreach ($r in $todo) {
                [void]$sb.AppendLine('  ' + $r.Kind.PadRight(14) + $r.User.PadRight(40) + $r.Result)
                if ($r.Error) { [void]$sb.AppendLine('      ' + $r.Error) }
            }
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("  $ok succeeded, $failed failed.")
            if ($failed) {
                [void]$sb.AppendLine('  The ones that succeeded are in place. Re-running skips them.')
            }
            $txtOut.Text = $sb.ToString()
            $lblStatus.Text = "$ok granted, $failed failed."

            $result = if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }
            Write-CompletedLog -Action 'CopyMailboxPermissions' `
                -Target (Get-LogTarget $p.TargetSmtp) -Result $result `
                -Detail "Copied from $($form.SourceSmtp): $ok granted, $failed failed"
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Applying the permission copy'
            $lblStatus.Text = 'Apply failed.'
            Write-FailedLog -Action 'CopyMailboxPermissions' -Target "$($txtTgt.Text)" -ErrorRecord $_
        } finally {
            # ALL THREE COME BACK, including Apply.
            #
            # The three were disabled together at the top of the run and only
            # two were restored, so the second Apply of a session pressed a dead
            # button -- the same "nothing happens and nothing says why" the rest
            # of this change is removing. Re-running Apply is a legitimate thing
            # to want: the first run may have partially failed, and a re-run
            # skips whatever already succeeded.
            $btnLoad.Enabled = $true
            $btnPreview.Enabled = $true
            $btnApply.Enabled = $true
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false

            # THE PLAN IS DROPPED AFTER AN APPLY, ALWAYS.
            #
            # It described the target as it was BEFORE these grants. Re-using it
            # for a second Apply would offer to make grants that have just been
            # made, and -- worse -- would report "already present" counts from a
            # stale reading. A re-run after a partial failure is a legitimate
            # thing to do and it must re-read.
            $form.Plan = $null
            $form.PlanKey = ''
        }
    })

    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Mailbox Permission Copy' | Out-Null

    [void]$form.ShowDialog()
}
#endregion

Show-PermDupeForm -Seed (Get-TargetSmtp -Text $SmtpAddress)
