#Requires -Version 5.1
<#
    Automap-Fix.ps1
    Remove Full Access, wait, re-add it -- the standard fix for a shared
    mailbox that will not automap into Outlook.

    THE REMOVE AND THE RE-ADD ARE ONE OPERATION. If the remove succeeds and the
    re-add fails, the user has LESS access than before this ran. Every failure
    path below says so in those words, because a half-finished run here is worse
    than not running at all.

    -Environment IS ACCEPTED AND USED, and passed to every Write-*Log call.
#>
# TOOLKIT-HELP-RUNS: WRITES mailbox permissions in Exchange Online. Removes and re-adds Full Access.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Fix a shared mailbox that will not appear in somebody's Outlook.
#
# It removes their Full Access, waits, and adds it back. That is the fix,
# because AUTOMAPPING IS DECIDED WHEN THE PERMISSION IS GRANTED and never
# revisited -- if it did not take the first time, nothing will change it except
# granting it again.
#
# THIS BRIEFLY REMOVES THEIR ACCESS. For about ten seconds they have none. If
# they are working in the mailbox right now, do it later.
#
# IT IS NOT INSTANT AFTERWARDS. Outlook picks up the new mapping when it next
# reconnects -- usually minutes, sometimes at the next restart. A user watching
# their Outlook while you press the button will see nothing happen and conclude
# it failed. Tell them to close and reopen Outlook.
#
# AUTOMAPPING ONLY WORKS WHEN THE PERMISSION IS ON THE PERSON. Access granted
# through a group never automaps, no matter how many times this is run. If they
# get Full Access via a group, this tool cannot help and the answer is a direct
# grant or a manually added mailbox in their profile.
#
# WHAT IT DOES NOT FIX: an Outlook profile that has the mailbox added manually
# already, a cached mode problem, or Send As. Only automapping.
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

# TEN SECONDS, and the number is not arbitrary. It is long enough for the
# removal to settle before the grant is written, which is the whole point of
# doing them as two steps rather than one Set-.
$script:WaitSeconds = 10

$script:Mbx = $null
$script:Usr = $null

#region Lookup ------------------------------------------------------------------
function Resolve-One {
    <#
        Resolve a typed string to exactly one recipient.

        EXACTLY ONE, OR IT IS AN ERROR. This tool writes permissions, and
        resolving "smith" to the first of four Smiths would grant a stranger
        access to a mailbox.
    #>
    param([Parameter(Mandatory)][string]$Identity, [switch]$MailboxOnly)

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Item = $null }
    if (-not "$Identity".Trim()) { $r.Reason = 'Nothing entered.'; return $r }

    $name = if ($MailboxOnly) { 'Get-Mailbox' } else { 'Get-Recipient' }
    $found = @()
    try {
        $found = @(Invoke-EXOCommand -Name $name -Parameters @{ Identity = $Identity.Trim(); ErrorAction = 'Stop' })
    } catch { $found = @() }

    if (@($found).Count -eq 0) {
        $r.Reason = "'$Identity' did not resolve. Use the alias or the full address -- this tool does not guess from a partial name, because granting mailbox access to the wrong person is not a recoverable mistake."
        return $r
    }
    if (@($found).Count -gt 1) {
        $r.Reason = "'$Identity' matched $(@($found).Count) objects. Be exact."
        return $r
    }
    $r.Item = $found[0]
    $r.Ok = $true
    return $r
}

function Get-FullAccessState {
    <#
        Does this user already hold Full Access, and is it automapped?

        AUTOMAPPING IS NOT READABLE FROM Get-MailboxPermission. Exchange Online
        does not return it, so the tool cannot show the current state -- and
        this function does not pretend to. It reports whether the permission
        EXISTS, which is what decides whether there is anything to fix.

        INHERITED AND SELF ENTRIES ARE EXCLUDED. They are not grants anybody
        made and removing one is not something this tool should attempt.
    #>
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][string]$User)

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Has = $false; Entries = @() }
    try {
        $perms = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{
            Identity = $Mailbox; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
            Where-Object {
                -not $_.IsInherited -and
                "$($_.User)" -notlike 'NT AUTHORITY\*' -and
                "$($_.AccessRights)" -like '*FullAccess*'
            })
        $r.Entries = @($perms | Where-Object { "$($_.User)" -eq $User -or "$($_.User)" -like "*$User*" })
        $r.Has = (@($r.Entries).Count -gt 0)
        $r.Ok = $true
    } catch {
        $r.Reason = Format-ErrorDetail -ErrorRecord $_ -Context "Get-MailboxPermission $Mailbox"
    }
    return $r
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Automap Fix'
$form.ClientSize      = New-Object System.Drawing.Size(920, 620)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$grp          = New-Object System.Windows.Forms.GroupBox
$grp.Text     = 'Who'
$grp.Location = New-Object System.Drawing.Point(14, 46)
$grp.Size     = New-Object System.Drawing.Size(892, 128)
$grp.Anchor   = 'Top,Left,Right'
[void]$form.Controls.Add($grp)

$lblMbx          = New-Object System.Windows.Forms.Label
$lblMbx.Text     = 'Shared mailbox'
$lblMbx.Location = New-Object System.Drawing.Point(14, 29)
$lblMbx.Size     = New-Object System.Drawing.Size(120, 20)
[void]$grp.Controls.Add($lblMbx)

$txtMbx          = New-Object System.Windows.Forms.TextBox
$txtMbx.Location = New-Object System.Drawing.Point(140, 26)
$txtMbx.Size     = New-Object System.Drawing.Size(320, 22)
[void]$grp.Controls.Add($txtMbx)

$lblUsr          = New-Object System.Windows.Forms.Label
$lblUsr.Text     = 'User'
$lblUsr.Location = New-Object System.Drawing.Point(14, 59)
$lblUsr.Size     = New-Object System.Drawing.Size(120, 20)
[void]$grp.Controls.Add($lblUsr)

$txtUsr          = New-Object System.Windows.Forms.TextBox
$txtUsr.Location = New-Object System.Drawing.Point(140, 56)
$txtUsr.Size     = New-Object System.Drawing.Size(320, 22)
[void]$grp.Controls.Add($txtUsr)

$lblTip           = New-Object System.Windows.Forms.Label
$lblTip.Text      = 'Alias or full address. Both must resolve to exactly one object -- this writes permissions, so it will not guess.'
$lblTip.Location  = New-Object System.Drawing.Point(140, 84)
$lblTip.Size      = New-Object System.Drawing.Size(600, 32)
$lblTip.ForeColor = [System.Drawing.Color]::Gray
$lblTip.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grp.Controls.Add($lblTip)

$btnCheck           = New-Object System.Windows.Forms.Button
$btnCheck.Text      = 'Check'
$btnCheck.Location  = New-Object System.Drawing.Point(480, 24)
$btnCheck.Size      = New-Object System.Drawing.Size(120, 26)
$btnCheck.FlatStyle = 'Flat'
[void]$grp.Controls.Add($btnCheck)

# DISABLED UNTIL CHECK HAS RUN. The fix removes access, and it must not be
# possible to remove access from a pair of names nobody has confirmed resolve.
$btnFix           = New-Object System.Windows.Forms.Button
$btnFix.Text      = 'Fix Automapping'
$btnFix.Location  = New-Object System.Drawing.Point(480, 54)
$btnFix.Size      = New-Object System.Drawing.Size(150, 26)
$btnFix.FlatStyle = 'Flat'
$btnFix.Enabled   = $false
[void]$grp.Controls.Add($btnFix)

$lblStage           = New-Object System.Windows.Forms.Label
$lblStage.Location  = New-Object System.Drawing.Point(14, 182)
$lblStage.Size      = New-Object System.Drawing.Size(892, 24)
$lblStage.Font      = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$lblStage.Anchor    = 'Top,Left,Right'
[void]$form.Controls.Add($lblStage)

$pb = Add-ReportProgressBar -Form $form -X 14 -Y 210 -Width 892 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(14, 232)
$txtOut.Size       = New-Object System.Drawing.Size(892, 338)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(786, 580)
$btnClose.Size      = New-Object System.Drawing.Size(120, 28)
$btnClose.FlatStyle = 'Flat'
$btnClose.Anchor    = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
[void]$form.Controls.Add($btnClose)
#endregion

#region Actions -----------------------------------------------------------------
function Invoke-Check {
    $btnFix.Enabled = $false
    $script:Mbx = $null
    $script:Usr = $null
    $btnCheck.Enabled = $false
    $lblStage.Text = 'Checking ...'
    $form.Refresh()

    try {
        $L = New-Object System.Collections.ArrayList

        $m = Resolve-One -Identity "$($txtMbx.Text)" -MailboxOnly
        if (-not $m.Ok) { $lblStage.Text = 'Mailbox not resolved'; $txtOut.Text = $m.Reason; return }
        $u = Resolve-One -Identity "$($txtUsr.Text)"
        if (-not $u.Ok) { $lblStage.Text = 'User not resolved'; $txtOut.Text = $u.Reason; return }

        $script:Mbx = $m.Item
        $script:Usr = $u.Item
        $mbxAddr = "$($script:Mbx.PrimarySmtpAddress)"
        $usrAddr = "$($script:Usr.PrimarySmtpAddress)"

        [void]$L.Add("Mailbox : $($script:Mbx.DisplayName)  <$mbxAddr>")
        [void]$L.Add("          type $($script:Mbx.RecipientTypeDetails)")
        [void]$L.Add("User    : $($script:Usr.DisplayName)  <$usrAddr>")
        [void]$L.Add('')

        $st = Get-FullAccessState -Mailbox $mbxAddr -User $usrAddr
        if (-not $st.Ok) { $lblStage.Text = 'Could not read permissions'; $txtOut.Text = $st.Reason; return }

        if (-not $st.Has) {
            # NOTHING TO REMOVE MEANS THIS IS THE WRONG TOOL. Adding the
            # permission here would be a different operation than the one the
            # button promises, so it is refused rather than quietly widened.
            $lblStage.Text = 'No direct Full Access to fix'
            $lblStage.ForeColor = [System.Drawing.Color]::Firebrick
            [void]$L.Add('THIS USER HAS NO DIRECT FULL ACCESS ON THAT MAILBOX.')
            [void]$L.Add('')
            [void]$L.Add('There is nothing to remove and re-add, so this tool cannot fix anything.')
            [void]$L.Add('Two common reasons:')
            [void]$L.Add('')
            [void]$L.Add('  They get access through a GROUP. Group-based Full Access never automaps,')
            [void]$L.Add('  and no amount of re-granting will change that. They need a direct grant,')
            [void]$L.Add('  or the mailbox added to their Outlook profile by hand.')
            [void]$L.Add('')
            [void]$L.Add('  They genuinely have no access. Grant it first -- Mailbox Permissions or')
            [void]$L.Add('  Mailbox / OneDrive Access -- and automapping will be set at that point.')
            $txtOut.Text = ($L -join [Environment]::NewLine)
            return
        }

        $lblStage.Text = 'Ready to fix'
        $lblStage.ForeColor = [System.Drawing.Color]::DarkGoldenrod
        [void]$L.Add("Direct Full Access found: $(@($st.Entries).Count) entry(ies).")
        foreach ($e in @($st.Entries)) { [void]$L.Add("   $($e.User)  --  $($e.AccessRights -join ', ')") }
        [void]$L.Add('')
        [void]$L.Add('AUTOMAPPING CANNOT BE READ BACK. Exchange Online does not report it, so')
        [void]$L.Add('this tool cannot tell you whether it is currently on -- only that the')
        [void]$L.Add('permission exists. That is also why the fix is a re-grant: the setting is')
        [void]$L.Add('decided when the permission is created and never revisited.')
        [void]$L.Add('')
        [void]$L.Add("PRESSING FIX REMOVES THEIR ACCESS FOR ABOUT $($script:WaitSeconds) SECONDS.")
        [void]$L.Add('If they are working in the mailbox right now, come back to it later.')
        $txtOut.Text = ($L -join [Environment]::NewLine)
        $btnFix.Enabled = $true
    } catch {
        $lblStage.Text = 'Check failed'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Check'
    } finally {
        $btnCheck.Enabled = $true
    }
}

function Invoke-Fix {
    if (-not $script:Mbx -or -not $script:Usr) { $txtOut.Text = 'Press Check first.'; return }

    $mbxAddr = "$($script:Mbx.PrimarySmtpAddress)"
    $usrAddr = "$($script:Usr.PrimarySmtpAddress)"

    $ans = Show-ToolDialog(
        "Remove and re-add Full Access?" + [Environment]::NewLine + [Environment]::NewLine +
        "Mailbox : $mbxAddr" + [Environment]::NewLine +
        "User    : $usrAddr" + [Environment]::NewLine + [Environment]::NewLine +
        "Their access is REMOVED for about $($script:WaitSeconds) seconds and then granted again with automapping on." + [Environment]::NewLine + [Environment]::NewLine +
        'Do not do this while they are working in the mailbox.',
        'Fix automapping', [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($ans -ne [System.Windows.Forms.DialogResult]::OK) { $txtOut.Text = 'Cancelled. Nothing was changed.'; return }

    $btnFix.Enabled = $false
    $btnCheck.Enabled = $false
    $pb.Maximum = $script:WaitSeconds + 2
    $pb.Value = 0
    $pb.Visible = $true

    $L = New-Object System.Collections.ArrayList
    $removed = $false

    try {
        # ---- 1. remove ---------------------------------------------------
        $lblStage.Text = 'Removing Full Access ...'
        $lblStage.ForeColor = [System.Drawing.Color]::DarkGoldenrod
        $form.Refresh()
        try {
            [void](Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{
                Identity = $mbxAddr; User = $usrAddr; AccessRights = 'FullAccess'
                Confirm = $false; ErrorAction = 'Stop' })
            $removed = $true
            [void]$L.Add("1. REMOVED Full Access for $usrAddr on $mbxAddr.")
        } catch {
            # FAILED HERE IS THE SAFE FAILURE. Nothing changed, so say so
            # plainly rather than leaving any doubt.
            $lblStage.Text = 'Failed -- nothing was changed'
            $lblStage.ForeColor = [System.Drawing.Color]::Firebrick
            [void]$L.Add('THE REMOVE FAILED, SO NOTHING WAS CHANGED.')
            [void]$L.Add('')
            [void]$L.Add((Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-MailboxPermission'))
            [void]$L.Add('')
            [void]$L.Add('Their access is exactly as it was. It is safe to try again.')
            $txtOut.Text = ($L -join [Environment]::NewLine)
            Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $mbxAddr 'Automap Fix') `
                -Detail "Remove-MailboxPermission failed for $usrAddr. No change made." -Environment $Environment
            return
        }
        $pb.Value = 1

        # ---- 2. wait -----------------------------------------------------
        # COUNTED DOWN ON SCREEN. A window that sits still for ten seconds looks
        # hung, and this is the exact moment the user has no access -- the
        # operator should be able to see it is deliberate and how long is left.
        for ($s = $script:WaitSeconds; $s -ge 1; $s--) {
            $lblStage.Text = "Waiting $s second(s) -- their access is removed right now ..."
            $txtOut.Text = ($L -join [Environment]::NewLine) + [Environment]::NewLine +
                "2. Waiting $s second(s) before re-adding."
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Seconds 1
            $pb.Value = [Math]::Min($pb.Maximum, $pb.Value + 1)
        }
        [void]$L.Add("2. Waited $($script:WaitSeconds) seconds.")

        # ---- 3. re-add ---------------------------------------------------
        $lblStage.Text = 'Re-adding Full Access ...'
        $form.Refresh()
        try {
            [void](Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
                Identity = $mbxAddr; User = $usrAddr; AccessRights = 'FullAccess'
                InheritanceType = 'All'
                # THE POINT OF THE WHOLE EXERCISE. AutoMapping is only honoured
                # at the moment the permission is created, which is why this
                # cannot be done with a Set- and why the remove was necessary.
                AutoMapping = $true
                ErrorAction = 'Stop' })
            [void]$L.Add("3. RE-ADDED Full Access with automapping ON.")
        } catch {
            # THE BAD FAILURE. Access was removed and not restored, and that
            # must be the loudest thing on the screen.
            $lblStage.Text = 'ACCESS REMOVED AND NOT RESTORED'
            $lblStage.ForeColor = [System.Drawing.Color]::Firebrick
            [void]$L.Add('')
            [void]$L.Add('*** THE RE-ADD FAILED. THEY NOW HAVE NO ACCESS. ***')
            [void]$L.Add('')
            [void]$L.Add((Format-ErrorDetail -ErrorRecord $_ -Context 'Add-MailboxPermission'))
            [void]$L.Add('')
            [void]$L.Add('THE REMOVE SUCCEEDED AND THE GRANT DID NOT, so this user has LESS access')
            [void]$L.Add('than before this ran. Fix it now:')
            [void]$L.Add('')
            [void]$L.Add("   Add-MailboxPermission -Identity '$mbxAddr' ``")
            [void]$L.Add("       -User '$usrAddr' -AccessRights FullAccess ``")
            [void]$L.Add('       -InheritanceType All -AutoMapping $true')
            [void]$L.Add('')
            [void]$L.Add('Or press Check and then Fix again -- Check will report no direct access,')
            [void]$L.Add('which confirms the grant is genuinely gone.')
            $txtOut.Text = ($L -join [Environment]::NewLine)
            Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $mbxAddr 'Automap Fix') `
                -Detail "REMOVED Full Access for $usrAddr but the re-add FAILED. User left without access." -Environment $Environment
            return
        }

        # ---- 4. verify ---------------------------------------------------
        # READ BACK. "No exception" is not evidence the grant landed, and this
        # is the run where being sure matters most.
        $after = Get-FullAccessState -Mailbox $mbxAddr -User $usrAddr
        [void]$L.Add('')
        if ($after.Ok -and $after.Has) {
            $lblStage.Text = 'Done -- Outlook will pick it up on next reconnect'
            $lblStage.ForeColor = [System.Drawing.Color]::ForestGreen
            [void]$L.Add('4. VERIFIED -- the permission is present again.')
        } else {
            $lblStage.Text = 'Granted, but could not verify'
            $lblStage.ForeColor = [System.Drawing.Color]::DarkGoldenrod
            [void]$L.Add('4. The grant reported success but the read-back did not show it yet.')
            [void]$L.Add('   That is usually replication. Press Check in a minute to confirm.')
        }

        [void]$L.Add('')
        [void]$L.Add('WHAT TO TELL THEM')
        [void]$L.Add('')
        [void]$L.Add('  Close Outlook completely and open it again. The mailbox appears when')
        [void]$L.Add('  Outlook next reconnects -- usually a few minutes, sometimes not until')
        [void]$L.Add('  a restart. Watching the folder list right now will show nothing.')
        [void]$L.Add('')
        [void]$L.Add('  If it still has not appeared after a restart and an hour, automapping is')
        [void]$L.Add('  not the problem -- check whether the mailbox was already added manually')
        [void]$L.Add('  to their profile, which blocks the automatic one from showing.')
        $txtOut.Text = ($L -join [Environment]::NewLine)

        Write-CompletedLog -Action 'Modify' -Target (Get-LogTarget $mbxAddr 'Automap Fix') -Result 'Success' `
            -Detail "Removed and re-added Full Access for $usrAddr on $mbxAddr with AutoMapping true." -Environment $Environment
    } catch {
        $lblStage.Text = 'Failed'
        $lblStage.ForeColor = [System.Drawing.Color]::Firebrick
        $extra = if ($removed) { [Environment]::NewLine + 'THE REMOVE HAD ALREADY SUCCEEDED -- check their access before leaving this.' } else { '' }
        $txtOut.Text = (Format-ErrorDetail -ErrorRecord $_ -Context 'Fix automapping') + $extra
        Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $mbxAddr 'Automap Fix') `
            -Detail "Automap fix failed. Removed=$removed. $($_.Exception.Message)" -Environment $Environment
    } finally {
        $pb.Visible = $false
        $btnCheck.Enabled = $true
        # NOT RE-ENABLED. The state has changed, so Check must run again before
        # another removal is possible.
        $btnFix.Enabled = $false
    }
}
#endregion

$btnCheck.Add_Click({ try { Invoke-Check } catch { Show-HandlerError -ErrorRecord $_ -Context 'Check' } })
$btnFix.Add_Click({ try { Invoke-Fix } catch { Show-HandlerError -ErrorRecord $_ -Context 'Fix automapping' } })
# Editing either name invalidates the check -- otherwise you could check one
# pair and fix another.
$txtMbx.Add_TextChanged({ try { $btnFix.Enabled = $false; $script:Mbx = $null } catch { } })
$txtUsr.Add_TextChanged({ try { $btnFix.Enabled = $false; $script:Usr = $null } catch { } })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Automap Fix' | Out-Null

$form.Add_Shown({
    try {
        if ($SmtpAddress) { $txtMbx.Text = "$SmtpAddress" }
        $lblStage.Text = ''
        $txtOut.Text = 'Enter the shared mailbox and the user, press Check, then Fix Automapping.' +
            [Environment]::NewLine + [Environment]::NewLine +
            "Environment: $(if ($Environment) { $Environment } else { '(none passed)' })" +
            [Environment]::NewLine + [Environment]::NewLine +
            'The fix removes Full Access, waits ' + $script:WaitSeconds + ' seconds, and grants it again with' + [Environment]::NewLine +
            'automapping on. Automapping is only decided when the permission is created,' + [Environment]::NewLine +
            'which is why re-granting is the fix and why there is no setting to toggle.' +
            [Environment]::NewLine + [Environment]::NewLine +
            'Their access is genuinely gone during those seconds. Do not run it while they' + [Environment]::NewLine +
            'are working in the mailbox.'
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Automap Fix' }
})

[void]$form.ShowDialog()
