#Requires -Version 5.1
# TOOLKIT-VARIANT: Dual
<#
    Hide-Unhide.ps1
    Hide a mailbox from the address lists, or show it again.

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    2026-09-16  A missing on-premises connection is reported as such instead of
                as 'the object is not there'.
#>
# TOOLKIT-HELP-RUNS: Hide and Unhide each change ONE attribute on ONE mailbox and ask first. Close changes nothing.
# TOOLKIT-HELP-BEGIN
# Hide a mailbox from the Global Address List, or put it back.
#
# Type the mailbox -- a bare alias ('glenn.armbruster') is completed with the
# tenant's mail domain from Tenant.json, or type the whole address. Pick where
# the attribute is mastered, then Hide or Unhide.
#
#   On-premises   For accounts synced from Active Directory. Writes
#                 HiddenFromAddressListsEnabled on the remote mailbox (or the
#                 on-premises mailbox, if it still has one). Entra Connect
#                 carries it to the cloud on the next sync -- allow 30 min.
#
#   Cloud         For cloud-only mailboxes (shared, room, or anything created
#                 in Exchange Online). Writes the same attribute with
#                 Set-Mailbox in Exchange Online. REFUSED for a synced mailbox:
#                 the cloud would accept the write and the next sync would
#                 put it back, so the tool tells you to use On-premises.
#
# Both buttons look first and say what the current state is; pressing Hide on
# a mailbox that is already hidden does nothing and says so.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')
if (-not (Test-ConsoleLaunch)) { return }

#region Data -------------------------------------------------------------------

function Get-HideState {
    <#
        Where the mailbox is and whether it is hidden, in the chosen world.

        Returns Ok / Smtp / Display / Kind / Hidden / Synced / Reason. Kind is
        which cmdlet pair applies: 'RemoteMailbox' or 'Mailbox' on-premises,
        'Mailbox' in the cloud. Reason is filled in when Ok is false and is
        written for the operator, not the log.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][ValidateSet('OnPrem', 'Cloud')][string]$World
    )
    $r = [PSCustomObject]@{ Ok = $false; Smtp = $Smtp; Display = ''; Kind = ''; Hidden = $null; Synced = $false; Reason = '' }

    if ($World -eq 'Cloud') {
        try {
            $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        } catch {
            $r.Reason = "not found in Exchange Online: $($_.Exception.Message)"
            return $r
        }
        $r.Smtp    = "$($mb.PrimarySmtpAddress)"
        $r.Display = "$($mb.DisplayName)"
        $r.Kind    = 'Mailbox'
        $r.Hidden  = [bool]$mb.HiddenFromAddressListsEnabled
        $r.Synced  = [bool]$mb.IsDirSynced
        if ($r.Synced) {
            # THE CLOUD WOULD TAKE THE WRITE AND THE NEXT SYNC WOULD UNDO IT.
            # That is the one failure here that looks like success for half an
            # hour, so it is refused rather than allowed with a warning.
            $r.Reason = 'is synced from Active Directory. The attribute is mastered on-premises; pick On-premises and try again.'
            return $r
        }
        $r.Ok = $true
        return $r
    }

    # ON-PREMISES: a synced account normally has a REMOTE mailbox object (the
    # cloud mailbox's on-premises shadow). One that has not been migrated still
    # has a real on-premises mailbox. Both carry the attribute; try in that order.
    # THE REASON THE LOOKUP FAILED IS KEPT, not thrown away. Both catches used
    # to swallow everything, so 'on-premises Exchange is not connected in this
    # process' -- which Get-OnPremCommand says in those words -- came out as
    # 'this object does not exist on-premises', and sent the operator to the
    # Cloud radio to solve a connection problem.
    $rm = $null; $why = ''
    try { $rm = Invoke-OnPremCommand -Name 'Get-RemoteMailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } }
    catch { $rm = $null; if ("$($_.Exception.Message)" -match '(?i)not connected|refusing to run') { $why = "$($_.Exception.Message)" } }
    if ($rm) {
        $r.Smtp    = "$($rm.PrimarySmtpAddress)"
        $r.Display = "$($rm.DisplayName)"
        $r.Kind    = 'RemoteMailbox'
        $r.Hidden  = [bool]$rm.HiddenFromAddressListsEnabled
        $r.Synced  = $true
        $r.Ok      = $true
        return $r
    }
    $mb = $null
    try { $mb = Invoke-OnPremCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } }
    catch { $mb = $null; if (-not $why -and "$($_.Exception.Message)" -match '(?i)not connected|refusing to run') { $why = "$($_.Exception.Message)" } }
    if ($mb) {
        $r.Smtp    = "$($mb.PrimarySmtpAddress)"
        $r.Display = "$($mb.DisplayName)"
        $r.Kind    = 'Mailbox'
        $r.Hidden  = [bool]$mb.HiddenFromAddressListsEnabled
        $r.Synced  = $true
        $r.Ok      = $true
        return $r
    }
    $r.Reason = $(if ($why) {
        "could not be looked up on-premises: $why"
    } else {
        'has no remote mailbox and no mailbox on-premises. A cloud-only mailbox is hidden from the Cloud side.'
    })
    return $r
}

function Set-HideState {
    param(
        [Parameter(Mandatory)]$State,
        [Parameter(Mandatory)][ValidateSet('OnPrem', 'Cloud')][string]$World,
        [Parameter(Mandatory)][bool]$Hidden
    )
    $p = @{ Identity = $State.Smtp; HiddenFromAddressListsEnabled = $Hidden; ErrorAction = 'Stop' }
    if ($World -eq 'Cloud') {
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters $p | Out-Null
        return
    }
    if ($State.Kind -eq 'RemoteMailbox') {
        Invoke-OnPremCommand -Name 'Set-RemoteMailbox' -Parameters $p | Out-Null
    } else {
        Invoke-OnPremCommand -Name 'Set-Mailbox' -Parameters $p | Out-Null
    }
}

#endregion

#region UI ---------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Hide / Unhide from Address Lists'
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblMbx = New-Object System.Windows.Forms.Label
$lblMbx.Text = 'Mailbox (alias or full address):'
$lblMbx.Location = New-Object System.Drawing.Point(12, 14)
$lblMbx.Size = New-Object System.Drawing.Size(400, 18)
$form.Controls.Add($lblMbx)

$txtMbx = New-Object System.Windows.Forms.TextBox
$txtMbx.Location = New-Object System.Drawing.Point(12, 34)
$txtMbx.Size = New-Object System.Drawing.Size(400, 24)
$txtMbx.Text = "$(Get-TargetSmtp -Text $SmtpAddress)"
$form.Controls.Add($txtMbx)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "A bare alias becomes alias@$(Get-MailDomain)"
$lblHint.Location = New-Object System.Drawing.Point(12, 60)
$lblHint.Size = New-Object System.Drawing.Size(400, 16)
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblHint)

$grpWorld = New-Object System.Windows.Forms.GroupBox
$grpWorld.Text = 'Where the attribute is mastered'
$grpWorld.Location = New-Object System.Drawing.Point(12, 84)
$grpWorld.Size = New-Object System.Drawing.Size(400, 52)
$form.Controls.Add($grpWorld)

$rbOnPrem = New-Object System.Windows.Forms.RadioButton
$rbOnPrem.Text = 'On-premises (synced accounts)'
$rbOnPrem.Location = New-Object System.Drawing.Point(12, 20)
$rbOnPrem.Size = New-Object System.Drawing.Size(200, 22)
$rbOnPrem.Checked = $true
$grpWorld.Controls.Add($rbOnPrem)

$rbCloud = New-Object System.Windows.Forms.RadioButton
$rbCloud.Text = 'Cloud (cloud-only mailboxes)'
$rbCloud.Location = New-Object System.Drawing.Point(218, 20)
$rbCloud.Size = New-Object System.Drawing.Size(176, 22)
$grpWorld.Controls.Add($rbCloud)

$btnHide = New-Object System.Windows.Forms.Button
$btnHide.Text = 'Hide'
$btnHide.Location = New-Object System.Drawing.Point(12, 148)
$btnHide.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnHide)

$btnUnhide = New-Object System.Windows.Forms.Button
$btnUnhide.Text = 'Unhide'
$btnUnhide.Location = New-Object System.Drawing.Point(152, 148)
$btnUnhide.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnUnhide)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(292, 148)
$btnClose.Size = New-Object System.Drawing.Size(120, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(12, 190)
$txtOut.Size = New-Object System.Drawing.Size(400, 110)
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$form.ClientSize = New-Object System.Drawing.Size(424, ($txtOut.Bottom + 12))
$form.AcceptButton = $btnHide
$form.CancelButton = $btnClose

# ONE HANDLER, ONE FLAG. Hide and Unhide are the same operation with the
# value flipped, and two copies of the look-confirm-write-log sequence would
# be two places for them to drift apart.
$apply = {
    param([bool]$Hidden)
    $verb  = $(if ($Hidden) { 'Hide' } else { 'Unhide' })
    $world = $(if ($rbCloud.Checked) { 'Cloud' } else { 'OnPrem' })
    $smtp  = Get-TargetSmtp -Text $txtMbx.Text
    if (-not $smtp) { $txtOut.Text = 'Type a mailbox first.'; return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnHide.Enabled = $false; $btnUnhide.Enabled = $false
    try {
        $txtOut.Text = "Looking up $smtp ($world)..."
        $txtOut.Refresh()
        $st = Get-HideState -Smtp $smtp -World $world
        if (-not $st.Ok) {
            $txtOut.Text = "$smtp $($st.Reason)"
            return
        }
        if ($st.Hidden -eq $Hidden) {
            $txtOut.Text = "$($st.Display) <$($st.Smtp)> is already " +
                           $(if ($Hidden) { 'hidden from' } else { 'visible in' }) + ' the address lists. Nothing to do.'
            return
        }

        $where = $(if ($world -eq 'Cloud') { 'Exchange Online (Set-Mailbox)' }
                   elseif ($st.Kind -eq 'RemoteMailbox') { 'on-premises (Set-RemoteMailbox), synced to the cloud within ~30 min' }
                   else { 'on-premises (Set-Mailbox), synced to the cloud within ~30 min' })
        $a = Show-ToolDialog(
            ("$verb this mailbox " + $(if ($Hidden) { 'FROM' } else { 'IN' }) + " the address lists?`r`n`r`n" +
             "  $($st.Display)`r`n  $($st.Smtp)`r`n`r`nWritten $where."),
            "Confirm $verb", 'OKCancel', 'Question')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { $txtOut.Text = 'Cancelled. Nothing changed.'; return }

        Set-HideState -State $st -World $world -Hidden $Hidden
        Write-CompletedLog -Action 'Configure' -Target $st.Smtp -Result 'Success' `
            -Detail "$verb from address lists ($world, $($st.Kind)): HiddenFromAddressListsEnabled = $Hidden" -Environment $Environment

        $txtOut.Text = "Done. $($st.Display) <$($st.Smtp)> is now " +
                       $(if ($Hidden) { 'HIDDEN from' } else { 'VISIBLE in' }) + " the address lists.`r`n" +
                       $(if ($world -eq 'OnPrem') { 'The cloud picks this up on the next directory sync (~30 min); Outlook cached address books update on their own schedule.' }
                         else { 'Outlook cached address books update on their own schedule, usually within a day.' })
    } catch {
        Write-FailedLog -Action 'Configure' -Target $smtp -ErrorRecord $_ -Detail "$verb from address lists ($world)"
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "$verb from address lists"
    } finally {
        $btnHide.Enabled = $true; $btnUnhide.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}
$btnHide.Add_Click({   & $apply $true  })
$btnUnhide.Add_Click({ & $apply $false })

Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Hide / Unhide' | Out-Null
Set-ConsoleTheme -Form $form
[void]$form.ShowDialog()

#endregion
