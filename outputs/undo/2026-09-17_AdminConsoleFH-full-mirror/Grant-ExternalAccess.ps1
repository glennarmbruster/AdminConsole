#Requires -Version 5.1
<#
    Grant-ExternalAccess.ps1
    Assign a Teams external access policy to a user.

    THE FIRST TEAMS TOOL IN THIS TOOLKIT, so it carries its own availability
    check rather than using an Invoke-*Command wrapper -- there is no Teams
    equivalent of Invoke-EXOCommand yet. If a second Teams tool arrives, that
    check belongs in Common and this one should call it.

    -Environment IS ACCEPTED AND USED, and passed to every Write-*Log call.
#>
# TOOLKIT-HELP-RUNS: WRITES a Teams external access policy assignment. Changes configuration.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Assign a Teams external access policy to one user.
#
# External access controls whether somebody can chat and meet with people
# outside the organisation. This grants the policy that allows it.
#
# THIS WRITES. It changes the user's Teams policy assignment, so it shows you
# what they have now, asks, and then verifies the change actually took.
#
# THE CURRENT POLICY IS SHOWN BEFORE YOU GRANT. That matters more than it
# looks: a blank policy means the user is on the ORGANISATION-WIDE DEFAULT, and
# granting an explicit policy takes them off it permanently. If the global
# default is later changed, this user will not follow it.
#
# ASSIGNMENT IS NOT INSTANT. Teams policy changes usually apply within minutes
# but can take several hours to reach the client. A user reporting no change
# ten minutes later is normal, and re-running this will not speed it up.
#
# THE DOMAIN IS A FIELD, NOT A CONSTANT. It defaults to the tenant's primary domain. If
# the user is on another accepted domain, change it rather than typing the
# whole address -- the field is what gets appended.
#
# It needs the MicrosoftTeams module and a connected Teams session.
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

# THE DOMAIN IS READ FROM Tenant.json, NOT WRITTEN HERE.
#
# It was a literal, which is the one thing a .ps1 in this toolkit may not
# contain -- Test-Toolkit fails the build on a domain in code, and the reason
# is that the same script ships to more than one place. Get-MailDomain returns
# the tenant's primary domain and '' if it cannot, which is the same answer an
# empty box gives: the operator types the full address instead.
$script:DefaultDomain = ''
try { $script:DefaultDomain = Get-MailDomain } catch { }
$script:DefaultPolicy = 'ExternalAccessAllowed'
$script:Current       = $null

#region Teams -------------------------------------------------------------------
function Test-TeamsAvailable {
    <#
        Is there a usable Teams session?

        TESTED BY CALLING, NOT BY Get-Command. The cmdlet exists as soon as the
        module is imported, whether or not anybody has connected -- so checking
        for the command reports "available" and the first real call then fails
        with an authentication error the operator has to decode.
    #>
    $r = [PSCustomObject]@{ Available = $false; Detail = '' }

    if (-not (Get-Command Grant-CsExternalAccessPolicy -ErrorAction SilentlyContinue)) {
        $r.Detail = 'The MicrosoftTeams module is not loaded. Run Connect-MicrosoftTeams in this session first.'
        return $r
    }
    try {
        [void](Get-CsExternalAccessPolicy -ErrorAction Stop | Select-Object -First 1)
        $r.Available = $true
    } catch {
        $r.Detail = "Teams is not connected: $($_.Exception.Message)"
    }
    return $r
}

function Get-UserPolicyState {
    <#
        What the user has now.

        A BLANK PolicyName IS NOT AN ERROR AND IS NOT 'none'. It means the user
        inherits the organisation-wide default, which is a meaningfully
        different state from having an explicit policy -- and it is the state
        this tool is about to take them out of.
    #>
    param([Parameter(Mandatory)][string]$Upn)

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Policy = ''; Inherited = $false; Identity = '' }
    try {
        $u = Get-CsOnlineUser -Identity $Upn -ErrorAction Stop
        if (-not $u) { $r.Reason = "No Teams user found for $Upn."; return $r }
        $r.Identity  = "$($u.UserPrincipalName)"
        $p = "$($u.ExternalAccessPolicy)"
        if (-not $p.Trim()) { $r.Inherited = $true; $r.Policy = '(organisation-wide default)' }
        else { $r.Policy = $p }
        $r.Ok = $true
    } catch {
        $r.Reason = Format-ErrorDetail -ErrorRecord $_ -Context "Get-CsOnlineUser $Upn"
    }
    return $r
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Grant External Access'
$form.ClientSize      = New-Object System.Drawing.Size(900, 560)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$grp          = New-Object System.Windows.Forms.GroupBox
$grp.Text     = 'User'
$grp.Location = New-Object System.Drawing.Point(14, 46)
$grp.Size     = New-Object System.Drawing.Size(872, 128)
$grp.Anchor   = 'Top,Left,Right'
[void]$form.Controls.Add($grp)

$lblUser          = New-Object System.Windows.Forms.Label
$lblUser.Text     = 'User name'
$lblUser.Location = New-Object System.Drawing.Point(14, 30)
$lblUser.Size     = New-Object System.Drawing.Size(80, 20)
[void]$grp.Controls.Add($lblUser)

$txtUser          = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(98, 27)
$txtUser.Size     = New-Object System.Drawing.Size(260, 22)
[void]$grp.Controls.Add($txtUser)

$lblAt          = New-Object System.Windows.Forms.Label
$lblAt.Text     = '@'
$lblAt.Location = New-Object System.Drawing.Point(364, 30)
$lblAt.Size     = New-Object System.Drawing.Size(16, 20)
[void]$grp.Controls.Add($lblAt)

$txtDomain          = New-Object System.Windows.Forms.TextBox
$txtDomain.Location = New-Object System.Drawing.Point(384, 27)
$txtDomain.Size     = New-Object System.Drawing.Size(200, 22)
$txtDomain.Text     = $script:DefaultDomain
[void]$grp.Controls.Add($txtDomain)

$lblResolved           = New-Object System.Windows.Forms.Label
$lblResolved.Location  = New-Object System.Drawing.Point(98, 54)
$lblResolved.Size      = New-Object System.Drawing.Size(486, 20)
$lblResolved.ForeColor = [System.Drawing.Color]::Gray
[void]$grp.Controls.Add($lblResolved)

$lblPolicy          = New-Object System.Windows.Forms.Label
$lblPolicy.Text     = 'Policy'
$lblPolicy.Location = New-Object System.Drawing.Point(14, 88)
$lblPolicy.Size     = New-Object System.Drawing.Size(80, 20)
[void]$grp.Controls.Add($lblPolicy)

$cboPolicy               = New-Object System.Windows.Forms.ComboBox
$cboPolicy.DropDownStyle = 'DropDownList'
$cboPolicy.Location      = New-Object System.Drawing.Point(98, 85)
$cboPolicy.Size          = New-Object System.Drawing.Size(300, 22)
[void]$grp.Controls.Add($cboPolicy)

$btnCheck           = New-Object System.Windows.Forms.Button
$btnCheck.Text      = 'Check user'
$btnCheck.Location  = New-Object System.Drawing.Point(600, 25)
$btnCheck.Size      = New-Object System.Drawing.Size(120, 28)
$btnCheck.FlatStyle = 'Flat'
[void]$grp.Controls.Add($btnCheck)

# DISABLED UNTIL THE USER HAS BEEN CHECKED. The panel has to have shown what
# they have now before it can change it -- that is the whole safeguard.
$btnGrant           = New-Object System.Windows.Forms.Button
$btnGrant.Text      = 'Grant policy'
$btnGrant.Location  = New-Object System.Drawing.Point(600, 83)
$btnGrant.Size      = New-Object System.Drawing.Size(120, 28)
$btnGrant.FlatStyle = 'Flat'
$btnGrant.Enabled   = $false
[void]$grp.Controls.Add($btnGrant)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(14, 184)
$txtOut.Size       = New-Object System.Drawing.Size(872, 326)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(766, 520)
$btnClose.Size      = New-Object System.Drawing.Size(120, 28)
$btnClose.FlatStyle = 'Flat'
$btnClose.Anchor    = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
[void]$form.Controls.Add($btnClose)
#endregion

#region Actions -----------------------------------------------------------------
function Get-TargetUpn {
    $u = "$($txtUser.Text)".Trim()
    $d = "$($txtDomain.Text)".Trim().TrimStart('@')
    if (-not $u) { return '' }
    # PASTING A FULL ADDRESS IS THE COMMON MISTAKE, and silently producing
    # bob@x.com@contoso.com would fail with an unhelpful "user not found".
    if ($u -like '*@*') { return $u }
    if (-not $d) { return '' }
    return "$u@$d"
}

function Update-Resolved {
    $upn = Get-TargetUpn
    if ($upn) { $lblResolved.Text = "Will act on: $upn" } else { $lblResolved.Text = '' }
    $btnGrant.Enabled = $false
    $script:Current = $null
}

function Invoke-Check {
    $upn = Get-TargetUpn
    if (-not $upn) { $txtOut.Text = 'Enter a user name.'; return }

    $t = Test-TeamsAvailable
    if (-not $t.Available) { $txtOut.Text = $t.Detail; return }

    $btnCheck.Enabled = $false
    try {
        $script:Current = Get-UserPolicyState -Upn $upn
        $L = New-Object System.Collections.ArrayList
        if (-not $script:Current.Ok) {
            [void]$L.Add($script:Current.Reason)
            $btnGrant.Enabled = $false
        } else {
            [void]$L.Add("User    : $($script:Current.Identity)")
            [void]$L.Add("Policy  : $($script:Current.Policy)")
            [void]$L.Add('')
            if ($script:Current.Inherited) {
                [void]$L.Add('THEY ARE ON THE ORGANISATION-WIDE DEFAULT at the moment.')
                [void]$L.Add('')
                [void]$L.Add('Granting an explicit policy takes them off it permanently. If the global')
                [void]$L.Add('default is changed later, this user will NOT follow it -- they will keep')
                [void]$L.Add('whatever is granted here until somebody removes it.')
            } elseif ("$($script:Current.Policy)" -eq "$($cboPolicy.SelectedItem)") {
                [void]$L.Add("THEY ALREADY HAVE '$($cboPolicy.SelectedItem)'. Granting it again changes")
                [void]$L.Add('nothing. If external access is not working, the cause is elsewhere --')
                [void]$L.Add('the org-wide federation settings, or the change simply has not reached')
                [void]$L.Add('the client yet.')
            } else {
                [void]$L.Add("They have an explicit policy already. Granting will REPLACE it.")
            }
            [void]$L.Add('')
            [void]$L.Add("Press Grant policy to assign '$($cboPolicy.SelectedItem)'.")
            $btnGrant.Enabled = $true
        }
        $txtOut.Text = ($L -join [Environment]::NewLine)
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Check user'
    } finally {
        $btnCheck.Enabled = $true
    }
}

function Invoke-Grant {
    $upn    = Get-TargetUpn
    $policy = "$($cboPolicy.SelectedItem)"
    if (-not $upn -or -not $policy) { return }
    if (-not $script:Current -or -not $script:Current.Ok) {
        $txtOut.Text = 'Check the user first.'; return
    }

    $was = $script:Current.Policy
    $msg = "Assign external access policy?" + [Environment]::NewLine + [Environment]::NewLine +
           "User   : $upn" + [Environment]::NewLine +
           "From   : $was" + [Environment]::NewLine +
           "To     : $policy"
    $ans = Show-ToolDialog($msg, 'Grant external access',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = 'Cancelled. Nothing was changed.'; return }

    $btnGrant.Enabled = $false
    try {
        Grant-CsExternalAccessPolicy -Identity $upn -PolicyName $policy -ErrorAction Stop

        # VERIFIED BY READING IT BACK. Grant-CsExternalAccessPolicy returns
        # nothing on success, so "no error" is the only signal it gives -- and
        # that is not the same as the assignment having landed.
        $after = Get-UserPolicyState -Upn $upn
        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("GRANTED '$policy' to $upn.")
        [void]$L.Add('')
        [void]$L.Add("Was : $was")
        if ($after.Ok) {
            [void]$L.Add("Now : $($after.Policy)")
            [void]$L.Add('')
            if ("$($after.Policy)" -ne $policy) {
                [void]$L.Add('THE READ-BACK DOES NOT MATCH YET. That is usually replication rather')
                [void]$L.Add('than failure -- check again in a few minutes before doing anything else.')
            } else {
                [void]$L.Add('Confirmed by reading the assignment back.')
            }
        } else {
            [void]$L.Add("Could not read the policy back: $($after.Reason)")
        }
        [void]$L.Add('')
        [void]$L.Add('IT WILL NOT BE INSTANT FOR THE USER. Teams policy changes usually apply')
        [void]$L.Add('within minutes but can take several hours to reach the client. Re-running')
        [void]$L.Add('this does not speed it up, and signing out and in is worth trying first.')
        $txtOut.Text = ($L -join [Environment]::NewLine)

        Write-CompletedLog -Action 'Modify' -Target (Get-LogTarget $upn 'Grant External Access') -Result 'Success' `
            -Detail "External access policy '$was' -> '$policy'." -Environment $Environment

        $script:Current = $after
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Grant-CsExternalAccessPolicy $upn"
        Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $upn 'Grant External Access') `
            -Detail "$($_.Exception.Message)" -Environment $Environment
    } finally {
        $btnGrant.Enabled = $true
    }
}
#endregion

$txtUser.Add_TextChanged({ try { Update-Resolved } catch { } })
$txtDomain.Add_TextChanged({ try { Update-Resolved } catch { } })
$btnCheck.Add_Click({ try { Invoke-Check } catch { Show-HandlerError -ErrorRecord $_ -Context 'Check' } })
$btnGrant.Add_Click({ try { Invoke-Grant } catch { Show-HandlerError -ErrorRecord $_ -Context 'Grant' } })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Grant External Access' | Out-Null

$form.Add_Shown({
    try {
        # POPULATED FROM THE TENANT, not hardcoded. A policy name typed into a
        # constant fails at the moment of the write, after the confirmation --
        # the worst possible time to discover it does not exist.
        $t = Test-TeamsAvailable
        if ($t.Available) {
            $names = @()
            try { $names = @(Get-CsExternalAccessPolicy -ErrorAction Stop | ForEach-Object { "$($_.Identity)" -replace '^Tag:', '' } | Sort-Object -Unique) } catch { }
            foreach ($n in $names) { if ($n) { [void]$cboPolicy.Items.Add($n) } }
            if ($cboPolicy.Items.Count -eq 0) { [void]$cboPolicy.Items.Add($script:DefaultPolicy) }
            $i = $cboPolicy.Items.IndexOf($script:DefaultPolicy)
            $cboPolicy.SelectedIndex = $(if ($i -ge 0) { $i } else { 0 })
            $txtOut.Text = "Enter a user name, press Check user, then Grant policy." + [Environment]::NewLine +
                "Environment: $(if ($Environment) { $Environment } else { '(none passed)' })"
        } else {
            [void]$cboPolicy.Items.Add($script:DefaultPolicy)
            $cboPolicy.SelectedIndex = 0
            $txtOut.Text = $t.Detail
        }

        if ($SmtpAddress) {
            # Launched against an object from the console -- prefill from it.
            $txtUser.Text = "$SmtpAddress"
        }
        Update-Resolved
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Grant External Access' }
})

[void]$form.ShowDialog()
