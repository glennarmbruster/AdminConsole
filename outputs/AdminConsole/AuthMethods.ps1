#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-GRAPH-SCOPES: UserAuthenticationMethod.ReadWrite.All, User.Read.All
<#
    AuthMethods.ps1
    What a user can sign in with -- registered MFA methods -- and the two things
    a helpdesk does about them: remove a stale one, and issue a Temporary Access
    Pass so they can re-enrol.

    Author   : Glenn Armbruster
    Created  : 2026-09-17
#>
# TOOLKIT-HELP-RUNS: Lists one user's authentication methods. Remove and Temporary Access Pass change things and each asks first.
# TOOLKIT-HELP-BEGIN
# The authentication methods registered on one account: Authenticator app,
# phone numbers, FIDO2 security keys, Windows Hello, software OATH tokens,
# email (for self-service password reset), Temporary Access Passes, and the
# password itself.
#
# THREE ACTIONS, AND WHAT EACH ONE IS FOR:
#
#   Remove selected      Deletes the ticked methods. The usual reason is a lost
#                        or replaced phone: remove the old Authenticator and
#                        phone entries, then issue a Temporary Access Pass so
#                        the person can register the new device. The password
#                        cannot be removed and is not offered.
#
#   Temporary Access Pass  A short-lived code that satisfies MFA on its own,
#                        for a user who has no working method. Shown ONCE, here,
#                        and never again -- copy it to the ticket or read it to
#                        the caller. One hour and single-use by default.
#
#   Reset MFA            Removes EVERY removable method at once. This is the
#                        "start over" for a compromised or fully lost set of
#                        methods, and it asks for the user's address to be typed
#                        back before it runs.
#
# Removing the method a user relies on locks them out until a new one exists,
# which is why Temporary Access Pass sits next to Remove.
#
# NEEDS: the Graph scope UserAuthenticationMethod.ReadWrite.All, and the
# Authentication Administrator role on YOUR account (Privileged Authentication
# Administrator to act on accounts that hold admin roles). Global Reader can
# list methods but not change them.
#
# Every remove and every pass is written to the audit log.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

$script:Scopes = @('UserAuthenticationMethod.ReadWrite.All', 'User.Read.All')

#region Data ------------------------------------------------------------------

<#
    Learn 2026-09-17 (authentication methods overview + each method's delete
    page). The @odata.type on every returned method decides which collection
    it is deleted from. Password is never deletable; the two at the end are
    manageable in the portal but their delete paths were not verified, so
    they are listed and left alone.
#>
$script:MethodMap = @{
    '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'   = @{ Label = 'Authenticator app';      Path = 'microsoftAuthenticatorMethods' }
    '#microsoft.graph.phoneAuthenticationMethod'                    = @{ Label = 'Phone';                  Path = 'phoneMethods' }
    '#microsoft.graph.fido2AuthenticationMethod'                    = @{ Label = 'FIDO2 security key';     Path = 'fido2Methods' }
    '#microsoft.graph.softwareOathAuthenticationMethod'             = @{ Label = 'Software OATH token';    Path = 'softwareOathMethods' }
    '#microsoft.graph.emailAuthenticationMethod'                    = @{ Label = 'Email (SSPR)';           Path = 'emailMethods' }
    '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod'  = @{ Label = 'Windows Hello';          Path = 'windowsHelloForBusinessMethods' }
    '#microsoft.graph.temporaryAccessPassAuthenticationMethod'      = @{ Label = 'Temporary Access Pass';  Path = 'temporaryAccessPassMethods' }
    '#microsoft.graph.platformCredentialAuthenticationMethod'       = @{ Label = 'Platform credential';    Path = 'platformCredentialMethods' }
    '#microsoft.graph.passwordAuthenticationMethod'                 = @{ Label = 'Password';               Path = '' }
    '#microsoft.graph.externalAuthenticationMethod'                 = @{ Label = 'External method';        Path = '' }
    '#microsoft.graph.qrCodePinAuthenticationMethod'                = @{ Label = 'QR code + PIN';          Path = '' }
}

function Get-UserAuthMethod {
    # Learn 2026-09-17 (List methods): GET /users/{id}/authentication/methods. No query parameters.
    param([Parameter(Mandatory)][string]$UserId)
    $r = Invoke-ToolkitGraph -Uri "users/$UserId/authentication/methods"
    if ($r -and $r.PSObject.Properties['value']) { return @($r.value) }
    return @()
}

function Get-AuthMethodRow {
    <#
        One grid row per method: a friendly type, the detail a helpdesk
        recognises (the phone number, the device name, the key model), and
        whether it can be removed here.
    #>
    param($m)
    $type = "$($m.'@odata.type')"
    $def  = $script:MethodMap[$type]
    $label = if ($def) { $def.Label } else { ($type -replace '^#microsoft\.graph\.', '') }
    $detail = ''
    switch ($type) {
        '#microsoft.graph.phoneAuthenticationMethod'                   { $detail = "$($m.phoneType): $($m.phoneNumber)   sms sign-in: $($m.smsSignInState)" }
        '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'  { $detail = "$($m.displayName)   $($m.deviceTag)   v$($m.phoneAppVersion)" }
        '#microsoft.graph.fido2AuthenticationMethod'                   { $detail = "$($m.displayName)   $($m.model)   attestation: $($m.attestationLevel)" }
        '#microsoft.graph.softwareOathAuthenticationMethod'            { $detail = "secret key on file" }
        '#microsoft.graph.emailAuthenticationMethod'                   { $detail = "$($m.emailAddress)" }
        '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' { $detail = "$($m.displayName)   key strength: $($m.keyStrength)   device: $($m.device.displayName)" }
        '#microsoft.graph.temporaryAccessPassAuthenticationMethod'     { $detail = "usable: $($m.isUsable)   once: $($m.isUsableOnce)   $($m.lifetimeInMinutes) min from $(ConvertTo-LocalGraphTime $m.startDateTime)   $($m.methodUsabilityReason)" }
        '#microsoft.graph.platformCredentialAuthenticationMethod'      { $detail = "$($m.displayName)   platform: $($m.platform)   $($m.device.displayName)" }
        '#microsoft.graph.passwordAuthenticationMethod'                { $detail = 'cannot be removed here (reset it through the normal password process)' }
        default                                                       { $detail = "$($m.displayName)" }
    }
    return [PSCustomObject]@{
        Type      = $label
        Detail    = "$detail".Trim()
        Created   = (ConvertTo-LocalGraphTime $m.createdDateTime)
        Removable = [bool]($def -and $def.Path)
        Path      = $(if ($def) { $def.Path } else { '' })
        Id        = "$($m.id)"
        Raw       = $m
    }
}

function Remove-UserAuthMethod {
    # Learn 2026-09-17: DELETE /users/{id}/authentication/<collection>/{methodId}, returns 204.
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$MethodId)
    Invoke-ToolkitGraph -Uri "users/$UserId/authentication/$Path/$MethodId" -Method DELETE | Out-Null
}

function New-UserTemporaryAccessPass {
    <#
        Learn 2026-09-17 (Create temporaryAccessPassAuthenticationMethod):
        POST /users/{id}/authentication/temporaryAccessPassMethods with
        lifetimeInMinutes (10-43200), isUsableOnce, optional startDateTime.
        Response carries temporaryAccessPass -- the only time it is readable.
    #>
    param([Parameter(Mandatory)][string]$UserId, [int]$Minutes = 60, [bool]$Once = $true)
    $body = @{ lifetimeInMinutes = $Minutes; isUsableOnce = $Once }
    return (Invoke-ToolkitGraph -Uri "users/$UserId/authentication/temporaryAccessPassMethods" -Method POST -Body $body)
}
#endregion

#region UI --------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Authentication Methods'
$form.ClientSize    = New-Object System.Drawing.Size(1060, 560)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$y = 12
$lblUser          = New-Object System.Windows.Forms.Label
$lblUser.Text     = 'User:'
$lblUser.Location = New-Object System.Drawing.Point(12, ($y + 3))
$lblUser.Size     = New-Object System.Drawing.Size(40, 18)
$form.Controls.Add($lblUser)

$txtUser          = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(54, $y)
$txtUser.Size     = New-Object System.Drawing.Size(360, 24)
$txtUser.Font     = New-Object System.Drawing.Font('Consolas', 9)
$txtUser.Text     = "$(Get-TargetSmtp -Text $SmtpAddress)"
$form.Controls.Add($txtUser)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load'
$btnLoad.Location = New-Object System.Drawing.Point(426, ($y - 2))
$btnLoad.Size     = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnLoad)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(948, ($y - 2))
$btnClose.Size     = New-Object System.Drawing.Size(100, 28)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)
$y += 34

$lblWho           = New-Object System.Windows.Forms.Label
$lblWho.Tag       = 'NoTheme'
$lblWho.Location  = New-Object System.Drawing.Point(12, $y)
$lblWho.Size      = New-Object System.Drawing.Size(1036, 18)
$lblWho.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWho.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblWho)
$y += 22

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(12, $y)
$lblStatus.Size      = New-Object System.Drawing.Size(1036, 18)
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$lblStatus.Text      = 'Enter a user and press Load.'
$form.Controls.Add($lblStatus)
$y += 22

$grid                     = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(12, $y)
$grid.Size                = New-Object System.Drawing.Size(1036, 330)
$grid.ReadOnly            = $false
$grid.AllowUserToAddRows  = $false
$grid.SelectionMode       = 'FullRowSelect'
$grid.MultiSelect         = $true
$grid.RowHeadersVisible   = $false
$grid.AutoSizeColumnsMode = 'None'
$grid.ScrollBars          = 'Both'
$grid.Font                = New-Object System.Drawing.Font('Consolas', 9)
$colPick = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colPick.Name = 'Pick'; $colPick.HeaderText = ''; $colPick.Width = 36
[void]$grid.Columns.Add($colPick)
foreach ($c in @(@('Type', 190), @('Detail', 560), @('Created', 118), @('Removable', 80))) {
    $idx = $grid.Columns.Add($c[0], $c[0])
    $grid.Columns[$idx].Width    = $c[1]
    $grid.Columns[$idx].ReadOnly = $true
    $grid.Columns[$idx].SortMode = 'Automatic'
}
$form.Controls.Add($grid)
$y += 340

$btnRemove          = New-Object System.Windows.Forms.Button
$btnRemove.Text     = 'Remove ticked'
$btnRemove.Location = New-Object System.Drawing.Point(12, $y)
$btnRemove.Size     = New-Object System.Drawing.Size(150, 28)
$btnRemove.Enabled  = $false
$form.Controls.Add($btnRemove)

$btnTap          = New-Object System.Windows.Forms.Button
$btnTap.Text     = 'Temporary Access Pass...'
$btnTap.Location = New-Object System.Drawing.Point(170, $y)
$btnTap.Size     = New-Object System.Drawing.Size(190, 28)
$btnTap.Enabled  = $false
$form.Controls.Add($btnTap)

$btnReset          = New-Object System.Windows.Forms.Button
$btnReset.Text     = 'Reset MFA (remove all)'
$btnReset.Location = New-Object System.Drawing.Point(368, $y)
$btnReset.Size     = New-Object System.Drawing.Size(180, 28)
$btnReset.Enabled  = $false
$form.Controls.Add($btnReset)

$lblNote           = New-Object System.Windows.Forms.Label
$lblNote.Text      = 'Remove the lost device, then issue a pass so the user can register the new one. The password is never removed here.'
$lblNote.Location  = New-Object System.Drawing.Point(560, ($y + 6))
$lblNote.Size      = New-Object System.Drawing.Size(488, 18)
$lblNote.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblNote)

$form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force
$script:User = $null
$script:Rows = @()

$load = {
    $grid.Rows.Clear(); $script:Rows = @()
    $btnRemove.Enabled = $false; $btnTap.Enabled = $false; $btnReset.Enabled = $false
    $t = "$($txtUser.Text)".Trim()
    if (-not $t) { $lblStatus.Text = 'Enter a user (UPN or email address).'; return }
    $lblStatus.Text = "Resolving $t..."
    $u = Resolve-GraphUser -Text $t
    if (-not $u) { $lblStatus.Text = "No user matches '$t'."; $lblWho.Text = ''; return }
    $script:User = $u
    $lblWho.Text = "$($u.displayName)  <$($u.userPrincipalName)>   account $(if ($u.accountEnabled) { 'ENABLED' } else { 'DISABLED' })"
    $lblWho.ForeColor = $(if ($u.accountEnabled) { [System.Drawing.Color]::FromArgb(0, 110, 60) } else { [System.Drawing.Color]::FromArgb(200, 0, 0) })
    $lblStatus.Text = 'Reading authentication methods...'
    $ms = @(Get-UserAuthMethod -UserId $u.id | ForEach-Object { Get-AuthMethodRow $_ })
    $script:Rows = $ms
    foreach ($r in $ms) {
        $i = $grid.Rows.Add($false, $r.Type, $r.Detail, $r.Created, $(if ($r.Removable) { 'yes' } else { 'no' }))
        $grid.Rows[$i].Tag = $r
        if (-not $r.Removable) {
            $grid.Rows[$i].Cells['Pick'].ReadOnly = $true
            $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DimGray
        }
    }
    $removable = @($ms | Where-Object { $_.Removable }).Count
    $lblStatus.Text = "$($ms.Count) method(s); $removable removable."
    $btnRemove.Enabled = ($removable -gt 0); $btnReset.Enabled = ($removable -gt 0); $btnTap.Enabled = $true
}

$run = {
    param([scriptblock]$Do, [string]$Context)
    if ($form.Busy) { return }
    $form.Busy = $true
    try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; & $Do }
    catch { $lblStatus.Text = "$Context failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context $Context }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
}

$btnLoad.Add_Click({ & $run $load 'Load authentication methods (Graph)' })
$txtUser.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $btnLoad.PerformClick(); $_.SuppressKeyPress = $true } })

$removeSet = {
    param([object[]]$Targets, [string]$Action, [string]$Why)
    if ($Targets.Count -eq 0) { $lblStatus.Text = 'Nothing ticked that can be removed.'; return }
    $names = @($Targets | ForEach-Object { "$($_.Type)  --  $($_.Detail)" })
    $a = Show-ToolDialog(
        ("Remove $($Targets.Count) method(s) from $($script:User.userPrincipalName)?`r`n`r`n  " + ($names -join "`r`n  ") + "`r`n`r`n$Why"),
        'Confirm', 'OKCancel', 'Warning')
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $ok = 0; $failed = 0; $errs = @()
    foreach ($r in $Targets) {
        try { Remove-UserAuthMethod -UserId $script:User.id -Path $r.Path -MethodId $r.Id; $ok++ }
        catch { $failed++; $errs += "$($r.Type) -- $($_.Exception.Message)" }
    }
    $result = if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }
    Write-CompletedLog -Action $Action -Target (Get-LogTarget $script:User.userPrincipalName) -Result $result `
        -Detail "$ok removed, $failed failed: $($names -join '; ')$(if ($errs) { ' | ' + ($errs -join ' | ') })" -Environment 'Cloud'
    & $load
    $lblStatus.Text = "Removed $ok, failed $failed.$(if ($errs) { '  ' + ($errs -join '  ') })"
}

$btnRemove.Add_Click({
    & $run {
        $picked = @()
        foreach ($row in $grid.Rows) {
            if ($row.Cells['Pick'].Value -eq $true -and $row.Tag -and $row.Tag.Removable) { $picked += $row.Tag }
        }
        & $removeSet $picked 'AuthMethodRemove' 'If this is the only method the user has, they cannot pass MFA until a new one is registered -- issue a Temporary Access Pass next.'
    } 'Remove authentication methods'
})

$btnReset.Add_Click({
    & $run {
        $all = @($script:Rows | Where-Object { $_.Removable })
        if ($all.Count -eq 0) { $lblStatus.Text = 'Nothing removable.'; return }
        # Typed confirmation, because this is the one that locks a person out
        # on purpose and there is no undo.
        $typed = Show-InputDialog -Title 'Reset MFA' -Width 560 -Prompt "This removes ALL $($all.Count) removable method(s) from $($script:User.userPrincipalName). Type the user's address to confirm:"
        if ("$typed".Trim() -ne "$($script:User.userPrincipalName)") { $lblStatus.Text = 'Reset MFA cancelled -- the address did not match.'; return }
        & $removeSet $all 'AuthMethodReset' 'Every method goes. Issue a Temporary Access Pass afterwards or the user is locked out.'
    } 'Reset MFA'
})

$btnTap.Add_Click({
    & $run {
        $mins = Show-InputDialog -Title 'Temporary Access Pass' -Width 560 -Prompt "Lifetime in minutes for $($script:User.userPrincipalName) (10 to 43200). Single-use." -Default '60'
        if ("$mins".Trim() -eq '') { return }
        $n = 0
        if (-not [int]::TryParse("$mins".Trim(), [ref]$n) -or $n -lt 10 -or $n -gt 43200) { $lblStatus.Text = 'Lifetime must be a number from 10 to 43200.'; return }
        $r = New-UserTemporaryAccessPass -UserId $script:User.id -Minutes $n -Once $true
        $pass = "$($r.temporaryAccessPass)"
        Write-CompletedLog -Action 'TapCreate' -Target (Get-LogTarget $script:User.userPrincipalName) -Result 'Success' `
            -Detail "Temporary Access Pass issued: $n min, single-use, starts $(ConvertTo-LocalGraphTime $r.startDateTime)." -Environment 'Cloud'
        try { [System.Windows.Forms.Clipboard]::SetText($pass) } catch { }
        $null = Show-ToolDialog(
            "Temporary Access Pass for $($script:User.userPrincipalName):`r`n`r`n        $pass`r`n`r`nValid $n minute(s) from $(ConvertTo-LocalGraphTime $r.startDateTime), single-use. It is on the clipboard.`r`n`r`nTHIS IS THE ONLY TIME IT CAN BE SEEN. It is not stored anywhere.",
            'Temporary Access Pass', 'OK', 'Information')
        & $load
    } 'Temporary Access Pass'
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Authentication Methods' | Out-Null

$form.Add_Shown({
    $gate = Assert-ToolkitGraphScope -Scopes $script:Scopes
    if (-not $gate.Ok) {
        $lblStatus.Text = $gate.Message
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
        $btnLoad.Enabled = $false
        return
    }
    if ("$($txtUser.Text)".Trim()) { $btnLoad.PerformClick() }
})

[void]$form.ShowDialog()
$form.Dispose()
#endregion
