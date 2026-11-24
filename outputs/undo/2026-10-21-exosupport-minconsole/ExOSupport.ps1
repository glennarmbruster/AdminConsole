#Requires -Version 5.1
<#
    ExOSupport.ps1  --  the ExO Support console.

    A narrow, cloud-only console for support staff who look after shared
    mailboxes and distribution groups once those are managed in Exchange
    Online. One window:

        top        Activate (PIM) and Connect ExO
        middle     Cloud Lookup -- the focus. Search, pick a row, open the
                   Shared Mailbox Editor or the Group Editor
        bottom     status strip: PIM activation time and time left, and who
                   is signed in

    Start it with ExOSupport.cmd, or:
        powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File ExOSupport.ps1

    Files beside this one:
        ExOSupport.Common.ps1    config, tenant, connection, visibility, PIM, log
        ExOSupport.Lookup.ps1    the Cloud Lookup panel in the main window
        ExOSupport.Editors.ps1   the Shared Mailbox Editor and the Group Editor
        Config\ExOSupport.json   which tenant, its domain, team groups, PIM
        Logs\                    one CSV a month of every change made

    Author   : Glenn Armbruster
    Created  : 2026-10-20
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

. (Join-Path $PSScriptRoot 'ExOSupport.Common.ps1')
. (Join-Path $PSScriptRoot 'ExOSupport.Lookup.ps1')
. (Join-Path $PSScriptRoot 'ExOSupport.Editors.ps1')

$script:ExsBuild = '2026.10.21.5'
$bad = @()
foreach ($part in 'Common', 'Lookup', 'Editors') {
    $have = $(if ($script:ExsParts) { "$($script:ExsParts[$part])" } else { '' })
    if ($have -ne $script:ExsBuild) { $bad += "ExOSupport.$part.ps1 is $(if ($have) { "build $have" } else { 'an older build' })" }
}
if ($bad.Count) {
    [void][System.Windows.Forms.MessageBox]::Show(("The ExO Support files are from different versions:`r`n`r`n  " + ($bad -join "`r`n  ") +
        "`r`n`r`nExOSupport.ps1 is build $($script:ExsBuild). Copy the WHOLE ExOSupport folder again, then start it."), 'ExO Support', 'OK', 'Error')
    return
}

try {
    $cfg = Read-ExsConfig
    $chosen = Select-ExsTenant -Config $cfg
    if (-not $chosen) { return }
    Initialize-ExsTenant -Tenant $chosen -Config $cfg
} catch {
    [void][System.Windows.Forms.MessageBox]::Show("ExO Support could not start:`r`n`r`n$($_.Exception.Message)", 'ExO Support', 'OK', 'Error')
    return
}

# WHO IS WORKING. Asked once at launch (pre-filled with last time's answer);
# used for the support team lookup and written to every log line.
if (-not (Show-ExsOperatorPrompt)) { return }

#region Status updaters ----------------------------------------------------------
# Read the script-scope controls of whichever window pass is current.

function Update-ExsPimLabel {
    $st = Get-ExsPimStatus
    $script:ssPim.Text = $st.Text
    $script:ssPim.ForeColor = $st.Colour
}

function Update-ExsConnLabel {
    $c = Get-ExsConnection
    if ($c.Connected) {
        $script:lblConn.Text = 'Connected to Exchange Online'   # who: status strip
        $script:lblConn.ForeColor = [System.Drawing.Color]::ForestGreen
        $script:btnConnect.Text = 'Reconnect ExO'
        $script:ssConn.Text = "$($script:OperatorEmail)   |   admin: $($c.Upn)"
        $script:ssConn.ForeColor = [System.Drawing.Color]::ForestGreen
    } else {
        $script:lblConn.Text = 'Not connected -- Activate (PIM), then Connect ExO.'
        $script:lblConn.ForeColor = [System.Drawing.Color]::Firebrick
        $script:btnConnect.Text = 'Connect ExO'
        $script:ssConn.Text = "$($script:OperatorEmail)   |   Exchange Online: not connected"
        $script:ssConn.ForeColor = [System.Drawing.Color]::Firebrick
        # A dropped session leaves no operator; the visibility line says so.
        if ($script:Operator -and $script:Operator.Upn) { $script:Operator = $null }
    }
    $script:Lk.Vis.Text = Get-ExsVisibilityText
    $script:Lk.Vis.ForeColor = $(if ((Test-ExsFilteringOn) -and $script:Operator -and $script:Operator.Upn -and @($script:Operator.Values).Count -eq 0) {
            [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::Black })
    return $c
}
#endregion

#region Main loop ---------------------------------------------------------------
<#
    THE WINDOW CLOSES TO SIGN IN, AND REOPENS AFTERWARDS.

    Connect ExO does not sign in from the button. It records the request and
    closes the window; the sign-in then runs below, after ShowDialog() has
    returned and no modal window owns this thread; then the loop builds the
    window again. Signing in from inside the open window is what made the
    button spin forever -- the sign-in prompt is a window on the same thread
    and cannot take input while a modal loop holds it. Same process
    throughout, so the session the sign-in creates is the one the console
    uses. This is the Admin Console's PendingConnect pattern, unchanged.
#>
$script:PendingConnect = $false
$script:Notice     = $null     # message to show once the window is back
$script:FormAt     = $null     # where the window was, so it reopens there
$script:FormSize   = $null     # and how big
$script:LastStatus = $null     # status text to show once the window is back
$script:LastSearch = ''        # re-run after a reconnect

while ($true) {
    $script:PendingConnect = $false

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "ExO Support  --  $($script:Tenant.DisplayName)"
    $form.ClientSize = New-Object System.Drawing.Size(980, 660)
    $form.MinimumSize = New-Object System.Drawing.Size(820, 540)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    if ($script:FormAt) {
        $form.StartPosition = 'Manual'; $form.Location = $script:FormAt
        if ($script:FormSize) { $form.Size = $script:FormSize }
    } else { $form.StartPosition = 'CenterScreen' }
    $W = $form.ClientSize.Width

    # Header band.
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Location = New-Object System.Drawing.Point(0, 0)
    $hdr.Size = New-Object System.Drawing.Size($W, 56)
    $hdr.Anchor = 'Top,Left,Right'
    $hdr.BackColor = [System.Drawing.Color]::FromArgb(31, 58, 95)
    $form.Controls.Add($hdr)
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = 'ExO Support'
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
    $lblTitle.Location = New-Object System.Drawing.Point(14, 5)
    $lblTitle.AutoSize = $true
    $hdr.Controls.Add($lblTitle)
    $lblTenant = New-Object System.Windows.Forms.Label
    $lblTenant.Text = "$($script:Tenant.DisplayName)   |   shared mailboxes and groups in Exchange Online"
    $lblTenant.ForeColor = [System.Drawing.Color]::Gainsboro
    $lblTenant.Location = New-Object System.Drawing.Point(16, 35)
    $lblTenant.Size = New-Object System.Drawing.Size(($W - 32), 18)
    $lblTenant.Anchor = 'Top,Left,Right'
    $hdr.Controls.Add($lblTenant)

    # Top bar: Activate and Connect, with the connection state beside them.
    $btnActivate = New-ExsButton $form 'Activate (PIM)' 16 68 150 34
    $script:btnConnect = New-ExsButton $form 'Connect ExO' 174 68 150 34
    $script:lblConn = New-ExsLabel $form '' 336 76 ($W - 470) 20 -Bold
    $script:lblConn.Anchor = 'Top,Left,Right'
    $lnkEmail = New-Object System.Windows.Forms.LinkLabel
    $lnkEmail.Text = 'Change my email'
    $lnkEmail.Location = New-Object System.Drawing.Point(($W - 126), 78)
    $lnkEmail.Size = New-Object System.Drawing.Size(110, 20)
    $lnkEmail.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $lnkEmail.Anchor = 'Top,Right'
    $form.Controls.Add($lnkEmail)
    $rule = New-Object System.Windows.Forms.Label
    $rule.BorderStyle = 'Fixed3D'
    $rule.Location = New-Object System.Drawing.Point(16, 112)
    $rule.Size = New-Object System.Drawing.Size(($W - 32), 2)
    $rule.Anchor = 'Top,Left,Right'
    $form.Controls.Add($rule)

    # Status strip at the bottom: PIM activation time on the left, the signed-in
    # account on the right.
    $ss = New-Object System.Windows.Forms.StatusStrip
    $ss.SizingGrip = $true
    $script:ssPim = New-Object System.Windows.Forms.ToolStripStatusLabel
    $script:ssPim.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $script:ssConn = New-Object System.Windows.Forms.ToolStripStatusLabel
    $script:ssConn.Spring = $true
    $script:ssConn.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    [void]$ss.Items.Add($script:ssPim)
    [void]$ss.Items.Add($script:ssConn)
    $form.Controls.Add($ss)

    # The middle: Cloud Lookup, which is what the window is for.
    $btnY = Add-ExsLookupPanel -Form $form -Top 124 -Bottom ($ss.Height + 4)
    $btnClose = New-ExsButton $form 'Close' ($W - 106) $btnY 90 36
    $btnClose.Anchor = 'Bottom,Right'
    $btnClose.Add_Click({ $form.Close() })

    $btnActivate.Add_Click({
            try { Start-Process $script:Tenant.PimUrl }
            catch { [void](Show-ExsMessage -Text "Could not open the PIM portal: $($_.Exception.Message)" -Icon 'Error'); return }
            $a = Show-ExsMessage -Title 'Activate (PIM)' -Icon 'Information' -Buttons 'OKCancel' -Text (
                "The PIM portal is opening in your browser.`r`n`r`n" +
                "Activate your Exchange role there. When it shows as active, press OK here -- that records the activation time " +
                "and starts the countdown at the bottom of this window.`r`n`r`nPress Cancel if you did not activate.")
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
            Set-ExsPimActivated
            Update-ExsPimLabel
            Write-ExsLog -Action 'PimMarked' -Target 'self' -Detail "activation recorded ($($script:Tenant.PimHours)h)" -Result 'Info'
            # Activate and Connect are independent (operator's call, 2026-10-21).
            # Only a hint: a session opened before activating lacks the role.
            if ((Get-ExsConnection).Connected) {
                $script:Lk.Count.Text = 'Activation recorded. Your current session was opened before it -- press Reconnect ExO if edits are refused.'
            }
        })

    $lnkEmail.Add_LinkClicked({
            $was = $script:OperatorEmail
            $now = Show-ExsOperatorPrompt -Owner $form
            if (-not $now -or $now -eq $was) { return }
            Write-ExsLog -Action 'OperatorEmail' -Target 'self' -Detail "changed from '$was' to '$now'" -Result 'Info'
            if ((Get-ExsConnection).Connected) {
                $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                try {
                    $op = Resolve-ExsOperatorScope
                    if (@($op.Notes).Count) { [void](Show-ExsMessage -Text (@($op.Notes) -join "`r`n`r`n") -Icon 'Warning') }
                } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
            }
            [void](Update-ExsConnLabel)
            $script:Lk.Grid.Rows.Clear()
            $script:Lk.Count.Text = 'Email changed. Search again.'
        })

    # RECORD AND CLOSE -- the sign-in itself runs after ShowDialog returns.
    $script:btnConnect.Add_Click({
            $script:PendingConnect = $true
            $form.Close()
        })

    # PIM countdown and session state every 30 seconds. Both are local reads.
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 30000
    $timer.Add_Tick({ try { Update-ExsPimLabel; [void](Update-ExsConnLabel) } catch { } })

    $form.Add_Shown({
            $form.Activate()
            $script:Lk.Txt.Focus()
            if ($script:Notice) {
                $n = $script:Notice; $script:Notice = $null
                [void](Show-ExsMessage -Text $n.Text -Title 'Connect ExO' -Icon $n.Icon)
            }
            # Back from a reconnect: put the last search back and run it again.
            if ($script:LastSearch -and (Get-ExsConnection).Connected) {
                $script:Lk.Txt.Text = $script:LastSearch
                Invoke-ExsLookupSearch
            }
            $script:Lk.Txt.Focus()
        })
    $form.Add_FormClosing({
            if ($form.WindowState -eq 'Normal') { $script:FormAt = $form.Location; $script:FormSize = $form.Size }
        })

    Update-ExsPimLabel
    $c = Update-ExsConnLabel
    if ($script:LastStatus) { $script:Lk.Count.Text = $script:LastStatus; $script:LastStatus = $null }
    elseif (-not $c.Connected) { $script:Lk.Count.Text = 'Activate (PIM), then Connect ExO, then search.' }
    if ($script:ExsScopeFilterLocked) { $form.Text = "$($form.Text)   [team filtering on]" }
    $timer.Start()
    [void]$form.ShowDialog()
    $timer.Stop(); $timer.Dispose()
    $form.Dispose()

    if (-not $script:PendingConnect) { break }

    # ---- No window is open from here to the end of the loop. ----
    Write-Host ''
    Write-Host '=== Connect to Exchange Online ===' -ForegroundColor Cyan
    $res = Connect-ExsExchangeOnline
    if ($res.Ok) {
        Write-Host $res.Detail -ForegroundColor Green
        Write-Host 'Checking your support team groups...' -ForegroundColor DarkGray
        $op = $null
        try { $op = Resolve-ExsOperatorScope } catch { }
        if ($op) {
            Write-ExsLog -Action 'Connect' -Target 'Exchange Online' -Result 'Info' `
                -Detail "teams: $(if (@($op.Teams).Count) { @($op.Teams) -join ', ' } else { 'none' }); filtering $(if (Test-ExsFilteringOn) { 'on' } else { 'off' })"
            if (@($op.Notes).Count) {
                $script:Notice = [PSCustomObject]@{ Text = (@($op.Notes) -join "`r`n`r`n"); Icon = 'Warning' }
                $script:LastStatus = 'Connected, with notes.'
            } else { $script:LastStatus = 'Connected. Type a name or address and press Enter.' }
        }
    } else {
        Write-Host $res.Detail -ForegroundColor Red
        Write-ExsLog -Action 'Connect' -Target 'Exchange Online' -Detail $res.Detail -Result 'Failed'
        $script:Notice = [PSCustomObject]@{ Text = $res.Detail; Icon = 'Warning' }
        $script:LastStatus = 'Not connected.'
    }
    Write-Host 'Reopening ExO Support...' -ForegroundColor DarkGray
}
#endregion
