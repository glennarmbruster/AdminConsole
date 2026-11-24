#Requires -Version 5.1
<#
    ExOSupport.ps1  --  the ExO Support console.

    A narrow, cloud-only console for support staff who look after shared
    mailboxes and distribution groups once those are managed in Exchange
    Online. It has four things on it:

        PIM countdown     activate the Exchange role in the portal, then mark it
        Connect ExO       sign this window in to Exchange Online as yourself
        Visibility        which objects you are responsible for, from your
                          support team group(s)
        Cloud Lookup      find a shared mailbox or distribution group and open
                          the Shared Mailbox Editor or the Group Editor

    Start it with ExOSupport.cmd, or:
        powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File ExOSupport.ps1

    Files beside this one:
        ExOSupport.Common.ps1    config, tenant, connection, visibility, PIM, log
        ExOSupport.Lookup.ps1    the Cloud Lookup window
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

try {
    $cfg = Read-ExsConfig
    $chosen = Select-ExsTenant -Config $cfg
    if (-not $chosen) { return }
    Initialize-ExsTenant -Tenant $chosen -Config $cfg
} catch {
    [void][System.Windows.Forms.MessageBox]::Show("ExO Support could not start:`r`n`r`n$($_.Exception.Message)", 'ExO Support', 'OK', 'Error')
    return
}

#region Window ------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "ExO Support  --  $($script:Tenant.DisplayName)"
$form.ClientSize = New-Object System.Drawing.Size(620, 404)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# Header band.
$hdr = New-Object System.Windows.Forms.Panel
$hdr.Location = New-Object System.Drawing.Point(0, 0)
$hdr.Size = New-Object System.Drawing.Size(620, 58)
$hdr.BackColor = [System.Drawing.Color]::FromArgb(31, 58, 95)
$form.Controls.Add($hdr)
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'ExO Support'
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(14, 6)
$lblTitle.AutoSize = $true
$hdr.Controls.Add($lblTitle)
$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Text = "$($script:Tenant.DisplayName)   |   cloud only   |   shared mailboxes and distribution groups"
$lblTenant.ForeColor = [System.Drawing.Color]::Gainsboro
$lblTenant.Location = New-Object System.Drawing.Point(16, 36)
$lblTenant.Size = New-Object System.Drawing.Size(590, 18)
$hdr.Controls.Add($lblTenant)

# Step 1 -- PIM.
[void](New-ExsLabel $form '1.  Activate your Exchange role (PIM)' 14 72 400 20 -Bold)
$lblPim = New-ExsLabel $form '' 34 96 570 20 -Bold
$btnPimOpen  = New-ExsButton $form 'Open PIM portal' 34 120 150 30
$btnPimMark  = New-ExsButton $form 'Mark activated' 192 120 130 30
$btnPimClear = New-ExsButton $form 'Clear' 330 120 80 30
[void](New-ExsLabel $form 'Activate first, then connect: a session opened before activating does not get the role.' 34 154 570 18 -Muted)

# Step 2 -- Connect.
[void](New-ExsLabel $form '2.  Connect to Exchange Online' 14 184 400 20 -Bold)
$lblConn = New-ExsLabel $form '' 34 208 420 20 -Bold
$btnConnect = New-ExsButton $form 'Connect ExO' 462 202 142 32
$lblVis = New-ExsLabel $form '' 34 236 570 36
$lblVis.Size = New-Object System.Drawing.Size(570, 36)

# Step 3 -- Lookup.
[void](New-ExsLabel $form '3.  Find a shared mailbox or distribution group' 14 280 450 20 -Bold)
$btnLookup = New-ExsButton $form 'Cloud Lookup' 34 304 570 44
$btnLookup.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)

$lblStatus = New-ExsLabel $form '' 14 362 500 20 -Muted
$btnClose = New-ExsButton $form 'Close' 524 358 80 30
$btnClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $btnClose
#endregion

#region Behaviour ---------------------------------------------------------------

function Update-ExsPimLabel {
    $s = Get-ExsPimStatus
    $lblPim.Text = $s.Text
    $lblPim.ForeColor = $s.Colour
}

function Update-ExsConnLabel {
    $c = Get-ExsConnection
    if ($c.Connected) {
        $lblConn.Text = $c.Detail
        $lblConn.ForeColor = [System.Drawing.Color]::ForestGreen
        $btnConnect.Text = 'Reconnect ExO'
    } else {
        $lblConn.Text = $c.Detail
        $lblConn.ForeColor = [System.Drawing.Color]::Firebrick
        $btnConnect.Text = 'Connect ExO'
        # A dropped session leaves no operator; the visibility line says so.
        if ($script:Operator -and $script:Operator.Upn) { $script:Operator = $null }
    }
    $lblVis.Text = Get-ExsVisibilityText
    $lblVis.ForeColor = $(if ((Test-ExsFilteringOn) -and $script:Operator -and $script:Operator.Upn -and @($script:Operator.Values).Count -eq 0) {
            [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::Black })
    $btnLookup.Enabled = $c.Connected
    return $c
}

$btnPimOpen.Add_Click({
        try { Start-Process $script:Tenant.PimUrl } catch { [void](Show-ExsMessage -Text "Could not open the portal: $($_.Exception.Message)" -Icon 'Error') }
    })
$btnPimMark.Add_Click({
        Set-ExsPimActivated
        Update-ExsPimLabel
        Write-ExsLog -Action 'PimMarked' -Target 'self' -Detail "countdown started ($($script:Tenant.PimHours)h)" -Result 'Info'
        if ((Get-ExsConnection).Connected) {
            $lblStatus.Text = 'Already connected? Press Reconnect ExO so the session picks up the activated role.'
        }
    })
$btnPimClear.Add_Click({ Set-ExsPimActivated -Clear; Update-ExsPimLabel })

$btnConnect.Add_Click({
        $btnConnect.Enabled = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $lblStatus.Text = 'Signing in -- complete the sign-in window, including MFA.'
            $form.Refresh()
            if ((Get-ExsConnection).Connected) {
                try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
            }
            $res = Connect-ExsExchangeOnline
            if (-not $res.Ok) {
                $lblStatus.Text = 'Not connected.'
                [void](Show-ExsMessage -Text $res.Detail -Title 'Connect ExO' -Icon 'Warning')
                Write-ExsLog -Action 'Connect' -Target 'Exchange Online' -Detail $res.Detail -Result 'Failed'
                return
            }
            $lblStatus.Text = 'Connected. Checking your support team groups...'
            $form.Refresh()
            $op = Resolve-ExsOperatorScope
            $lblStatus.Text = $(if ($op.Notes.Count) { "Connected, with notes: $($op.Notes -join ' ')" } else { 'Connected.' })
            Write-ExsLog -Action 'Connect' -Target 'Exchange Online' -Result 'Info' `
                -Detail "teams: $(if ($op.Teams.Count) { $op.Teams -join ', ' } else { 'none' }); filtering $(if (Test-ExsFilteringOn) { 'on' } else { 'off' })"
            if ($op.Notes.Count) { [void](Show-ExsMessage -Text ($op.Notes -join "`r`n`r`n") -Title 'Connect ExO' -Icon 'Warning') }
        } finally {
            [void](Update-ExsConnLabel)
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnConnect.Enabled = $true
        }
    })

$btnLookup.Add_Click({
        $c = Update-ExsConnLabel
        if (-not $c.Connected) { [void](Show-ExsMessage -Text 'Connect to Exchange Online first.' -Icon 'Warning'); return }
        if (-not $script:Operator -or $script:Operator.Upn -ne $c.Upn) { [void](Resolve-ExsOperatorScope); [void](Update-ExsConnLabel) }
        Show-ExsLookup -Owner $form
    })

# The PIM countdown and the session state, every 30 seconds. Both are local
# reads (a file and Get-ConnectionInformation), so the timer costs nothing.
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 30000
$timer.Add_Tick({ try { Update-ExsPimLabel; [void](Update-ExsConnLabel) } catch { } })
$form.Add_FormClosed({ $timer.Stop(); $timer.Dispose() })

Update-ExsPimLabel
[void](Update-ExsConnLabel)
if ($script:ExsScopeFilterLocked) { $lblStatus.Text = 'Team filtering is locked on in this build.' }
$timer.Start()
[void]$form.ShowDialog()
#endregion
