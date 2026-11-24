#Requires -Version 5.1
<#
    Forget-ApiCredentials.ps1
    Clear every API credential held in memory for this session.

    IT EXISTS SO THE "HELD FOR THIS SESSION" PROMISE IS SOMETHING THE OPERATOR
    CAN ACT ON rather than just be told. Somebody else needs the workstation, or
    a key was rotated mid-session -- without this the only answer is to close
    the console.

    THIS TOOL MUST NOT REQUEST ITS OWN PROCESS. In its own process it would
    clear a different global scope from the one the console and its tools
    share, and would silently do nothing.

    The directive is not spelled out here -- the console scanner matches the
    literal string even inside a comment forbidding it.

    -Environment IS ACCEPTED AND USED.
#>
# TOOLKIT-HELP-RUNS: NOTHING. Clears in-memory credentials. No system is contacted.
# TOOLKIT-VARIANT: Any
# TOOLKIT-HELP-BEGIN
# Forget the API credentials held in memory for this session.
#
# Proofpoint, SER and SendGrid credentials are asked for once per session and
# kept in memory. This clears them. The next tool that needs one asks again.
#
# IT ALSO CLEARS ANY CACHED BEARER TOKEN. A token minted from a cleared client
# secret is still a live credential -- forgetting the secret and leaving the
# token would make the promise false.
#
# NOTHING WAS EVER ON DISK, so there is nothing to clean up anywhere else. This
# only affects the running console and the tools it launches.
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

$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Forget API Credentials'
$form.ClientSize      = New-Object System.Drawing.Size(760, 420)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$grid                     = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(14, 52)
$grid.Size                = New-Object System.Drawing.Size(732, 150)
$grid.ReadOnly            = $true
$grid.AllowUserToAddRows  = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.Anchor              = 'Top,Left,Right'
[void]$form.Controls.Add($grid)

$btnForget           = New-Object System.Windows.Forms.Button
$btnForget.Text      = 'Forget all'
$btnForget.Location  = New-Object System.Drawing.Point(14, 14)
$btnForget.Size      = New-Object System.Drawing.Size(130, 28)
$btnForget.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnForget)

$btnRefresh           = New-Object System.Windows.Forms.Button
$btnRefresh.Text      = 'Refresh'
$btnRefresh.Location  = New-Object System.Drawing.Point(154, 14)
$btnRefresh.Size      = New-Object System.Drawing.Size(110, 28)
$btnRefresh.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnRefresh)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(14, 212)
$txtOut.Size       = New-Object System.Drawing.Size(732, 158)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(626, 380)
$btnClose.Size      = New-Object System.Drawing.Size(120, 28)
$btnClose.FlatStyle = 'Flat'
$btnClose.Anchor    = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
[void]$form.Controls.Add($btnClose)

function Update-View {
    # Get-ApiSecretState NEVER RETURNS A VALUE and has no parameter that could
    # make it. The console can show what is held with no code path that touches
    # a secret.
    $st = @(Get-ApiSecretState)
    $t = New-Object System.Data.DataTable
    foreach ($c in @('Service', 'Label', 'Loaded', 'Acquired', 'Token held')) { [void]$t.Columns.Add($c) }
    foreach ($r in $st) {
        [void]$t.Rows.Add($r.Service, $r.Label, $(if ($r.Loaded) { 'yes' } else { 'no' }),
            $(if ($r.AcquiredAt) { ([datetime]$r.AcquiredAt).ToString('HH:mm:ss') } else { '' }),
            $(if ($r.HasToken) { 'yes' } else { 'no' }))
    }
    $grid.DataSource = $t
    foreach ($gr in $grid.Rows) {
        if ("$($gr.Cells['Loaded'].Value)" -eq 'yes') { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGoldenrod }
    }
    return @($st | Where-Object { $_.Loaded }).Count
}

$btnRefresh.Add_Click({ try { [void](Update-View) } catch { Show-HandlerError -ErrorRecord $_ -Context 'Refresh' } })

$btnForget.Add_Click({
    try {
        $before = Update-View
        if ($before -eq 0) { $txtOut.Text = 'Nothing is held. There is nothing to clear.'; return }

        $n = Clear-ApiSecret
        [void](Update-View)

        # THE COUNT IS THE POINT. "Cleared" on its own does not tell the
        # operator whether it did anything.
        $txtOut.Text = "$n credential(s) cleared." + [Environment]::NewLine +
            'Any cached bearer token was cleared with them.' + [Environment]::NewLine + [Environment]::NewLine +
            'The next tool that needs one will ask again. Nothing was ever written to' + [Environment]::NewLine +
            'disk, so there is nothing to clean up anywhere else.'

        # The audit row records the action, not what was held.
        Write-CompletedLog -Action 'Modify' -Target (Get-LogTarget 'Forget API Credentials') -Result 'Success' `
            -Detail "Cleared $n in-memory API credential(s) and any cached tokens." -Environment $Environment
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Forget all' }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Forget API Credentials' | Out-Null

$form.Add_Shown({
    try {
        $n = Update-View
        $txtOut.Text = "$n credential(s) currently held in memory for this session." + [Environment]::NewLine +
            [Environment]::NewLine +
            'Credentials are asked for once per session and kept in memory only. This' + [Environment]::NewLine +
            'clears them, along with any cached bearer token -- a token minted from a' + [Environment]::NewLine +
            'cleared secret is still a live credential.' + [Environment]::NewLine +
            [Environment]::NewLine +
            "Environment: $(if ($Environment) { $Environment } else { '(none passed)' })"
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Forget API Credentials' }
})

[void]$form.ShowDialog()
