#Requires -Version 5.1
<#
    SRGAdminConsole.ps1
    SRG Admin Console -- Layer 1 launcher for the SRG (cloud-only,
    srgpros.com) tenant. Mailbox / Group lookup + dynamic tool menu.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

param(
    [string]$ToolsConfigPath = (Join-Path $PSScriptRoot 'SRGTools.json')
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'SRGCommon.ps1')

#region SRGTools.json load -------------------------------------------------------
function Get-ToolDefinitions {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "SRGTools.json not found at '$Path'. Use SRGToolManager.ps1 to create one, or copy the seed SRGTools.json shipped with this console."
    }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $tools = @($raw | ConvertFrom-Json -ErrorAction Stop)
    } catch {
        throw "SRGTools.json at '$Path' could not be parsed as JSON: $($_.Exception.Message)"
    }
    return @($tools | Sort-Object { if ($_.PSObject.Properties['Order']) { [int]$_.Order } else { 999 } })
}
#endregion

#region Main window ------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "SRG Admin Console"
$form.Size = New-Object System.Drawing.Size(900, 580)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Connect to Exchange Online for the SRG tenant before running."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(600, 20)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Location = New-Object System.Drawing.Point(620, 10)
$lblTenant.Size = New-Object System.Drawing.Size(252, 20)
$lblTenant.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblTenant.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

function Update-TenantLabel {
    $d = Get-ConnectedTenantDomain
    if ($d) {
        $lblTenant.Text = "Tenant: $d"
        $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
        if ($d -ine $script:MailDomain) {
            $lblTenant.Text = "Tenant: $d  (box appends @$($script:MailDomain))"
            $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(180, 40, 30)
        }
    } else {
        $lblTenant.Text = "Tenant: not connected"
        $lblTenant.ForeColor = [System.Drawing.Color]::DimGray
    }
}

$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Enter mailbox or group email address ($($script:SrgMailDomain) optional):"
$lblInput.Location = New-Object System.Drawing.Point(12, 40)
$lblInput.AutoSize = $true
$form.Controls.Add($lblInput)
$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(12, 60)
$txtInput.Size = New-Object System.Drawing.Size(860, 20)
$txtInput.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtInput)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Look Up"
$btnRun.Location = New-Object System.Drawing.Point(12, 88)
$btnRun.Size = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnRun)
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object System.Drawing.Point(130, 88)
$btnClear.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClear)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(228, 88)
$btnClose.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)
$form.CancelButton = $btnClose

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Not looked up yet."
$lblResult.Location = New-Object System.Drawing.Point(12, 126)
$lblResult.Size = New-Object System.Drawing.Size(860, 20)
$lblResult.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblResult)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.Location = New-Object System.Drawing.Point(12, 150)
$txtOut.Size = New-Object System.Drawing.Size(860, 140)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtOut.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($txtOut)

$lblTools = New-Object System.Windows.Forms.Label
$lblTools.Text = "Tools"
$lblTools.Location = New-Object System.Drawing.Point(12, 298)
$lblTools.Size = New-Object System.Drawing.Size(860, 20)
$form.Controls.Add($lblTools)

$flowTools = New-Object System.Windows.Forms.FlowLayoutPanel
$flowTools.Location = New-Object System.Drawing.Point(12, 320)
$flowTools.Size = New-Object System.Drawing.Size(860, 172)
$flowTools.AutoScroll = $true
$flowTools.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($flowTools)

$toolTip = New-Object System.Windows.Forms.ToolTip
#endregion

#region State + lookup ----------------------------------------------------------
$script:LastSmtp = $null
$script:LastType = $null
$script:AllTools = @()

try {
    $script:AllTools = Get-ToolDefinitions -Path $ToolsConfigPath
} catch {
    $script:AllTools = @()
    $txtOut.Text = $_.Exception.Message
}

function Start-ToolScript {
    param($Tool)
    $scriptPath = Join-Path $PSScriptRoot $Tool.Script
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "'$($Tool.Name)' points to '$($Tool.Script)', which does not exist under this console's folder.`n`nCheck SRGTools.json (SRGToolManager.ps1) for a stale/typo'd path.",
            'Tool Script Not Found', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    try {
        & $scriptPath -SmtpAddress $script:LastSmtp
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Launching $($Tool.Name)"
    }
}

$script:KnownObjectTypes = @('Mailbox', 'UnifiedGroup', 'DistributionGroup')

function Build-ToolButtons {
    $flowTools.Controls.Clear()
    if (@($script:AllTools).Count -eq 0) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "No tools found in SRGTools.json. Use SRGToolManager.ps1 to add some."
        $lbl.AutoSize = $true
        $flowTools.Controls.Add($lbl)
        return
    }
    foreach ($tool in $script:AllTools) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $tool.Name
        $btn.Size = New-Object System.Drawing.Size(150, 32)
        $btn.Margin = New-Object System.Windows.Forms.Padding(4)
        $btn.Tag = $tool
        $btn.Add_Click({ Start-ToolScript -Tool $this.Tag })
        $flowTools.Controls.Add($btn)
    }
}

function Update-ToolButtonState {
    param([string]$DetectedType)
    $showDimming = ($DetectedType -and ($script:KnownObjectTypes -contains $DetectedType))
    foreach ($ctrl in @($flowTools.Controls)) {
        if (-not $ctrl.Tag) { continue }
        $tool = $ctrl.Tag
        $matches = @($tool.ObjectTypes) -contains $DetectedType
        if ($showDimming -and -not $matches) {
            $ctrl.Enabled = $false
            $toolTip.SetToolTip($ctrl, "'$($tool.Name)' applies to: $(@($tool.ObjectTypes) -join ', '). The last lookup resolved to '$DetectedType', so this is disabled for it. Look up a matching object to enable it, or click Clear to release the grey-out on every tool.")
        } else {
            $ctrl.Enabled = $true
            $toolTip.SetToolTip($ctrl, '')
        }
    }
}

$btnRun.Add_Click({
    Update-TenantLabel
    $btnRun.Enabled = $false
    $txtOut.Text = ''
    $lblResult.Text = 'Looking up...'
    try {
        $smtp = Get-TargetSmtp -Text $txtInput.Text
        if (-not $smtp) {
            $script:LastSmtp = $null
            $script:LastType = $null
            $lblResult.Text = 'Enter an email address first. (Tools grey-out has been cleared.)'
            Update-ToolButtonState -DetectedType $null
            return
        }
        $target = Resolve-Target -Smtp $smtp
        $script:LastSmtp = $smtp
        $script:LastType = $target.Type
        switch ($target.Type) {
            'Mailbox'           { $lblResult.Text = "Found: Mailbox -- $smtp" }
            'UnifiedGroup'      { $lblResult.Text = "Found: Microsoft 365 Group -- $smtp" }
            'DistributionGroup' { $lblResult.Text = "Found: Distribution Group -- $smtp" }
            default {
                $lblResult.Text = "Not found as a Mailbox, Microsoft 365 Group, or Distribution Group -- $smtp"
                if ($target.ResolutionErrors.Count -gt 0) {
                    $txtOut.Text = ($target.ResolutionErrors -join "`r`n`r`n")
                }
            }
        }
        Update-ToolButtonState -DetectedType $target.Type
    } catch {
        $lblResult.Text = 'Lookup failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'SRG Admin Console lookup'
        Update-ToolButtonState -DetectedType $null
    } finally {
        $btnRun.Enabled = $true
    }
})
$btnClear.Add_Click({
    $txtInput.Text = ''
    $lblResult.Text = 'Not looked up yet.'
    $txtOut.Text = ''
    $script:LastSmtp = $null
    $script:LastType = $null
    Update-ToolButtonState -DetectedType $null
})
$btnClose.Add_Click({ $form.Close() })

Build-ToolButtons
Update-ToolButtonState -DetectedType $null
Update-TenantLabel
[void]$form.ShowDialog()
#endregion
