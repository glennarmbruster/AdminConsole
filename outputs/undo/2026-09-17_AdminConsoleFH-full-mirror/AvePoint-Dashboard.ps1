#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-RUNS: AvePoint
# TOOLKIT-HELP-BEGIN
# Executive and compliance view built from the AvePoint API.
# Read-only: it reports what AvePoint knows and changes nothing.
# TOOLKIT-HELP-END
<#
    AvePoint-Dashboard.ps1
    AvePoint Executive Dashboard and Management Tool
    
    Provides a dedicated interface for AvePoint Graph API queries including:
    - Executive Dashboard overview
    - Tenant management
    - Compliance status
    - Site management
    
    Credentials stored in environment variables:
    - AVEPOINT_CLIENT_ID
    - AVEPOINT_CLIENT_SECRET
    
    Author: Alan Smith (with Glenn Armbruster's AdminConsole)
    Date: August 27, 2026
#>

param(
    [string]$Environment
)

#region Load Common
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

# Load AvePoint query module
. (Join-Path $PSScriptRoot 'AvePointQuery.ps1')

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Form Setup
$form = New-Object System.Windows.Forms.Form
$form.Text = 'AvePoint Dashboard'
$form.ClientSize = New-Object System.Drawing.Size(1000, 700)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition = 'Manual'

# Main container
$mainContainer = New-Object System.Windows.Forms.SplitContainer
$mainContainer.Orientation = 'Vertical'
$mainContainer.Dock = 'Fill'
$mainContainer.SplitterDistance = 200
[void]$form.Controls.Add($mainContainer)

# Left panel - Navigation
$leftPanel = $mainContainer.Panel1
$leftPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

$lblNav = New-Object System.Windows.Forms.Label
$lblNav.Text = 'AvePoint'
$lblNav.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$lblNav.Location = New-Object System.Drawing.Point(10, 10)
$lblNav.Size = New-Object System.Drawing.Size(180, 25)
[void]$leftPanel.Controls.Add($lblNav)

# Navigation buttons
$buttons = @(
    @{ Text = 'Dashboard'; Tag = 'dashboard'; Y = 50 }
    @{ Text = 'Tenants'; Tag = 'tenants'; Y = 90 }
    @{ Text = 'Compliance'; Tag = 'compliance'; Y = 130 }
    @{ Text = 'Sites'; Tag = 'sites'; Y = 170 }
    @{ Text = 'Refresh'; Tag = 'refresh'; Y = 230 }
)

$navButtons = @{}
foreach ($btn in $buttons) {
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $btn.Text
    $button.Location = New-Object System.Drawing.Point(10, $btn.Y)
    $button.Size = New-Object System.Drawing.Size(170, 30)
    $button.FlatStyle = 'Flat'
    $button.Tag = $btn.Tag
    [void]$leftPanel.Controls.Add($button)
    $navButtons[$btn.Tag] = $button
}

# Right panel - Content
$rightPanel = $mainContainer.Panel2
$rightPanel.BackColor = [System.Drawing.Color]::White

# Content area
$contentBox = New-Object System.Windows.Forms.TextBox
$contentBox.Multiline = $true
$contentBox.ReadOnly = $true
$contentBox.ScrollBars = 'Both'
$contentBox.WordWrap = $false
$contentBox.Font = New-Object System.Drawing.Font('Consolas', 10)
$contentBox.Dock = 'Fill'
[void]$rightPanel.Controls.Add($contentBox)

# Status bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = 'Ready'
[void]$statusBar.Items.Add($statusLabel)
[void]$form.Controls.Add($statusBar)

#endregion

#region Functions
function Show-Dashboard {
    $statusLabel.Text = 'Loading Executive Dashboard...'
    $form.Refresh()
    
    try {
        $dashboard = Get-AvePointExecutiveDashboard
        
        if ($dashboard) {
            $output = @"
========================================================================
  AvePoint Executive Dashboard
========================================================================

Dashboard Data:
$(ConvertTo-Json $dashboard -Depth 3)

Retrieved: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@
            $contentBox.Text = $output
            $statusLabel.Text = 'Dashboard loaded successfully'
        } else {
            $contentBox.Text = "Failed to load dashboard. Check credentials and API access."
            $statusLabel.Text = 'Error loading dashboard'
        }
    } catch {
        $contentBox.Text = "Error: $($_.Exception.Message)"
        $statusLabel.Text = 'Error'
    }
}

function Show-Tenants {
    $statusLabel.Text = 'Loading Tenants...'
    $form.Refresh()
    
    try {
        $tenants = Get-AvePointTenants
        
        if ($tenants) {
            $output = @"
========================================================================
  AvePoint Managed Tenants
========================================================================

Tenants:
$(ConvertTo-Json $tenants -Depth 3)

Count: $(@($tenants).Count)
Retrieved: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@
            $contentBox.Text = $output
            $statusLabel.Text = "Loaded $(@($tenants).Count) tenant(s)"
        } else {
            $contentBox.Text = "No tenants found or error retrieving tenants."
            $statusLabel.Text = 'No tenants'
        }
    } catch {
        $contentBox.Text = "Error: $($_.Exception.Message)"
        $statusLabel.Text = 'Error'
    }
}

function Show-Compliance {
    $statusLabel.Text = 'Loading Compliance Status...'
    $form.Refresh()
    
    try {
        $tenants = Get-AvePointTenants
        
        if ($tenants) {
            $output = "========================================================================`n"
            $output += "  AvePoint Compliance Status`n"
            $output += "========================================================================`n`n"
            
            foreach ($tenant in $tenants) {
                $compliance = Get-AvePointComplianceStatus -TenantId $tenant.id
                $output += "Tenant: $($tenant.name)`n"
                $output += "ID: $($tenant.id)`n"
                if ($compliance) {
                    $output += "$(ConvertTo-Json $compliance -Depth 2)`n`n"
                } else {
                    $output += "No compliance data available`n`n"
                }
            }
            
            $output += "Retrieved: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            $contentBox.Text = $output
            $statusLabel.Text = 'Compliance status loaded'
        } else {
            $contentBox.Text = "No tenants found."
            $statusLabel.Text = 'No tenants'
        }
    } catch {
        $contentBox.Text = "Error: $($_.Exception.Message)"
        $statusLabel.Text = 'Error'
    }
}

function Show-Sites {
    $statusLabel.Text = 'Loading Sites...'
    $form.Refresh()
    
    try {
        $sites = Get-AvePointSites
        
        if ($sites) {
            $output = @"
========================================================================
  AvePoint Managed SharePoint Sites
========================================================================

Sites:
$(ConvertTo-Json $sites -Depth 3)

Count: $(@($sites).Count)
Retrieved: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@
            $contentBox.Text = $output
            $statusLabel.Text = "Loaded $(@($sites).Count) site(s)"
        } else {
            $contentBox.Text = "No sites found or error retrieving sites."
            $statusLabel.Text = 'No sites'
        }
    } catch {
        $contentBox.Text = "Error: $($_.Exception.Message)"
        $statusLabel.Text = 'Error'
    }
}

#endregion

#region Event Handlers
$navButtons['dashboard'].Add_Click({ Show-Dashboard })
$navButtons['tenants'].Add_Click({ Show-Tenants })
$navButtons['compliance'].Add_Click({ Show-Compliance })
$navButtons['sites'].Add_Click({ Show-Sites })
$navButtons['refresh'].Add_Click({
    $statusLabel.Text = 'Refreshing...'
    $contentBox.Text = ''
    $statusLabel.Text = 'Ready'
})

#endregion

#region Theme and Display
Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'AvePoint Dashboard' | Out-Null

$form.Add_Shown({
    try {
        $credStatus = if ($env:AVEPOINT_CLIENT_ID) { "[OK] AvePoint credentials configured" } else { "[ERROR] AvePoint credentials NOT configured" }
        $contentBox.Text = @"
========================================================================
  AvePoint Executive Dashboard
  FH-PR2323-ExecDashboard
========================================================================

Welcome to the AvePoint Dashboard tool.

Select an option from the left panel:

  Dashboard    - View executive dashboard overview
  Tenants      - List all managed tenants
  Compliance   - Check compliance status
  Sites        - View managed SharePoint sites
  Refresh      - Clear current view

Credentials Status:
  $credStatus

For more information, see:
  https://learn.avepoint.com/graphapi/docs/Overview.html
"@
        Write-CompletedLog -Action 'Open' -Target (Get-LogTarget 'AvePoint Dashboard') -Result 'Success' `
            -Detail 'AvePoint Dashboard opened' -Environment $Environment
    } catch {
        Show-HandlerError -ErrorRecord $_ -Context 'Form initialization'
    }
})

$form.ShowDialog() | Out-Null

#endregion
