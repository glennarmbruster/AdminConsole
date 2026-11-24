#Requires -Version 5.1
<#
    M365Reports.ps1
    Launcher for the Microsoft 365 admin portal reports -- the ones Microsoft already computes better than a script could. Author : Glenn Armbruster

    Reasoning: ScriptNotes/M365Reports.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Buttons that open the Microsoft 365 admin portal reports -- usage, activity,
# licensing, security -- in the browser.
#
# These are reports Microsoft already computes, from data a script cannot get at
# directly. Rather than reproduce them badly, this puts them one click away.
#
# It opens each in the elevated Edge profile, so you are signed in as the right
# account when you land.
#
# Nothing is read or changed by this tool. It opens links.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration --------------------------------------------------------
$Links = @(
    # --- Usage and storage ------------------------------------------------
    [PSCustomObject]@{
        Group = 'Usage and storage'; Label = 'Usage reports (all)'
        Url = 'https://go.microsoft.com/fwlink/p/?linkid=2074756'
        Durability = 'fwlink'
        Note = 'Root of every usage report. Exchange, OneDrive, SharePoint, Teams.'
    }
    [PSCustomObject]@{
        Group = 'Usage and storage'; Label = 'Mailbox usage and quota'
        Url = 'https://admin.microsoft.com/#/reportsUsage/MailboxUsage'
        Durability = 'direct'
        Note = 'Storage used, quota band, item count, archive state, per mailbox. Exportable.'
    }
    [PSCustomObject]@{
        Group = 'Usage and storage'; Label = 'Email activity'
        Url = 'https://admin.microsoft.com/#/reportsUsage/EmailActivity'
        Durability = 'direct'
        Note = 'Send/read/meeting activity per user over 7-180 days.'
    }
    [PSCustomObject]@{
        Group = 'Usage and storage'; Label = 'Email apps usage'
        Url = 'https://admin.microsoft.com/#/reportsUsage/EmailAppUsage'
        Durability = 'direct'
        Note = 'Which clients people actually use -- pairs with the Protocols report.'
    }
    [PSCustomObject]@{
        Group = 'Usage and storage'; Label = 'OneDrive usage'
        Url = 'https://admin.microsoft.com/#/reportsUsage/ODBSiteUsage'
        Durability = 'direct'
        Note = 'Per-site storage and file counts.'
    }
    [PSCustomObject]@{
        Group = 'Usage and storage'; Label = 'SharePoint usage'
        Url = 'https://admin.microsoft.com/#/reportsUsage/SPSiteUsage'
        Durability = 'direct'
        Note = 'Site storage, activity, external sharing.'
    }

    # --- Licences and accounts -------------------------------------------
    [PSCustomObject]@{
        Group = 'Licences and accounts'; Label = 'Licences'
        Url = 'https://admin.microsoft.com/#/licenses'
        Durability = 'direct'
        Note = 'Assigned vs available per SKU. The authoritative count.'
    }
    [PSCustomObject]@{
        Group = 'Licences and accounts'; Label = 'Active users'
        Url = 'https://admin.microsoft.com/#/reportsUsage/ActiveUsers'
        Durability = 'direct'
        Note = 'Activity by service. Cross-check against the stale accounts report.'
    }
    [PSCustomObject]@{
        Group = 'Licences and accounts'; Label = 'All users'
        Url = 'https://admin.microsoft.com/#/users'
        Durability = 'direct'
        Note = 'User list with licence and sign-in status.'
    }
    [PSCustomObject]@{
        Group = 'Licences and accounts'; Label = 'Entra sign-in logs'
        Url = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/SignInEventsV3'
        Durability = 'direct'
        Note = 'The accurate sign-in source. Better than lastLogonTimestamp.'
    }
    [PSCustomObject]@{
        Group = 'Licences and accounts'; Label = 'Restricted users'
        Url = 'https://security.microsoft.com/restrictedentities'
        Durability = 'direct'
        Note = 'Accounts blocked from SENDING mail after tripping outbound limits -- the first sign of a compromise. Microsoft titles the page "Restricted entities". Defender portal, not admin center.'
    }

    # --- Mail flow and security ------------------------------------------
    [PSCustomObject]@{
        Group = 'Mail flow and security'; Label = 'Exchange admin center'
        Url = 'https://admin.exchange.microsoft.com'
        Durability = 'direct'
        Note = 'EAC home.'
    }
    [PSCustomObject]@{
        Group = 'Mail flow and security'; Label = 'Message trace'
        Url = 'https://admin.exchange.microsoft.com/#/messagetrace'
        Durability = 'direct'
        Note = 'Longer retention than the PowerShell trace tool.'
    }
    [PSCustomObject]@{
        Group = 'Mail flow and security'; Label = 'Threat Explorer'
        Url = 'https://security.microsoft.com/threatexplorerv3'
        Durability = 'direct'
        Note = 'Where message trace stops. Delivery verdict, phish/malware detections, URL clicks. Defender for Office 365 Plan 2 -- P1 tenants land on Real-time detections instead.'
    }
    [PSCustomObject]@{
        Group = 'Mail flow and security'; Label = 'Mail flow reports'
        Url = 'https://admin.exchange.microsoft.com/#/reports/mailflowreportsmain'
        Durability = 'direct'
        Note = 'Volume, top senders, forwarding, spoof.'
    }
    [PSCustomObject]@{
        Group = 'Mail flow and security'; Label = 'Defender portal'
        Url = 'https://security.microsoft.com'
        Durability = 'direct'
        Note = 'Threat Explorer, quarantine, incidents.'
    }
    [PSCustomObject]@{
        Group = 'Mail flow and security'; Label = 'Quarantine'
        Url = 'https://security.microsoft.com/quarantine'
        Durability = 'direct'
        Note = 'Tenant-wide quarantine.'
    }

    # --- Service and audit -----------------------------------------------
    [PSCustomObject]@{
        Group = 'Service and audit'; Label = 'Service health'
        Url = 'https://admin.microsoft.com/#/servicehealth'
        Durability = 'direct'
        Note = 'Check before debugging anything that broke this morning.'
    }
    [PSCustomObject]@{
        Group = 'Service and audit'; Label = 'Audit log search'
        Url = 'https://purview.microsoft.com/audit/auditsearch'
        Durability = 'direct'
        Note = 'Who changed what. Finds who set a forward, or when.'
    }
    [PSCustomObject]@{
        Group = 'Service and audit'; Label = 'Tenant Allow/Block List'
        Url = 'https://security.microsoft.com/tenantAllowBlockList'
        Durability = 'direct'
        Note = 'Manual overrides of the filtering verdict -- senders, domains, URLs, files, IPv6. Blocks beat allows. Check here first when legitimate mail is vanishing.'
    }
    [PSCustomObject]@{
        Group = 'Service and audit'; Label = 'Message center'
        Url = 'https://admin.microsoft.com/#/MessageCenter'
        Durability = 'direct'
        Note = 'Upcoming changes. Where deprecations are announced.'
    }
    [PSCustomObject]@{
        Group = 'Service and audit'; Label = 'M365 admin center'
        Url = 'https://go.microsoft.com/fwlink/p/?linkid=2024339'
        Durability = 'fwlink'
        Note = 'Portal home.'
    }
)
#endregion

#region UI -------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Microsoft 365 Reports'
$form.Size            = New-Object System.Drawing.Size(700, 700)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Opens the portal reports Microsoft already computes -- mailbox sizes, quotas, activity and audit.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(660, 18)
$form.Controls.Add($lblInfo)

$lblTenant           = New-Object System.Windows.Forms.Label
$lblTenant.Tag       = 'NoTheme'
$lblTenant.Location  = New-Object System.Drawing.Point(12, 32)
$lblTenant.Size      = New-Object System.Drawing.Size(660, 18)
$lblTenant.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

$lblProfile           = New-Object System.Windows.Forms.Label
$lblProfile.Tag       = 'NoTheme'
$lblProfile.Location  = New-Object System.Drawing.Point(12, 52)
$lblProfile.Size      = New-Object System.Drawing.Size(660, 18)
$lblProfile.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblProfile)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 20000

$y = 78
foreach ($g in ($Links | Select-Object -ExpandProperty Group -Unique)) {
    $items = @($Links | Where-Object { $_.Group -eq $g })
    $rows  = [Math]::Ceiling($items.Count / 3.0)

    $grp          = New-Object System.Windows.Forms.GroupBox
    $grp.Text     = $g
    $grp.Location = New-Object System.Drawing.Point(12, $y)
    $grp.Size     = New-Object System.Drawing.Size(660, (26 + ($rows * 36)))
    $form.Controls.Add($grp)

    $i = 0
    foreach ($lnk in $items) {
        $col = $i % 3
        $row = [Math]::Floor($i / 3)
        $b          = New-Object System.Windows.Forms.Button
        $b.Text     = $lnk.Label
        $b.Location = New-Object System.Drawing.Point((12 + ($col * 214)), (22 + ($row * 36)))
        $b.Size     = New-Object System.Drawing.Size(206, 30)
        $b.Tag      = $lnk.Url
        $tip = $lnk.Note
        if ($lnk.Durability -eq 'direct') {
            $tip += '  [direct portal URL -- may change]'
        }
        $toolTip.SetToolTip($b, "$tip`r`n$($lnk.Url)")
        $b.Add_Click({
            try {
    # Purpose 'Admin': these portals need the elevated account, and Start-Process
    # would land in whichever profile Edge used last.
                $st = Start-UrlInBrowserProfile -Url "$($this.Tag)" -Purpose 'Admin'
                $lblStatus.Text = "$($this.Text) -- $st"
            } catch {
                $lblStatus.Text = "Could not open: $($_.Exception.Message)"
            }
        })
        $grp.Controls.Add($b)
        $i++
    }
    $y += $grp.Height + 8
}

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Text      = 'Hover a button to see what it shows and where it goes.'
$lblStatus.Location  = New-Object System.Drawing.Point(12, ($y + 4))
$lblStatus.Size      = New-Object System.Drawing.Size(540, 18)
$form.Controls.Add($lblStatus)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(562, $y)
$btnClose.Size     = New-Object System.Drawing.Size(110, 28)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$form.ClientSize = New-Object System.Drawing.Size(684, ($y + 40))

$profName = $script:BrowserProfiles['Admin']
$profDir  = Get-EdgeProfileDirectory -Name $profName
if ($profDir) {
    $lblProfile.Text      = "Browser: Edge profile '$profName'"
    $lblProfile.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
} else {
    $found = (Get-EdgeProfiles | ForEach-Object { $_.DisplayName }) -join ', '
    $lblProfile.Text      = "Edge profile '$profName' not found -- links will use the last-used profile. Available: $found"
    $lblProfile.ForeColor = [System.Drawing.Color]::FromArgb(200, 120, 0)
}

$d = Get-ConnectedTenantDomain
if ($d) {
    $lblTenant.Text      = "Tenant: $d   --   sign in to the portal as the matching admin account"
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
} else {
    $lblTenant.Text      = 'Not connected to Exchange Online -- these links still work, but check which tenant the browser is signed in to.'
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
}

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'M365 Reports' | Out-Null
[void]$form.ShowDialog()
#endregion
