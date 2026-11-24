#Requires -Version 5.1
# TOOLKIT-VARIANT: Any
<#
    Customizations.ps1
    Save this console's arrangement to a file, and load one back.

    Author   : Glenn Armbruster
    Revised  : 2026-09-15
#>
# TOOLKIT-HELP-RUNS: Reads the config files and writes one file you choose, or your own profile. Touches no mailbox and no other person's settings.
# TOOLKIT-HELP-BEGIN
# Your console arrangement, in a file you keep.
#
# WHAT IS IN IT. Which sections exist, how many columns, how many rows each
# section gets, which tool sits where, which tools are grouped behind one
# button, which are hidden, and your theme. NOT the tools themselves -- those
# live in Tools.json and are shared, so a tool renamed or retired there is
# renamed or retired for everyone whatever this file says.
#
# SAVE is for three moments:
#   Before an upgrade    A new build may replace the shared Sections.json and
#                        Tools.json. Save first and your arrangement survives
#                        it, even though those files did not.
#   Moving machines      Same layout on the new laptop in one Load.
#   Before experimenting Try a different shape knowing the old one is on disk.
#
# LOAD puts a saved file back as YOUR layout, for the tenant you are in or for
# every tenant. It never touches the shared files and never touches anyone
# else's profile.
#
# IF YOU HAVE NEVER CUSTOMISED ANYTHING, Save still works: it captures what
# the shared files currently give you. That is the one to run before an
# upgrade on a console somebody arranged the old way, by editing the shared
# files -- it turns that arrangement into something a profile can hold.
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

$script:FileKind = 'AdminConsoleCustomizations'

function Get-CurrentLayout {
    <#
        WHAT THIS CONSOLE SHOWS YOU RIGHT NOW, as a layout block.

        Profile first, shared files second, key by key -- so somebody who has
        never customised anything still gets a complete, usable file rather
        than an empty one. That is the case this tool exists for: an operator
        who arranged their console by editing the shared files, about to take
        a build that replaces them.
    #>
    $layout = [PSCustomObject][ordered]@{}

    # ---- sections and columns ----
    $secs = @()
    $per  = 0
    $uSec = Get-UserLayoutProperty -Name 'Sections'
    $uPer = Get-UserLayoutProperty -Name 'ButtonsPerRow'
    if ($uSec) {
        foreach ($sc in @($uSec)) {
            $secs += [PSCustomObject][ordered]@{ Name = "$($sc.Name)"; Rows = [int]$sc.Rows; Accent = "$($sc.Accent)" }
        }
    }
    if ($uPer) { $per = [int]$uPer }
    if ($secs.Count -eq 0 -or $per -lt 1) {
        try {
            $sp = Get-ConfigFilePath -Name 'Sections.json'
            if (Test-Path -LiteralPath $sp) {
                $sj = Get-Content -LiteralPath $sp -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($secs.Count -eq 0) {
                    foreach ($sc in @($sj.Sections)) {
                        $secs += [PSCustomObject][ordered]@{ Name = "$($sc.Name)"; Rows = [int]$sc.Rows; Accent = "$($sc.Accent)" }
                    }
                }
                if ($per -lt 1) { $per = [int]$sj.ButtonsPerRow }
            }
        } catch { }
    }
    if ($per -lt 1) { $per = 5 }

    # ---- placement ----
    $placement = Get-UserLayoutProperty -Name 'Placement'
    if (-not $placement) {
        # Built from Tools.json: every tool, where it sits today. This is the
        # snapshot that rescues an arrangement made by editing the shared file.
        $placement = [PSCustomObject]@{}
        try {
            $tp = Get-ConfigFilePath -Name 'Tools.json'
            foreach ($t in @(Get-Content -LiteralPath $tp -Raw -Encoding UTF8 | ConvertFrom-Json)) {
                $k = Get-ParkedKey -Script "$($t.Script)"
                if (-not $k) { continue }
                $e = [ordered]@{ Section = "$($t.Section)".Trim(); Order = [int]$t.Order }
                if ($t.PSObject.Properties['Group'] -and "$($t.Group)".Trim()) { $e['Group'] = "$($t.Group)".Trim() }
                Add-Member -InputObject $placement -NotePropertyName $k -NotePropertyValue ([PSCustomObject]$e) -Force
            }
        } catch { }
    }

    Add-Member -InputObject $layout -NotePropertyName 'ButtonsPerRow' -NotePropertyValue $per -Force
    Add-Member -InputObject $layout -NotePropertyName 'Sections'      -NotePropertyValue $secs -Force
    Add-Member -InputObject $layout -NotePropertyName 'Placement'     -NotePropertyValue $placement -Force
    foreach ($carry in @('Theme', 'Parked')) {
        $v = Get-UserLayoutProperty -Name $carry
        if ($null -ne $v) { Add-Member -InputObject $layout -NotePropertyName $carry -NotePropertyValue $v -Force }
    }
    return $layout
}

function Measure-Layout {
    <# A layout in one sentence, for a label or a confirmation. #>
    param($Layout)
    if (-not $Layout) { return '(nothing)' }
    $ns = @($Layout.Sections).Count
    $np = 0
    try { $np = @($Layout.Placement.PSObject.Properties).Count } catch { }
    $bits = @("$ns section(s)", "$np tool placement(s)", "$([int]$Layout.ButtonsPerRow) per row")
    if ($Layout.PSObject.Properties['Theme'] -and "$($Layout.Theme)".Trim()) { $bits += "theme '$($Layout.Theme)'" }
    if ($Layout.PSObject.Properties['Parked'] -and $Layout.Parked) {
        $ps = 0; $pc = 0
        try { $ps = @($Layout.Parked.Scripts).Count } catch { }
        try { $pc = @($Layout.Parked.Sections).Count } catch { }
        if ($ps -or $pc) { $bits += "$ps parked tool(s), $pc parked section(s)" }
    }
    return ($bits -join ', ')
}

#endregion

#region UI ---------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Save / Load Customizations'
$form.ClientSize = New-Object System.Drawing.Size(720, 420)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWho = New-Object System.Windows.Forms.Label
$lblWho.Location = New-Object System.Drawing.Point(12, 12)
$lblWho.Size = New-Object System.Drawing.Size(696, 20)
$lblWho.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWho.Text = $(if (Get-ActiveTenantLabel) {
                    "Layout for '$(Get-UserProfileKey)' in '$(Get-ActiveTenantLabel)'"
                } else {
                    "Layout for '$(Get-UserProfileKey)'"
                })
$form.Controls.Add($lblWho)

$lblNow = New-Object System.Windows.Forms.Label
$lblNow.Location = New-Object System.Drawing.Point(12, 34)
$lblNow.Size = New-Object System.Drawing.Size(696, 34)
$lblNow.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblNow)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = 'Save customizations...'
$btnSave.Location = New-Object System.Drawing.Point(12, 76)
$btnSave.Size = New-Object System.Drawing.Size(170, 30)
$form.Controls.Add($btnSave)

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = 'Load customizations...'
$btnLoad.Location = New-Object System.Drawing.Point(192, 76)
$btnLoad.Size = New-Object System.Drawing.Size(170, 30)
$form.Controls.Add($btnLoad)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = 'Reset to shared layout'
$btnReset.Location = New-Object System.Drawing.Point(372, 76)
$btnReset.Size = New-Object System.Drawing.Size(170, 30)
$form.Controls.Add($btnReset)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(598, 76)
$btnClose.Size = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(12, 118)
$txtOut.Size = New-Object System.Drawing.Size(696, 290)
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$form.CancelButton = $btnClose

$script:Refresh = {
    $mine = Test-UserProfile
    $lblNow.Text = "Now: $(Measure-Layout -Layout (Get-CurrentLayout))" + [Environment]::NewLine +
                   $(if ($mine) { "You have a profile: $(Get-UserProfilePath)" }
                     else       { 'You have no profile yet -- this is the shared layout from Sections.json and Tools.json.' })
    $btnReset.Enabled = $mine
}

$btnSave.Add_Click({
    try {
        $layout = Get-CurrentLayout
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = 'Console layout (*.json)|*.json'
        $sfd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $tenantBit = $(if (Get-ActiveTenantLabel) { "-$((Get-ActiveTenantLabel) -replace '[^A-Za-z0-9._-]', '_')" } else { '' })
        $sfd.FileName = "ConsoleLayout-$(Get-UserProfileKey)$tenantBit-$(Get-Date -Format 'yyyyMMdd').json"
        if ($sfd.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $doc = [PSCustomObject][ordered]@{
            _comment = @(
                'A saved Admin Console arrangement. Load it back with the Save /',
                'Load Customizations tool on the Dashboard.',
                '',
                'This is a LAYOUT, not a toolkit: it says where things sit, not',
                'what they are. Loading it cannot add, remove or change a tool.'
            )
            Kind      = $script:FileKind
            Version   = 1
            SavedBy   = (Get-UserProfileKey)
            SavedFrom = "$(Get-ActiveTenantLabel)"
            SavedOn   = (Get-Date -Format 'yyyy-MM-dd HH:mm')
            Layout    = $layout
        }
        $json = ($doc | ConvertTo-Json -Depth 12)
        $enc  = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($sfd.FileName, ($json -replace "`r?`n", "`r`n"), $enc)

        $txtOut.Text = "SAVED" + [Environment]::NewLine + [Environment]::NewLine +
            "  $($sfd.FileName)" + [Environment]::NewLine + [Environment]::NewLine +
            "  $(Measure-Layout -Layout $layout)" + [Environment]::NewLine + [Environment]::NewLine +
            'Keep this somewhere that survives a toolkit upgrade -- not inside the toolkit folder.'
        Write-CompletedLog -Action 'Configure' -Target $sfd.FileName -Result 'Success' `
            -Detail "Console layout saved for $(Get-UserProfileKey)" -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Saving the layout'
        Write-FailedLog -Action 'Configure' -Target 'Console layout' -ErrorRecord $_ -Detail 'Saving customizations'
    }
})

$btnLoad.Add_Click({
    try {
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = 'Console layout (*.json)|*.json'
        $ofd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        if ($ofd.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $doc = $null
        try { $doc = Get-Content -LiteralPath $ofd.FileName -Raw -Encoding UTF8 | ConvertFrom-Json } catch { }
        # CHECKED BEFORE IT IS TRUSTED. This file came off a desktop and may be
        # any JSON at all; applying the wrong one would leave a console nobody
        # recognises and no obvious way back.
        if (-not $doc -or "$($doc.Kind)" -ne $script:FileKind -or -not $doc.Layout) {
            $txtOut.Text = "That file is not a saved console layout." + [Environment]::NewLine +
                           "  $($ofd.FileName)" + [Environment]::NewLine + [Environment]::NewLine +
                           "A layout file says Kind = $script:FileKind and carries a Layout block."
            return
        }

        $summary = "Saved by : $($doc.SavedBy)" + [Environment]::NewLine +
                   "From     : $(if ("$($doc.SavedFrom)") { $doc.SavedFrom } else { '(single tenant)' })" + [Environment]::NewLine +
                   "On       : $($doc.SavedOn)" + [Environment]::NewLine +
                   "Contains : $(Measure-Layout -Layout $doc.Layout)"

        $here = $(if (Get-ActiveTenantLabel) { "'$(Get-ActiveTenantLabel)'" } else { 'this console' })
        $a = Show-ToolDialog(
            ($summary + "`r`n`r`n" +
             "Load this as YOUR layout?" + "`r`n`r`n" +
             "YES     for $here only" + "`r`n" +
             "NO      for every tenant (your default)" + "`r`n" +
             "CANCEL  change nothing" + "`r`n`r`n" +
             "Your current arrangement is replaced. Nobody else's console changes."),
            'Load customizations', 'YesNoCancel', 'Question')
        if ($a -eq [System.Windows.Forms.DialogResult]::Cancel) { $txtOut.Text = 'Nothing loaded.'; return }

        $path = Save-UserLayout -Layout $doc.Layout -AllTenants:($a -eq [System.Windows.Forms.DialogResult]::No)
        & $script:Refresh
        $txtOut.Text = "LOADED" + [Environment]::NewLine + [Environment]::NewLine +
            $summary + [Environment]::NewLine + [Environment]::NewLine +
            "Written to your profile:" + [Environment]::NewLine + "  $path" + [Environment]::NewLine + [Environment]::NewLine +
            'Close and reopen the Admin Console to see it.'
        Write-CompletedLog -Action 'Configure' -Target $path -Result 'Success' `
            -Detail "Console layout loaded from $($ofd.FileName)" -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Loading the layout'
        Write-FailedLog -Action 'Configure' -Target 'Console layout' -ErrorRecord $_ -Detail 'Loading customizations'
    }
})

$btnReset.Add_Click({
    try {
        $here = $(if (Get-ActiveTenantLabel) { "'$(Get-ActiveTenantLabel)'" } else { 'this console' })
        $a = Show-ToolDialog(
            ("Go back to the shared layout?" + "`r`n`r`n" +
             "YES     for $here only" + "`r`n" +
             "NO      everything -- your whole profile" + "`r`n" +
             "CANCEL  keep what you have" + "`r`n`r`n" +
             "The whole-profile option renames the file rather than deleting it, so an afternoon of arranging is recoverable."),
            'Reset layout', 'YesNoCancel', 'Warning', 'Button3')
        if ($a -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
        $p = Remove-UserLayout -All:($a -eq [System.Windows.Forms.DialogResult]::No)
        & $script:Refresh
        $txtOut.Text = "RESET" + [Environment]::NewLine + [Environment]::NewLine +
            $(if ($p) { "  $p" } else { "  Nothing was saved for $here on its own, so nothing was reset. Your layout is in the every-tenant block -- use NO to clear that." }) + [Environment]::NewLine + [Environment]::NewLine +
            'Close and reopen the Admin Console to see it.'
        Write-CompletedLog -Action 'Configure' -Target "$p" -Result 'Success' `
            -Detail 'Console layout reset to the shared files' -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Resetting the layout'
    }
})

& $script:Refresh
$txtOut.Text = 'Save writes a file you keep. Load puts one back as your own layout.' + [Environment]::NewLine +
               'Neither touches Tools.json, Sections.json, or anybody else.'

Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Save / Load Customizations' | Out-Null
Set-ConsoleTheme -Form $form
[void]$form.ShowDialog()

#endregion
