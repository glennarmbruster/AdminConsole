#Requires -Version 5.1
<#
    TenantManager.ps1
    Creates and edits Config\Tenant.json -- the file that decides what this
    deployment is and therefore which tools exist.

    Reasoning: ScriptNotes/TenantManager.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-08
#>
# TOOLKIT-HELP-BEGIN
# Tenant.json holds the FACTS about this tenant: its name, its domains, which
# services it has, its directories. Every other tool reads it. It is the only
# file in the toolkit allowed to contain a domain or server name.
#
# WHY THIS TOOL EXISTS. Capabilities decide which buttons appear at all, and an
# absent Tenant.json means every capability reads as false -- so a new tenant
# opened a console with the on-premises tools simply missing and nothing saying
# why. Every other config file had an editor. This one, the one that decides
# what the console IS, had none.
#
# TWO MODES, chosen automatically:
#   No Tenant.json  -- first-run setup, started from the shipped template.
#   Tenant.json     -- editor, showing what is there now.
#
# 'Detect' fills in what it can from the connections this window already has,
# so most of the work is confirming rather than typing.
#
# CAPABILITY IS A CLAIM, NOT A PROBE. Ticking Exchange Online says this tenant
# HAS it, not that you are connected right now. Untick something the tenant
# genuinely does not have and its tools disappear rather than sitting greyed.
# TOOLKIT-HELP-END

param([switch]$Detect)

. (Join-Path $PSScriptRoot 'Common.ps1')


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# EDIT THE TENANT THAT IS ACTUALLY LOADED, not always Config\Tenant.json.
#
# Get-TenantConfigPath is the same resolver the console and every tool use, so
# this window edits exactly the file they read -- including the right one out of
# Config\Tenants when the console was launched with -Tenant. Hardcoding
# Tenant.json here would have made this the one tool that silently edits a
# different tenant from the one on screen, which is the worst possible place for
# that bug to live.
$script:ConfigPath   = Get-TenantConfigPath
$script:TemplatePath = Get-ConfigFilePath -Name 'Tenant.TEMPLATE.json'
$script:IsFirstRun   = -not (Test-Path -LiteralPath $script:ConfigPath)

#region Load --------------------------------------------------------------------
function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try { return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json) }
    catch { return $null }
}

# The template is the skeleton on first run, so the shape of a new file matches
# the shape the toolkit expects -- including the _comment blocks, which are the
# only documentation of this file that travels with it.
$script:Cfg = Read-JsonFile -Path $script:ConfigPath
if (-not $script:Cfg) { $script:Cfg = Read-JsonFile -Path $script:TemplatePath }
if (-not $script:Cfg) {
    # Last resort. A missing template must not stop a tenant being configured;
    # it only means the comments and defaults are absent.
    $script:Cfg = [PSCustomObject]@{
        TenantName = ''; PrimaryDomain = ''; OnMicrosoftDomain = ''
        Capabilities = [PSCustomObject]@{
            ExchangeOnline = $true; SharePointOnline = $true
            OnPremAD = $false; OnPremExchange = $false; Pim = $false; PimTimer = $false
            MicrosoftGraph = $false }
        Directories = @()
        OnPremExchange = [PSCustomObject]@{ ServerUri = ''; PreferredDomainController = '' }
        Pim = [PSCustomObject]@{ DurationHours = 4 }
        LicenseGroups = @()
        SharePoint = [PSCustomObject]@{ AdminUrl = ''; MySiteHost = '' }
        Graph = [PSCustomObject]@{ ClientId = ''; TenantId = '' }
        Theme = ''
    }
}

# THIS FILE'S OWN COPY, DELIBERATELY NAMED DIFFERENTLY FROM Common.ps1's.
#
# It used to be called Get-Prop, the same as the function in
# Common.ps1 -- which this file dot-sources. The local definition wins, so
# the calls below reached THIS one and always had; but the shadowing was a
# trap with two edges. A fix made in Common never arrived here, and anyone
# tidying up an apparent duplicate would have deleted this and left the
# calls binding silently to Common's version, which does not take the same
# parameters. That is a run-time failure at the click, on a Saturday.
#
# The two are genuinely different functions that happened to share a name,
# so they now have different names. Renamed 2026-09-12; behaviour unchanged.
function Get-TenantProp {
    param($Object, [string]$Name, $Default = '')
    if ($Object -and $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $Default
}
#endregion

#region Form --------------------------------------------------------------------

# SET BEFORE THE FORM IS BUILT, because $form.Size reads it. It was declared
# down with the picker controls, seventeen lines after the first use -- which in
# PowerShell is not an error, just $null, so the window silently came back the
# short size with the Save button off the bottom.
$script:MultiTenant = $false
try { $script:MultiTenant = [bool](Test-MultiTenantMode) } catch { }

$form = New-Object System.Windows.Forms.Form
$form.Text = if ($script:IsFirstRun) { 'Tenant Setup -- first run' } else { 'Tenant Manager' }
# WIDER THAN IT WAS. At 760 the Directories grid had seven columns sharing
# 714px -- Domain, the domain controllers, three OU paths -- and every one of
# them was an ellipsis. A config editor you cannot read the current value in
# is a config editor that gets guessed at. Reported 2026-09-12.
# 30px taller in multi-tenant mode, because the picker row pushes everything
# below it down by that much and the Save button lives at the bottom. A window
# that grows its content without growing itself hides the button you need.
$form.Size = New-Object System.Drawing.Size(1100, 844)   # top row + the taller Directories grid
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable

# ---- tenant picker -------------------------------------------------------------
#
# ONLY VISIBLE IN MULTI-TENANT MODE. With no Config\Tenants folder there is one
# tenant and nothing to choose, and a dropdown with one entry is furniture that
# has to be explained. It appears when it means something.
#
# WHY A NEW PROCESS TO SWITCH. The tenant is a per-process global read by a
# config cache that a hundred functions have already consulted by the time this
# window is open. Re-pointing it in place would leave every one of those caches
# holding the old tenant -- the exact class of bug that made Server Health read
# an empty server list. Relaunching is the only honest way to change it, and it
# takes a second.

$lblTenantPick = New-Object System.Windows.Forms.Label
$lblTenantPick.Text = 'Tenant file:'
$lblTenantPick.Location = New-Object System.Drawing.Point(12, 12)
$lblTenantPick.Size = New-Object System.Drawing.Size(80, 20)
$lblTenantPick.Visible = $script:MultiTenant
$form.Controls.Add($lblTenantPick)

$cboTenant = New-Object System.Windows.Forms.ComboBox
$cboTenant.DropDownStyle = 'DropDownList'
$cboTenant.Location = New-Object System.Drawing.Point(96, 9)
$cboTenant.Size = New-Object System.Drawing.Size(300, 22)
$cboTenant.Visible = $script:MultiTenant
$form.Controls.Add($cboTenant)

$btnTenantNew = New-Object System.Windows.Forms.Button
$btnTenantNew.Text = 'New / Clone'
$btnTenantNew.Location = New-Object System.Drawing.Point(404, 8)
$btnTenantNew.Size = New-Object System.Drawing.Size(100, 24)
$btnTenantNew.Visible = $script:MultiTenant
$form.Controls.Add($btnTenantNew)

$btnTenantDefault = New-Object System.Windows.Forms.Button
$btnTenantDefault.Text = 'Make default'
$btnTenantDefault.Location = New-Object System.Drawing.Point(510, 8)
$btnTenantDefault.Size = New-Object System.Drawing.Size(100, 24)
$btnTenantDefault.Visible = $script:MultiTenant
$form.Controls.Add($btnTenantDefault)

$btnTenantNick = New-Object System.Windows.Forms.Button
$btnTenantNick.Text = 'Nickname'
$btnTenantNick.Location = New-Object System.Drawing.Point(616, 8)
$btnTenantNick.Size = New-Object System.Drawing.Size(90, 24)
$btnTenantNick.Visible = $script:MultiTenant
$form.Controls.Add($btnTenantNick)

# OPEN IN A NEW POWERSHELL WINDOW -- the supported way to change tenant.
#
# Not 'switch to', because there is nothing to switch. The Exchange Online
# session, the on-premises PSSession, SPO, Graph and any PIM activation belong
# to this powershell.exe and outlive the console window that created them. A
# second tenant in the same process means tenant B's config driving tenant A's
# connections, with nothing on screen saying so -- Set-ActiveTenant refuses it
# outright for that reason.
#
# So this starts a PROCESS. The button exists so that is one click rather than
# a command somebody has to remember correctly at five o'clock.
$btnTenantOpen = New-Object System.Windows.Forms.Button
$btnTenantOpen.Text = 'Open in new window'
$btnTenantOpen.Location = New-Object System.Drawing.Point(712, 8)
$btnTenantOpen.Size = New-Object System.Drawing.Size(140, 24)
$btnTenantOpen.Visible = $script:MultiTenant
$form.Controls.Add($btnTenantOpen)

function Update-TenantPicker {
    if (-not $script:MultiTenant) { return }
    $cboTenant.Items.Clear()
    $active = ''
    try { $active = Get-ActiveTenantKey } catch { }
    $sel = 0; $i = 0
    foreach ($t in @(Get-TenantRegistry)) {
        $label = $t.Key
        if ($t.Nickname) { $label += "   [$($t.Nickname)]" }
        if ($t.IsDefault) { $label += '   (default)' }
        [void]$cboTenant.Items.Add($label)
        if ($t.Key -eq $active) { $sel = $i }
        $i++
    }
    if ($cboTenant.Items.Count -gt 0) { $cboTenant.SelectedIndex = $sel }
}

function Get-PickedTenant {
    $i = $cboTenant.SelectedIndex
    $reg = @(Get-TenantRegistry)
    if ($i -lt 0 -or $i -ge $reg.Count) { return $null }
    return $reg[$i]
}

function Write-TenantIndex {
    <#
        Rewrite Tenants.json from the registry plus one change.

        THE INDEX IS REBUILT FROM THE FILES ON DISK, never edited in place, so
        a tenant somebody deleted by hand cannot linger in it and a tenant
        somebody dropped in cannot be missing from it. The index only ever adds
        nicknames and marks a default -- it is not allowed to be the list.
    #>
    param([string]$DefaultKey, [hashtable]$Nicknames)

    $reg = @(Get-TenantRegistry)
    $entries = @()
    foreach ($t in $reg) {
        $nick = $t.Nickname
        if ($Nicknames -and $Nicknames.ContainsKey($t.Key)) { $nick = $Nicknames[$t.Key] }
        $e = [ordered]@{ Key = $t.Key }
        if ($nick) { $e['Nickname'] = $nick }
        $entries += [PSCustomObject]$e
    }
    $def = $DefaultKey
    if (-not $def) {
        $d = @($reg | Where-Object { $_.IsDefault })
        if ($d.Count -gt 0) { $def = $d[0].Key }
    }
    $obj = [ordered]@{
        _comment = @(
            'Nicknames and the default tenant. The FILES in this folder are the',
            'list -- this only decorates them, so deleting this file loses the',
            'nicknames and nothing else.',
            '',
            'AdminConsole.ps1 -Tenant <Key or Nickname>'
        )
        Default = $def
        Tenants = $entries
    }
    $path = Join-Path (Get-TenantsFolder) 'Tenants.json'
    $json = ([PSCustomObject]$obj | ConvertTo-Json -Depth 6)
    $enc  = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($path, ($json -replace "`r?`n", "`r`n"), $enc)
}

$cboTenant.Add_SelectedIndexChanged({
    # Selection alone changes nothing. Switching is a relaunch (see above), and
    # a dropdown that silently re-pointed the editor would edit one tenant while
    # showing another's values.
})

$btnTenantNew.Add_Click({
    try {
        $name = Show-InputDialog -Title 'New tenant' -Prompt ('File name for the new tenant, without .json.' + [Environment]::NewLine + 'It becomes the -Tenant value, e.g. wxyz')
        if (-not "$name".Trim()) { return }
        $key = ("$name".Trim() -replace '[^\w\.\-]', '')
        if (-not $key) { Write-Out 'That name has no usable characters.'; return }
        $dest = Join-Path (Get-TenantsFolder) "$key.json"
        if (Test-Path -LiteralPath $dest) { Write-Out "$key.json already exists."; return }

        $ans = Show-ToolDialog(
            ("Start from the tenant currently open ($([System.IO.Path]::GetFileNameWithoutExtension($script:ConfigPath)))?" + [Environment]::NewLine + [Environment]::NewLine +
             'Yes  -- copy it, then edit the domains.' + [Environment]::NewLine +
             'No   -- start from the blank template.'),
            'New tenant', 'YesNoCancel', 'Question')
        if ($ans -eq [System.Windows.Forms.DialogResult]::Cancel) { return }

        $src = if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) { $script:ConfigPath } else { $script:TemplatePath }
        if (-not (Test-Path -LiteralPath $src)) { Write-Out "Nothing to copy from ($src)."; return }
        Copy-Item -LiteralPath $src -Destination $dest -Force
        Write-TenantIndex
        Update-TenantPicker
        Write-Out ("Created $dest." + [Environment]::NewLine +
                   "Open it with:  .\AdminConsole.ps1 -Tenant $key")
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Creating a tenant' }
})

$btnTenantOpen.Add_Click({
    try {
        $sel = Get-PickedTenant
        if (-not $sel) { return }

        # NAME THE CONNECTIONS THIS WINDOW ALREADY HOLDS. The new process gets
        # a clean slate either way, but seeing 'you are signed in to X here' is
        # what makes the new-window rule read as a reason rather than ceremony.
        $live = @()
        try { $live = @(Get-ToolkitLiveConnection) } catch { }
        $note = ''
        if ($live.Count -gt 0) {
            $note = [Environment]::NewLine + [Environment]::NewLine +
                    "This window is signed in to $($live -join ', '). The new one starts with no connections, which is the point."
        }

        $go = Show-ToolDialog(
            ("Open '$($sel.Key)' in a new PowerShell window?" + $note),
            'Open tenant', 'OKCancel', 'Question')
        if ($go -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $null = Start-TenantConsole -Tenant $sel.Key
        Write-Out "Opening $($sel.Key) in a new window. This one keeps its own tenant and connections."
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening a tenant in a new window' }
})

$btnTenantDefault.Add_Click({
    try {
        $t = Get-PickedTenant
        if (-not $t) { return }
        Write-TenantIndex -DefaultKey $t.Key
        Update-TenantPicker
        Write-Out "$($t.Key) is now the default -- AdminConsole.ps1 with no -Tenant opens it."
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Setting the default tenant' }
})

$btnTenantNick.Add_Click({
    try {
        $t = Get-PickedTenant
        if (-not $t) { return }
        $nick = Show-InputDialog -Title 'Nickname' -Prompt ("Short name for $($t.Key). Blank removes it." + [Environment]::NewLine + 'Used as -Tenant <nickname>.') -Default $t.Nickname
        if ($null -eq $nick) { return }
        $clean = ("$nick".Trim() -replace '[^\w\.\-]', '')
        Write-TenantIndex -Nicknames @{ $t.Key = $clean }
        Update-TenantPicker
        Write-Out $(if ($clean) { "$($t.Key) is now also '$clean'." } else { "Nickname removed from $($t.Key)." })
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Setting a nickname' }
})

# ---- bootstrap ------------------------------------------------------------------
#
# THE ONLY BUTTON THAT SHOWS WHEN MULTI-TENANT MODE IS **OFF**, and it exists
# because without it the feature could not be switched on from the UI at all:
# New / Clone is hidden until Config\Tenants exists, and nothing creates
# Config\Tenants. The first tenant file had to be made by hand, which is a poor
# advertisement for a tool whose whole job is editing that file.
#
# It does the smallest honest thing: makes the folder, copies the tenant that is
# open right now into it as the first entry, and marks it default. Nothing is
# moved or deleted -- Config\Tenant.json stays exactly where it is, which is
# what makes REVERT-MultiTenant.ps1 a rename rather than a restore.
$btnTenantSetup = New-Object System.Windows.Forms.Button
$btnTenantSetup.Text = 'Set up multiple tenants'
$btnTenantSetup.Location = New-Object System.Drawing.Point(12, 8)
$btnTenantSetup.Size = New-Object System.Drawing.Size(170, 24)
$btnTenantSetup.Visible = (-not $script:MultiTenant)
$form.Controls.Add($btnTenantSetup)

$btnTenantSetup.Add_Click({
    try {
        if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
            Write-Out 'Save this tenant first -- there is nothing to copy yet.'
            return
        }

        # Name the first file after the primary domain's leading label, because
        # that is what somebody will type after -Tenant. Falls back to 'default'
        # rather than refusing: a name can be changed, a blocked setup cannot.
        $suggest = 'default'
        $dom = "$(Get-TenantProp $script:Cfg 'PrimaryDomain')".Trim()
        if ($dom) { $suggest = (($dom -split '\.')[0] -replace '[^\w\-]', '') }
        if (-not $suggest) { $suggest = 'default' }

        $key = Show-InputDialog -Title 'Set up multiple tenants' -Default $suggest -Prompt (
            'This creates Config\Tenants and copies the tenant open now into it as the first one.' + [Environment]::NewLine + [Environment]::NewLine +
            'Config\Tenant.json is left exactly where it is, so this is undoable by renaming one folder (REVERT-MultiTenant.ps1).' + [Environment]::NewLine + [Environment]::NewLine +
            'File name for this tenant, without .json:')
        if (-not "$key".Trim()) { return }
        $key = ("$key".Trim() -replace '[^\w\.\-]', '')
        if (-not $key) { Write-Out 'That name has no usable characters.'; return }

        $dir = Get-TenantsFolder
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
        $dest = Join-Path $dir "$key.json"
        if (Test-Path -LiteralPath $dest) { Write-Out "$key.json already exists."; return }
        Copy-Item -LiteralPath $script:ConfigPath -Destination $dest -Force

        Write-TenantIndex -DefaultKey $key

        Write-Out (
            "Multi-tenant mode is ON." + [Environment]::NewLine + [Environment]::NewLine +
            "  Created $dest" + [Environment]::NewLine +
            "  '$key' is the default, so AdminConsole.ps1 with no switch opens it." + [Environment]::NewLine + [Environment]::NewLine +
            'CLOSE AND REOPEN this window to see the tenant picker, then use New / Clone' + [Environment]::NewLine +
            'for the second tenant.' + [Environment]::NewLine + [Environment]::NewLine +
            'To undo: run REVERT-MultiTenant.ps1, or rename Config\Tenants.')
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Setting up multiple tenants' }
})

$lblBanner = New-Object System.Windows.Forms.Label
# Below the picker row when there is one, at the top when there is not. The
# picker occupies y=8..32, so a banner still at y=10 would sit on top of it.
$script:BannerY = 40   # the picker row (multi) or the setup button (single) occupies y=8..32
$lblBanner.Location = New-Object System.Drawing.Point(12, $script:BannerY)
$lblBanner.Size = New-Object System.Drawing.Size(1040, 34)
$lblBanner.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblBanner.Text = if ($script:IsFirstRun) {
    "No Tenant.json yet. Until this is saved, every capability reads as OFF and the tools that need one are hidden."
} else {
    "Editing Config\Tenant.json. Changes take effect when a tool next starts."
}
$lblBanner.ForeColor = if ($script:IsFirstRun) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::Black }
$form.Controls.Add($lblBanner)

function New-Field {
    param([string]$Label, [int]$Y, [string]$Value, [string]$Hint)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Label; $l.Location = New-Object System.Drawing.Point(12, ($Y + 3))
    $l.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($l)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point(168, $Y)
    $t.Size = New-Object System.Drawing.Size(560, 22)
    $t.Text = "$Value"
    $form.Controls.Add($t)
    if ($Hint) {
        $h = New-Object System.Windows.Forms.Label
        $h.Text = $Hint; $h.Location = New-Object System.Drawing.Point(736, ($Y + 3))
        $h.Size = New-Object System.Drawing.Size(330, 20)
        $h.ForeColor = [System.Drawing.Color]::DimGray
        $form.Controls.Add($h)
    }
    return $t
}

# EVERY FIELD BELOW IS PLACED FROM $y, so shifting it here is the whole layout
# adjustment for the picker row -- rather than 30 hardcoded offsets to find and
# get wrong. The banner is 34 tall; 52 clears it from y=10, so it has to clear
# it from y=40 too.
$y = $script:BannerY + 42
$txtName  = New-Field -Label 'Tenant name:'        -Y $y -Value (Get-TenantProp $script:Cfg 'TenantName')        -Hint 'Shown in the console header'
$y += 28
$txtDom   = New-Field -Label 'Primary domain:'     -Y $y -Value (Get-TenantProp $script:Cfg 'PrimaryDomain')     -Hint 'e.g. contoso.com'
$y += 28
$txtOnMs  = New-Field -Label 'onmicrosoft domain:' -Y $y -Value (Get-TenantProp $script:Cfg 'OnMicrosoftDomain') -Hint 'e.g. contoso.onmicrosoft.com'
$y += 28
$txtTheme = New-Field -Label 'Theme:'              -Y $y -Value (Get-TenantProp $script:Cfg 'Theme')             -Hint 'Optional, from Themes.json'
$y += 28
# THE BAND is the strip of colour along the top of every window this tenant
# opens, so two consoles side by side are never mistaken for each other.
# Empty means a colour is picked for the tenant automatically; set one to
# choose it.
$txtBand  = New-Field -Label 'Band colour:'        -Y $y -Value (Get-TenantProp $script:Cfg 'Band')              -Hint ''

# THE FIELD SHOWS THE COLOUR AND OPENS A PICKER, same as the section accents
# in Tool Manager. A hex value is not a colour anyone can see, and nobody
# should have to know that #D64545 is the red one. The box is painted in the
# colour with black or white text by luminance, so the field IS the preview;
# the hex stays as the value underneath and can still be typed by hand.
$txtBand.Size = New-Object System.Drawing.Size(220, 22)
# NoTheme: this box paints itself. Without the tag the theme walk repaints it
# on every Set-ConsoleTheme and the preview is lost.
$txtBand.Tag = 'NoTheme'

$btnBand = New-Object System.Windows.Forms.Button
$btnBand.Text = 'Pick...'
$btnBand.Location = New-Object System.Drawing.Point(396, ($y - 1))
$btnBand.Size = New-Object System.Drawing.Size(80, 24)
$form.Controls.Add($btnBand)

$btnBandAuto = New-Object System.Windows.Forms.Button
$btnBandAuto.Text = 'Automatic'
$btnBandAuto.Location = New-Object System.Drawing.Point(482, ($y - 1))
$btnBandAuto.Size = New-Object System.Drawing.Size(90, 24)
$form.Controls.Add($btnBandAuto)

$lblBandHint = New-Object System.Windows.Forms.Label
$lblBandHint.Location = New-Object System.Drawing.Point(584, ($y + 3))
$lblBandHint.Size = New-Object System.Drawing.Size(482, 20)
$lblBandHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblBandHint)

# Painting and the hint are one function, called from every path that can
# change the value -- picker, Automatic, and hand-typing -- so the three can
# never disagree about what the box is showing.
$script:PaintBand = {
    $hex = "$($txtBand.Text)".Trim()
    if ($hex -match '^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$') {
        $c = ConvertTo-ThemeColor $hex
        if ($c) {
            $txtBand.BackColor = $c
            $lum = (0.299 * $c.R + 0.587 * $c.G + 0.114 * $c.B) / 255
            $txtBand.ForeColor = $(if ($lum -lt 0.55) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black })
            $lblBandHint.Text = 'Strip along the top of every window for this tenant.'
            return
        }
    }
    # Not a colour: the theme's own input colours, captured once after the
    # theme ran, rather than a half-painted box. Re-running Set-ConsoleTheme
    # here would repaint the whole form on every keystroke.
    if ($script:BandBoxBack)  { $txtBand.BackColor = $script:BandBoxBack }
    if ($script:BandBoxFore)  { $txtBand.ForeColor = $script:BandBoxFore }
    $lblBandHint.Text = $(if ($hex) { 'Not a hex colour (#RGB or #RRGGBB). A colour will be picked automatically.' }
                          else      { 'Empty: a colour is picked automatically from the tenant key.' })
}

$btnBand.Add_Click({
    try {
        $cd = New-Object System.Windows.Forms.ColorDialog
        $cd.FullOpen = $true
        $cd.AnyColor = $true
        # Seeded with the current band -- or, when there is none, with the
        # colour this tenant is getting automatically, so 'a bit darker'
        # starts from what is actually on screen rather than from black.
        $cur = $null
        try { $cur = ConvertTo-ThemeColor "$($txtBand.Text)".Trim() } catch { }
        if (-not $cur) { try { $cur = Get-TenantBandColor } catch { } }
        if ($cur) { $cd.Color = $cur }
        if ($cd.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $txtBand.Text = '#{0:X2}{1:X2}{2:X2}' -f $cd.Color.R, $cd.Color.G, $cd.Color.B
        & $script:PaintBand
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Picking the band colour' }
})
$btnBandAuto.Add_Click({ $txtBand.Text = ''; & $script:PaintBand })
$txtBand.Add_TextChanged({ & $script:PaintBand })

$y += 36

# ---- capabilities -------------------------------------------------------------
$grpCap = New-Object System.Windows.Forms.GroupBox
$grpCap.Text = 'Capabilities -- what this tenant HAS. Unticked means the tools disappear, not grey out.'
$grpCap.Location = New-Object System.Drawing.Point(12, $y)
$grpCap.Size = New-Object System.Drawing.Size(1054, 78)
$form.Controls.Add($grpCap)

$caps = @('ExchangeOnline', 'SharePointOnline', 'OnPremAD', 'OnPremExchange', 'Pim', 'PimTimer', 'MicrosoftGraph')
$capBoxes = @{}
$cx = 14; $cy = 22
foreach ($c in $caps) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $c
    $cb.Location = New-Object System.Drawing.Point($cx, $cy)
    $cb.Size = New-Object System.Drawing.Size(160, 22)
    $cb.Checked = [bool](Get-TenantProp (Get-TenantProp $script:Cfg 'Capabilities' $null) $c $false)
    $grpCap.Controls.Add($cb)
    $capBoxes[$c] = $cb
    $cx += 172
    if ($cx -gt 600) { $cx = 14; $cy += 26 }
}
$y += 88

# ---- directories --------------------------------------------------------------
$lblDir = New-Object System.Windows.Forms.Label
$lblDir.Text = 'Directories -- one row per AD directory. Domain controllers separated by semicolons. Leave empty for a cloud-only tenant.'
$lblDir.Location = New-Object System.Drawing.Point(12, $y)
$lblDir.Size = New-Object System.Drawing.Size(1054, 18)
$lblDir.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblDir)
$y += 20

$gridDir = New-Object System.Windows.Forms.DataGridView
$gridDir.Location = New-Object System.Drawing.Point(12, $y)
$gridDir.Size = New-Object System.Drawing.Size(1054, 170)
# THE HEADER ROW SIZES ITSELF TO ITS TEXT.
#
# 'DCs (semicolon separated)' and 'Distribution grp OU' both wrap to two lines
# in the width Fill gives them, and the default fixed header height showed the
# first line and half of the second. AutoSize measures the wrapped text and
# makes the row tall enough -- the alternative is guessing a pixel height that
# is wrong again the next time a column is added or a font changes.
$gridDir.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$gridDir.AllowUserToAddRows = $true
$gridDir.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
foreach ($col in @(
    @{ N = 'Id';                 H = 'Id';                 W = 15 },
    @{ N = 'Label';              H = 'Label';              W = 18 },
    @{ N = 'Domain';             H = 'Domain';             W = 27 },
    @{ N = 'DomainControllers';  H = 'DCs (semicolon separated)'; W = 24 },
    # WHERE NEW OBJECTS GO. Blank means "ask me each time", which is the right
    # answer for a directory whose franchises each have their own OU. Filling
    # it in is what stops the tool asking.
    #
    # Canonical form, the shape ADUC shows:
    #   contoso.com/Accounts/Resource Mailbox
    @{ N = 'SharedMailboxOu';    H = 'Shared mbx OU';      W = 22 },
    @{ N = 'DistributionGroupOu';H = 'Distribution grp OU'; W = 22 },
    @{ N = 'MailContactOu';      H = 'Mail contact OU';    W = 20 },
    # OUT OF SYNC SCOPE. Where Projects\Migration step 5 moves objects so the
    # next directory sync deletes their cloud copies. Blank = that step
    # refuses to run. Same canonical form as the others.
    @{ N = 'NonSyncOu';          H = 'NonSync OU';         W = 20 })) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = $col.N; $c.HeaderText = $col.H; $c.FillWeight = $col.W
    [void]$gridDir.Columns.Add($c)
}
$cDef = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$cDef.Name = 'IsDefault'; $cDef.HeaderText = 'Default'; $cDef.FillWeight = 8
[void]$gridDir.Columns.Add($cDef)
foreach ($d in @(Get-TenantProp $script:Cfg 'Directories' @())) {
    # By name. A positional Rows.Add here is what shifted every value in the
    # Tools.json grid when a column was inserted.
    $i = $gridDir.Rows.Add()
    $r = $gridDir.Rows[$i]
    $r.Cells['Id'].Value                  = "$(Get-TenantProp $d 'Id')"
    $r.Cells['Label'].Value               = "$(Get-TenantProp $d 'Label')"
    $r.Cells['Domain'].Value              = "$(Get-TenantProp $d 'Domain')"
    $r.Cells['DomainControllers'].Value   = ((@(Get-TenantProp $d 'DomainControllers' @()) | ForEach-Object { "$_" }) -join '; ')
    $r.Cells['SharedMailboxOu'].Value     = "$(Get-TenantProp $d 'SharedMailboxOu')"
    $r.Cells['DistributionGroupOu'].Value = "$(Get-TenantProp $d 'DistributionGroupOu')"
    $r.Cells['MailContactOu'].Value       = "$(Get-TenantProp $d 'MailContactOu')"
    $r.Cells['NonSyncOu'].Value           = "$(Get-TenantProp $d 'NonSyncOu')"
    $r.Cells['IsDefault'].Value           = [bool](Get-TenantProp $d 'IsDefault' $false)
}
$form.Controls.Add($gridDir)
# 170 (the grid) + 12, NOT 128.
#
# The grid was 140 tall and this advanced 128, so the next label was drawn
# TWELVE PIXELS INSIDE the grid and its bottom half was covered -- which is the
# 'Exchange server URI:' row nobody could read. Derived from the height above
# rather than written as a second literal, so the two cannot disagree again.
$y += 182

# ---- on-prem exchange + sharepoint --------------------------------------------
$txtExs = New-Field -Label 'Exchange server URI:' -Y $y -Value (Get-TenantProp (Get-TenantProp $script:Cfg 'OnPremExchange' $null) 'ServerUri') -Hint 'On-prem only'
$y += 28
# THE SERVERS THE ON-PREMISES TOOLS WORK AGAINST -- Check Queues, Clear IIS
# Logs, Exchange Health Check, onprem mailtest. Until 2026-09-12 those four
# carried the hostnames as literals in their own code; they read this now.
#
# SEMICOLON-SEPARATED, like the domain controllers column above, because that
# is the separator this file already uses for a list in a text box and a
# second convention would be one to remember.
#
# BLANK IS A REAL ANSWER AND USUALLY THE RIGHT ONE: leave it empty and the
# tools ask the directory (Get-ExchangeServer) with an on-premises session
# open, which needs no maintenance when a server is added. Fill it in only to
# pin a subset.
# NAME, OR NAME:SITE:ROLE. Since 2026-09-17 an entry may carry the site and
# role ServerReboot.ps1 needs to honour its one-per-site rule -- declared
# here, never guessed from the name. The field round-trips both shapes: an
# object renders as SERVER01:SiteA:Mailbox, a bare name as SERVER01, and Save
# writes back whichever it was given. Before this the field flattened every
# entry to a name on Save, so one visit to this window silently deleted the
# site and role of every server and the reboot tool greyed them all out.
function ConvertTo-ServerFieldText {
    param($Entries)
    $parts = @()
    foreach ($e in @($Entries)) {
        if ($null -eq $e) { continue }
        if ($e -is [string]) { $t = "$e".Trim(); if ($t) { $parts += $t }; continue }
        $n = ''; $si = ''; $ro = ''
        try { $n = "$($e.Name)".Trim(); $si = "$($e.Site)".Trim(); $ro = "$($e.Role)".Trim() } catch { }
        if (-not $n) { $t = "$e".Trim(); if ($t) { $parts += $t }; continue }
        $parts += $(if ($si -or $ro) { "${n}:${si}:${ro}" } else { $n })
    }
    return ($parts -join '; ')
}
function ConvertFrom-ServerFieldText {
    param([string]$Text)
    $out = @()
    foreach ($tok in @("$Text" -split ';')) {
        $t = "$tok".Trim()
        if (-not $t) { continue }
        if ($t -notmatch ':') { $out += $t; continue }
        $bits = @($t -split ':', 3 | ForEach-Object { "$_".Trim() })
        $n = $bits[0]; if (-not $n) { continue }
        $si = $(if ($bits.Count -gt 1) { $bits[1] } else { '' })
        $ro = $(if ($bits.Count -gt 2) { $bits[2] } else { '' })
        if ($ro -match '^(?i)mailbox$') { $ro = 'Mailbox' } elseif ($ro -match '^(?i)edge$') { $ro = 'Edge' }
        $o = [ordered]@{ Name = $n }
        if ($si) { $o['Site'] = $si }
        if ($ro) { $o['Role'] = $ro }
        $out += [PSCustomObject]$o
    }
    return $out
}

$txtSrv = New-Field -Label 'Exchange servers:' -Y $y `
            -Value (ConvertTo-ServerFieldText -Entries (Get-TenantProp (Get-TenantProp $script:Cfg 'OnPremExchange' $null) 'Servers' @())) `
            -Hint 'Blank = ask the directory. Semicolons between entries. NAME, or NAME:Site:Role (Role = Mailbox or Edge) for Server Reboot.'
$y += 30

# SERVER HEALTH IS A SEPARATE LIST, AND THAT IS THE POINT.
#
# 'Exchange servers' above means "servers the Exchange tools may run cmdlets
# against". Server Health asks operating-system questions -- disk, patches,
# reboots -- of a WIDER set: domain controllers, file servers, a jump box.
# Overloading one list would either drag those into Exchange tooling or keep
# them out of the health check.
#
# ONE LINE, SEMICOLONS BETWEEN SERVERS, and a trailing * marks a server that
# should have Exchange on it. A grid would be the obvious choice for three
# fields, but this window is a column of text boxes and one grid in the middle
# of it reads as a different program -- and the flag is the only extra fact.
# Server Health reports Exchange wherever it FINDS it, so the star only decides
# whether ABSENCE is worth flagging.
$shServers = @()
try {
    $shCfg = Get-TenantProp $script:Cfg 'ServerHealth' $null
    if ($shCfg) {
        foreach ($e in @(Get-TenantProp $shCfg 'Servers' @())) {
            if ($null -eq $e) { continue }
            if ($e.PSObject.Properties['Name']) {
                $n = "$($e.Name)".Trim()
                if (-not $n) { continue }
                $isEx = $false
                if ($e.PSObject.Properties['Exchange']) { $isEx = ("$($e.Exchange)".Trim() -match '(?i)^(true|yes|1)$') }
                $shServers += $(if ($isEx) { "$n*" } else { $n })
            } else {
                $n = "$e".Trim()
                if ($n) { $shServers += $n }
            }
        }
    }
} catch { }
$txtHealth = New-Field -Label 'Server Health list:' -Y $y `
                -Value ($shServers -join '; ') `
                -Hint 'Semicolons. Trailing * = expect Exchange.'
$y += 30

$txtSpo = New-Field -Label 'SPO admin URL:' -Y $y -Value (Get-TenantProp (Get-TenantProp $script:Cfg 'SharePoint' $null) 'AdminUrl') -Hint 'Blank = derived from the tenant'
$y += 28
$txtMy  = New-Field -Label 'MySite host:' -Y $y -Value (Get-TenantProp (Get-TenantProp $script:Cfg 'SharePoint' $null) 'MySiteHost') -Hint 'e.g. https://contoso-my.sharepoint.com'
$y += 28

# The tenant's Graph app (2026-09-17). Two GUIDs from the app registration;
# the MicrosoftGraph capability above is what turns the Graph tools on, and a
# ticked capability with a blank client id reads as 'not configured'.
$txtGraphClient = New-Field -Label 'Graph app client ID:' -Y $y -Value (Get-TenantProp (Get-TenantProp $script:Cfg 'Graph' $null) 'ClientId') -Hint 'Application (client) ID of the delegated Graph app'
$y += 28
$txtGraphTenant = New-Field -Label 'Graph tenant ID:' -Y $y -Value (Get-TenantProp (Get-TenantProp $script:Cfg 'Graph' $null) 'TenantId') -Hint 'Directory (tenant) ID. Blank = let sign-in pick it'
$y += 36

# ---- licence groups -----------------------------------------------------------
$lblLic = New-Object System.Windows.Forms.Label
$lblLic.Text = 'Licence groups -- referenced by the disable tools. Name must match the AD group exactly.'
$lblLic.Location = New-Object System.Drawing.Point(12, $y)
$lblLic.Size = New-Object System.Drawing.Size(1054, 18)
$lblLic.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblLic)
$y += 20

$gridLic = New-Object System.Windows.Forms.DataGridView
$gridLic.Location = New-Object System.Drawing.Point(12, $y)
$gridLic.Size = New-Object System.Drawing.Size(1054, 100)
$gridLic.AllowUserToAddRows = $true
$gridLic.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
foreach ($col in @(
    @{ N = 'Name';      H = 'Group name'; W = 44 },
    @{ N = 'Sku';       H = 'Sku';        W = 16 },
    @{ N = 'Directory'; H = 'Directory';  W = 24 })) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = $col.N; $c.HeaderText = $col.H; $c.FillWeight = $col.W
    [void]$gridLic.Columns.Add($c)
}

# WHICH GROUPS A TERMINATION REMOVES FROM.
#
# This list serves two unrelated purposes. Reports read all of it; a
# termination must only touch the group the leaver actually holds. One tenant
# has E1 groups here purely so the Disabled/Licensed report can find people
# who are disabled but still licensed -- removing anyone from those breaks
# the report and does nothing about the licence.
#
# Nothing in the file distinguished them, so the disable pipeline guessed,
# took the first entry, and stripped terminated staff from an E1 group they
# were never in while their E5 licence stayed assigned. It reported success.
#
# A checkbox, not free text: this decides whether a script removes a licence,
# and 'yes'/'true'/'Y' typed into a cell is three ways to get that wrong.
$cTerm = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$cTerm.Name = 'RemoveOnTerm'
$cTerm.HeaderText = 'Remove on term'
$cTerm.FillWeight = 16
[void]$gridLic.Columns.Add($cTerm)

foreach ($l in @(Get-TenantProp $script:Cfg 'LicenseGroups' @())) {
    # By name, not by position -- adding this column to a positional Rows.Add
    # is exactly what shifted every value in the Tools.json grid.
    $i = $gridLic.Rows.Add()
    $r = $gridLic.Rows[$i]
    $r.Cells['Name'].Value         = "$(Get-TenantProp $l 'Name')"
    $r.Cells['Sku'].Value          = "$(Get-TenantProp $l 'Sku')"
    $r.Cells['Directory'].Value    = "$(Get-TenantProp $l 'Directory')"
    $r.Cells['RemoveOnTerm'].Value = [bool](Get-TenantProp $l 'RemoveOnTerm' $false)
}
$form.Controls.Add($gridLic)
$y += 110

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(12, $y)
$txtOut.Size = New-Object System.Drawing.Size(1054, 70)
$txtOut.Multiline = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.ReadOnly = $true
$txtOut.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($txtOut)
$y += 78
#endregion

#region Detect / validate / save ------------------------------------------------
function Write-Out { param([string]$Text) $txtOut.Text = $Text }

$btnDetect = New-Object System.Windows.Forms.Button
$btnDetect.Text = 'Detect'
$btnDetect.Location = New-Object System.Drawing.Point(12, $y)
$btnDetect.Size = New-Object System.Drawing.Size(110, 30)
$btnDetect.Add_Click({
    # Fills only what is EMPTY. Overwriting a value somebody deliberately set,
    # because a probe happened to return something else, is how a config tool
    # loses trust -- and this file is the one nobody can afford to distrust.
    $notes = New-Object System.Collections.ArrayList
    try {
        $domains = @(Invoke-EXOCommand -Name 'Get-AcceptedDomain' -Parameters @{ ErrorAction = 'Stop' })
        $default = @($domains | Where-Object { $_.Default })
        if (-not $txtDom.Text.Trim() -and $default.Count -gt 0) {
            $txtDom.Text = "$($default[0].DomainName)"
            [void]$notes.Add("Primary domain from the default accepted domain: $($txtDom.Text)")
        }
        $onms = @($domains | Where-Object { "$($_.DomainName)" -match '(?i)\.onmicrosoft\.com$' -and "$($_.DomainName)" -notmatch '(?i)\.mail\.onmicrosoft\.com$' })
        if (-not $txtOnMs.Text.Trim() -and $onms.Count -gt 0) {
            $txtOnMs.Text = "$($onms[0].DomainName)"
            [void]$notes.Add("onmicrosoft domain: $($txtOnMs.Text)")
        }
        if (-not $capBoxes['ExchangeOnline'].Checked) {
            $capBoxes['ExchangeOnline'].Checked = $true
            [void]$notes.Add('Exchange Online ticked -- this window is connected to it')
        }
        if (-not $txtName.Text.Trim() -and $txtDom.Text.Trim()) {
            $txtName.Text = ($txtDom.Text.Trim() -split '\.')[0]
            [void]$notes.Add("Tenant name guessed from the domain: $($txtName.Text) -- change it if that is wrong")
        }
    } catch {
        [void]$notes.Add("Exchange Online not reachable from this window, so cloud values were not detected: $($_.Exception.Message)")
    }

    try {
        $root = Get-ADRootDSE -ErrorAction Stop
        $forest = "$($root.rootDomainNamingContext)" -replace '(?i)DC=', '' -replace ',', '.'
        if ($gridDir.Rows.Count -le 1 -and $forest) {
            [void]$gridDir.Rows.Add('Primary', 'Primary directory', $forest, '', $true)
            [void]$notes.Add("Directory added from the forest root: $forest -- add domain controllers if you pin them")
        }
        if (-not $capBoxes['OnPremAD'].Checked) {
            $capBoxes['OnPremAD'].Checked = $true
            [void]$notes.Add('On-prem AD ticked -- a directory answered')
        }
    } catch {
        [void]$notes.Add('No on-premises AD reachable, so OnPremAD was left as it is.')
    }

    if ($notes.Count -eq 0) { Write-Out 'Nothing to fill in -- every field this can detect already has a value.' }
    else { Write-Out (($notes | ForEach-Object { "* $_" }) -join [Environment]::NewLine) }
})
$form.Controls.Add($btnDetect)

function Get-FormConfig {
    $dirs = @()
    foreach ($row in $gridDir.Rows) {
        if ($row.IsNewRow) { continue }
        $id = "$($row.Cells['Id'].Value)".Trim()
        if (-not $id) { continue }
        $dcs = @("$($row.Cells['DomainControllers'].Value)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $dirs += [PSCustomObject]@{
            Id                  = $id
            Label               = "$($row.Cells['Label'].Value)".Trim()
            Domain              = "$($row.Cells['Domain'].Value)".Trim()
            DomainControllers   = $dcs
            SharedMailboxOu     = "$($row.Cells['SharedMailboxOu'].Value)".Trim()
            DistributionGroupOu = "$($row.Cells['DistributionGroupOu'].Value)".Trim()
            MailContactOu       = "$($row.Cells['MailContactOu'].Value)".Trim()
            NonSyncOu           = "$($row.Cells['NonSyncOu'].Value)".Trim()
            IsDefault           = [bool]$row.Cells['IsDefault'].Value
        }
    }
    $lics = @()
    foreach ($row in $gridLic.Rows) {
        if ($row.IsNewRow) { continue }
        $n = "$($row.Cells['Name'].Value)".Trim()
        if (-not $n) { continue }
        $lics += [PSCustomObject]@{
            Name      = $n
            Sku       = "$($row.Cells['Sku'].Value)".Trim()
            Directory = "$($row.Cells['Directory'].Value)".Trim()
            # An unticked checkbox cell reads back as $null, not $false, so the
            # cast is doing real work. Always written, unlike optional string
            # fields, because "absent" and "false" must not be distinguishable
            # here -- a reader treating absent as true is how this started.
            RemoveOnTerm = [bool]$row.Cells['RemoveOnTerm'].Value
        }
    }
    $capObj = [ordered]@{}
    foreach ($c in $caps) { $capObj[$c] = [bool]$capBoxes[$c].Checked }

    # _comment and Pim are carried through from what was loaded rather than
    # rebuilt. This tool does not edit them, and a config editor that silently
    # drops the fields it does not understand is a config editor that destroys
    # data the next version will need.
    $out = [ordered]@{}
    $existingComment = Get-TenantProp $script:Cfg '_comment' $null
    if ($existingComment) { $out['_comment'] = $existingComment }
    $out['TenantName']        = $txtName.Text.Trim()
    $out['PrimaryDomain']     = $txtDom.Text.Trim()
    $out['OnMicrosoftDomain'] = $txtOnMs.Text.Trim()
    $out['Capabilities']      = [PSCustomObject]$capObj
    $out['Directories']       = $dirs
    $out['OnPremExchange']    = [PSCustomObject]@{
        ServerUri                 = $txtExs.Text.Trim()
        PreferredDomainController = "$(Get-TenantProp (Get-TenantProp $script:Cfg 'OnPremExchange' $null) 'PreferredDomainController')"
    }
    # Written only when there is something to write. An empty Servers array is
    # not the same as no key: absent means "ask the directory", and a key
    # holding [] would read as "there are none" to anyone editing the file by
    # hand. Omit it and the absent-means-derive rule keeps working.
    $srvList = @(ConvertFrom-ServerFieldText -Text "$($txtSrv.Text)")
    if ($srvList.Count -gt 0) {
        $out['OnPremExchange'] | Add-Member -NotePropertyName 'Servers' -NotePropertyValue $srvList -Force
    }

    # ServerHealth, written only when there is something to write -- same rule
    # as Servers above. An empty block would read as "checked, nothing there".
    #
    # The trailing * is parsed off the name, never stored: the file keeps a
    # real JSON boolean so anything reading it does not have to know about a
    # convention invented for one text box. Notes on existing entries are
    # carried through by name, because this field cannot edit them and
    # silently dropping somebody's note would be worse than not offering it.
    $oldNotes = @{}
    try {
        $shOld = Get-TenantProp $script:Cfg 'ServerHealth' $null
        if ($shOld) {
            foreach ($e in @(Get-TenantProp $shOld 'Servers' @())) {
                if ($e -and $e.PSObject.Properties['Name'] -and $e.PSObject.Properties['Notes']) {
                    $k = (("$($e.Name)".Trim() -split '\.')[0]).ToLowerInvariant()
                    if ($k) { $oldNotes[$k] = "$($e.Notes)" }
                }
            }
        }
    } catch { }

    $healthList = @()
    foreach ($raw in @("$($txtHealth.Text)" -split ';')) {
        $n = "$raw".Trim()
        if (-not $n) { continue }
        $isEx = $false
        if ($n.EndsWith('*')) { $isEx = $true; $n = $n.TrimEnd('*').Trim() }
        if (-not $n) { continue }
        $entry = [ordered]@{ Name = $n; Exchange = $isEx }
        $k = (($n -split '\.')[0]).ToLowerInvariant()
        if ($oldNotes.ContainsKey($k) -and $oldNotes[$k]) { $entry['Notes'] = $oldNotes[$k] }
        $healthList += [PSCustomObject]$entry
    }
    if ($healthList.Count -gt 0) {
        $out['ServerHealth'] = [PSCustomObject]@{ Servers = $healthList }
    }
    $pim = Get-TenantProp $script:Cfg 'Pim' $null
    if ($pim) { $out['Pim'] = $pim }
    $out['LicenseGroups'] = $lics

    # Carried through, exactly like Pim above, and for the same reason: this
    # tool has no field for it, and rebuilding the object without it would
    # silently delete the Feature Request recipient the first time anybody
    # opened Tenant Manager and pressed Save.
    $fr = Get-TenantProp $script:Cfg 'FeatureRequest' $null
    if ($fr) { $out['FeatureRequest'] = $fr }
    $out['SharePoint']    = [PSCustomObject]@{
        AdminUrl   = $txtSpo.Text.Trim()
        MySiteHost = $txtMy.Text.Trim()
    }
    $out['Graph'] = [PSCustomObject]@{
        ClientId = $txtGraphClient.Text.Trim()
        TenantId = $txtGraphTenant.Text.Trim()
    }
    # The template ships a _comment under Graph; keep it, and anything else
    # under that key this editor has no field for.
    $oldGraph = Get-TenantProp $script:Cfg 'Graph' $null
    if ($oldGraph) {
        foreach ($p in @($oldGraph.PSObject.Properties)) {
            if ($out['Graph'].PSObject.Properties[$p.Name]) { continue }
            $out['Graph'] | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force
        }
    }
    $out['Theme'] = $txtTheme.Text.Trim()
    $out['Band']  = $txtBand.Text.Trim()

    # EVERYTHING ELSE IS CARRIED THROUGH TOO. Pim and FeatureRequest above were
    # the two cases somebody had already been burned by; the rule they follow
    # is general. A Tenant.json can carry per-section comment blocks
    # (_comment_Pim, _comment_Directories ...) and keys a newer toolkit reads
    # that this editor has no field for. Any top-level key this tool did not
    # just write is copied from what was loaded, so Save never deletes what it
    # does not understand. Order: known keys first, then the rest as found.
    foreach ($p in @($script:Cfg.PSObject.Properties)) {
        if ($out.Contains($p.Name)) { continue }
        $out[$p.Name] = $p.Value
    }

    # OnPremExchange is rebuilt from two fields above; keys under it that this
    # tool does not edit (ExcludedServers, for one) come back the same way.
    $oldExs = Get-TenantProp $script:Cfg 'OnPremExchange' $null
    if ($oldExs) {
        foreach ($p in @($oldExs.PSObject.Properties)) {
            if ($out['OnPremExchange'].PSObject.Properties[$p.Name]) { continue }
            $out['OnPremExchange'] | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force
        }
    }
    return [PSCustomObject]$out
}

function Test-FormConfig {
    param($Cfg)
    $problems = New-Object System.Collections.ArrayList

    if (-not $Cfg.TenantName)    { [void]$problems.Add('ERROR  Tenant name is empty -- the console header will show nothing.') }
    if (-not $Cfg.PrimaryDomain) { [void]$problems.Add('ERROR  Primary domain is empty -- Get-MailDomain returns blank and every "@domain optional" label breaks.') }

    $on = @($caps | Where-Object { $Cfg.Capabilities.$_ })
    if ($on.Count -eq 0) {
        [void]$problems.Add('ERROR  No capabilities ticked. Every capability-gated tool will be HIDDEN and the console will look empty.')
    }

    if (($Cfg.Capabilities.OnPremAD -or $Cfg.Capabilities.OnPremExchange) -and @($Cfg.Directories).Count -eq 0) {
        [void]$problems.Add('ERROR  An on-premises capability is ticked but no directory is defined -- nothing can resolve a server.')
    }
    if (-not $Cfg.Capabilities.OnPremAD -and @($Cfg.Directories).Count -gt 0) {
        [void]$problems.Add('WARN   Directories are defined but OnPremAD is unticked, so they will never be used.')
    }
    if ($Cfg.Capabilities.OnPremExchange -and -not $Cfg.OnPremExchange.ServerUri) {
        [void]$problems.Add('WARN   On-prem Exchange is ticked with no server URI -- Connect ExSE will have to ask.')
    }
    if (@($Cfg.Directories | Where-Object { $_.IsDefault }).Count -gt 1) {
        [void]$problems.Add('WARN   More than one directory is marked default.')
    }
    foreach ($d in @($Cfg.Directories)) {
        if (-not $d.Domain) { [void]$problems.Add("WARN   Directory '$($d.Id)' has no domain, so groups cannot be classified by environment.") }
    }
    foreach ($l in @($Cfg.LicenseGroups)) {
        if ($l.Directory -and @($Cfg.Directories | Where-Object { $_.Id -ieq $l.Directory }).Count -eq 0) {
            [void]$problems.Add("WARN   Licence group '$($l.Name)' names directory '$($l.Directory)', which is not defined above.")
        }
    }
    $guid = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    $gc = "$($Cfg.Graph.ClientId)".Trim(); $gt = "$($Cfg.Graph.TenantId)".Trim()
    if ($Cfg.Capabilities.MicrosoftGraph -and -not $gc) {
        [void]$problems.Add('ERROR  MicrosoftGraph is ticked but the Graph app client ID is blank -- the Graph tools will read as not configured.')
    }
    if ($gc -and $gc -notmatch $guid) { [void]$problems.Add('ERROR  The Graph app client ID is not a GUID.') }
    if ($gt -and $gt -notmatch $guid) { [void]$problems.Add('ERROR  The Graph tenant ID is not a GUID.') }
    if ($gc -and -not $Cfg.Capabilities.MicrosoftGraph) {
        [void]$problems.Add('WARN   A Graph app client ID is set but MicrosoftGraph is unticked, so the Graph tools stay hidden.')
    }
    if ($Cfg.OnMicrosoftDomain -and $Cfg.OnMicrosoftDomain -notmatch '(?i)\.onmicrosoft\.com$') {
        [void]$problems.Add('WARN   The onmicrosoft domain does not end in .onmicrosoft.com.')
    }
    if ($Cfg.PSObject.Properties['Band'] -and "$($Cfg.Band)".Trim() -and "$($Cfg.Band)".Trim() -notmatch '^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$') {
        [void]$problems.Add('WARN   Band colour is not a hex colour (#RGB or #RRGGBB); the automatic colour will be used.')
    }
    return $problems
}

$btnValidate = New-Object System.Windows.Forms.Button
$btnValidate.Text = 'Validate'
$btnValidate.Location = New-Object System.Drawing.Point(130, $y)
$btnValidate.Size = New-Object System.Drawing.Size(110, 30)
$btnValidate.Add_Click({
    $p = Test-FormConfig -Cfg (Get-FormConfig)
    if ($p.Count -eq 0) { Write-Out 'No problems found.' }
    else { Write-Out ($p -join [Environment]::NewLine) }
})
$form.Controls.Add($btnValidate)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = 'Save'
$btnSave.Location = New-Object System.Drawing.Point(248, $y)
$btnSave.Size = New-Object System.Drawing.Size(110, 30)
$btnSave.Add_Click({
    $cfg = Get-FormConfig
    $problems = Test-FormConfig -Cfg $cfg
    $errors = @($problems | Where-Object { $_ -like 'ERROR*' })
    if ($errors.Count -gt 0) {
        $ans = Show-ToolDialog(
            (($problems -join [Environment]::NewLine) + [Environment]::NewLine + [Environment]::NewLine + 'Save anyway?'),
            'Tenant Manager', 'YesNo', 'Warning', 'Button2')
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { Write-Out 'Not saved.'; return }
    }
    try {
        # UTF-8 WITH BOM and CRLF, matching every other config file. Without a
        # BOM PowerShell 5.1 reads this through the ANSI codepage and a domain
        # with a non-ASCII character becomes a different domain.
        $json = $cfg | ConvertTo-Json -Depth 6
        $enc  = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($script:ConfigPath, ($json -replace "`r?`n", "`r`n"), $enc)

        # The cache is per-process and other windows hold their own. Saying so
        # is better than letting somebody wonder why the console did not change.
        try { $null = Get-TenantConfig -Refresh } catch { }
        $msg = "Saved to $($script:ConfigPath)."
        if ($problems.Count -gt 0) { $msg += [Environment]::NewLine + ($problems -join [Environment]::NewLine) }
        $msg += [Environment]::NewLine + 'Windows that are already open keep the old values until they are restarted.'
        Write-Out $msg
        $script:IsFirstRun = $false
        $lblBanner.Text = 'Editing Config\Tenant.json. Changes take effect when a tool next starts.'
        $lblBanner.ForeColor = [System.Drawing.Color]::Black
        try { Write-CompletedLog -Action 'TenantManager' -Target $cfg.TenantName -Result 'Success' -Detail 'Tenant.json saved' } catch { }
    } catch {
        Write-Out (Format-ErrorDetail -ErrorRecord $_ -Context 'Saving Tenant.json')
    }
})
$form.Controls.Add($btnSave)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(956, $y)
$btnClose.Size = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$form.ClientSize = New-Object System.Drawing.Size(1078, ($y + 44))
#endregion

if ($script:IsFirstRun) {
    Write-Out @"
No Tenant.json was found, so this started from the shipped template.

Fill in the tenant name and primary domain, tick what this tenant has, add a
directory row for each AD directory, then Save. Detect fills in what it can
from the connections this window already has.

Nothing else in the toolkit works properly until this file exists.
"@
}

try { Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath | Out-Null } catch { }
try { Set-ConsoleTheme -Form $form } catch { }

# AFTER THE THEME, NOT BEFORE. The band box paints itself, so it needs the
# theme's input colours to fall back to -- and they only exist once the theme
# has run. First paint happens here for the same reason.
$script:BandBoxBack = $txtName.BackColor
$script:BandBoxFore = $txtName.ForeColor
try { & $script:PaintBand } catch { }
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
if ($Detect) { $btnDetect.PerformClick() }
# AFTER every handler is wired, BEFORE the window shows: the picker reads the
# registry from disk and the buttons that rewrite it must already exist.
try { Update-TenantPicker } catch { }

[void]$form.ShowDialog()
