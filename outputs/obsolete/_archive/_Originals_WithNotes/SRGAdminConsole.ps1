#Requires -Version 5.1
<#
.SYNOPSIS
    SRG Admin Console -- Layer 1 launcher for the SRG (cloud-only,
    srgpros.com) tenant. Mailbox / Group lookup + dynamic tool menu.
.DESCRIPTION
    A SEPARATE, PARALLEL console to the original AdminConsole.ps1 --
    operator request (2026-07-21): "I don't want to change AdminConsole.ps1
    at all." SRG is a different, cloud-only engagement with its own mail
    domain (srgpros.com) and no on-prem Exchange/AD footprint whatsoever.
    This file, SRGCommon.ps1, SRGTools.json, SRGToolManager.ps1, and every
    SRG*.ps1 tool script are a complete, independent copy of the
    launcher/tool-script architecture -- nothing here is dot-sourced by, or
    dot-sources, the original suite's files, and nothing in the original
    suite was modified to build this.

    Same four jobs as the original AdminConsole.ps1:
      1. Collect the target mailbox/group email address.
      2. Detect the object type (Mailbox / Microsoft 365 Group /
         Distribution Group) via SRGCommon.ps1's Resolve-Target.
      3. Read SRGTools.json and render one button per entry, ALWAYS --
         lookup is optional, not a gate.
      4. Launch the selected tool script IN THIS SAME POWERSHELL SESSION
         (call operator, not a new process) -- same reasoning as the
         original: Connect-ExchangeOnline is per-session state, and a tool
         needs to see whatever's already connected here.

    WHAT'S DIFFERENT FROM AdminConsole.ps1 (this is the whole point of a
    separate file rather than a parameterized version of the original):
      - No Environment radio group, no Domain Controller selector -- not
        commented out, never built at all. SRG has nothing for them to
        choose between (see SRGCommon.ps1 header). Every SRG tool script's
        param block is just -SmtpAddress; Start-ToolScript below launches
        with only that one parameter.
      - Object types known: Mailbox, UnifiedGroup, DistributionGroup (not
        the original's Mailbox/UnifiedGroup/OnpremGroup) -- see
        SRGCommon.ps1's Resolve-Target for why DistributionGroup replaces
        OnpremGroup here.
      - ONE "Tools" button panel, not the original's Admin Tools / Daily
        Scripts split. Only 3 tools exist in SRGTools.json today (the
        operator said most of the original suite's scripts won't carry
        over) -- a single 5-wide x 4-tall FlowLayoutPanel (20-tool nominal
        capacity, same AutoScroll safety net as the original) covers that
        comfortably. If SRG's tool count grows enough to want two
        categories later, split $flowTools into two panels the same way
        the original's REVISED 2026-07-21 note describes, and add a Section
        column back to SRGTools.json/SRGToolManager.ps1.
      - Layout sized tight to its content from the start, not grown-then-
        trimmed like the original went through -- there's no leftover grey
        strip on the right or a dead zone at the bottom to fix later. See
        the exact control coordinates below if the tool list grows enough
        to need more room.
      - Look Up / Clear / Close button row (three buttons) from the start --
        the original only grew a Clear button after the fact; this console
        starts with the same reset behavior Clear gives it (blank the input,
        clear the result/output boxes, release any button grey-out) as a
        first-class button, not something bolted on later.

    ASSUMPTIONS MADE WITHOUT EXPLICIT CONFIRMATION (see SRGCommon.ps1
    header for the same list): single Tools section instead of a two-way
    split; the Invoke-EXOCommand collision guard was kept. Both are easy to
    revisit.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21
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
# REVISED 2026-07-24 -- narrowed 860 -> 600 to make room for the connected-
# tenant label on the same row. See $lblTenant just below.
$lblWarn.Size = New-Object System.Drawing.Size(600, 20)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

# ADDED 2026-07-24 -- shows WHICH tenant Exchange Online is connected to, the
# same information M365_Host_Web.ps1 puts in its top bar. "Connected" on its own
# doesn't say connected to what, and this is one of several consoles used in the
# same day. Right-aligned on the warning row so nothing else had to move.
#
# The value is the tenant's LIVE default accepted domain, which is deliberately
# NOT the same thing as this suite's own $script:MailDomain constant -- if the
# two disagree, the box is appending a domain from a tenant you are not
# connected to, and that is exactly what this makes visible.
$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Location = New-Object System.Drawing.Point(620, 10)
$lblTenant.Size = New-Object System.Drawing.Size(252, 20)
$lblTenant.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblTenant.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

# Repaints the tenant label. Green when known, grey when not -- a red "unknown"
# would read as an error when the honest state is simply "nothing connected
# yet". Cheap after the first call (the lookup is cached), so it is safe to call
# from the Look Up handler as well as at startup.
function Update-TenantLabel {
    $d = Get-ConnectedTenantDomain
    if ($d) {
        $lblTenant.Text = "Tenant: $d"
        $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
        if ($d -ine $script:MailDomain) {
            # Worth flagging loudly: names typed in the box get
            # $script:MailDomain appended, so a mismatch means this console is
            # resolving into a different tenant than the one it is talking to.
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

# Dynamic button surface -- FlowLayoutPanel auto-wraps buttons to new rows,
# so this doesn't need to know in advance how many tools are configured.
# Sized to fit exactly 5 columns (button 150 + Margin 4*2 = 158px each,
# 5*158=790 < 860 panel width) x 4 rows (32 + 4*2 = 40px each, 4*40=160 <
# 172 panel height) before wrapping/scrolling. AutoScroll stays on as a
# safety net if more than 20 tools ever land here.
$flowTools = New-Object System.Windows.Forms.FlowLayoutPanel
$flowTools.Location = New-Object System.Drawing.Point(12, 320)
$flowTools.Size = New-Object System.Drawing.Size(860, 172)
$flowTools.AutoScroll = $true
$flowTools.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($flowTools)

# One shared ToolTip instance services every tool button -- SetToolTip is
# called per-button in Update-ToolButtonState to explain why a button is
# greyed out (or cleared when it isn't).
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

# Launches the tool script with the call operator (&), not Start-Process, so
# it runs in this same PowerShell session/runspace and sees whatever
# Exchange Online connection is already established -- same reasoning as
# the original AdminConsole.ps1's Start-ToolScript (see that file's header
# for the full trade-off discussion; it applies identically here). Only
# -SmtpAddress is passed -- SRG tool scripts have no -Environment /
# -DomainController parameters to fill in.
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

# Set of object types SRGAdminConsole/Resolve-Target actually knows how to
# detect. Anything outside this set (most notably 'NotFound', but also plain
# $null before any lookup has happened) means "no positively-detected type
# to compare a tool's ObjectTypes against" -- see Update-ToolButtonState.
$script:KnownObjectTypes = @('Mailbox', 'UnifiedGroup', 'DistributionGroup')

# Builds the button surface ONCE from $script:AllTools -- every tool gets a
# button, always, regardless of any lookup. Greying out buttons that don't
# match a resolved type is Update-ToolButtonState's job below, run
# separately (and repeatedly, once per lookup) against these same button
# instances.
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
        # NOTE: no .GetNewClosure() here -- WinForms Add_Click scriptblocks
        # get $this bound to whichever control actually raised the event, so
        # each button already resolves its own .Tag correctly with no
        # closure capture required (same reasoning as the original
        # AdminConsole.ps1's Add-ToolButtonsToPanel).
        $btn.Add_Click({ Start-ToolScript -Tool $this.Tag })
        $flowTools.Controls.Add($btn)
    }
}

# Toggles Enabled/tooltip on the already-built buttons -- never adds,
# removes, or reorders a button. $DetectedType outside
# $script:KnownObjectTypes (covers $null before any lookup, and
# 'NotFound'/anything else a failed or no-match lookup produces) means every
# button is re-enabled: there's no positively-known type to declare a
# mismatch against, so nothing is greyed.
function Update-ToolButtonState {
    param([string]$DetectedType)
    $showDimming = ($DetectedType -and ($script:KnownObjectTypes -contains $DetectedType))
    foreach ($ctrl in @($flowTools.Controls)) {
        if (-not $ctrl.Tag) { continue }   # the "no tools found" fallback label, if present
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
            # Blank input doubles as "clear the last lookup" -- releases any
            # grey-out from a previous lookup instead of just erroring (same
            # as Clear below, just reachable via Look Up too).
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
        # Nothing positively resolved -- release any grey-out from a prior
        # lookup rather than leaving it stale against a failed attempt.
        Update-ToolButtonState -DetectedType $null
    } finally {
        $btnRun.Enabled = $true
    }
})
$btnClear.Add_Click({
    # Same reset $btnRun.Add_Click performs when handed blank input, exposed
    # as its own button.
    $txtInput.Text = ''
    $lblResult.Text = 'Not looked up yet.'
    $txtOut.Text = ''
    $script:LastSmtp = $null
    $script:LastType = $null
    Update-ToolButtonState -DetectedType $null
})
$btnClose.Add_Click({ $form.Close() })

# Build the full button surface up front -- no lookup has happened yet, so
# every tool starts enabled.
Build-ToolButtons
Update-ToolButtonState -DetectedType $null
Update-TenantLabel
[void]$form.ShowDialog()
#endregion
