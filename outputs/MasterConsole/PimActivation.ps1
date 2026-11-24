#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- Connect-MgGraph raises an interactive sign-in
#                                 window, which cannot accept input while the
#                                 console's own modal window is open (14d).
# TOOLKIT-HELP-BEGIN
# NOT IN USE. Kept because it works, and only consent is missing.
#
# This tenant will not consent to the Graph permissions self-activation needs
# (PrivilegedEligibilitySchedule.Read.AzureADGroup and
# PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup are admin-consent
# delegated permissions on the Microsoft Graph PowerShell app). Without them
# there is no supported programmatic route -- PIM for Groups is a Graph-only
# API, and the retired AzureADPreview cmdlets covered directory roles, not
# groups.
#
# Use activate.ps1 instead, which opens the portal in the admin profile.
#
# If consent is ever granted, this file should work as written -- it was
# verified against the Graph v1.0 reference on 2026-08-06.
#
# ---------------------------------------------------------------------------
#
# Activates your eligible PIM group assignments, without going to the portal.
#
# Lists the groups you are eligible for, shows which are already active, and
# lets you activate the ones you need with a duration and a justification.
#
#   * Justification is required -- your PIM policy may reject a request
#     without one, and it is recorded either way.
#   * A ticket number can be attached if your policy asks for one.
#   * Some assignments need approval. Those come back as PENDING APPROVAL, not
#     activated, and this tool says so rather than pretending otherwise.
#
# IMPORTANT: activating changes your group membership, but sessions you have
# ALREADY connected -- Exchange Online, SharePoint -- are still using the token
# they were issued at connect time. Reconnect them afterwards or the new access
# will not appear.
#
# Nothing here grants access you did not already have. It activates eligibility
# that has already been assigned to you.
# TOOLKIT-HELP-END
<#
    PimActivation.ps1
    Self-activation of eligible PIM for Groups assignments, via Microsoft Graph.

    Reasoning: ScriptNotes/PimActivation.ps1.notes.md

    Author   : Glenn Armbruster
    Created  : 2026-08-06
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Graph -------------------------------------------------------------------
<#
    Raw REST through Invoke-MgGraphRequest, not the Identity.Governance cmdlets.

    Two reasons, and the second is the one that matters:

    1. Microsoft.Graph.Authentication is one small module. The full SDK is
       ~40 modules and several seconds of import, which is the cold-launch cost
       this toolkit already fought once (14c).

    2. The cmdlet names are generated from the API and have been renamed
       between SDK versions; the REST shape is versioned and stable. Every URI
       and property below was read from the v1.0 reference on 2026-08-06:

         POST /identityGovernance/privilegedAccess/group/assignmentScheduleRequests
         GET  /identityGovernance/privilegedAccess/group/eligibilityScheduleInstances?$filter=principalId eq '{id}'
         GET  /identityGovernance/privilegedAccess/group/assignmentScheduleInstances?$filter=principalId eq '{id}'

       Coding against the documented contract rather than a generated wrapper
       means the docs page IS the spec for this file.
#>

$script:GraphScopes = @(
    'PrivilegedEligibilitySchedule.Read.AzureADGroup'
    'PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup'
    'User.Read'
)

function Test-GraphModule {
    if (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication') { return $true }
    return $false
}

function Connect-GraphForPim {
    <#
        Connects only if the current context is missing a scope we need.

        Reconnecting when already connected is not free -- it can raise another
        sign-in prompt -- and this runs every time the form opens.
    #>
    $ctx = $null
    try { $ctx = Get-MgContext } catch { $ctx = $null }

    if ($ctx -and $ctx.Scopes) {
        $missing = @($script:GraphScopes | Where-Object { $ctx.Scopes -notcontains $_ })
        if ($missing.Count -eq 0) { return $ctx }
    }

    # -NoWelcome only exists in Graph SDK v2. Splatting it conditionally rather
    # than assuming, because on v1 an unknown parameter is a hard binding error
    # and the tool would fail at the one step everything else depends on.
    $p = @{ Scopes = $script:GraphScopes; ErrorAction = 'Stop' }
    $cmd = Get-Command Connect-MgGraph -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Parameters.ContainsKey('NoWelcome')) { $p.NoWelcome = $true }

    Connect-MgGraph @p | Out-Null
    return (Get-MgContext)
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('GET', 'POST')][string]$Method = 'GET',
        $Body
    )
    $p = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
    if ($Body) {
        $p.Body        = ($Body | ConvertTo-Json -Depth 6)
        $p.ContentType = 'application/json'
    }
    return Invoke-MgGraphRequest @p
}

function Get-MyPrincipalId {
    $me = Invoke-Graph -Uri 'https://graph.microsoft.com/v1.0/me?$select=id,displayName,userPrincipalName'
    return [PSCustomObject]@{
        Id   = "$($me.id)"
        Upn  = "$($me.userPrincipalName)"
        Name = "$($me.displayName)"
    }
}

function Get-GroupDisplayName {
    <#
        eligibilityScheduleInstances returns groupId and nothing else readable,
        so the name has to be fetched. Best-effort on purpose: reading a group
        needs a directory scope this tool does not ask for, and a missing name
        is a cosmetic problem while a failed connection is not. Falls back to
        the raw id, which still activates correctly.
    #>
    param([Parameter(Mandatory)][string]$GroupId)
    if (-not $script:GroupNameCache) { $script:GroupNameCache = @{} }
    if ($script:GroupNameCache.ContainsKey($GroupId)) { return $script:GroupNameCache[$GroupId] }

    $name = $GroupId
    try {
        $g = Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId`?`$select=displayName"
        if ($g -and $g.displayName) { $name = "$($g.displayName)" }
    } catch { }
    $script:GroupNameCache[$GroupId] = $name
    return $name
}

function Get-PimEligibility {
    <#
        What I am eligible for, and what is already active.

        Both lists are needed. Activating something already active is not
        harmful, but it is a wasted request and a confusing result, and the
        list is far more readable when it says which is which.
    #>
    param([Parameter(Mandatory)][string]$PrincipalId)

    $base = 'https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group'
    $elig = @()
    $act  = @()

    # $expand=group is the point here. eligibilityScheduleInstances returns only
    # groupId, and resolving each one via /groups/{id} needs a directory read
    # scope this tool deliberately does not request -- so without the expand,
    # the list would show GUIDs and be unusable on a first run where consent
    # covers only the PIM scopes.
    #
    # Falls back to the unexpanded call, because $expand is the newer of the
    # two and a tool that shows GUIDs still activates correctly.
    $r = $null
    try {
        $r = Invoke-Graph -Uri "$base/eligibilityScheduleInstances?`$filter=principalId eq '$PrincipalId'&`$expand=group"
    } catch {
        $r = Invoke-Graph -Uri "$base/eligibilityScheduleInstances?`$filter=principalId eq '$PrincipalId'"
    }
    if ($r.value) { $elig = @($r.value) }

    try {
        $r2 = Invoke-Graph -Uri "$base/assignmentScheduleInstances?`$filter=principalId eq '$PrincipalId'"
        if ($r2.value) { $act = @($r2.value) }
    } catch {
        # Non-fatal. Without it the list simply cannot say "already active",
        # which is worth less than the whole tool refusing to open.
    }

    $out = New-Object System.Collections.ArrayList
    foreach ($e in $elig) {
        $gid = "$($e.groupId)"
        $aid = "$($e.accessId)"
        $live = @($act | Where-Object { "$($_.groupId)" -eq $gid -and "$($_.accessId)" -eq $aid })
        $until = $null
        if ($live.Count -gt 0 -and $live[0].endDateTime) {
            try { $until = [DateTimeOffset]::Parse("$($live[0].endDateTime)").ToLocalTime() } catch { }
        }
        # Expanded name first -- it arrived with the list and costs nothing.
        $gname = $null
        if ($e.group -and $e.group.displayName) { $gname = "$($e.group.displayName)" }
        if (-not $gname) { $gname = Get-GroupDisplayName -GroupId $gid }

        [void]$out.Add([PSCustomObject]@{
            GroupId   = $gid
            GroupName = $gname
            AccessId  = $aid
            Active    = ($live.Count -gt 0)
            ActiveUntil = $until
        })
    }
    return @($out.ToArray() | Sort-Object GroupName, AccessId)
}

function Request-PimActivation {
    <#
        One selfActivate request.

        Returns the request's own status rather than a boolean. PIM can accept
        a request and NOT activate it -- an assignment behind approval comes
        back PendingApproval -- and reporting that as success would be the
        worst kind of wrong: you would go and act as though you had access.
    #>
    param(
        [Parameter(Mandatory)][string]$PrincipalId,
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$AccessId,
        [Parameter(Mandatory)][int]$Hours,
        [Parameter(Mandatory)][string]$Justification,
        [string]$TicketNumber,
        [string]$TicketSystem
    )

    $body = @{
        accessId      = $AccessId
        principalId   = $PrincipalId
        groupId       = $GroupId
        action        = 'selfActivate'
        justification = $Justification
        scheduleInfo  = @{
            # Round-trip format, and UTC. PIM rejects an ambiguous local time,
            # and a start time in the past by more than a moment.
            startDateTime = (Get-Date).ToUniversalTime().ToString('o')
            expiration    = @{
                type     = 'afterDuration'
                duration = "PT$($Hours)H"
            }
        }
    }
    if ($TicketNumber -or $TicketSystem) {
        $body.ticketInfo = @{ ticketNumber = "$TicketNumber"; ticketSystem = "$TicketSystem" }
    }

    $r = Invoke-Graph -Method POST -Body $body `
            -Uri 'https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/assignmentScheduleRequests'
    return "$($r.status)"
}
#endregion

#region UI ----------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'PIM -- activate eligible access'
$form.Size          = New-Object System.Drawing.Size(820, 640)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWho           = New-Object System.Windows.Forms.Label
$lblWho.Tag       = 'NoTheme'
$lblWho.Location  = New-Object System.Drawing.Point(12, 12)
$lblWho.Size      = New-Object System.Drawing.Size(700, 18)
$lblWho.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWho.Text      = 'Not connected.'
$lblWho.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblWho)

$lblList          = New-Object System.Windows.Forms.Label
$lblList.Text     = 'Eligible assignments (tick the ones to activate):'
$lblList.Location = New-Object System.Drawing.Point(12, 38)
$lblList.Size     = New-Object System.Drawing.Size(400, 18)
$form.Controls.Add($lblList)

$clb            = New-Object System.Windows.Forms.CheckedListBox
$clb.Location   = New-Object System.Drawing.Point(12, 58)
$clb.Size       = New-Object System.Drawing.Size(780, 190)
$clb.CheckOnClick = $true
$clb.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($clb)

$btnRefresh          = New-Object System.Windows.Forms.Button
$btnRefresh.Text     = 'Refresh'
$btnRefresh.Location = New-Object System.Drawing.Point(12, 256)
$btnRefresh.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnRefresh)

$btnAll          = New-Object System.Windows.Forms.Button
$btnAll.Text     = 'Select inactive'
$btnAll.Location = New-Object System.Drawing.Point(110, 256)
$btnAll.Size     = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnAll)

$lblDur          = New-Object System.Windows.Forms.Label
$lblDur.Text     = 'Duration:'
$lblDur.Location = New-Object System.Drawing.Point(12, 300)
$lblDur.Size     = New-Object System.Drawing.Size(60, 18)
$form.Controls.Add($lblDur)

$cmbDur                = New-Object System.Windows.Forms.ComboBox
$cmbDur.Location       = New-Object System.Drawing.Point(76, 297)
$cmbDur.Size           = New-Object System.Drawing.Size(110, 24)
$cmbDur.DropDownStyle  = 'DropDownList'
foreach ($h in 1, 2, 4, 8) { [void]$cmbDur.Items.Add("$h hour$(if ($h -gt 1) { 's' })") }
$cmbDur.SelectedIndex  = 3
$form.Controls.Add($cmbDur)

$lblDurNote           = New-Object System.Windows.Forms.Label
$lblDurNote.Tag       = 'NoTheme'
$lblDurNote.Text      = 'Your PIM policy caps this. Asking for more than the cap fails, and the error names the maximum.'
$lblDurNote.Location  = New-Object System.Drawing.Point(196, 300)
$lblDurNote.Size      = New-Object System.Drawing.Size(600, 18)
$lblDurNote.ForeColor = [System.Drawing.Color]::DimGray
$lblDurNote.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Italic)
$form.Controls.Add($lblDurNote)

$lblJust          = New-Object System.Windows.Forms.Label
$lblJust.Text     = 'Justification (required):'
$lblJust.Location = New-Object System.Drawing.Point(12, 330)
$lblJust.Size     = New-Object System.Drawing.Size(200, 18)
$form.Controls.Add($lblJust)

$txtJust          = New-Object System.Windows.Forms.TextBox
$txtJust.Location = New-Object System.Drawing.Point(12, 350)
$txtJust.Size     = New-Object System.Drawing.Size(780, 24)
$form.Controls.Add($txtJust)

$lblTicket          = New-Object System.Windows.Forms.Label
$lblTicket.Text     = 'Ticket number (optional):'
$lblTicket.Location = New-Object System.Drawing.Point(12, 384)
$lblTicket.Size     = New-Object System.Drawing.Size(150, 18)
$form.Controls.Add($lblTicket)

$txtTicket          = New-Object System.Windows.Forms.TextBox
$txtTicket.Location = New-Object System.Drawing.Point(166, 381)
$txtTicket.Size     = New-Object System.Drawing.Size(180, 24)
$form.Controls.Add($txtTicket)

$lblSys          = New-Object System.Windows.Forms.Label
$lblSys.Text     = 'System:'
$lblSys.Location = New-Object System.Drawing.Point(360, 384)
$lblSys.Size     = New-Object System.Drawing.Size(50, 18)
$form.Controls.Add($lblSys)

$txtSys          = New-Object System.Windows.Forms.TextBox
$txtSys.Location = New-Object System.Drawing.Point(412, 381)
$txtSys.Size     = New-Object System.Drawing.Size(180, 24)
$txtSys.Text     = 'ServiceNow'
$form.Controls.Add($txtSys)

$btnGo          = New-Object System.Windows.Forms.Button
$btnGo.Text     = 'Activate'
$btnGo.Location = New-Object System.Drawing.Point(12, 416)
$btnGo.Size     = New-Object System.Drawing.Size(120, 32)
$btnGo.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnGo)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(692, 416)
$btnClose.Size     = New-Object System.Drawing.Size(100, 32)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.WordWrap   = $true
$txtOut.Location   = New-Object System.Drawing.Point(12, 458)
$txtOut.Size       = New-Object System.Drawing.Size(780, 128)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.BackColor  = [System.Drawing.Color]::WhiteSmoke
$txtOut.ForeColor  = [System.Drawing.Color]::Black
$txtOut.Tag        = 'NoTheme'
$form.Controls.Add($txtOut)

$script:Me    = $null
$script:Items = @()

function Write-Out {
    param([string]$Text)
    $txtOut.AppendText("$Text`r`n")
    $txtOut.SelectionStart = $txtOut.TextLength
    $txtOut.ScrollToCaret()
}

function Update-List {
    $clb.Items.Clear()
    $script:Items = @()
    if (-not $script:Me) { return }
    try {
        $script:Items = @(Get-PimEligibility -PrincipalId $script:Me.Id)
    } catch {
        Write-Out "Could not read your eligible assignments: $($_.Exception.Message)"
        return
    }
    if ($script:Items.Count -eq 0) {
        Write-Out 'No eligible PIM group assignments found for your account.'
        return
    }
    foreach ($i in $script:Items) {
        $state = if ($i.Active) {
                     if ($i.ActiveUntil) { "ACTIVE until $($i.ActiveUntil.ToString('HH:mm'))" } else { 'ACTIVE' }
                 } else { 'eligible' }
        [void]$clb.Items.Add(("{0,-12} {1,-9} {2}" -f $state, $i.AccessId, $i.GroupName))
    }
}

$btnRefresh.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try { Update-List } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnAll.Add_Click({
    # Inactive only. "Select all" would tick things already active, and the
    # requests for those are wasted at best and confusing at worst.
    for ($i = 0; $i -lt $clb.Items.Count; $i++) {
        $clb.SetItemChecked($i, (-not $script:Items[$i].Active))
    }
})

$btnGo.Add_Click({
    $just = "$($txtJust.Text)".Trim()
    if (-not $just) {
        [void][System.Windows.Forms.MessageBox]::Show(
            'A justification is required. Your PIM policy may reject the request without one, and it is recorded on the request either way.',
            'Justification needed', 'OK', 'Warning')
        $txtJust.Focus(); return
    }

    $picked = @()
    for ($i = 0; $i -lt $clb.Items.Count; $i++) {
        if ($clb.GetItemChecked($i)) { $picked += $script:Items[$i] }
    }
    if ($picked.Count -eq 0) { Write-Out 'Nothing selected.'; return }

    $hours = @(1, 2, 4, 8)[$cmbDur.SelectedIndex]
    $names = ($picked | ForEach-Object { "$($_.AccessId) of $($_.GroupName)" }) -join "`r`n    "
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Activate for $hours hour(s)?`r`n`r`n    $names`r`n`r`nJustification: $just",
        'Confirm activation', 'YesNo', 'Question')
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $btnGo.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $ok = 0; $pending = 0; $fail = 0
    Write-Out ''
    Write-Out "--- $(Get-Date -Format 'HH:mm:ss')  activating $($picked.Count) assignment(s) for $hours h ---"
    try {
        foreach ($p in $picked) {
            $label = "$($p.AccessId) of $($p.GroupName)"
            try {
                $status = Request-PimActivation -PrincipalId $script:Me.Id -GroupId $p.GroupId `
                            -AccessId $p.AccessId -Hours $hours -Justification $just `
                            -TicketNumber "$($txtTicket.Text)".Trim() -TicketSystem "$($txtSys.Text)".Trim()
                if ($status -match '(?i)provisioned|granted') {
                    Write-Out "  ACTIVATED        $label"; $ok++
                } else {
                    # Most often PendingApproval. Reported as-is -- an approval
                    # request is NOT access, and treating it as success is how
                    # somebody goes and tries to use it.
                    Write-Out "  $($status.ToUpper().PadRight(16)) $label   <-- not active yet"; $pending++
                }
            } catch {
                Write-Out "  FAILED           $label"
                Write-Out "      $($_.Exception.Message)"
                $fail++
            }
        }
    } finally {
        $btnGo.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }

    Write-Out "Done. Activated: $ok   Pending: $pending   Failed: $fail"
    if ($ok -gt 0) {
        Write-Out ''
        Write-Out 'Reconnect Exchange Online / SharePoint before using the new access.'
        Write-Out 'Existing sessions still hold the token they were issued at connect time,'
        Write-Out 'which does not include the membership you just activated.'
    }

    try {
        Write-CompletedLog -Action 'PIM Activate' -Target $script:Me.Upn `
            -Result $(if ($fail -gt 0) { 'Failed' } elseif ($pending -gt 0) { 'Partial' } else { 'Success' }) `
            -Detail "$hours h; $($picked.Count) requested; $ok active, $pending pending, $fail failed; justification: $just" `
            -Environment 'Cloud'
    } catch { }

    Update-List
})

# --- startup ------------------------------------------------------------------
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'PIM Activate' | Out-Null

$form.Add_Shown({
    if (-not (Test-GraphModule)) {
        $lblWho.Text = 'Microsoft.Graph.Authentication is not installed.'
        $lblWho.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
        Write-Out 'The Microsoft.Graph.Authentication module is required and was not found.'
        Write-Out ''
        Write-Out '    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
        Write-Out ''
        Write-Out 'Only that one module is needed -- not the full Microsoft.Graph SDK.'
        $btnGo.Enabled = $false; $btnRefresh.Enabled = $false; $btnAll.Enabled = $false
        return
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Write-Out 'Connecting to Microsoft Graph...'
        Connect-GraphForPim | Out-Null
        $script:Me = Get-MyPrincipalId
        $lblWho.Text = "Signed in as $($script:Me.Upn)"
        $lblWho.ForeColor = [System.Drawing.Color]::FromArgb(0, 110, 60)
        Write-Out "Connected as $($script:Me.Upn)."
        Update-List
    } catch {
        $lblWho.Text = 'Not connected to Microsoft Graph.'
        $lblWho.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
        Write-Out "Could not connect: $($_.Exception.Message)"
        Write-Out ''
        Write-Out 'If consent was refused, these are the scopes requested and why:'
        foreach ($s in $script:GraphScopes) { Write-Out "    $s" }
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

[void]$form.ShowDialog()
$form.Dispose()
#endregion
