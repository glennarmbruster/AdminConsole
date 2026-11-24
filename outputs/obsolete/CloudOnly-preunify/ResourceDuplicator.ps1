#Requires -Version 5.1
<#
    ResourceDuplicator.ps1
    Duplicates a room or equipment mailbox: reads every setting off an existing
    resource, creates a new one, and applies the captured settings to it.

    Cloud-only. The hybrid variant lives in the tenant-specific toolkit.

    Reasoning: ScriptNotes/ResourceDuplicator.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Copies a room or equipment mailbox: reads every setting off an existing
# resource, creates a new one, and applies the same settings.
#
# For when a new meeting room should behave exactly like an existing one --
# booking rules, delegates, capacity, auto-accept behaviour. Copying is more
# reliable than setting a dozen options by hand and hoping you remembered them
# all.
#
# Check the source resource is actually configured the way you want before
# copying it. Its mistakes get copied too.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration --------------------------------------------------------

# Under the reports root, not beside it. Captures are OUTPUT, and putting them
# here means they inherit the Reports folder setting -- point reports at a share
# and a duplication started on one machine can be finished on another. A
# subfolder rather than loose in reports\ because ReportBrowser lists *.csv, and
# a folder of JSON alongside would make that view mixed-purpose for no gain.
$CaptureDir = Join-Path (Get-ReportPath) 'captures'
if (-not (Test-Path -LiteralPath $CaptureDir)) {
    New-Item -Path $CaptureDir -ItemType Directory -Force | Out-Null
}

# The full Set-Place surface. All of it works on a cloud mailbox.
$PlaceCloudProps = @(
    'Building', 'Capacity', 'Floor', 'FloorLabel', 'Label', 'Tags',
    'AudioDeviceName', 'VideoDeviceName', 'DisplayDeviceName',
    'IsWheelChairAccessible', 'MTREnabled',
    'City', 'State', 'PostalCode', 'Street', 'Phone', 'CountryOrRegion'
)

# Booking behaviour. Identity and the resource's own address are excluded --
# copying those would either fail or point the new room at the old one.
$CalendarProps = @(
    'AutomateProcessing', 'AllowConflicts', 'BookingWindowInDays',
    'MaximumDurationInMinutes', 'MinimumDurationInMinutes', 'AllowRecurringMeetings',
    'EnforceSchedulingHorizon', 'ScheduleOnlyDuringWorkHours', 'ConflictPercentageAllowed',
    'MaximumConflictInstances', 'ForwardRequestsToDelegates', 'DeleteAttachments',
    'DeleteComments', 'RemovePrivateProperty', 'DeleteSubject', 'AddOrganizerToSubject',
    'DeleteNonCalendarItems', 'TentativePendingApproval', 'EnableResponseDetails',
    'OrganizerInfo', 'ResourceDelegates', 'RequestOutOfPolicy', 'AllRequestOutOfPolicy',
    'BookInPolicy', 'AllBookInPolicy', 'RequestInPolicy', 'AllRequestInPolicy',
    'AddAdditionalResponse', 'AdditionalResponse', 'RemoveOldMeetingMessages',
    'AddNewRequestsTentatively', 'ProcessExternalMeetingMessages',
    'RemoveForwardedMeetingNotifications', 'EnforceCapacity'
)
#endregion

#region Capture --------------------------------------------------------------
function Get-ResourceCapture {
    # Every CLOUD call goes through Invoke-EXOCommand: Get-Mailbox exists in
    # both worlds and resolves by module load order (ARCHITECTURE 2).
    param([Parameter(Mandatory)][string]$Identity, [scriptblock]$Progress)

    & $Progress 'Reading the source mailbox...' 10
    $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Identity; ErrorAction = 'Stop' }
    if (-not $mbx) { throw "'$Identity' was not found in Exchange Online." }

    $type = "$($mbx.RecipientTypeDetails)"
    if ($type -notin @('RoomMailbox', 'EquipmentMailbox')) {
        throw "'$Identity' is a $type. This tool duplicates room and equipment mailboxes only."
    }

    & $Progress 'Reading calendar processing...' 30
    $cal = Invoke-EXOCommand -Name 'Get-CalendarProcessing' -Parameters @{ Identity = $Identity; ErrorAction = 'Stop' }

    $place = $null
    if ($type -eq 'RoomMailbox') {
        & $Progress 'Reading room metadata...' 45
        try {
            $place = Invoke-EXOCommand -Name 'Get-Place' -Parameters @{ Identity = $Identity; ErrorAction = 'Stop' }
        } catch { }
    }

    & $Progress 'Reading calendar permissions...' 60
    $perms = @()
    try {
        $folder = "$($mbx.PrimarySmtpAddress):\Calendar"
        $perms = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' `
                        -Parameters @{ Identity = $folder; ErrorAction = 'Stop' } |
                   Where-Object { "$($_.User)" -notin @('Default', 'Anonymous') -or
                                  "$($_.AccessRights)" -notin @('None', 'AvailabilityOnly') })
    } catch { }

    & $Progress 'Reading regional settings...' 75
    $regional = $null
    try {
        $regional = Invoke-EXOCommand -Name 'Get-MailboxRegionalConfiguration' `
                        -Parameters @{ Identity = $Identity; ErrorAction = 'Stop' }
    } catch { }

    & $Progress 'Checking room list membership...' 88
    $roomLists = @()
    try {
        $addr = "$($mbx.PrimarySmtpAddress)"
        foreach ($rl in @(Invoke-EXOCommand -Name 'Get-DistributionGroup' `
                            -Parameters @{ RecipientTypeDetails = 'RoomList'; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })) {
            $members = @(Invoke-EXOCommand -Name 'Get-DistributionGroupMember' `
                            -Parameters @{ Identity = "$($rl.Identity)"; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
            if ($members | Where-Object { "$($_.PrimarySmtpAddress)" -eq $addr }) {
                $roomLists += "$($rl.PrimarySmtpAddress)"
            }
        }
    } catch { }

    $capture = [ordered]@{
        CapturedAt          = (Get-Date).ToString('s')
        SourceIdentity      = "$($mbx.PrimarySmtpAddress)"
        SourceDisplayName   = "$($mbx.DisplayName)"
        RecipientTypeDetails= $type
        ResourceCapacity    = $mbx.ResourceCapacity
        HiddenFromGAL       = [bool]$mbx.HiddenFromAddressListsEnabled
        CustomAttributes    = [ordered]@{}
        CalendarProcessing  = [ordered]@{}
        Place               = [ordered]@{}
        CalendarPermissions = @()
        Regional            = [ordered]@{}
        RoomLists           = @($roomLists)
    }

    for ($i = 1; $i -le 15; $i++) {
        $n = "CustomAttribute$i"
        if ($mbx.$n) { $capture.CustomAttributes[$n] = "$($mbx.$n)" }
    }
    foreach ($p in $CalendarProps) {
        if ($cal.PSObject.Properties[$p]) { $capture.CalendarProcessing[$p] = $cal.$p }
    }
    if ($place) {
        foreach ($p in ($PlaceCloudProps + 'GeoCoordinates')) {
            if ($place.PSObject.Properties[$p]) { $capture.Place[$p] = $place.$p }
        }
    }
    foreach ($p in $perms) {
        $capture.CalendarPermissions += [ordered]@{
            User = "$($p.User)"; AccessRights = @($p.AccessRights | ForEach-Object { "$_" })
        }
    }
    if ($regional) {
        foreach ($p in @('Language','DateFormat','TimeFormat','TimeZone')) {
            if ($regional.PSObject.Properties[$p] -and $regional.$p) { $capture.Regional[$p] = "$($regional.$p)" }
        }
    }
    return [PSCustomObject]$capture
}

function Get-CaptureFiles {
    if (-not (Test-Path -LiteralPath $CaptureDir)) { return @() }
    return @(Get-ChildItem -LiteralPath $CaptureDir -Filter 'Resource_*.json' -File -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending)
}

function Import-Capture {
    param([Parameter(Mandatory)][string]$Path)
    return ((Get-Content -LiteralPath $Path -Raw -ErrorAction Stop) | ConvertFrom-Json -ErrorAction Stop)
}

function Save-Capture {
    param($Capture, [string]$Path)
    if ($Path) {
        $path = $Path
    } else {
        $safe  = ($Capture.SourceDisplayName -replace '[^\w\-]', '_')
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $path  = Join-Path $CaptureDir "Resource_${safe}_$stamp.json"
    }
    $json  = $Capture | ConvertTo-Json -Depth 8
    $enc   = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($path, ($json -replace "`r?`n", "`r`n"), $enc)
    return $path
}
#endregion

#region Create ---------------------------------------------------------------
function New-DuplicateResource {
    param(
        [Parameter(Mandatory)]$Capture,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Alias,
        [Parameter(Mandatory)][string]$Upn,
        [string]$DisplayName,
        [scriptblock]$Progress
    )

    & $Progress 'Creating the resource mailbox...' 30

    $p = @{
        Name               = $Name
        Alias              = $Alias
        PrimarySmtpAddress = $Upn
        DisplayName        = if ($DisplayName) { $DisplayName } else { $Name }
        ErrorAction        = 'Stop'
    }

    # Only parameters the installed cmdlet declares -- the surface varies by version.
    $declared = Get-CmdletParameterNames -Name 'New-Mailbox'
    if ($Capture.RecipientTypeDetails -eq 'RoomMailbox') {
        if ($declared -notcontains 'Room') { throw 'New-Mailbox does not declare -Room on this build.' }
        $p['Room'] = $true
    } else {
        if ($declared -notcontains 'Equipment') { throw 'New-Mailbox does not declare -Equipment on this build.' }
        $p['Equipment'] = $true
    }
    foreach ($k in @($p.Keys)) {
        if ($k -notin @('ErrorAction') -and $declared -notcontains $k) { $p.Remove($k) }
    }

    $new = Invoke-EXOCommand -Name 'New-Mailbox' -Parameters $p

    & $Progress 'Waiting for the mailbox to become addressable...' 70
    $addr = if ($new.PrimarySmtpAddress) { "$($new.PrimarySmtpAddress)" } else { $Upn }
    $ready = $false
    for ($i = 0; $i -lt 10; $i++) {
        Start-Sleep -Seconds 3
        try {
            $null = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $addr; ErrorAction = 'Stop' }
            $ready = $true; break
        } catch { }
        & $Progress "Waiting for the mailbox to become addressable... ($(($i + 1) * 3)s)" (70 + $i)
    }

    return [PSCustomObject]@{
        PrimarySmtp = $addr
        Ready       = $ready
        Warnings    = @(if (-not $ready) {
            'The mailbox was created but is not resolvable yet. Wait a minute and click Apply Settings.'
        })
    }
}
#endregion

#region Apply ----------------------------------------------------------------
function Set-DuplicateResourceSettings {
    param(
        [Parameter(Mandatory)]$Capture,
        [Parameter(Mandatory)][string]$Identity,
        [scriptblock]$Progress
    )

    $applied = @(); $failed = @(); $skipped = @()

    & $Progress 'Confirming the mailbox exists...' 5
    $mbx = $null
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Identity; ErrorAction = 'Stop' }
    } catch { }
    if (-not $mbx) {
        throw "'$Identity' was not found in Exchange Online. If you have just created it, give it a moment and click Apply again."
    }

    & $Progress 'Applying calendar processing...' 20
    $cp = @{ Identity = $Identity; ErrorAction = 'Stop' }
    $declared = Get-CmdletParameterNames -Name 'Set-CalendarProcessing'
    foreach ($k in $Capture.CalendarProcessing.PSObject.Properties.Name) {
        if ($declared -contains $k) { $cp[$k] = $Capture.CalendarProcessing.$k }
    }
    try {
        Invoke-EXOCommand -Name 'Set-CalendarProcessing' -Parameters $cp | Out-Null
        $applied += "Calendar processing ($($cp.Count - 2) properties)"
    } catch { $failed += "Calendar processing: $($_.Exception.Message)" }

    if ($Capture.ResourceCapacity) {
        try {
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
                Identity = $Identity; ResourceCapacity = [int]$Capture.ResourceCapacity; ErrorAction = 'Stop' } | Out-Null
            $applied += "Resource capacity ($($Capture.ResourceCapacity))"
        } catch { $failed += "Resource capacity: $($_.Exception.Message)" }
    }

    if ($Capture.RecipientTypeDetails -eq 'RoomMailbox' -and $Capture.Place.PSObject.Properties.Name.Count -gt 0) {
        & $Progress 'Applying room metadata...' 45
        $pl = @{ Identity = $Identity; ErrorAction = 'Stop' }
        $plDeclared = Get-CmdletParameterNames -Name 'Set-Place'
        foreach ($k in $PlaceCloudProps) {
            if ($Capture.Place.PSObject.Properties[$k] -and
                $null -ne $Capture.Place.$k -and
                "$($Capture.Place.$k)" -ne '' -and
                $plDeclared -contains $k) {
                $pl[$k] = $Capture.Place.$k
            }
        }
        if ($pl.Count -gt 2) {
            try {
                Invoke-EXOCommand -Name 'Set-Place' -Parameters $pl | Out-Null
                $applied += "Room metadata ($($pl.Count - 2) properties)"
            } catch { $failed += "Room metadata: $($_.Exception.Message)" }
        }
        if ($Capture.Place.PSObject.Properties['GeoCoordinates'] -and $Capture.Place.GeoCoordinates) {
            $skipped += "GeoCoordinates '$($Capture.Place.GeoCoordinates)' -- Set-Place accepts it but it is captured separately; check it landed if location search matters."
        }
    }

    if (@($Capture.CalendarPermissions).Count -gt 0) {
        & $Progress 'Applying calendar permissions...' 62
        $folder = "$($mbx.PrimarySmtpAddress):\Calendar"
        foreach ($perm in @($Capture.CalendarPermissions)) {
            $u = "$($perm.User)"; $r = @($perm.AccessRights)
            try {
                Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' `
                    -Parameters @{ Identity = $folder; User = $u; AccessRights = $r; ErrorAction = 'Stop' } | Out-Null
                $applied += "Calendar permission: $u"
            } catch {
        # Set fails when no entry exists yet; Add is the other half.
                try {
                    Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' `
                        -Parameters @{ Identity = $folder; User = $u; AccessRights = $r; ErrorAction = 'Stop' } | Out-Null
                    $applied += "Calendar permission: $u"
                } catch { $failed += "Calendar permission $u : $($_.Exception.Message)" }
            }
        }
    }

    if ($Capture.Regional.PSObject.Properties.Name.Count -gt 0) {
        & $Progress 'Applying regional settings...' 78
        $rg = @{ Identity = $Identity; ErrorAction = 'Stop' }
        $rgDeclared = Get-CmdletParameterNames -Name 'Set-MailboxRegionalConfiguration'
        foreach ($k in $Capture.Regional.PSObject.Properties.Name) {
            if ($rgDeclared -contains $k) { $rg[$k] = $Capture.Regional.$k }
        }
        try {
            Invoke-EXOCommand -Name 'Set-MailboxRegionalConfiguration' -Parameters $rg | Out-Null
            $applied += "Regional settings ($($Capture.Regional.TimeZone))"
        } catch { $failed += "Regional settings: $($_.Exception.Message)" }
    }

    if (@($Capture.RoomLists).Count -gt 0) {
        & $Progress 'Adding to room lists...' 90
        foreach ($rl in @($Capture.RoomLists)) {
            try {
                Invoke-EXOCommand -Name 'Add-DistributionGroupMember' `
                    -Parameters @{ Identity = $rl; Member = $Identity; ErrorAction = 'Stop' } | Out-Null
                $applied += "Room list: $rl"
            } catch { $failed += "Room list $rl : $($_.Exception.Message)" }
        }
    }

    if ($Capture.HiddenFromGAL) {
        try {
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
                Identity = $Identity; HiddenFromAddressListsEnabled = $true; ErrorAction = 'Stop' } | Out-Null
            $applied += 'Hidden from address lists'
        } catch { $failed += "Hide from GAL: $($_.Exception.Message)" }
    }

    return [PSCustomObject]@{ Applied = $applied; Failed = $failed; Skipped = $skipped }
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Resource Duplicator -- rooms and equipment'
$form.Size            = New-Object System.Drawing.Size(820, 780)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Copies a room or equipment mailbox: booking policy, room metadata, calendar permissions, time zone and room lists.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(780, 18)
$form.Controls.Add($lblInfo)

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Text      = 'Cloud-only: needs Exchange Online connected in this window. Every call is forced to the EXO module, so an on-prem session loaded alongside cannot capture it.'
$lblWarn.Location  = New-Object System.Drawing.Point(12, 32)
$lblWarn.Size      = New-Object System.Drawing.Size(780, 18)
$lblWarn.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblWarn)

$grpSrc          = New-Object System.Windows.Forms.GroupBox
$grpSrc.Text     = 'Step 1  --  source resource'
$grpSrc.Location = New-Object System.Drawing.Point(12, 56)
$grpSrc.Size     = New-Object System.Drawing.Size(780, 62)
$form.Controls.Add($grpSrc)

$txtSource          = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(12, 24)
$txtSource.Size     = New-Object System.Drawing.Size(520, 22)
$txtSource.Text     = $SmtpAddress
$grpSrc.Controls.Add($txtSource)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load Source'
$btnLoad.Location = New-Object System.Drawing.Point(542, 22)
$btnLoad.Size     = New-Object System.Drawing.Size(120, 26)
$grpSrc.Controls.Add($btnLoad)

$grpNew          = New-Object System.Windows.Forms.GroupBox
$grpNew.Text     = 'Step 2  --  new resource  (created in Exchange Online)'
$grpNew.Location = New-Object System.Drawing.Point(12, 126)
$grpNew.Size     = New-Object System.Drawing.Size(780, 166)
$form.Controls.Add($grpNew)

$fields = @{}
$y = 22
foreach ($f in @(
    @{ Key='Name';    Label='Name:';         Hint='AD object name, e.g. Conference Room A' }
    @{ Key='Display'; Label='Display name:'; Hint='Blank = same as Name' }
    @{ Key='Alias';   Label='Alias:';        Hint='No spaces, e.g. ConferenceRoomA' }
    @{ Key='Upn';     Label='Email / UPN:';  Hint='Full address' }
)) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $f.Label
    $l.Location = New-Object System.Drawing.Point(12, ($y + 3))
    $l.Size = New-Object System.Drawing.Size(90, 20)
    $grpNew.Controls.Add($l)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(106, $y)
    $tb.Size = New-Object System.Drawing.Size(300, 22)
    $grpNew.Controls.Add($tb)
    $fields[$f.Key] = $tb

    $h = New-Object System.Windows.Forms.Label
    $h.Text = $f.Hint
    $h.Location = New-Object System.Drawing.Point(414, ($y + 3))
    $h.Size = New-Object System.Drawing.Size(350, 18)
    $h.ForeColor = [System.Drawing.Color]::DimGray
    $grpNew.Controls.Add($h)
    $y += 26
}

$btnCreate          = New-Object System.Windows.Forms.Button
$btnCreate.Text     = 'Create Mailbox'
$btnCreate.Location = New-Object System.Drawing.Point(106, 128)
$btnCreate.Size     = New-Object System.Drawing.Size(130, 28)
$btnCreate.Enabled  = $false
$grpNew.Controls.Add($btnCreate)

$grpApply          = New-Object System.Windows.Forms.GroupBox
$grpApply.Text     = 'Step 3  --  apply settings  (run straight after Create, or in a later session)'
$grpApply.Location = New-Object System.Drawing.Point(12, 300)
$grpApply.Size     = New-Object System.Drawing.Size(780, 96)
$form.Controls.Add($grpApply)

$lblCap          = New-Object System.Windows.Forms.Label
$lblCap.Text     = 'Capture:'
$lblCap.Location = New-Object System.Drawing.Point(12, 27)
$lblCap.Size     = New-Object System.Drawing.Size(90, 20)
$grpApply.Controls.Add($lblCap)

$cboCapture               = New-Object System.Windows.Forms.ComboBox
$cboCapture.DropDownStyle = 'DropDownList'
$cboCapture.Location      = New-Object System.Drawing.Point(106, 24)
$cboCapture.Size          = New-Object System.Drawing.Size(560, 22)
$grpApply.Controls.Add($cboCapture)

$lblTgt          = New-Object System.Windows.Forms.Label
$lblTgt.Text     = 'New resource:'
$lblTgt.Location = New-Object System.Drawing.Point(12, 59)
$lblTgt.Size     = New-Object System.Drawing.Size(90, 20)
$grpApply.Controls.Add($lblTgt)

$txtTarget          = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(106, 56)
$txtTarget.Size     = New-Object System.Drawing.Size(300, 22)
$grpApply.Controls.Add($txtTarget)

$btnApply          = New-Object System.Windows.Forms.Button
$btnApply.Text     = 'Apply Settings'
$btnApply.Location = New-Object System.Drawing.Point(420, 54)
$btnApply.Size     = New-Object System.Drawing.Size(140, 28)
$btnApply.Enabled  = $false
$grpApply.Controls.Add($btnApply)

$btnFolder          = New-Object System.Windows.Forms.Button
$btnFolder.Text     = 'Open Captures'
$btnFolder.Location = New-Object System.Drawing.Point(572, 54)
$btnFolder.Size     = New-Object System.Drawing.Size(120, 28)
$grpApply.Controls.Add($btnFolder)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(682, 404)
$btnClose.Size     = New-Object System.Drawing.Size(110, 28)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Enter an existing room or equipment mailbox and click Load Source -- or pick a capture below to finish an earlier one.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 410)
$lblStatus.Size     = New-Object System.Drawing.Size(660, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 438 -Width 780 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 456)
$txtOut.Size       = New-Object System.Drawing.Size(780, 268)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$script:Capture     = $null
$script:CapturePath = $null
$script:CaptureFiles = @()

function Update-CaptureList {
    # Also runs on Form.Activated, so a capture written by another run appears.
    $keep = $null
    if ($cboCapture.SelectedIndex -ge 0 -and $cboCapture.SelectedIndex -lt @($script:CaptureFiles).Count) {
        $keep = $script:CaptureFiles[$cboCapture.SelectedIndex].FullName
    }
    $script:CaptureFiles = @(Get-CaptureFiles)
    $cboCapture.Items.Clear()

    if (@($script:CaptureFiles).Count -eq 0) {
        [void]$cboCapture.Items.Add('(no captures yet -- run Step 1)')
        $cboCapture.SelectedIndex = 0
        $cboCapture.Enabled = $false
        $btnApply.Enabled   = $false
        return
    }
    foreach ($f in $script:CaptureFiles) {
        [void]$cboCapture.Items.Add(('{0}   {1}' -f $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm'), $f.Name))
    }
    $cboCapture.Enabled = $true

    $idx = 0
    if ($keep) {
        for ($i = 0; $i -lt @($script:CaptureFiles).Count; $i++) {
            if ($script:CaptureFiles[$i].FullName -eq $keep) { $idx = $i; break }
        }
    }
    $cboCapture.SelectedIndex = $idx
}

$cboCapture.Add_SelectedIndexChanged({
    if (@($script:CaptureFiles).Count -eq 0) { return }
    $i = $cboCapture.SelectedIndex
    if ($i -lt 0 -or $i -ge @($script:CaptureFiles).Count) { return }
    try {
        $script:CapturePath = $script:CaptureFiles[$i].FullName
        $script:Capture     = Import-Capture -Path $script:CapturePath
        if ($script:Capture.PSObject.Properties['NewIdentity'] -and $script:Capture.NewIdentity) {
            $txtTarget.Text = "$($script:Capture.NewIdentity)"
        }
        $btnApply.Enabled = $true
        $lblStatus.Text = "Capture loaded: $($script:Capture.SourceDisplayName) ($($script:Capture.RecipientTypeDetails))."
    } catch {
        $btnApply.Enabled = $false
        $lblStatus.Text = "That capture file could not be read: $($_.Exception.Message)"
    }
})

$form.Add_Activated({ Update-CaptureList })

$btnFolder.Add_Click({ Open-ReportFolder -Path $CaptureDir })

$btnLoad.Add_Click({
    $src = "$($txtSource.Text)".Trim()
    if (-not $src) { $lblStatus.Text = 'Enter the source resource first.'; return }
    $btnLoad.Enabled = $false
    $prev = $form.Cursor; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Reading source'
        $script:Capture = Get-ResourceCapture -Identity (Get-TargetSmtp -Text $src) -Progress $progress
        Complete-ProgressReporter -Activity 'Reading source' -Bar $pbar

        $script:CapturePath = Save-Capture -Capture $script:Capture
        Update-CaptureList
        for ($i = 0; $i -lt @($script:CaptureFiles).Count; $i++) {
            if ($script:CaptureFiles[$i].FullName -eq $script:CapturePath) { $cboCapture.SelectedIndex = $i; break }
        }
        $c = $script:Capture

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("Source            : $($c.SourceDisplayName)  <$($c.SourceIdentity)>")
        [void]$sb.AppendLine("Type              : $($c.RecipientTypeDetails)")
        [void]$sb.AppendLine("Capacity          : $(if ($c.ResourceCapacity) { $c.ResourceCapacity } else { '(not set)' })")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Calendar processing: $($c.CalendarProcessing.PSObject.Properties.Name.Count) properties")
        [void]$sb.AppendLine("  AutomateProcessing = $($c.CalendarProcessing.AutomateProcessing)")
        [void]$sb.AppendLine("  AllBookInPolicy    = $($c.CalendarProcessing.AllBookInPolicy)")
        [void]$sb.AppendLine("  BookingWindowInDays= $($c.CalendarProcessing.BookingWindowInDays)")
        $del = @($c.CalendarProcessing.ResourceDelegates)
        [void]$sb.AppendLine("  ResourceDelegates  = $(if ($del.Count) { $del -join ', ' } else { '(none)' })")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Room metadata      : $($c.Place.PSObject.Properties.Name.Count) properties")
        [void]$sb.AppendLine("Calendar perms     : $(@($c.CalendarPermissions).Count) non-default entries")
        [void]$sb.AppendLine("Regional           : $(if ($c.Regional.TimeZone) { $c.Regional.TimeZone } else { '(not read)' })")
        [void]$sb.AppendLine("Room lists         : $(if (@($c.RoomLists).Count) { @($c.RoomLists) -join ', ' } else { '(none)' })")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Captured to:")
        [void]$sb.AppendLine("  $script:CapturePath")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('--- Fill in Step 2 and click Create Mailbox. ---')
        $txtOut.Text = $sb.ToString()

        $btnCreate.Enabled = $true
        $lblStatus.Text = "Source captured. $($c.RecipientTypeDetails), $(@($c.CalendarPermissions).Count) permission entries."
    } catch {
        $lblStatus.Text = 'Could not read the source.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Reading source resource'
    } finally {
        Complete-ProgressReporter -Activity 'Reading source'
        $form.Cursor = $prev; $btnLoad.Enabled = $true
    }
})

$btnCreate.Add_Click({
    if (-not $script:Capture) { return }
    $name  = "$($fields['Name'].Text)".Trim()
    $alias = "$($fields['Alias'].Text)".Trim()
    $upn   = "$($fields['Upn'].Text)".Trim()
    $disp  = "$($fields['Display'].Text)".Trim()
    if (-not $name -or -not $alias -or -not $upn) {
        $lblStatus.Text = 'Name, Alias and Email are all required.'; return
    }

    # GUARD: the only step here that creates a directory object.
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Create a new $($script:Capture.RecipientTypeDetails) in Exchange Online?`r`n`r`n" +
        "  Name  : $name`r`n  Alias : $alias`r`n  Email : $upn`r`n`r`n" +
        "Settings are applied in Step 3, which you can run straight away.",
        'Create resource mailbox', 'YesNo', 'Question', 'Button2')
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $btnCreate.Enabled = $false
    $prev = $form.Cursor; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Creating resource'
        $r = New-DuplicateResource -Capture $script:Capture -Name $name -Alias $alias `
                -Upn $upn -DisplayName $disp -Progress $progress
        Complete-ProgressReporter -Activity 'Creating resource' -Bar $pbar

        $newId = if ($r.PrimarySmtp) { $r.PrimarySmtp } else { $upn }
        $txtTarget.Text = $newId
        Add-Member -InputObject $script:Capture -NotePropertyName 'NewIdentity' -NotePropertyValue $newId -Force
        [void](Save-Capture -Capture $script:Capture -Path $script:CapturePath)
        $btnApply.Enabled = $true
        $txtOut.Text += [Environment]::NewLine +
            "=== Created ===" + [Environment]::NewLine +
            "  sAMAccountName : $($r.SamAccountName)" + [Environment]::NewLine +
            "  Address        : $newId" + [Environment]::NewLine
        foreach ($w in $r.Warnings) { $txtOut.Text += "  WARNING: $w" + [Environment]::NewLine }
        $txtOut.Text += [Environment]::NewLine +
            'Created in Exchange Online. Click "Apply Settings" now -- no sync wait.' + [Environment]::NewLine +
            'If anything fails you can close this window and retry later: Step 3' + [Environment]::NewLine +
            'reads the capture file, not this session.' + [Environment]::NewLine

        $btnApply.Enabled = $true
        $lblStatus.Text = 'Created in Exchange Online. Click Apply Settings.'
        Write-CompletedLog -Task "Created $($script:Capture.RecipientTypeDetails) $newId from $($script:Capture.SourceIdentity)"
    } catch {
        $lblStatus.Text = 'Create failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Creating resource mailbox'
        # $upn, not $newId: $newId is assigned inside the try, so a throw from
        # the create cmdlet itself -- the failure most worth recording -- would
        # log a blank target.
        Write-FailedLog -Action 'Create' -Target $upn `
            -Detail "from $($script:Capture.SourceIdentity)" -ErrorRecord $_
        $btnCreate.Enabled = $true
    } finally {
        Complete-ProgressReporter -Activity 'Creating resource'
        $form.Cursor = $prev
    }
})

$btnApply.Add_Click({
    $target = "$($txtTarget.Text)".Trim()
    if (-not $script:Capture) { $lblStatus.Text = 'Pick a capture file first.'; return }
    if (-not $target)         { $lblStatus.Text = 'Enter the new resource address.'; return }
    $btnApply.Enabled = $false
    $prev = $form.Cursor; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Applying settings'
        $r = Set-DuplicateResourceSettings -Capture $script:Capture -Identity $target -Progress $progress
        Complete-ProgressReporter -Activity 'Applying settings' -Bar $pbar

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('=== Applied ===')
        foreach ($a in $r.Applied) { [void]$sb.AppendLine("  OK    $a") }
        if ($r.Skipped.Count) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('=== Not applicable ===')
            foreach ($s in $r.Skipped) { [void]$sb.AppendLine("  --    $s") }
        }
        if ($r.Failed.Count) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('=== Failed ===')
            foreach ($f in $r.Failed) { [void]$sb.AppendLine("  FAIL  $f") }
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Fix the cause and click Apply Settings again -- every step is')
            [void]$sb.AppendLine('  idempotent, so re-running will not duplicate anything.')
        }
        $txtOut.Text = $sb.ToString()

        $note = if ($r.Failed.Count) { ' -- PARTIAL' } else { '' }
        $lblStatus.Text = "Applied $($r.Applied.Count), failed $($r.Failed.Count)."
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($note) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "Configured $target from $($script:Capture.SourceIdentity) -- $($r.Applied.Count) applied$note" -Result $logResult
    } catch {
        $lblStatus.Text = 'Apply failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Applying resource settings'
        Write-FailedLog -Action 'Configure' -Target $target `
            -Detail "from $($script:Capture.SourceIdentity)" -ErrorRecord $_
    } finally {
        Complete-ProgressReporter -Activity 'Applying settings'
        $form.Cursor = $prev; $btnApply.Enabled = $true
    }
})

Update-CaptureList

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'ResourceDuplicator' | Out-Null
[void]$form.ShowDialog()
#endregion
