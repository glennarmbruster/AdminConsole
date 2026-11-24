#Requires -Version 5.1
<#
    ExOSupport.Common.ps1  --  everything ExOSupport needs that is not a window.

    A NARROW, CLOUD-ONLY CONSOLE. It finds cloud-managed shared mailboxes and
    distribution groups and lets support staff add and remove people. Nothing
    on-premises, no Active Directory, no Graph. One module:
    ExchangeOnlineManagement.

    DELIBERATELY NOT BUILT ON THE ADMIN CONSOLE'S Common.ps1. That file is
    nineteen thousand lines of hybrid machinery this tool must never call; a
    support operator's copy should contain only what it uses, so it can be
    read, reviewed and handed over in an afternoon.

    WHO SEES WHAT. The operator signs in to Exchange Online as themselves. That
    sign-in -- not anything typed into this tool -- is their identity. Their
    membership of ONE authorization group (Scope.Group in the config) decides
    access: a member sees and edits every object whose custom attribute
    (CustomAttribute10 by default) equals Scope.Value; a non-member sees and
    edits nothing. (Changed 2026-10-21 from per-team groups to one group.)

    THIS FILTER IS FOR CONVENIENCE. Enforcement belongs in Exchange RBAC
    (management scopes on the same attribute) -- see Setup-ScopedRbac.ps1. The
    console shows people what they are responsible for; Exchange refuses what
    they are not allowed to change.

    Author   : Glenn Armbruster
    Created  : 2026-10-20
#>

$script:ExsRoot = $PSScriptRoot

# BUILD STAMP. Every ExOSupport file carries one; ExOSupport.ps1 refuses to start
# unless all of them match its own, so a half-copied folder fails with a
# sentence instead of a missing-function error. Bump all four together.
if (-not $script:ExsParts) { $script:ExsParts = @{} }
$script:ExsParts['Common'] = '2026.10.21.7'

# ---------------------------------------------------------------------------
# DELIVERY SWITCH. While this is $false, Scope.Enabled in the config turns
# visibility filtering on and off. Set it to $true for the build handed to
# operators: filtering is then always on and the config switch is ignored, so
# an operator cannot widen what they see by editing a JSON file.
# ---------------------------------------------------------------------------
$script:ExsScopeFilterLocked = $false

# The only two kinds of object this console ever returns or edits.
$script:ExsTypeNames = @{
    'SharedMailbox'                  = 'Shared mailbox'
    'MailUniversalDistributionGroup' = 'Distribution group'
    'MailUniversalSecurityGroup'     = 'Mail-enabled security group'
}

$script:Tenant   = $null      # the active tenant block from the config
$script:Operator = $null      # who is signed in, and whether they are authorized
$script:OperatorEmail = ''    # the operator's own address, typed at launch
$script:ResultCap = 200

#region Config and tenant -------------------------------------------------------

function Get-ExsConfigPath {
    return (Join-Path (Join-Path $script:ExsRoot 'Config') 'ExOSupport.json')
}

function Read-ExsConfig {
    $p = Get-ExsConfigPath
    if (-not (Test-Path -LiteralPath $p)) { throw "The config file was not found: $p" }
    $raw = Get-Content -LiteralPath $p -Raw -Encoding UTF8
    try { return ($raw | ConvertFrom-Json -ErrorAction Stop) }
    catch { throw "The config file is not valid JSON ($p): $($_.Exception.Message)" }
}

function Show-ExsTenantPicker {
    <#
        Shown only when the config lists more than one tenant and ActiveTenant
        is blank. Numbered, so the list reads the way the config describes it:
        1) <first tenant>, 2) ...
    #>
    param([object[]]$Tenants)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'ExO Support -- choose a tenant'
    $f.ClientSize = New-Object System.Drawing.Size(360, 230)
    $f.StartPosition = 'CenterScreen'
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Which tenant?'; $l.Location = New-Object System.Drawing.Point(12, 12); $l.AutoSize = $true
    $f.Controls.Add($l)
    $lb = New-Object System.Windows.Forms.ListBox
    $lb.Location = New-Object System.Drawing.Point(12, 36); $lb.Size = New-Object System.Drawing.Size(336, 140)
    $i = 0
    foreach ($t in $Tenants) { $i++; [void]$lb.Items.Add("$i) $($t.DisplayName)") }
    $lb.SelectedIndex = 0
    $f.Controls.Add($lb)
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'OK'; $ok.Location = New-Object System.Drawing.Point(192, 188); $ok.Size = New-Object System.Drawing.Size(75, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($ok); $f.AcceptButton = $ok
    $cn = New-Object System.Windows.Forms.Button
    $cn.Text = 'Cancel'; $cn.Location = New-Object System.Drawing.Point(273, 188); $cn.Size = New-Object System.Drawing.Size(75, 28)
    $cn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.Controls.Add($cn); $f.CancelButton = $cn
    $lb.Add_DoubleClick({ $f.DialogResult = [System.Windows.Forms.DialogResult]::OK; $f.Close() })
    if ($f.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or $lb.SelectedIndex -lt 0) { return $null }
    return $Tenants[$lb.SelectedIndex]
}

function Select-ExsTenant {
    <#
        ActiveTenant in the config names the tenant. Blank and only one tenant
        listed: that one. Blank and several listed: ask. Adding a tenant later
        is one more block in Tenants -- nothing in the scripts changes.
    #>
    param([Parameter(Mandatory)]$Config)
    $list = @($Config.Tenants)
    if ($list.Count -eq 0) { throw 'The config file lists no tenants.' }
    $want = "$($Config.ActiveTenant)".Trim()
    if ($want) {
        $t = @($list | Where-Object { "$($_.Id)" -ieq $want }) | Select-Object -First 1
        if ($t) { return $t }
        throw "ActiveTenant '$want' is not one of the tenants in the config file ($((@($list | ForEach-Object { "$($_.Id)" })) -join ', '))."
    }
    if ($list.Count -eq 1) { return $list[0] }
    return (Show-ExsTenantPicker -Tenants $list)
}

function Initialize-ExsTenant {
    <#
        Validates the chosen tenant block and normalises it into
        $script:Tenant. Every check here fails at launch with a sentence,
        rather than later with a cmdlet error nobody can map back to the file.
    #>
    param([Parameter(Mandatory)]$Tenant, $Config)
    foreach ($k in 'Id', 'DisplayName', 'PrimaryDomain') {
        if (-not "$($Tenant.$k)".Trim()) { throw "The tenant block is missing '$k'." }
    }
    $scope = $Tenant.Scope
    $attr = 'CustomAttribute10'
    if ($scope -and "$($scope.Attribute)".Trim()) { $attr = "$($scope.Attribute)".Trim() }
    if ($attr -notmatch '^CustomAttribute([1-9]|1[0-5])$') {
        throw "Scope.Attribute must be CustomAttribute1 to CustomAttribute15 (it is '$attr')."
    }
    # ONE AUTHORIZATION GROUP (2026-10-21): Scope.Group + Scope.Value. The
    # older Scope.Teams list is still read if present, so an old config keeps
    # working, but new configs use the single pair.
    $teams = @()
    if ($scope -and "$($scope.Group)".Trim()) {
        if (-not "$($scope.Value)".Trim()) { throw 'Scope.Group is set but Scope.Value is empty.' }
        $teams += [PSCustomObject]@{ Name = "$($scope.Group)".Trim(); Group = "$($scope.Group)".Trim(); Value = "$($scope.Value)".Trim() }
    }
    foreach ($tm in @($(if ($scope -and -not $teams.Count) { $scope.Teams }))) {
        if (-not $tm) { continue }
        $g = "$($tm.Group)".Trim(); $v = "$($tm.Value)".Trim()
        if (-not $g -or -not $v) { throw "Every Scope.Teams entry needs a Group and a Value." }
        $teams += [PSCustomObject]@{ Name = $(if ("$($tm.Name)".Trim()) { "$($tm.Name)".Trim() } else { $v }); Group = $g; Value = $v }
    }
    $pim = $Tenant.Pim
    $hours = 4
    if ($pim -and $pim.PSObject.Properties['DurationHours']) {
        $h = 0; if ([int]::TryParse("$($pim.DurationHours)", [ref]$h) -and $h -gt 0) { $hours = $h }
    }
    $url = 'https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadmigratedroles'
    if ($pim -and "$($pim.PortalUrl)".Trim()) { $url = "$($pim.PortalUrl)".Trim() }

    $script:Tenant = [PSCustomObject]@{
        Id            = "$($Tenant.Id)".Trim()
        DisplayName   = "$($Tenant.DisplayName)".Trim()
        PrimaryDomain = "$($Tenant.PrimaryDomain)".Trim().TrimStart('@')
        TenantId      = "$($Tenant.TenantId)".Trim()
        PimUrl        = $url
        PimHours      = $hours
        ScopeEnabled  = [bool]$(if ($scope) { $scope.Enabled } else { $false })
        Attribute     = $attr
        Teams         = @($teams)
    }
    if ($Config -and $Config.PSObject.Properties['ResultCap']) {
        $c = 0; if ([int]::TryParse("$($Config.ResultCap)", [ref]$c) -and $c -gt 0) { $script:ResultCap = $c }
    }
}

function Test-ExsFilteringOn {
    return ($script:ExsScopeFilterLocked -or [bool]$script:Tenant.ScopeEnabled)
}
#endregion

#region Logging -----------------------------------------------------------------

function Write-ExsLog {
    <#
        One CSV line per change, in Logs\ beside the scripts, one file a month.
        LOCAL FOR NOW -- point Logs at a share later and every operator's
        changes land in one place. Never throws: a log that cannot be written
        must not undo a change that already happened.
    #>
    param([string]$Action, [string]$Target, [string]$Detail, [ValidateSet('Success', 'Failed', 'Skipped', 'Info')][string]$Result = 'Success')
    try {
        $dir = Join-Path $script:ExsRoot 'Logs'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null }
        # v2 in the name: the columns changed (Operator is the operator's own
        # address, AdminAccount the Exchange sign-in). Export-Csv -Append will
        # not add columns to an existing file, so a new schema gets a new file.
        $file = Join-Path $dir ("ExOSupport-v2-{0}.csv" -f (Get-Date -Format 'yyyy-MM'))
        $row = [PSCustomObject]@{
            WhenLocal   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            WhenUtc     = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
            Operator    = "$script:OperatorEmail"
            AdminAccount = $(if ($script:Operator) { "$($script:Operator.Upn)" } else { "$((Get-ExsConnection).Upn)" })
            WindowsUser = "$env:USERDOMAIN\$env:USERNAME"
            Machine     = "$env:COMPUTERNAME"
            Tenant      = $(if ($script:Tenant) { "$($script:Tenant.Id)" } else { '' })
            Action      = $Action
            Target      = $Target
            Result      = $Result
            Detail      = $Detail
        }
        $row | Export-Csv -LiteralPath $file -NoTypeInformation -Append -Encoding UTF8 -ErrorAction Stop
    } catch { }
}
#endregion

#region Exchange Online connection ----------------------------------------------

function Get-ExsConnection {
    <#
        The live Exchange Online session in THIS process, read locally from
        Get-ConnectionInformation -- no network call. A session started in
        another PowerShell window does not count; Connect-ExchangeOnline is
        per-process.
    #>
    $r = [PSCustomObject]@{ Connected = $false; Upn = ''; TenantId = ''; Detail = 'Not connected.' }
    if (-not (Get-Module -Name 'ExchangeOnlineManagement')) { return $r }
    $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
    if (-not $gci) { $r.Detail = 'This ExchangeOnlineManagement version is too old (no Get-ConnectionInformation). Update the module.'; return $r }
    $conns = @()
    try { $conns = @(& $gci -ErrorAction Stop) } catch { $r.Detail = "Could not read the session: $($_.Exception.Message)"; return $r }
    $ok = @($conns | Where-Object {
            (-not $_.PSObject.Properties['State'] -or "$($_.State)" -eq 'Connected') -and
            (-not $_.PSObject.Properties['TokenStatus'] -or "$($_.TokenStatus)" -ne 'Expired') }) | Select-Object -First 1
    if (-not $ok) {
        $r.Detail = $(if ($conns.Count) { 'The session has expired. Press Connect ExO again.' } else { 'Not connected.' })
        return $r
    }
    $r.Connected = $true
    $r.Upn       = "$($ok.UserPrincipalName)"
    $r.TenantId  = "$($ok.TenantID)"
    $r.Detail    = "Connected as $($r.Upn)"
    return $r
}

function Connect-ExsExchangeOnline {
    <#
        Sign THIS process in to Exchange Online. Returns Ok + Detail; never
        throws.

        NEVER CALL THIS FROM INSIDE AN OPEN WINDOW. The console's Connect
        button only records the request and closes the window; ExOSupport.ps1
        calls this after ShowDialog() has returned. A modal window owns this
        thread's message loop, and the sign-in prompt -- a window on the same
        thread -- then either never appears or never accepts input. The button
        just spins. That is the Admin Console's lesson (ARCHITECTURE 14d, the
        PendingConnect pattern) and this console follows it exactly.

        Three more guards from the same source:
          * the WinForms synchronisation context a closed dialog leaves behind
            is cleared, or the prompt is posted to a pump that no longer runs
          * the thread must be STA, or no sign-in window can be created
          * dead (expired/broken) sessions are cleared first, so the new
            sign-in is not shadowed by an old token
        and one retry with a clean context if the first attempt fails on the
        apartment error.

        ORDER MATTERS: activate PIM first, connect second.
    #>
    $res = [PSCustomObject]@{ Ok = $false; Detail = '' }

    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
        $res.Detail = 'This PowerShell window is not STA, so no sign-in window can open. Start ExO Support with ExOSupport.cmd.'
        return $res
    }

    if (-not (Get-Module -Name 'ExchangeOnlineManagement')) {
        $avail = @(Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' | Sort-Object Version -Descending) | Select-Object -First 1
        if (-not $avail) {
            $res.Detail = "The ExchangeOnlineManagement module is not installed for this account.`r`n`r`nIn a PowerShell window run:`r`n    Install-Module ExchangeOnlineManagement -Scope CurrentUser`r`n`r`nthen start ExO Support again."
            return $res
        }
        Write-Host "Loading ExchangeOnlineManagement $($avail.Version)..." -ForegroundColor DarkGray
        try { Import-Module -Name $avail.Path -ErrorAction Stop }
        catch { $res.Detail = "The ExchangeOnlineManagement module would not load: $($_.Exception.Message)"; return $res }
    }

    # Clear every existing session -- live or dead. Connect is also Reconnect
    # (after a PIM activation the old token lacks the role), and a dead
    # session left beside a new one can still be picked up by cmdlets.
    try {
        $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
        if ($gci -and @(& $gci -ErrorAction SilentlyContinue).Count -gt 0) {
            Write-Host 'Clearing the previous Exchange Online session...' -ForegroundColor DarkGray
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue 2>$null | Out-Null
        }
    } catch { }

    $cx = Get-Command -Name 'Connect-ExchangeOnline' -ErrorAction SilentlyContinue
    if (-not $cx) { $res.Detail = 'Connect-ExchangeOnline is not available from the loaded module.'; return $res }
    $p = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    # Browser sign-in rather than the Windows broker -- the same choice the
    # Admin Console makes.
    if ($cx.Parameters.ContainsKey('DisableWAM')) { $p['DisableWAM'] = $true }

    [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
    Write-Host ''
    Write-Host 'A sign-in window will open. Complete the sign-in and MFA there.' -ForegroundColor Yellow
    Write-Host '(If you do not see it, check behind this window and the taskbar.)' -ForegroundColor DarkGray
    $msg = $null
    try {
        & $cx @p 2>$null | Out-Null
    } catch {
        $msg = "$($_.Exception.Message)"
        if ($msg -match 'single-threaded apartment|8856f961') {
            Write-Host 'Retrying once with a clean thread context...' -ForegroundColor Yellow
            [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
            try { & $cx @p 2>$null | Out-Null; $msg = $null } catch { $msg = "$($_.Exception.Message)" }
        }
    }
    if ($msg) { $res.Detail = "Sign-in did not complete: $msg"; return $res }

    $c = Get-ExsConnection
    if (-not $c.Connected) { $res.Detail = "Sign-in finished but no session is present: $($c.Detail)"; return $res }

    # WRONG-TENANT GUARD, when the config names a tenant id.
    if ($script:Tenant.TenantId -and $c.TenantId -and ($c.TenantId -ine $script:Tenant.TenantId)) {
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue 2>$null | Out-Null } catch { }
        $res.Detail = "You signed in to a different tenant ($($c.TenantId)) than this console is set up for ($($script:Tenant.DisplayName)). You have been signed out; sign in again with your $($script:Tenant.DisplayName) account."
        return $res
    }
    $res.Ok = $true
    $res.Detail = $c.Detail
    return $res
}
#endregion

#region Operator email ----------------------------------------------------------
<#
    WHO THE OPERATOR IS, as opposed to which admin account they signed in with.

    Operators sign in to Exchange Online with an admin account (for example
    a-someone@tenant.onmicrosoft.com), which is not a member of anything. Their
    authorization group membership is on their NORMAL account. So the console asks for
    their own email address when it opens, remembers it per Windows user, and
    uses it for the authorization check. Both the address and the admin sign-in are
    written to every log line.
#>

function Get-ExsOperatorPath {
    $dir = Join-Path $env:LOCALAPPDATA 'ExOSupport'
    if (-not (Test-Path -LiteralPath $dir)) { try { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null } catch { } }
    return (Join-Path $dir 'operator.json')
}

function Get-ExsSavedOperatorEmail {
    try {
        $p = Get-ExsOperatorPath
        if (Test-Path -LiteralPath $p) { return "$((Get-Content -LiteralPath $p -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop).Email)".Trim() }
    } catch { }
    return ''
}

function Save-ExsOperatorEmail {
    param([string]$Email)
    try { [PSCustomObject]@{ Email = $Email } | ConvertTo-Json | Set-Content -LiteralPath (Get-ExsOperatorPath) -Encoding UTF8 -ErrorAction Stop } catch { }
}

function Show-ExsOperatorPrompt {
    <#
        Ask for the operator's own address. Pre-filled with the last one used.
        A bare name gets the tenant's primary domain. Returns the address, or
        '' if cancelled.
    #>
    param($Owner)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'ExO Support -- your email address'
    $f.ClientSize = New-Object System.Drawing.Size(460, 196)
    $f.StartPosition = $(if ($Owner) { 'CenterParent' } else { 'CenterScreen' })
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $f.TopMost = (-not $Owner)
    [void](New-ExsLabel $f 'Your email address' 14 14 430 20 -Bold)
    [void](New-ExsLabel $f ("Your normal mailbox -- not your admin account. It is checked against the authorization group, " +
        'and it is written to the change log alongside your admin sign-in.') 14 36 430 36 -Muted)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point(14, 80)
    $t.Size = New-Object System.Drawing.Size(430, 28)
    $t.Font = New-Object System.Drawing.Font('Segoe UI', 11)
    $t.Text = $(if ($script:OperatorEmail) { $script:OperatorEmail } else { Get-ExsSavedOperatorEmail })
    $f.Controls.Add($t)
    $err = New-ExsLabel $f '' 14 116 430 20
    $err.ForeColor = [System.Drawing.Color]::Firebrick
    $ok = New-ExsButton $f 'OK' 270 150 84 32
    $cn = New-ExsButton $f 'Cancel' 360 150 84 32
    $cn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.AcceptButton = $ok; $f.CancelButton = $cn
    $ok.Add_Click({
            $v = "$($t.Text)".Trim()
            if ($v -and $v -notmatch '@') { $v = "$v@$($script:Tenant.PrimaryDomain)" }
            if ($v -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { $err.Text = 'Enter an email address, e.g. first.last@' + $script:Tenant.PrimaryDomain; return }
            $t.Text = $v
            $f.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $f.Close()
        })
    $f.Add_Shown({ $t.SelectAll(); $t.Focus() })
    $r = $(if ($Owner) { $f.ShowDialog($Owner) } else { $f.ShowDialog() })
    if ($r -ne [System.Windows.Forms.DialogResult]::OK) { return '' }
    $email = "$($t.Text)".Trim()
    $script:OperatorEmail = $email
    Save-ExsOperatorEmail -Email $email
    return $email
}
#endregion

#region Operator authorization and visibility -------------------------------------------

function Test-ExsGroupContains {
    <#
        Is any of $Keys a member of $Group, directly or through nested groups?
        Keys are the operator's UPN, object id, GUID and address -- whichever
        of them a member record happens to carry. Five levels deep at most,
        and each group once, so a loop in nesting cannot hang the console.
    #>
    param([Parameter(Mandatory)][string]$Group, [Parameter(Mandatory)]$Keys, $Seen, [int]$Depth = 0)
    if ($null -eq $Seen) { $Seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase) }
    if ($Depth -gt 5 -or -not $Seen.Add($Group)) { return $false }
    $members = @(Get-DistributionGroupMember -Identity $Group -ResultSize Unlimited -ErrorAction Stop)
    foreach ($m in $members) {
        foreach ($v in @("$($m.ExternalDirectoryObjectId)", "$($m.Guid)", "$($m.PrimarySmtpAddress)", "$($m.WindowsLiveID)")) {
            if ($v.Trim() -and $Keys.Contains($v.Trim())) { return $true }
        }
    }
    foreach ($m in $members) {
        $rtd = "$($m.RecipientTypeDetails)"
        if ($rtd -match 'Group$' -and $rtd -ne 'GroupMailbox') {
            $id = $(if ("$($m.PrimarySmtpAddress)".Trim()) { "$($m.PrimarySmtpAddress)" } else { "$($m.Guid)" })
            if ($id -and (Test-ExsGroupContains -Group $id -Keys $Keys -Seen $Seen -Depth ($Depth + 1))) { return $true }
        }
    }
    return $false
}

function Resolve-ExsOperatorScope {
    <#
        Who is working, and which team values they may see. Fills
        $script:Operator; returns it.

          Upn    the Exchange Online sign-in (usually an admin account)
          Email  the operator's own address, typed at launch -- the account
                 the team groups are checked against

        The address is checked to exist in Exchange Online on every connect,
        whether or not filtering is on, so a typo is caught at once. A team
        group that cannot be read counts as NOT a member, with a note --
        failing closed, never open.
    #>
    $c = Get-ExsConnection
    $op = [PSCustomObject]@{ Upn = $c.Upn; Email = "$script:OperatorEmail"; Found = $false; Teams = @(); Values = @(); Notes = @() }
    if (-not $c.Connected) { $script:Operator = $op; return $op }

    $keys = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    if ($op.Email) {
        [void]$keys.Add($op.Email)
        try {
            $me = Get-Recipient -Identity $op.Email -ErrorAction Stop
            $op.Found = $true
            foreach ($v in @("$($me.ExternalDirectoryObjectId)", "$($me.Guid)", "$($me.PrimarySmtpAddress)", "$($me.WindowsLiveID)")) {
                if ($v.Trim()) { [void]$keys.Add($v.Trim()) }
            }
        } catch {
            $op.Notes += "No mailbox was found for '$($op.Email)'. Check the address (Change my email, top right) -- the authorization group is checked against it."
        }
    } else {
        $op.Notes += 'No email address was given, so authorization cannot be checked.'
    }

    # Filtering off: everyone sees everything, so the team groups are not read
    # at all -- placeholder or missing groups then cost nothing and raise no
    # warning on every connect.
    if (-not (Test-ExsFilteringOn)) { $script:Operator = $op; return $op }

    if ($op.Found) {
        foreach ($t in @($script:Tenant.Teams)) {
            try {
                if (Test-ExsGroupContains -Group $t.Group -Keys $keys) {
                    $op.Teams += $t.Name
                    if (@($op.Values) -notcontains $t.Value) { $op.Values += $t.Value }
                }
            } catch {
                $op.Notes += "Team group '$($t.Group)' could not be read ($($_.Exception.Message)); treated as not a member."
            }
        }
    }
    $script:Operator = $op
    return $op
}

function Get-ExsVisibilityText {
    if (-not (Test-ExsFilteringOn)) {
        return 'Visibility: ALL cloud shared mailboxes, distribution groups and mail-enabled security groups (filtering is off in the config).'
    }
    if (-not $script:Operator -or -not $script:Operator.Upn) { return 'Visibility: connect to see which objects you are responsible for.' }
    if (@($script:Operator.Values).Count -eq 0) {
        return "Visibility: NONE -- $($script:Operator.Email) is not in $((@($script:Tenant.Teams | ForEach-Object { $_.Group })) -join ' or '), so nothing is shown or editable."
    }
    return "Visibility: objects where $($script:Tenant.Attribute) is $((@($script:Operator.Values)) -join ' or ')   (authorized through $((@($script:Operator.Teams)) -join ', '))"
}

function Test-ExsRecipientAllowed {
    <#
        THE ONE GATE every lookup row and every write passes through:
          * one of the two supported types
          * cloud-managed (not directory-synced)
          * when filtering is on, the object's attribute is one of the
            operator's team values
    #>
    param([Parameter(Mandatory)]$Recipient)
    $rtd = "$($Recipient.RecipientTypeDetails)"
    if (-not $script:ExsTypeNames.ContainsKey($rtd)) {
        return [PSCustomObject]@{ Ok = $false; Why = "it is a $rtd; this console handles shared mailboxes, distribution groups and mail-enabled security groups only" }
    }
    if ("$($Recipient.IsDirSynced)" -eq 'True') {
        return [PSCustomObject]@{ Ok = $false; Why = 'it is still synced from on-premises, so it cannot be changed in the cloud' }
    }
    if (Test-ExsFilteringOn) {
        $val = "$($Recipient.($script:Tenant.Attribute))".Trim()
        $mine = @($(if ($script:Operator) { $script:Operator.Values }))
        if ($mine.Count -eq 0) { return [PSCustomObject]@{ Ok = $false; Why = 'you are not in the authorization group' } }
        if (-not $val -or ($mine -notcontains $val)) {
            return [PSCustomObject]@{ Ok = $false; Why = "its $($script:Tenant.Attribute) is not one this console manages" }
        }
    }
    return [PSCustomObject]@{ Ok = $true; Why = '' }
}
#endregion

#region Lookup ------------------------------------------------------------------

function ConvertTo-ExsOpath { param([string]$Value) return ("$Value" -replace "'", "''") }

function Find-ExsRecipient {
    <#
        Name, address or part of either. No asterisk: matches anywhere. With
        an asterisk: used as typed ('sales*' is starts-with). 'name@' is
        completed with the tenant's primary domain. A lone '*' is EVERYTHING
        the operator may see, with no result cap.

        Synced objects are excluded IN THE QUERY (IsDirSynced -eq $false, a
        documented filterable property), not only afterwards -- otherwise a
        tenant with many synced objects fills the result cap with rows that
        are then hidden, and cloud objects never appear.

        Returns Rows, Note, Capped. Synced objects and other teams' objects
        are counted in Note but never returned.
    #>
    param([Parameter(Mandatory)][string]$Text)
    $out = [PSCustomObject]@{ Rows = @(); Note = ''; Capped = $false }
    $t = "$Text".Trim()
    if ($t -match '@$') { $t = $t + $script:Tenant.PrimaryDomain }
    $all = ($t -match '^\*+$')
    $types = "(RecipientTypeDetails -eq 'SharedMailbox') -or (RecipientTypeDetails -eq 'MailUniversalDistributionGroup') -or (RecipientTypeDetails -eq 'MailUniversalSecurityGroup')"
    $filter = "(($types) -and (IsDirSynced -eq `$false))"
    if (-not $all) {
        $pat = $(if ($t -match '\*') { $t } else { "*$t*" })
        $v = ConvertTo-ExsOpath $pat
        $filter = "$filter -and ((Name -like '$v') -or (DisplayName -like '$v') -or (Alias -like '$v') -or (EmailAddresses -like '$v'))"
    }

    if (Test-ExsFilteringOn) {
        $vals = @($(if ($script:Operator) { $script:Operator.Values }))
        if ($vals.Count -eq 0) { $out.Note = 'You are not in the authorization group, so nothing is shown.'; return $out }
        $attr = $script:Tenant.Attribute
        $sc = (@($vals | ForEach-Object { "($attr -eq '$(ConvertTo-ExsOpath $_)')" })) -join ' -or '
        $filter = "$filter -and ($sc)"
    }

    if ($all) {
        $found = @(Get-Recipient -Filter $filter -ResultSize Unlimited -ErrorAction Stop)
    } else {
        $found = @(Get-Recipient -Filter $filter -ResultSize ($script:ResultCap + 1) -ErrorAction Stop)
        if ($found.Count -gt $script:ResultCap) { $out.Capped = $true; $found = @($found | Select-Object -First $script:ResultCap) }
    }

    $rows = @(); $synced = 0; $other = 0
    foreach ($r in $found) {
        $ok = Test-ExsRecipientAllowed -Recipient $r
        if (-not $ok.Ok) {
            if ("$($r.IsDirSynced)" -eq 'True') { $synced++ } else { $other++ }
            continue
        }
        $rows += [PSCustomObject]@{
            DisplayName = "$($r.DisplayName)"
            Address     = "$($r.PrimarySmtpAddress)"
            Type        = $script:ExsTypeNames["$($r.RecipientTypeDetails)"]
            Kind        = $(if ("$($r.RecipientTypeDetails)" -eq 'SharedMailbox') { 'Mailbox' } else { 'Group' })
            Team        = "$($r.($script:Tenant.Attribute))"
        }
    }
    $notes = @()
    if ($synced) { $notes += "$synced still synced from on-premises (hidden)" }
    if ($other)  { $notes += "$other not shown" }
    if ($out.Capped) { $notes += "first $($script:ResultCap) only -- narrow the search, or search * for everything" }
    $out.Rows = @($rows | Sort-Object DisplayName)
    $out.Note = ($notes -join '; ')
    return $out
}

function Get-ExsLiveRecipient {
    <#
        Re-read an object and re-check the gate. Called when an editor opens
        and again before every write, so a row found a minute ago that has
        since been re-tagged to another team cannot be changed.
    #>
    param([Parameter(Mandatory)][string]$Identity)
    $r = Get-Recipient -Identity $Identity -ErrorAction Stop
    $ok = Test-ExsRecipientAllowed -Recipient $r
    if (-not $ok.Ok) { throw "'$($r.DisplayName)' cannot be edited here: $($ok.Why)." }
    return $r
}

function Resolve-ExsPerson {
    <#
        One line an operator typed -> one recipient, or a reason. A bare alias
        (no @, no space) gets the tenant's primary domain.
    #>
    param([Parameter(Mandatory)][string]$Text)
    $t = "$Text".Trim().Trim(';', ',')
    $res = [PSCustomObject]@{ Ok = $false; Input = $t; Smtp = ''; Display = ''; Type = ''; Reason = '' }
    if (-not $t) { $res.Reason = 'blank'; return $res }
    $id = $(if ($t -notmatch '@' -and $t -notmatch '\s') { "$t@$($script:Tenant.PrimaryDomain)" } else { $t })
    try {
        $r = @(Get-Recipient -Identity $id -ErrorAction Stop)
        if ($r.Count -ne 1) { $res.Reason = "$($r.Count) matches -- use the full address"; return $res }
        $res.Ok = $true
        $res.Smtp = "$($r[0].PrimarySmtpAddress)"
        $res.Display = "$($r[0].DisplayName)"
        $res.Type = "$($r[0].RecipientTypeDetails)"
    } catch {
        $m = "$($_.Exception.Message)"
        $res.Reason = $(if ($m -match '(?i)matches multiple') { 'more than one match -- use the full address' }
                        elseif ($m -match "(?i)couldn't be found|could not be found|wasn't found") { 'not found' }
                        else { $m })
    }
    return $res
}
#endregion

#region Shared mailbox access ---------------------------------------------------

function Get-ExsMailboxAccess {
    <#
        Who has Full Access and who has Send As, one row per person. System
        entries (NT AUTHORITY\SELF), inherited and deny entries, and orphaned
        SIDs are left out -- they are not people and cannot be managed here.
    #>
    param([Parameter(Mandatory)][string]$Smtp)
    $map = [ordered]@{}
    foreach ($p in @(Get-MailboxPermission -Identity $Smtp -ErrorAction Stop)) {
        $u = "$($p.User)"
        if ($p.IsInherited -or $p.Deny -or $u -match '^(NT AUTHORITY\\|S-1-5-)') { continue }
        if (@($p.AccessRights | ForEach-Object { "$_" }) -notcontains 'FullAccess') { continue }
        $k = $u.ToLowerInvariant()
        if (-not $map.Contains($k)) { $map[$k] = [PSCustomObject]@{ User = $u; FullAccess = $false; SendAs = $false } }
        $map[$k].FullAccess = $true
    }
    foreach ($p in @(Get-RecipientPermission -Identity $Smtp -AccessRights SendAs -ErrorAction Stop)) {
        $u = "$($p.Trustee)"
        if ($u -match '^(NT AUTHORITY\\|S-1-5-)') { continue }
        $k = $u.ToLowerInvariant()
        if (-not $map.Contains($k)) { $map[$k] = [PSCustomObject]@{ User = $u; FullAccess = $false; SendAs = $false } }
        $map[$k].SendAs = $true
    }
    return @($map.Values | Sort-Object User)
}

function Add-ExsMailboxAccess {
    # Automapping is ALWAYS on in this console -- the mailbox opens in the
    # person's Outlook by itself, which is what support requests ask for.
    param([Parameter(Mandatory)][string]$Smtp, [Parameter(Mandatory)][string]$User, [switch]$SendAs)
    Add-MailboxPermission -Identity $Smtp -User $User -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$false -ErrorAction Stop | Out-Null
    if ($SendAs) {
        Add-RecipientPermission -Identity $Smtp -Trustee $User -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
    }
}

function Remove-ExsMailboxAccess {
    param([Parameter(Mandatory)][string]$Smtp, [Parameter(Mandatory)][string]$User, [bool]$FullAccess, [bool]$SendAs)
    if ($FullAccess) {
        Remove-MailboxPermission -Identity $Smtp -User $User -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
    }
    if ($SendAs) {
        Remove-RecipientPermission -Identity $Smtp -Trustee $User -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
    }
}
#endregion

#region Distribution group membership -------------------------------------------

function Get-ExsGroupMembers {
    param([Parameter(Mandatory)][string]$Group)
    return @(Get-DistributionGroupMember -Identity $Group -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
            [PSCustomObject]@{
                Name    = "$($_.DisplayName)"
                Address = "$($_.PrimarySmtpAddress)"
                Type    = "$($_.RecipientTypeDetails)"
                Id      = $(if ("$($_.PrimarySmtpAddress)".Trim()) { "$($_.PrimarySmtpAddress)" } else { "$($_.Guid)" })
            } } | Sort-Object Name)
}

function Set-ExsGroupMember {
    <#
        Add or remove one member.

        NOT AN OWNER IS THE NORMAL CASE HERE. Support staff are not in the
        group's ManagedBy, and Exchange then requires
        -BypassSecurityGroupManagerCheck, which in turn needs the 'Security
        Group Creation and Membership' role (Learn, Add-DistributionGroupMember).
        The switch is passed when the session offers it; if Exchange rejects
        the switch itself, the call is retried without it, which succeeds for
        a group the operator does own.
    #>
    param([Parameter(Mandatory)][ValidateSet('Add', 'Remove')][string]$Verb,
          [Parameter(Mandatory)][string]$Group, [Parameter(Mandatory)][string]$Member)
    $name = $(if ($Verb -eq 'Add') { 'Add-DistributionGroupMember' } else { 'Remove-DistributionGroupMember' })
    $cmd = Get-Command -Name $name -ErrorAction Stop
    $p = @{ Identity = $Group; Member = $Member; Confirm = $false; ErrorAction = 'Stop' }
    if ($cmd.Parameters.ContainsKey('BypassSecurityGroupManagerCheck')) { $p['BypassSecurityGroupManagerCheck'] = $true }
    try {
        & $cmd @p | Out-Null
    } catch {
        if ($p.ContainsKey('BypassSecurityGroupManagerCheck') -and "$($_.Exception.Message)" -match 'BypassSecurityGroupManagerCheck') {
            [void]$p.Remove('BypassSecurityGroupManagerCheck')
            & $cmd @p | Out-Null
        } else { throw }
    }
}
#endregion

#region PIM countdown -----------------------------------------------------------
<#
    A NOTE, NOT A LIVE READING -- the same design as the admin console. The
    operator activates in the portal, presses 'Mark activated', and this
    counts down the tenant's activation length from that moment. Stored per
    Windows user under LOCALAPPDATA, never beside the scripts.
#>

function Get-ExsPimPath {
    $dir = Join-Path $env:LOCALAPPDATA 'ExOSupport'
    if (-not (Test-Path -LiteralPath $dir)) { try { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null } catch { } }
    return (Join-Path $dir 'pim.json')
}

function Get-ExsPimActivated {
    $p = Get-ExsPimPath
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    try {
        $j = Get-Content -LiteralPath $p -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if (-not "$($j.Activated)".Trim()) { return $null }
        $dt = [DateTimeOffset]::MinValue
        if ([DateTimeOffset]::TryParse("$($j.Activated)", [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$dt)) { return $dt }
    } catch { }
    return $null
}

function Set-ExsPimActivated {
    param([switch]$Clear)
    $v = $(if ($Clear) { $null } else { [DateTimeOffset]::Now.ToString('o') })
    try { [PSCustomObject]@{ Activated = $v } | ConvertTo-Json | Set-Content -LiteralPath (Get-ExsPimPath) -Encoding UTF8 -ErrorAction Stop } catch { }
}

function Get-ExsPimStatus {
    <# Text and colour for the console's PIM line. #>
    $at = Get-ExsPimActivated
    if (-not $at) {
        return [PSCustomObject]@{ Text = 'PIM: not activated -- press Activate (PIM)'; Colour = [System.Drawing.Color]::DimGray }
    }
    $left = $at.AddHours($script:Tenant.PimHours) - [DateTimeOffset]::Now
    $since = $at.LocalDateTime.ToString('HH:mm')
    $until = $at.AddHours($script:Tenant.PimHours).LocalDateTime.ToString('HH:mm')
    if ($left.TotalSeconds -le 0) {
        return [PSCustomObject]@{ Text = "PIM expired at $until (activated $since) -- Activate again, then Reconnect ExO"; Colour = [System.Drawing.Color]::Firebrick }
    }
    $m = [int][Math]::Floor($left.TotalMinutes)
    $txt = $(if ($m -ge 60) { "{0}h {1:00}m" -f [int][Math]::Floor($m / 60), ($m % 60) } else { "{0}m" -f $m })
    $col = $(if ($left.TotalMinutes -le 30) { [System.Drawing.Color]::DarkOrange } else { [System.Drawing.Color]::ForestGreen })
    return [PSCustomObject]@{ Text = "PIM activated $since   --   $txt left (until $until)"; Colour = $col }
}
#endregion

#region Small UI helpers --------------------------------------------------------

function Set-ExsConsoleWindow {
    <#
        Minimise or restore the PowerShell window this console runs in.

        WHY NOT 'start /min' OR -WindowStyle Minimized: those set the process's
        start-up show state, which Windows applies to its FIRST top-level
        window -- the console FORM, not the black window -- so the console
        itself opened minimised (Admin Console, Dashboard notes, 2026-08-20).

        So the black window is minimised from inside, AFTER the form is up,
        with SW_SHOWMINNOACTIVE (7): minimise WITHOUT changing which window is
        active, so the form keeps the foreground. Before a sign-in it is
        restored (SW_RESTORE, 9), because the sign-in runs there with the form
        closed and needs a visible window to hand its prompt the foreground --
        the sign-in path stays exactly as it was proven.

        Never throws: a window that will not minimise is cosmetic.
    #>
    param([ValidateSet('Minimize', 'Restore')][string]$State)
    try {
        if (-not ('ExsNative.Win32' -as [type])) {
            Add-Type -Namespace 'ExsNative' -Name 'Win32' -ErrorAction Stop -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
'@
        }
        $h = [ExsNative.Win32]::GetConsoleWindow()
        if ($h -eq [System.IntPtr]::Zero) { return }
        [void][ExsNative.Win32]::ShowWindow($h, $(if ($State -eq 'Minimize') { 7 } else { 9 }))
    } catch { }
}

function New-ExsLabel {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 300, [int]$H = 20, [switch]$Bold, [switch]$Muted)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, $H)
    if ($Bold) { $l.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold) }
    if ($Muted) { $l.ForeColor = [System.Drawing.Color]::DimGray }
    $Parent.Controls.Add($l)
    return $l
}

function New-ExsButton {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 130, [int]$H = 30)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($W, $H)
    $Parent.Controls.Add($b)
    return $b
}

function New-ExsGrid {
    param($Parent, [object[]]$Columns, [int]$X, [int]$Y, [int]$W, [int]$H)
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Location = New-Object System.Drawing.Point($X, $Y)
    $g.Size = New-Object System.Drawing.Size($W, $H)
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.SelectionMode = 'FullRowSelect'
    $g.MultiSelect = $true
    $g.RowHeadersVisible = $false
    $g.BackgroundColor = [System.Drawing.Color]::White
    foreach ($c in $Columns) {
        $i = $g.Columns.Add($c[0], $c[0])
        $g.Columns[$i].Width = $c[1]
        $g.Columns[$i].SortMode = 'Automatic'
    }
    $g.Columns[$g.Columns.Count - 1].AutoSizeMode = 'Fill'
    $Parent.Controls.Add($g)
    return $g
}

function New-ExsListBox {
    # One entry per line. AcceptsReturn so Enter makes a new line instead of
    # pressing the form's default button.
    param($Parent, [int]$X, [int]$Y, [int]$W, [int]$H)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Multiline = $true
    $t.AcceptsReturn = $true
    $t.ScrollBars = 'Vertical'
    $t.WordWrap = $false
    $t.Font = New-Object System.Drawing.Font('Consolas', 9)
    $t.Location = New-Object System.Drawing.Point($X, $Y)
    $t.Size = New-Object System.Drawing.Size($W, $H)
    $Parent.Controls.Add($t)
    return $t
}

function Get-ExsLines {
    param([string]$Text)
    # Lines and semicolons only. NOT commas: display names here are
    # 'Last, First', and splitting on the comma turns one person into two.
    return @("$Text" -split "[\r\n;]+" | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Select-Object -Unique)
}

function Show-ExsResultList {
    param([string]$Title, [string[]]$Lines)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title
    $f.ClientSize = New-Object System.Drawing.Size(720, 380)
    $f.StartPosition = 'CenterParent'
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Multiline = $true; $t.ReadOnly = $true; $t.ScrollBars = 'Both'; $t.WordWrap = $false
    $t.Font = New-Object System.Drawing.Font('Consolas', 9)
    $t.BackColor = [System.Drawing.Color]::White
    $t.Location = New-Object System.Drawing.Point(12, 12); $t.Size = New-Object System.Drawing.Size(696, 316)
    $t.Anchor = 'Top,Bottom,Left,Right'
    $t.Text = ($Lines -join "`r`n")
    $f.Controls.Add($t)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = 'Close'; $b.Location = New-Object System.Drawing.Point(618, 338); $b.Size = New-Object System.Drawing.Size(90, 30)
    $b.Anchor = 'Bottom,Right'
    $b.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($b); $f.AcceptButton = $b; $f.CancelButton = $b
    [void]$f.ShowDialog()
}

function Show-ExsMessage {
    param([string]$Text, [string]$Title = 'ExO Support', [string]$Icon = 'Information', [string]$Buttons = 'OK')
    return [System.Windows.Forms.MessageBox]::Show($Text, $Title, $Buttons, $Icon)
}
#endregion
