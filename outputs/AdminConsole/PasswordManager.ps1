#Requires -Version 5.1
# TOOLKIT-VARIANT: OnPremAD
#
# ONE WINDOW FOR A PASSWORD -- the three things a password call is actually
# about: is it expired, is the account locked (and on WHICH controller), and
# reset it.
#
# WHY THIS IS NOT A TAB ON THE PERSON EDITOR. The Person Editor's Account tab
# already says locked-or-not and when the password expires, and that is the
# right amount for a window about the whole person. The full job -- every
# domain controller asked one at a time, a typed-twice reset with the rules
# checked before anything is sent -- is a different pace of work and a
# credential operation, and it gets its own window with its own audit rows.
# The Person Editor's help still says it will not reset a password; it points
# here instead.
#
# THIS REPLACES A STANDALONE TOOL that did the same job on its own theme, its
# own log file and its own directory layer (System.DirectoryServices, no
# RSAT). The search grid from that tool is NOT here: finding a person is what
# Unified Lookup does, and this window is opened from a row there or from the
# console with one identity. Everything else it did is here, on Common.ps1 and
# the ActiveDirectory module, so it looks like the rest of the toolkit and
# logs where the rest of the toolkit logs.
#
# THREE FACTS ABOUT ACTIVE DIRECTORY THAT SHAPE THIS WINDOW:
#
#   1. badPwdCount and badPasswordTime ARE NOT REPLICATED, and lockoutTime
#      replicates urgently but not instantly. So 'is this account locked' is a
#      per-controller question, and 'why does it keep locking' is answered by
#      whichever controller recorded the bad passwords. The lockout grid asks
#      EVERY controller and says which one.
#   2. LOCKED MEANS THE COMPUTED FLAG, not a non-zero lockoutTime. A
#      lockoutTime whose lockout duration has elapsed is a stale number on an
#      account that is not locked; msDS-User-Account-Control-Computed bit 0x10
#      is the controller's own answer.
#   3. THE EXPIRY IS msDS-UserPasswordExpiryTimeComputed, the domain's own
#      answer with any fine-grained policy applied. 'pwdLastSet plus the
#      default maximum age' is wrong for anyone under a PSO.
#
# THE PASSWORD IS NEVER LOGGED, NEVER SHOWN IN A DIALOG, AND CLEARED FROM THE
# BOXES THE MOMENT THE RESET RETURNS. The audit row says a reset happened, on
# which controller, and which options were ticked. That is all it says.
<#
    PasswordManager.ps1
    Password expiry, per-controller lockout, and a checked reset for one
    account, on Common.ps1 and the ActiveDirectory module.  Author : Glenn Armbruster
    Created  : 2026-09-19
#>
#
# TOOLKIT-HELP-RUNS: Reads one account from every domain controller in its domain. Writes go to Active Directory: unlock, must-change-at-next-logon, and a password reset written to the PDC emulator. Every write asks first and is logged; the password itself never is.
# TOOLKIT-HELP-BEGIN
# Everything about one PASSWORD, in one window. Type an address, UPN or
# account name, press Load -- or arrive here from a row in Unified Lookup.
#
#   PASSWORD          The verdict in one line -- expired N days ago, expires
#                     in N days, must change at next logon, never expires --
#                     with the dates behind it and the policy that decided it
#                     (fine-grained if one applies, otherwise the domain
#                     default). Tick or untick 'must change at next logon'
#                     and Apply.
#
#   LOCKOUT           Press 'Check every DC'. Every domain controller in the
#                     account's domain is asked in turn -- the PDC emulator
#                     first and marked -- for the locked flag, the bad
#                     password count, when the last bad password arrived, and
#                     the last logon it saw. The controller with the bad
#                     password count is the one the stale credential is
#                     hitting; that is usually the whole answer. 'Unlock on
#                     all DCs' clears the lock everywhere at once rather than
#                     on one controller and waiting for replication.
#
#   RESET             Type the new password twice. The rules grid goes green
#                     as each rule is met -- length (the domain's minimum, or
#                     12, whichever is higher), a capital, a number, a special
#                     character, both entries match, and the Windows rule
#                     nobody remembers: not containing the account name or any
#                     three-plus-letter part of the display name. Reset stays
#                     grey until every row is green, so the controller's vague
#                     'does not meet the requirements' is caught here with a
#                     reason. The reset is written to the PDC emulator, with
#                     'must change at next logon' and 'unlock on all DCs' as
#                     ticks.
#
# WHAT IS LOGGED. That an unlock, a must-change or a reset happened, for whom,
# on which controller, with which ticks. The password is never written
# anywhere -- not the log, not a dialog, not the status line.
#
# WHAT IT NEEDS. The ActiveDirectory module (RSAT) and a reachable controller.
# A controller that does not answer on the LDAP port within two seconds is
# marked 'Unreachable' in the grid and skipped, so one dead controller does
# not hang the scan.
#
# WHAT IT WILL NOT DO. Search. Finding the person is Unified Lookup's job;
# this window takes one identity. It will not change the password policy, and
# it will not act on Exchange Online -- a synced account's password is
# mastered here and flows to the cloud through the sync.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

$script:Target    = $null     # the resolved account -- see Resolve-PasswordTarget
$script:Busy      = $false
$script:Closing   = $false
$script:Cancel    = $false
$script:AdReady   = $null
$script:Dcs       = @()       # the last scan's rows
$script:Policy    = $null     # the resultant password policy
$script:MinLength = 12        # the tool's own floor; a domain that demands more wins

# msDS-User-Account-Control-Computed bits. Only meaningful on the COMPUTED
# attribute -- on userAccountControl itself the lockout bit is never set.
$script:UAC_LOCKOUT     = 0x00000010
$script:UAC_PWD_EXPIRED = 0x00800000

#region Resolve -----------------------------------------------------------------
function Get-PasswordAdProperties {
    @('DisplayName', 'UserPrincipalName', 'SamAccountName', 'mail', 'proxyAddresses',
      'Enabled', 'LockedOut', 'PasswordLastSet', 'PasswordNeverExpires', 'PasswordExpired', 'CannotChangePassword',
      'pwdLastSet', 'lockoutTime', 'badPwdCount', 'badPasswordTime', 'lastLogon', 'LastLogonDate',
      'msDS-UserPasswordExpiryTimeComputed', 'msDS-User-Account-Control-Computed', 'userAccountControl',
      'DistinguishedName', 'ObjectGUID')
}

function Resolve-PasswordTarget {
    <#
        Find the account, in every configured directory in turn, the way the
        Person Editor does. The CONTROLLER that answered is kept: the verdict
        and the policy are read back from it, and the scan starts from its
        domain.

        NOTHING IS READ FROM EXCHANGE ONLINE. A password is mastered in Active
        Directory or this window has nothing to do.
    #>
    param([Parameter(Mandatory)][string]$Identity)
    $r = [PSCustomObject]@{
        Ok = $false; Smtp = ''; Display = ''; Upn = ''; Sam = ''
        Directory = ''; Server = ''; Dn = ''; DomainDns = ''; Pdc = ''
        Ad = $null; Reason = ''
    }
    $t = "$Identity".Trim()
    if (-not $t) { $r.Reason = 'Enter an address, a UPN or an account name.'; return $r }

    if ($null -eq $script:AdReady) { $script:AdReady = Test-AdAvailable }
    if (-not $script:AdReady.Available) {
        $r.Reason = "Active Directory is not available: $($script:AdReady.Detail)"
        return $r
    }

    $esc = ConvertTo-LdapFilterValue -Value $t
    $filter = "(&(objectCategory=person)(objectClass=user)(|(mail=$esc)(userPrincipalName=$esc)(sAMAccountName=$esc)(proxyAddresses=smtp:$esc)))"
    $errors = @()
    foreach ($d in @(Get-TenantDirectory)) {
        $srv = Get-TenantDirectoryServer -Id $d.Id
        if (-not $srv) { continue }
        try {
            $hits = @(Get-ADUser -LDAPFilter $filter -Server $srv -Properties (Get-PasswordAdProperties) -ErrorAction Stop)
        } catch { $errors += "$($d.Id): $($_.Exception.Message)"; continue }
        if ($hits.Count -eq 0) { continue }
        if ($hits.Count -gt 1) {
            $r.Reason = "'$t' matches $($hits.Count) accounts in $($d.Id). Use the UPN or the account name."
            return $r
        }
        $u = $hits[0]
        $r.Ad = $u; $r.Directory = "$($d.Id)"; $r.Server = $srv
        $r.Display = "$($u.DisplayName)"; $r.Upn = "$($u.UserPrincipalName)"; $r.Sam = "$($u.SamAccountName)"
        $r.Dn = "$($u.DistinguishedName)"
        $r.Smtp = "$($u.mail)"
        if (-not $r.Smtp) {
            $p = @($u.proxyAddresses | Where-Object { "$_" -cmatch '^SMTP:' } | Select-Object -First 1)
            if ($p.Count) { $r.Smtp = "$($p[0])".Substring(5) }
        }
        # THE DOMAIN AND ITS PDC EMULATOR, from the controller that answered.
        # The reset is written to the PDC on purpose: Set-ADAccountPassword
        # will not run against a read-only controller, and the PDC is the one
        # every other controller consults on a bad-password check, so the new
        # password works everywhere the moment it is set.
        try {
            $dom = Get-ADDomain -Server $srv -ErrorAction Stop
            $r.DomainDns = "$($dom.DNSRoot)"
            $r.Pdc       = "$($dom.PDCEmulator)"
        } catch { $r.DomainDns = ''; $r.Pdc = '' }
        if (-not $r.Pdc) { $r.Pdc = $srv }
        break
    }
    if (-not $r.Ad) {
        $r.Reason = $(if ($errors.Count -gt 0 -and $errors.Count -eq @(Get-TenantDirectory).Count) { "Active Directory could not be asked: $($errors[0])" }
                      else { "'$t' was not found in Active Directory." })
        return $r
    }
    if (-not $r.Display) { $r.Display = $r.Sam }
    $r.Ok = $true
    return $r
}

function Invoke-PasswordAd {
    # Every AD cmdlet goes through here, pinned to a named controller: the
    # one the account was found on unless the caller names another (the PDC
    # for a reset, each controller in turn for the scan).
    param([Parameter(Mandatory)][string]$Name, [hashtable]$Parameters = @{}, [string]$Server)
    $p = @{} + $Parameters
    $srv = $Server
    if (-not $srv -and $script:Target -and $script:Target.Server) { $srv = $script:Target.Server }
    if ($srv) { $p['Server'] = $srv }
    $p['ErrorAction'] = 'Stop'
    return (& $Name @p)
}

function Test-PasswordLoaded {
    if (-not $script:Target -or -not $script:Target.Ok) {
        Set-PwStatus -Text 'Load an account first.' -Bad
        return $false
    }
    return $true
}

function Write-PasswordAudit {
    param([string]$Action, [string]$Detail, [string]$Result = 'Success')
    try {
        $o = $script:Target
        Write-CompletedLog -Action $Action -Target (Get-LogTarget "$(if ($o) { $o.Smtp })" "$(if ($o) { $o.Upn })" "$(if ($o) { $o.Sam })") `
            -Result $Result -Detail $Detail -Environment 'OnPrem'
    } catch { }
}
#endregion

#region Verdict and policy ------------------------------------------------------
function ConvertFrom-AdFileTime {
    param($Value)
    try {
        $v = [int64]$Value
        if ($v -le 0 -or $v -ge [int64]::MaxValue) { return $null }
        return [datetime]::FromFileTime($v)
    } catch { return $null }
}

function Get-PwdLastSetRaw {
    # pwdLastSet as the raw number: 0 is 'must change at next logon', which
    # the AD module reports as a $null PasswordLastSet -- indistinguishable
    # from 'never set' without the raw value.
    param($User)
    try { return [int64]$User.pwdLastSet } catch { return 1 }
}

function Get-PasswordVerdict {
    <#
        The one sentence the caller rang up about, plus the severity that
        colours it. Reads the AD module's friendly properties where they are
        honest and the raw attributes where they are not: PasswordLastSet is
        $null when pwdLastSet is 0, which is the 'must change' state and not
        'never set'.
    #>
    param($User)
    $r = [PSCustomObject]@{ Status = 'Unknown'; ExpiresOn = $null; DaysLeft = $null; Severity = 'Warn' }
    if (-not $User) { return $r }
    $pwdLast = 0; try { $pwdLast = [int64]$User.pwdLastSet } catch { }
    $uacComp = 0; try { $uacComp = [int]$User.'msDS-User-Account-Control-Computed' } catch { }
    $expRaw  = -1; try { if ($null -ne $User.'msDS-UserPasswordExpiryTimeComputed') { $expRaw = [int64]$User.'msDS-UserPasswordExpiryTimeComputed' } } catch { }

    if ($pwdLast -eq 0) { $r.Status = 'Must change at next logon'; $r.Severity = 'Warn'; return $r }
    if ($User.PasswordNeverExpires) { $r.Status = 'Never expires'; $r.Severity = 'Warn'; return $r }
    if ((($uacComp -band $script:UAC_PWD_EXPIRED) -ne 0) -and $expRaw -le 0) { $r.Status = 'EXPIRED'; $r.Severity = 'Bad'; return $r }
    if ($expRaw -ge [int64]::MaxValue) { $r.Status = 'Never expires'; $r.Severity = 'Warn'; return $r }
    if ($expRaw -le 0) { $r.Status = 'Unknown (the controller returned no expiry value)'; $r.Severity = 'Warn'; return $r }

    $exp = ConvertFrom-AdFileTime $expRaw
    if (-not $exp) { $r.Status = 'Unknown (the controller returned no expiry value)'; return $r }
    $r.ExpiresOn = $exp
    $now = Get-Date
    $days = [int][Math]::Floor(($exp - $now).TotalDays)
    $r.DaysLeft = $days
    if ($exp -le $now) {
        $ago = [int][Math]::Floor(($now - $exp).TotalDays)
        $r.Status = $(if ($ago -le 0) { 'EXPIRED today' } else { "EXPIRED $ago day(s) ago" })
        $r.Severity = 'Bad'
    } elseif ($days -le 0) { $r.Status = 'Expires today'; $r.Severity = 'Warn' }
    elseif ($days -le 7)   { $r.Status = "Expires in $days day(s)"; $r.Severity = 'Warn' }
    else                   { $r.Status = "Expires in $days day(s)"; $r.Severity = 'Good' }
    return $r
}

function Get-PasswordPolicy {
    <#
        The policy that actually applies to THIS account: a fine-grained
        policy if one covers them, otherwise the domain default. Both cmdlets
        return the same shape (MinPasswordLength, MaxPasswordAge,
        LockoutThreshold, LockoutDuration, ComplexityEnabled), so the caller
        does not care which answered -- only the label does.
    #>
    param($Dn, $Server)
    $r = [PSCustomObject]@{ Source = ''; Name = ''; MinLength = 0; MaxAge = $null; LockoutThreshold = 0; LockoutDuration = $null; Complexity = $true }
    $pol = $null
    try { $pol = Get-ADUserResultantPasswordPolicy -Identity $Dn -Server $Server -ErrorAction Stop } catch { $pol = $null }
    if ($pol) { $r.Source = 'fine-grained'; $r.Name = "$($pol.Name)" }
    else {
        try { $pol = Get-ADDefaultDomainPasswordPolicy -Server $Server -ErrorAction Stop; $r.Source = 'domain default' } catch { return $r }
    }
    try { $r.MinLength        = [int]$pol.MinPasswordLength } catch { }
    try { $r.MaxAge           = $pol.MaxPasswordAge } catch { }
    try { $r.LockoutThreshold = [int]$pol.LockoutThreshold } catch { }
    try { $r.LockoutDuration  = $pol.LockoutDuration } catch { }
    try { $r.Complexity       = [bool]$pol.ComplexityEnabled } catch { }
    return $r
}

function Get-RequiredLength {
    $min = $script:MinLength
    if ($script:Policy -and $script:Policy.MinLength -gt $min) { $min = $script:Policy.MinLength }
    return $min
}
#endregion

#region Reset rules -------------------------------------------------------------
function Test-ContainsAccountName {
    <#
        The Windows complexity rule nobody remembers: the password may not
        contain the sAMAccountName, nor any three-plus character piece of the
        display name split on spaces, commas, dots, hyphens or underscores.
        Checking it here turns the controller's 'does not meet the
        requirements' into a reason before anything is sent.
    #>
    param([string]$Password)
    if (-not $Password -or -not $script:Target) { return $false }
    $p = $Password.ToLowerInvariant()
    $sam = "$($script:Target.Sam)"
    if ($sam -and $sam.Length -ge 3 -and $p.Contains($sam.ToLowerInvariant())) { return $true }
    foreach ($token in ("$($script:Target.Display)" -split '[\s,\.\-_#]+')) {
        if ($token -and $token.Length -ge 3 -and $p.Contains($token.ToLowerInvariant())) { return $true }
    }
    return $false
}

function Get-PasswordRuleRows {
    # One row per rule. Test-PasswordAcceptable reads the same rows, so the
    # grid and the Reset button cannot disagree.
    param([string]$Password, [string]$Confirm)
    $min   = Get-RequiredLength
    $typed = [bool]$Password
    $rows  = @()
    $rows += [PSCustomObject]@{ Requirement = "At least $min characters";               Met = ($Password.Length -ge $min); Detail = $(if ($typed) { "$($Password.Length) typed" } else { '' }) }
    $rows += [PSCustomObject]@{ Requirement = 'At least 1 capital letter (A-Z)';        Met = ($Password -cmatch '[A-Z]');   Detail = '' }
    $rows += [PSCustomObject]@{ Requirement = 'At least 1 number (0-9)';                Met = ($Password -match '[0-9]');    Detail = '' }
    $rows += [PSCustomObject]@{ Requirement = 'At least 1 special character';           Met = ($Password -match '[^a-zA-Z0-9]'); Detail = '! @ # $ % ^ & * and so on' }
    $rows += [PSCustomObject]@{ Requirement = 'Both entries match';                     Met = ($typed -and ($Password -ceq $Confirm)); Detail = '' }
    $rows += [PSCustomObject]@{ Requirement = 'Does not contain the account or name';   Met = (-not (Test-ContainsAccountName -Password $Password)); Detail = 'Windows complexity rule' }
    foreach ($x in $rows) {
        $x | Add-Member -NotePropertyName 'Status' -NotePropertyValue $(if (-not $typed) { '' } elseif ($x.Met) { 'OK' } else { 'Missing' })
    }
    return $rows
}

function Test-PasswordAcceptable {
    param([string]$Password, [string]$Confirm)
    if (-not $Password) { return $false }
    $rows = Get-PasswordRuleRows -Password $Password -Confirm $Confirm
    return (@($rows | Where-Object { -not $_.Met }).Count -eq 0)
}
#endregion

#region Domain controllers ------------------------------------------------------
function Test-DcReachable {
    # A two-second knock on the LDAP port before the AD module is asked. The
    # module has no timeout of its own, and a controller that is down takes
    # the better part of a minute to say so -- times every dead controller.
    param([string]$HostName, [int]$Milliseconds = 2000)
    $c = New-Object System.Net.Sockets.TcpClient
    try {
        $ar = $c.BeginConnect($HostName, 389, $null, $null)
        if (-not $ar.AsyncWaitHandle.WaitOne($Milliseconds, $false)) { return $false }
        $c.EndConnect($ar)
        return $true
    } catch { return $false }
    finally { try { $c.Close() } catch { } }
}

function Get-DomainControllerList {
    <#
        Every controller in the account's domain, PDC emulator first, then by
        site and name. Read-only controllers are listed too -- they answer
        the lockout question like any other -- but marked, because a reset
        cannot be written to one.
    #>
    param([string]$Server)
    $list = @()
    $dcs = @(Get-ADDomainController -Filter * -Server $Server -ErrorAction Stop)
    foreach ($dc in $dcs) {
        $roles = @($dc.OperationMasterRoles | ForEach-Object { "$_" })
        $list += [PSCustomObject]@{
            Name   = "$($dc.HostName)"
            Site   = "$($dc.Site)"
            IsPdc  = ($roles -contains 'PDCEmulator')
            IsRodc = [bool]$dc.IsReadOnly
        }
    }
    return @($list | Sort-Object -Property @{ Expression = 'IsPdc'; Descending = $true }, Site, Name)
}

function Get-AccountOnDc {
    # The unreplicated attributes, from ONE controller. Locked comes from the
    # computed flag -- see the header.
    param([string]$Dn, [string]$Server)
    $r = [PSCustomObject]@{ Ok = $false; Locked = $false; LockoutTime = $null; BadPwdCount = 0; BadPasswordTime = $null; PasswordLastSet = $null; LastLogon = $null; Error = '' }
    try {
        $u = Get-ADUser -Identity $Dn -Server $Server -Properties badPwdCount, badPasswordTime, lockoutTime, 'msDS-User-Account-Control-Computed', pwdLastSet, lastLogon -ErrorAction Stop
        $uacComp = 0; try { $uacComp = [int]$u.'msDS-User-Account-Control-Computed' } catch { }
        $r.Locked          = (($uacComp -band $script:UAC_LOCKOUT) -ne 0)
        $r.LockoutTime     = ConvertFrom-AdFileTime $u.lockoutTime
        $r.BadPwdCount     = $(try { [int]$u.badPwdCount } catch { 0 })
        $r.BadPasswordTime = ConvertFrom-AdFileTime $u.badPasswordTime
        $r.PasswordLastSet = ConvertFrom-AdFileTime $u.pwdLastSet
        $r.LastLogon       = ConvertFrom-AdFileTime $u.lastLogon
        $r.Ok = $true
    } catch { $r.Error = $_.Exception.Message }
    return $r
}
#endregion

#region Window ------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Password Manager'
$form.ClientSize    = New-Object System.Drawing.Size(1120, 760)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$y = 12
$lblWho          = New-Object System.Windows.Forms.Label
$lblWho.Text     = 'Account:'
$lblWho.Location = New-Object System.Drawing.Point(12, ($y + 3))
$lblWho.Size     = New-Object System.Drawing.Size(58, 18)
$form.Controls.Add($lblWho)

$txtWho          = New-Object System.Windows.Forms.TextBox
$txtWho.Location = New-Object System.Drawing.Point(72, $y)
$txtWho.Size     = New-Object System.Drawing.Size(330, 24)
$txtWho.Font     = New-Object System.Drawing.Font('Consolas', 9)
$txtWho.Text     = "$SmtpAddress".Trim()
$form.Controls.Add($txtWho)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load'
$btnLoad.Location = New-Object System.Drawing.Point(412, ($y - 2))
$btnLoad.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnLoad)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(952, ($y - 2))
$btnClose.Size     = New-Object System.Drawing.Size(90, 28)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)
$y += 34

$lblIdent           = New-Object System.Windows.Forms.Label
$lblIdent.Tag       = 'NoTheme'
$lblIdent.Location  = New-Object System.Drawing.Point(12, $y)
$lblIdent.Size      = New-Object System.Drawing.Size(1096, 18)
$lblIdent.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblIdent.ForeColor = (Get-ToolStatusColour -Kind Muted)
$lblIdent.Text      = 'No account loaded.'
$form.Controls.Add($lblIdent)
$y += 20

$lblSync           = New-Object System.Windows.Forms.Label
$lblSync.Tag       = 'NoTheme'
$lblSync.Location  = New-Object System.Drawing.Point(12, $y)
$lblSync.Size      = New-Object System.Drawing.Size(1096, 18)
$lblSync.ForeColor = (Get-ToolStatusColour -Kind Muted)
$lblSync.Text      = ''
$form.Controls.Add($lblSync)
$y += 24

# ---- left column: PASSWORD -----------------------------------------------------
$lx = 12; $ly = $y; $lw = 530
[void](New-TabLabel $form 'PASSWORD' $lx $ly -Bold)
$ly += 22
$pwVerdict = New-Object System.Windows.Forms.Label
$pwVerdict.Tag       = 'NoTheme'
$pwVerdict.Location  = New-Object System.Drawing.Point($lx, $ly)
$pwVerdict.Size      = New-Object System.Drawing.Size($lw, 22)
$pwVerdict.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Bold)
$pwVerdict.ForeColor = (Get-ToolStatusColour -Kind Muted)
$pwVerdict.Text      = ''
$form.Controls.Add($pwVerdict)
$ly += 28
$pwLabels = @('Last set', 'Expires', 'Account', 'Never expires', 'Policy', 'Lockout policy')
$pwVals = @{}
foreach ($n in $pwLabels) {
    [void](New-TabLabel $form ($n + ':') $lx $ly 110)
    $pwVals[$n] = New-TabValue $form ($lx + 114) $ly ($lw - 114)
    $ly += 24
}
$ly += 4
$chkMustChange          = New-Object System.Windows.Forms.CheckBox
$chkMustChange.Text     = 'Must change password at next logon'
$chkMustChange.Location = New-Object System.Drawing.Point($lx, $ly)
$chkMustChange.Size     = New-Object System.Drawing.Size(260, 22)
$chkMustChange.Enabled  = $false
$form.Controls.Add($chkMustChange)
$btnMustChange = New-TabButton $form 'Apply' ($lx + 270) ($ly - 3) 90
$btnMustChange.Enabled = $false
$ly += 36

# ---- left column: RESET --------------------------------------------------------
[void](New-TabLabel $form 'RESET PASSWORD' $lx $ly -Bold)
$ly += 22
[void](New-TabLabel $form 'New password:' $lx ($ly + 3) 110)
$txtPw1 = New-Object System.Windows.Forms.TextBox
$txtPw1.Location = New-Object System.Drawing.Point(($lx + 114), $ly)
$txtPw1.Size     = New-Object System.Drawing.Size(300, 24)
$txtPw1.Font     = New-Object System.Drawing.Font('Consolas', 10)
$txtPw1.UseSystemPasswordChar = $true
$txtPw1.Enabled  = $false
$form.Controls.Add($txtPw1)
$chkShow          = New-Object System.Windows.Forms.CheckBox
$chkShow.Text     = 'Show'
$chkShow.Location = New-Object System.Drawing.Point(($lx + 424), ($ly + 1))
$chkShow.Size     = New-Object System.Drawing.Size(70, 22)
$form.Controls.Add($chkShow)
$ly += 30
[void](New-TabLabel $form 'Type it again:' $lx ($ly + 3) 110)
$txtPw2 = New-Object System.Windows.Forms.TextBox
$txtPw2.Location = New-Object System.Drawing.Point(($lx + 114), $ly)
$txtPw2.Size     = New-Object System.Drawing.Size(300, 24)
$txtPw2.Font     = New-Object System.Drawing.Font('Consolas', 10)
$txtPw2.UseSystemPasswordChar = $true
$txtPw2.Enabled  = $false
$form.Controls.Add($txtPw2)
$ly += 32
$gridRules = New-TabGrid $form @(@('Requirement', 290), @('Status', 70), @('Detail', 150)) $lx $ly $lw 160
$gridRules.MultiSelect = $false
$gridRules.ColumnHeadersVisible = $true
$ly += 168
$chkResetMustChange          = New-Object System.Windows.Forms.CheckBox
$chkResetMustChange.Text     = 'Must change at next logon'
$chkResetMustChange.Location = New-Object System.Drawing.Point($lx, $ly)
$chkResetMustChange.Size     = New-Object System.Drawing.Size(200, 22)
$chkResetMustChange.Checked  = $true
$form.Controls.Add($chkResetMustChange)
$chkResetUnlock          = New-Object System.Windows.Forms.CheckBox
$chkResetUnlock.Text     = 'Unlock on all DCs'
$chkResetUnlock.Location = New-Object System.Drawing.Point(($lx + 210), $ly)
$chkResetUnlock.Size     = New-Object System.Drawing.Size(150, 22)
$chkResetUnlock.Checked  = $true
$form.Controls.Add($chkResetUnlock)
$btnReset = New-TabButton $form 'Reset password' ($lx + 400) ($ly - 3) 130
$btnReset.Enabled = $false
$ly += 34
$lblResetNote = New-TabLabel $form '' $lx $ly $lw -Muted
$lblResetNote.Size = New-Object System.Drawing.Size($lw, 34)

# ---- right column: LOCKOUT -----------------------------------------------------
$rx = 560; $ry = $y; $rw = 548
[void](New-TabLabel $form 'LOCKOUT ON EVERY DOMAIN CONTROLLER' $rx $ry -Bold)
$ry += 22
$lockSummary = New-Object System.Windows.Forms.Label
$lockSummary.Tag       = 'NoTheme'
$lockSummary.Location  = New-Object System.Drawing.Point($rx, $ry)
$lockSummary.Size      = New-Object System.Drawing.Size($rw, 22)
$lockSummary.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Bold)
$lockSummary.ForeColor = (Get-ToolStatusColour -Kind Muted)
$lockSummary.Text      = ''
$form.Controls.Add($lockSummary)
$ry += 28
$lockNote = New-TabLabel $form 'Press Check every DC. Bad-password counts are not replicated, so the controller that shows them is the one the stale credential is hitting.' $rx $ry $rw -Muted
$lockNote.Size = New-Object System.Drawing.Size($rw, 34)
$ry += 38
$btnScan      = New-TabButton $form 'Check every DC' $rx $ry 130
$btnCancel    = New-TabButton $form 'Stop' ($rx + 138) $ry 70
$btnUnlockAll = New-TabButton $form 'Unlock on all DCs' ($rx + 216) $ry 140
$btnScan.Enabled = $false; $btnCancel.Enabled = $false; $btnUnlockAll.Enabled = $false
$ry += 36
$gridDcs = New-TabGrid $form @(@('Controller', 170), @('Site', 80), @('Role', 50), @('Locked', 72), @('Bad', 40), @('Last bad password', 118), @('Last logon here', 118), @('Note', 220)) $rx $ry $rw 400
$ry += 408
$lockFoot = New-TabLabel $form '' $rx $ry $rw -Muted
$lockFoot.Size = New-Object System.Drawing.Size($rw, 34)

$bottom = [Math]::Max($ly + 40, $ry + 40)
$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(12, $bottom)
$lblStatus.Size      = New-Object System.Drawing.Size(1096, 18)
$lblStatus.ForeColor = (Get-ToolStatusColour -Kind Muted)
$lblStatus.Text      = 'Enter an address, UPN or account name and press Load.'
$form.Controls.Add($lblStatus)
$form.ClientSize = New-Object System.Drawing.Size(1120, ($lblStatus.Bottom + 12))

function Set-PwStatus {
    param([string]$Text, [switch]$Bad)
    $lblStatus.Text = $Text
    $lblStatus.ForeColor = $(if ($Bad) { (Get-ToolStatusColour -Kind Bad) } else { (Get-ToolStatusColour -Kind Muted) })
    $lblStatus.Refresh()
}

function Invoke-PwAction {
    # One busy flag for the window. A second action started while a scan is
    # mid-loop is how the grid gets rebuilt under the loop reading it.
    param([scriptblock]$Do, [string]$Context)
    if ($script:Busy) { return }
    $script:Busy = $true
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        & $Do
    } catch {
        if ($script:Closing) { return }
        Set-PwStatus -Text "$Context failed: $($_.Exception.Message)" -Bad
        $o = $script:Target
        try {
            Write-FailedLog -Action 'PasswordManagerAction' -Environment 'OnPrem' -ErrorRecord $_ -Detail "$Context failed" `
                -Target (Get-LogTarget "$(if ($o) { $o.Smtp })" "$(if ($o) { $o.Upn })" "$(if ($o) { $o.Sam })")
        } catch { }
        Show-HandlerError -ErrorRecord $_ -Context $Context
    } finally {
        if (-not $script:Closing) { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
        $script:Busy = $false
    }
}
#endregion

#region Paint -------------------------------------------------------------------
function Format-PwDate {
    param($Value, [string]$Empty = '')
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm') }
    return $Empty
}

function Format-PwSpan {
    param($Span)
    if (-not ($Span -is [timespan])) { return '' }
    if ($Span.Ticks -le 0) { return 'never' }
    if ($Span.TotalDays -ge 1) { return ('{0} day(s)' -f [int]$Span.TotalDays) }
    return ('{0} minute(s)' -f [int]$Span.TotalMinutes)
}

function Update-PasswordSection {
    $t = $script:Target
    foreach ($k in $pwVals.Keys) { $pwVals[$k].Text = '' }
    if (-not $t -or -not $t.Ok) {
        $pwVerdict.Text = ''
        $chkMustChange.Checked = $false; $chkMustChange.Enabled = $false; $btnMustChange.Enabled = $false
        return
    }
    $u = $t.Ad
    $v = Get-PasswordVerdict -User $u
    $pwVerdict.Text      = $v.Status
    $pwVerdict.ForeColor = (Get-ToolStatusColour -Kind $v.Severity)
    $mustNow = (Get-PwdLastSetRaw -User $u) -eq 0
    $pwVals['Last set'].Text = $(if ($u.PasswordLastSet) { Format-PwDate $u.PasswordLastSet } elseif ($mustNow) { '(must change at next logon)' } else { '(never)' })
    $pwVals['Expires'].Text  = $(if ($v.ExpiresOn) { Format-PwDate $v.ExpiresOn } elseif ($u.PasswordNeverExpires) { 'never' } else { '' })
    $acct = @()
    $acct += $(if ($u.Enabled) { 'enabled' } else { 'DISABLED' })
    $acct += $(if ($u.LockedOut) { 'LOCKED (on the controller that answered)' } else { 'not locked on the controller that answered' })
    if ($u.CannotChangePassword) { $acct += 'user cannot change password' }
    $pwVals['Account'].Text = ($acct -join '; ')
    $pwVals['Never expires'].Text = $(if ($u.PasswordNeverExpires) { 'YES -- PasswordNeverExpires is set' } else { 'no' })
    $p = $script:Policy
    if ($p -and $p.Source) {
        $pwVals['Policy'].Text = ('{0}{1}: minimum {2} characters, maximum age {3}, complexity {4}' -f $p.Source,
            $(if ($p.Name) { " ($($p.Name))" } else { '' }), $p.MinLength, (Format-PwSpan $p.MaxAge), $(if ($p.Complexity) { 'on' } else { 'off' }))
        $pwVals['Lockout policy'].Text = $(if ($p.LockoutThreshold -le 0) { 'accounts do not lock out' }
                                           else { ('{0} bad passwords locks the account for {1}' -f $p.LockoutThreshold, (Format-PwSpan $p.LockoutDuration)) })
    } else {
        $pwVals['Policy'].Text = '(policy could not be read)'
    }
    $chkMustChange.Checked = $mustNow
    $chkMustChange.Enabled = $true
    $btnMustChange.Enabled = $true
    $chkMustChange.Tag     = $mustNow
}

function Update-RulesGrid {
    $gridRules.Rows.Clear()
    $rows = Get-PasswordRuleRows -Password $txtPw1.Text -Confirm $txtPw2.Text
    $good = (Get-ToolStatusColour -Kind Good); $bad = (Get-ToolStatusColour -Kind Bad)
    foreach ($r in $rows) {
        $i = $gridRules.Rows.Add($r.Requirement, $r.Status, $r.Detail)
        if ($r.Status -eq 'OK')      { $gridRules.Rows[$i].Cells[1].Style.ForeColor = $good }
        if ($r.Status -eq 'Missing') { $gridRules.Rows[$i].Cells[1].Style.ForeColor = $bad }
    }
    $gridRules.ClearSelection()
    $btnReset.Enabled = ($script:Target -and $script:Target.Ok -and (Test-PasswordAcceptable -Password $txtPw1.Text -Confirm $txtPw2.Text))
}

function Update-LockoutGrid {
    $gridDcs.Rows.Clear()
    $bad = (Get-ToolStatusColour -Kind Bad); $warn = (Get-ToolStatusColour -Kind Warn); $muted = (Get-ToolStatusColour -Kind Muted)
    foreach ($d in $script:Dcs) {
        $role = @()
        if ($d.IsPdc)  { $role += 'PDC' }
        if ($d.IsRodc) { $role += 'RODC' }
        $i = $gridDcs.Rows.Add($d.Name, $d.Site, ($role -join ' '), $d.LockState, $(if ($d.Reachable) { "$($d.BadPwdCount)" } else { '' }),
                               (Format-PwDate $d.BadPasswordTime), (Format-PwDate $d.LastLogon), $d.Note)
        $row = $gridDcs.Rows[$i]
        if ($d.LockState -eq 'LOCKED')  { $row.Cells[3].Style.ForeColor = $bad }
        if ($d.BadPwdCount -gt 0)       { $row.Cells[4].Style.ForeColor = $warn }
        if (-not $d.Reachable)          { $row.DefaultCellStyle.ForeColor = $muted }
    }
    $gridDcs.ClearSelection()
}

function Update-LockoutSummary {
    $answered = @($script:Dcs | Where-Object { $_.Reachable })
    $locked   = @($answered | Where-Object { $_.LockState -eq 'LOCKED' })
    $silent   = @($script:Dcs | Where-Object { -not $_.Reachable })
    if ($script:Dcs.Count -eq 0) { $lockSummary.Text = ''; $lockFoot.Text = ''; $btnUnlockAll.Enabled = $false; return }
    if ($answered.Count -eq 0) {
        $lockSummary.Text = 'No controller answered.'; $lockSummary.ForeColor = (Get-ToolStatusColour -Kind Bad)
    } elseif ($locked.Count -gt 0) {
        $lockSummary.Text = ('LOCKED on {0} of {1} controller(s) that answered.' -f $locked.Count, $answered.Count)
        $lockSummary.ForeColor = (Get-ToolStatusColour -Kind Bad)
    } else {
        $lockSummary.Text = ('Not locked on any of the {0} controller(s) that answered.' -f $answered.Count)
        $lockSummary.ForeColor = (Get-ToolStatusColour -Kind Good)
    }
    $hot = @($answered | Where-Object { $_.BadPwdCount -gt 0 } | Sort-Object -Property BadPwdCount -Descending)
    $foot = @()
    if ($hot.Count) { $foot += ('Bad passwords recorded on: ' + (($hot | ForEach-Object { "$($_.Name) ($($_.BadPwdCount))" }) -join ', ') + '.') }
    if ($silent.Count) { $foot += ('{0} controller(s) did not answer.' -f $silent.Count) }
    $lockFoot.Text = ($foot -join ' ')
    $btnUnlockAll.Enabled = ($answered.Count -gt 0)
}

function Update-Identity {
    $t = $script:Target
    if (-not $t -or -not $t.Ok) {
        $lblIdent.Text = 'No account loaded.'; $lblIdent.ForeColor = (Get-ToolStatusColour -Kind Muted)
        $lblSync.Text = ''
        return
    }
    $lblIdent.Text = ('{0}  <{1}>  {2}  [{3}]' -f $t.Display, $(if ($t.Smtp) { $t.Smtp } else { $t.Upn }), $t.Sam, $t.Directory)
    $lblIdent.ForeColor = (Get-ToolStatusColour -Kind Good)
    $lblSync.Text = ('Read from {0}. Domain {1}; the PDC emulator is {2}, and a reset is written there.' -f $t.Server, $t.DomainDns, $t.Pdc)
    $lblSync.ForeColor = (Get-ToolStatusColour -Kind Muted)
}

function Set-LoadedState {
    param([bool]$On)
    $txtPw1.Enabled = $On; $txtPw2.Enabled = $On
    $btnScan.Enabled = $On
    if (-not $On) { $btnUnlockAll.Enabled = $false; $btnReset.Enabled = $false }
}
#endregion

#region Load --------------------------------------------------------------------
function Invoke-PasswordReload {
    # Re-read from the SAME controller and repaint. The lockout grid is left
    # alone: it is a scan result, not a property of the object.
    if (-not $script:Target) { return }
    $t = $script:Target.Upn; if (-not $t) { $t = $script:Target.Sam }
    $n = Resolve-PasswordTarget -Identity $t
    if ($n.Ok) {
        $script:Target = $n
        $script:Policy = Get-PasswordPolicy -Dn $n.Dn -Server $n.Server
    }
    Update-Identity
    Update-PasswordSection
    Update-RulesGrid
}

function Invoke-PasswordLoad {
    $t = "$($txtWho.Text)".Trim()
    if (-not $t) { Set-PwStatus -Text 'Enter an address, UPN or account name.' -Bad; return }
    Set-PwStatus -Text "Resolving $t..."
    $script:Target = $null; $script:Policy = $null; $script:Dcs = @()
    $txtPw1.Text = ''; $txtPw2.Text = ''
    Update-LockoutGrid; Update-LockoutSummary
    $r = Resolve-PasswordTarget -Identity $t
    if (-not $r.Ok) {
        Update-Identity; Update-PasswordSection; Update-RulesGrid
        Set-LoadedState $false
        Set-PwStatus -Text $r.Reason -Bad
        return
    }
    $script:Target = $r
    $script:Policy = Get-PasswordPolicy -Dn $r.Dn -Server $r.Server
    Update-Identity
    Update-PasswordSection
    Update-RulesGrid
    Set-LoadedState $true
    $lblResetNote.Text = ('The reset is written to {0}. Minimum length here is {1}: the domain minimum, or 12, whichever is higher.' -f $r.Pdc, (Get-RequiredLength))
    Set-PwStatus -Text "Loaded $($r.Display). Press Check every DC for the lockout picture."
}

$btnLoad.Add_Click({ Invoke-PwAction { Invoke-PasswordLoad } 'Load account' })
$txtWho.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $btnLoad.PerformClick(); $_.SuppressKeyPress = $true } })
#endregion

#region Must change -------------------------------------------------------------
$btnMustChange.Add_Click({
    if (-not (Test-PasswordLoaded)) { return }
    Invoke-PwAction {
        $want = [bool]$chkMustChange.Checked
        $was  = [bool]$chkMustChange.Tag
        if ($want -eq $was) { Set-PwStatus -Text 'Nothing to change.'; return }
        $u = $script:Target.Ad
        if ($want -and $u.PasswordNeverExpires) {
            Set-PwStatus -Text "Active Directory will not set 'must change at next logon' while PasswordNeverExpires is set. Clear that first." -Bad
            $chkMustChange.Checked = $was
            return
        }
        $msg = $(if ($want) { "Require $($script:Target.Display) to change their password at next logon?`r`n`r`nThey will be prompted at their next interactive sign-in. Cached credentials on phones and other devices will start failing once the password changes." }
                 else { "Clear 'must change at next logon' for $($script:Target.Display)?`r`n`r`nTheir current password stays in force." })
        $a = Show-ToolDialog($msg, 'Must change at next logon', 'OKCancel', 'Question')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { $chkMustChange.Checked = $was; return }
        Invoke-PasswordAd -Name 'Set-ADUser' -Parameters @{ Identity = $script:Target.Dn; ChangePasswordAtLogon = $want } | Out-Null
        Write-PasswordAudit -Action 'PasswordMustChange' -Detail $(if ($want) { 'Must change at next logon: set.' } else { 'Must change at next logon: cleared.' })
        Invoke-PasswordReload
        Set-PwStatus -Text $(if ($want) { 'Must change at next logon is set.' } else { 'Must change at next logon is cleared.' })
    } 'Must change at next logon'
})
#endregion

#region Lockout scan ------------------------------------------------------------
function Invoke-LockoutScan {
    $t = $script:Target
    $script:Cancel = $false
    $script:Dcs = @()
    Update-LockoutGrid; Update-LockoutSummary
    Set-PwStatus -Text "Listing the domain controllers of $($t.DomainDns)..."
    $list = @(Get-DomainControllerList -Server $t.Server)
    if ($list.Count -eq 0) { Set-PwStatus -Text "No domain controller could be listed for $($t.DomainDns)." -Bad; return }
    $btnCancel.Enabled = $true
    $rows = New-Object System.Collections.Generic.List[object]
    $i = 0
    try {
        foreach ($dc in $list) {
            if ($script:Cancel -or $script:Closing) { break }
            $i++
            Set-PwStatus -Text ('Asking {0} ({1} of {2})...' -f $dc.Name, $i, $list.Count)
            [System.Windows.Forms.Application]::DoEvents()
            $row = [PSCustomObject]@{
                Name = $dc.Name; Site = $dc.Site; IsPdc = $dc.IsPdc; IsRodc = $dc.IsRodc
                Reachable = $false; LockState = 'Unknown'; LockoutTime = $null; BadPwdCount = 0
                BadPasswordTime = $null; LastLogon = $null; PasswordLastSet = $null; Note = ''
            }
            if (-not (Test-DcReachable -HostName $dc.Name)) {
                $row.Note = 'Unreachable (no answer on the LDAP port within 2 seconds)'
            } else {
                $res = Get-AccountOnDc -Dn $t.Dn -Server $dc.Name
                if ($res.Ok) {
                    $row.Reachable = $true
                    $row.LockState = $(if ($res.Locked) { 'LOCKED' } else { 'Not locked' })
                    $row.LockoutTime = $res.LockoutTime; $row.BadPwdCount = $res.BadPwdCount
                    $row.BadPasswordTime = $res.BadPasswordTime; $row.LastLogon = $res.LastLogon; $row.PasswordLastSet = $res.PasswordLastSet
                    if ($res.LockoutTime -and -not $res.Locked) { $row.Note = 'lockoutTime set but the lock has expired' }
                } else { $row.Note = $res.Error }
            }
            $rows.Add($row)
            # Render as we go: on a slow link the first answers are useful
            # long before the last controller replies.
            $script:Dcs = $rows.ToArray()
            Update-LockoutGrid
            [System.Windows.Forms.Application]::DoEvents()
        }
    } finally {
        $script:Dcs = $rows.ToArray()
        # THE WINDOW MAY HAVE CLOSED UNDER THE LOOP -- DoEvents lets it. A
        # disposed grid throws on Rows.Clear, so the repaint is skipped, not
        # attempted and caught.
        if (-not $script:Closing) {
            $btnCancel.Enabled = $false
            Update-LockoutGrid
            Update-LockoutSummary
        }
    }
    if ($script:Closing) { return }
    if ($script:Cancel) { Set-PwStatus -Text ('Stopped after {0} of {1} controller(s).' -f $rows.Count, $list.Count) }
    else { Set-PwStatus -Text ('Asked {0} controller(s).' -f $rows.Count) }
}

$btnScan.Add_Click({
    if (-not (Test-PasswordLoaded)) { return }
    Invoke-PwAction { Invoke-LockoutScan } 'Check every DC'
})
# STOP IS THE ONE BUTTON THAT WORKS WHILE BUSY. It only sets a flag; the loop
# reads it between controllers.
$btnCancel.Add_Click({ $script:Cancel = $true; $btnCancel.Enabled = $false })

function Invoke-UnlockEverywhere {
    # Unlock on every controller that answered the last scan -- or, with no
    # scan yet, on every controller that answers now. Returns the summary
    # line for the caller to log and show.
    $t = $script:Target
    $targets = @($script:Dcs | Where-Object { $_.Reachable } | ForEach-Object { $_.Name })
    if ($targets.Count -eq 0) {
        $targets = @(Get-DomainControllerList -Server $t.Server | ForEach-Object { $_.Name } | Where-Object { Test-DcReachable -HostName $_ })
    }
    $ok = @(); $failed = @()
    foreach ($dc in $targets) {
        Set-PwStatus -Text "Unlocking on $dc..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            Invoke-PasswordAd -Name 'Unlock-ADAccount' -Parameters @{ Identity = $t.Dn } -Server $dc | Out-Null
            $ok += $dc
        } catch { $failed += "$dc ($($_.Exception.Message))" }
    }
    return [PSCustomObject]@{ Ok = $ok; Failed = $failed }
}

$btnUnlockAll.Add_Click({
    if (-not (Test-PasswordLoaded)) { return }
    Invoke-PwAction {
        $a = Show-ToolDialog("Unlock $($script:Target.Display) on every domain controller that answers?`r`n`r`nIf it locks again within minutes, something is still trying an old password -- the grid says which controller is seeing it.",
                             'Unlock on all DCs', 'OKCancel', 'Question')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $r = Invoke-UnlockEverywhere
        $detail = ('Unlocked on {0} controller(s){1}.' -f $r.Ok.Count, $(if ($r.Failed.Count) { "; failed on $($r.Failed.Count): $($r.Failed -join '; ')" } else { '' }))
        Write-PasswordAudit -Action 'PasswordUnlockAll' -Detail $detail -Result $(if ($r.Failed.Count -and $r.Ok.Count) { 'Partial' } elseif ($r.Failed.Count) { 'Failed' } else { 'Success' })
        Invoke-PasswordReload
        if ($script:Dcs.Count) { Invoke-LockoutScan }
        Set-PwStatus -Text $detail -Bad:($r.Failed.Count -gt 0)
    } 'Unlock on all DCs'
})
#endregion

#region Reset -------------------------------------------------------------------
$txtPw1.Add_TextChanged({ Update-RulesGrid })
$txtPw2.Add_TextChanged({ Update-RulesGrid })
$chkShow.Add_CheckedChanged({
    $txtPw1.UseSystemPasswordChar = -not $chkShow.Checked
    $txtPw2.UseSystemPasswordChar = -not $chkShow.Checked
})

$btnReset.Add_Click({
    if (-not (Test-PasswordLoaded)) { return }
    Invoke-PwAction {
        $pw = $txtPw1.Text
        if (-not (Test-PasswordAcceptable -Password $pw -Confirm $txtPw2.Text)) { Set-PwStatus -Text 'The password does not meet every rule yet.' -Bad; return }
        $t = $script:Target
        $must   = [bool]$chkResetMustChange.Checked
        $unlock = [bool]$chkResetUnlock.Checked
        if ($must -and $t.Ad.PasswordNeverExpires) {
            $a = Show-ToolDialog("PasswordNeverExpires is set on this account, so Active Directory will refuse 'must change at next logon'.`r`n`r`nReset the password WITHOUT that tick?",
                                 'Reset password', 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $must = $false
        }
        $msg = ("Reset the password for $($t.Display) ($($t.Sam))?" + "`r`n`r`n" +
                "Written to $($t.Pdc)." + "`r`n" +
                "Must change at next logon: $(if ($must) { 'yes' } else { 'no' })" + "`r`n" +
                "Unlock on all DCs: $(if ($unlock) { 'yes' } else { 'no' })" + "`r`n`r`n" +
                "Anything signed in with the old password -- phones, mapped drives, scheduled tasks -- will start failing and can lock the account again.")
        $a = Show-ToolDialog($msg, 'Reset password', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $secure = ConvertTo-SecureString -String $pw -AsPlainText -Force
        $pw = $null
        Set-PwStatus -Text "Writing the new password to $($t.Pdc)..."
        Invoke-PasswordAd -Name 'Set-ADAccountPassword' -Parameters @{ Identity = $t.Dn; Reset = $true; NewPassword = $secure } -Server $t.Pdc | Out-Null
        # THE BOXES ARE CLEARED THE MOMENT THE WRITE RETURNS, before anything
        # that could fail. A failed must-change should not leave the password
        # sitting on screen.
        $txtPw1.Text = ''; $txtPw2.Text = ''; $chkShow.Checked = $false
        $warnings = @()
        if ($must) {
            try { Invoke-PasswordAd -Name 'Set-ADUser' -Parameters @{ Identity = $t.Dn; ChangePasswordAtLogon = $true } -Server $t.Pdc | Out-Null }
            catch { $warnings += "must-change could not be set: $($_.Exception.Message)" }
        }
        $unlockText = ''
        if ($unlock) {
            $r = Invoke-UnlockEverywhere
            $unlockText = ('unlocked on {0} controller(s)' -f $r.Ok.Count)
            if ($r.Failed.Count) { $warnings += "unlock failed on $($r.Failed.Count): $($r.Failed -join '; ')" }
        }
        $detail = ('Password reset on {0}; must change at next logon: {1}{2}{3}' -f $t.Pdc, $(if ($must) { 'yes' } else { 'no' }),
                   $(if ($unlockText) { "; $unlockText" } else { '' }), $(if ($warnings.Count) { '; WARNINGS: ' + ($warnings -join ' | ') } else { '' }))
        Write-PasswordAudit -Action 'PasswordReset' -Detail $detail -Result $(if ($warnings.Count) { 'Partial' } else { 'Success' })
        Invoke-PasswordReload
        if ($script:Dcs.Count) { Invoke-LockoutScan }
        if ($warnings.Count) { Set-PwStatus -Text ('Password reset, with warnings: ' + ($warnings -join ' | ')) -Bad }
        else { Set-PwStatus -Text "Password reset on $($t.Pdc)." }
    } 'Reset password'
})
#endregion

$form.Add_FormClosing({ $script:Closing = $true; $script:Cancel = $true })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Password Manager' | Out-Null

$form.Add_Shown({
    if ($null -eq $script:AdReady) { $script:AdReady = Test-AdAvailable }
    if (-not $script:AdReady.Available) {
        $lblSync.Text = "Read-only: $($script:AdReady.Detail)"
        $lblSync.ForeColor = (Get-ToolStatusColour -Kind Warn)
        $btnLoad.Enabled = $false
        Set-PwStatus -Text 'This window needs the ActiveDirectory module and a reachable controller.' -Bad
        return
    }
    Update-RulesGrid
    if ("$($txtWho.Text)".Trim()) { $btnLoad.PerformClick() }
})

[void]$form.ShowDialog()
try { $form.Dispose() } catch { }
