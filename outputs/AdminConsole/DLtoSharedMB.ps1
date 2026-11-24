#Requires -Version 5.1
# TOOLKIT-VARIANT: Dual
#
# RUNS AGAINST ON-PREMISES for steps 1-3 and against EXCHANGE ONLINE for step
# 4, from ONE window in the console's own process. Nothing here is called
# unqualified: every on-premises call goes through Invoke-OnPremCommand and
# every cloud call through Invoke-EXOCommand (see Unified-Lookup.ps1 for why
# that is the trick, and why this header must not name the own-process
# marker). Connect both before starting; the window says which is missing.
<#
    DLtoSharedMB.ps1
    One distribution list becomes one shared mailbox, in five numbered steps.

    Author   : Glenn Armbruster
    Revised  : 2026-09-09
#>
# TOOLKIT-HELP-RUNS: step 1 reads. Steps 2, 3, 4 and 5 change things; 2, 3 and 5 ask first.
# TOOLKIT-HELP-BEGIN
# "Make legal@ a shared inbox instead of a distribution list."
#
# The address has to move from a group to a mailbox, and the two cannot exist
# at once. So: capture everything about the list, delete it, wait for a sync
# so the cloud forgets the address, create the mailbox on-premises with the
# same address, wait for a sync so the cloud has the mailbox, then grant the
# old members access in the cloud.
#
#   1  EXPORT (on-premises)   Every attribute of the list -- owners, members,
#                             every proxy address, LegacyExchangeDN (the X500
#                             that keeps replies to old mail working), sender
#                             restrictions, custom attributes, Send As -- to a
#                             CSV in the reports folder. Read only.
#   2  DELETE (on-premises)   Distribution group: Remove-DistributionGroup.
#                             Mail-enabled SECURITY group: Disable- only -- the
#                             group may secure shares or apps, so it is kept and
#                             just loses its address. Needs step 1 done in
#                             this window. Then WAIT for a sync (30 min) and
#                             press CHECK SYNC: one Get-Recipient in the cloud
#                             on the list's address. "Still exists" -- wait;
#                             "deleted" -- Create unlocks. After step 3 the same
#                             button reports when the new mailbox has arrived.
#   3  CREATE (on-premises)   Load the report (newest for this list is picked
#                             for you), confirm the OU, account name and OWNER,
#                             press Create. New-RemoteMailbox -Shared with the
#                             list's address, then every proxy, the X500, the
#                             hidden/restriction/custom attributes, and the
#                             owner as the account's Manager -- all pinned to
#                             a writable DC in the OBJECT'S OWN DOMAIN (from
#                             its DN via Tenant.json), so a session parked on
#                             another domain's DC can neither lose it nor get
#                             a referral. WAIT for a sync.
#   4  STAMP (on-premises)    The same stamp, on its own button, repeatable:
#                             if step 3 reported a FAIL line, press this.
#                             Adds only what is missing.
#   5  PERMISSIONS (cloud)    Load the same report. Every former member and
#                             owner gets Full Access and Send As; Send on Behalf
#                             carries over; sent-items copy is turned on for
#                             both Send As and Send on Behalf. Idempotent --
#                             grants already present are reported, not
#                             repeated.
#
# OWNER IS REQUIRED. If the list had owners they are pre-filled; if not, you
# must name one before Create is allowed. A shared mailbox with nobody
# responsible for it is how a shared mailbox becomes nobody's.
#
# The OU defaults to the directory's shared-mailbox OU from Tenant.json
# (Tenant Manager, "Shared mbx OU"); it is editable here.
#
# Steps can be run in different sessions: 3 and 4 start from the report file,
# not from memory.
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

$script:Target = Get-TargetSmtp -Text $SmtpAddress
$script:ListSep = ';'

#region Helpers ---------------------------------------------------------------
function onprem { param([string]$n, [hashtable]$p) Invoke-OnPremCommand -Name $n -Parameters $p }
function exo    { param([string]$n, [hashtable]$p) Invoke-EXOCommand    -Name $n -Parameters $p }

function ConvertTo-CsvList   { param([object]$Value) if ($null -eq $Value) { return '' }; return (@($Value | ForEach-Object { "$_" } | Where-Object { $_ }) -join $script:ListSep) }
function ConvertFrom-CsvList { param([string]$Text) if (-not "$Text".Trim()) { return @() }; return @("$Text" -split [regex]::Escape($script:ListSep) | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

function Test-OnPremReady { try { return [bool](Test-EmsAvailable).Available } catch { return $false } }
function Test-CloudReady  { try { return [bool](Test-ExoInProcess).Present }  catch { return $false } }

function Get-OnPremSmtp {
    # Recipient references (ManagedBy etc.) come back as canonical names;
    # the report stores addresses because only an address survives.
    param([object]$Refs)
    $out = @()
    foreach ($r in @($Refs)) {
        $id = "$r"; if (-not $id) { continue }
        $x = $null
        try { $x = onprem 'Get-Recipient' @{ Identity = $id; ErrorAction = 'Stop' } } catch { }
        $out += $(if ($x -and "$($x.PrimarySmtpAddress)") { "$($x.PrimarySmtpAddress)" } else { $id })
    }
    return $out
}

function Get-AccountNameFromAlias {
    # sAMAccountName is capped at 20; the same rule ManageMailObject uses.
    param([string]$Alias)
    $s = ("$Alias".Trim() -replace '[^A-Za-z0-9 .!#$%^&\-_{}~]', '')
    if ($s.Length -gt 20) { $s = $s.Substring(0, 20) }
    return $s.TrimEnd('.')
}

function Find-AccountNameHolder {
    <#
        Who holds this sAMAccountName -- ANY object, not just users. Get-User
        misses groups, and a mail-disabled security group that kept the list's
        old account name is exactly the collision this tool creates in step 2.
        AD module when present (one LDAP query), Exchange Get-User + Get-Group
        otherwise. Returns a short description, or '' when free.
    #>
    param([Parameter(Mandatory)][string]$Sam)
    try {
        if (Get-Command -Name 'Get-ADObject' -ErrorAction SilentlyContinue) {
            $esc = $Sam -replace '\\', '\5c' -replace '\*', '\2a' -replace '\(', '\28' -replace '\)', '\29'
            $o = @(Get-ADObject -LDAPFilter "(sAMAccountName=$esc)" -Properties objectClass -ErrorAction Stop | Select-Object -First 1)
            if ($o.Count) { return "$($o[0].ObjectClass) '$($o[0].Name)'" }
            return ''
        }
    } catch { }
    try { $u = onprem 'Get-User'  @{ Identity = $Sam; ErrorAction = 'Stop' }; if ($u) { return "user '$($u.DisplayName)'" } } catch { }
    try { $g = onprem 'Get-Group' @{ Identity = $Sam; ErrorAction = 'Stop' }; if ($g) { return "group '$($g.Name)'" } } catch { }
    return ''
}

function Read-AccountName {
    # Small prompt: the account name is taken, type another. Loops until the
    # name is free or the operator cancels. Returns '' on cancel.
    param([string]$Suggested, [string]$TakenBy)
    while ($true) {
        $f = New-Object System.Windows.Forms.Form
        $f.Text = 'Account name is taken'; $f.ClientSize = New-Object System.Drawing.Size(460, 140)
        $f.FormBorderStyle = 'FixedDialog'; $f.StartPosition = 'CenterParent'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
        $f.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
        $l = New-Object System.Windows.Forms.Label; $l.Location = New-Object System.Drawing.Point(12, 12); $l.Size = New-Object System.Drawing.Size(436, 40)
        $l.Text = "'$Suggested' is used by $TakenBy.`r`nOnly the account name changes -- alias and address stay as they are. Enter another (max 20):"
        $f.Controls.Add($l)
        $t = New-Object System.Windows.Forms.TextBox; $t.Location = New-Object System.Drawing.Point(12, 60); $t.Size = New-Object System.Drawing.Size(300, 24); $t.MaxLength = 20
        $t.Font = New-Object System.Drawing.Font('Consolas', 9); $t.Text = (Get-AccountNameFromAlias ($Suggested + 'SMB')); $f.Controls.Add($t)
        $ok = New-Object System.Windows.Forms.Button; $ok.Text = 'Use this'; $ok.Location = New-Object System.Drawing.Point(252, 98); $ok.Size = New-Object System.Drawing.Size(94, 28); $ok.DialogResult = 'OK'; $ok.Tag = 'Primary'; $f.Controls.Add($ok)
        $no = New-Object System.Windows.Forms.Button; $no.Text = 'Cancel'; $no.Location = New-Object System.Drawing.Point(354, 98); $no.Size = New-Object System.Drawing.Size(94, 28); $no.DialogResult = 'Cancel'; $f.Controls.Add($no)
        $f.AcceptButton = $ok; $f.CancelButton = $no
        Set-ConsoleTheme -Form $f
        $t.SelectAll()
        if ($f.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return '' }
        $new = "$($t.Text)".Trim()
        if (-not $new -or $new.Length -gt 20 -or $new.EndsWith('.')) { continue }
        $holder = Find-AccountNameHolder -Sam $new
        if (-not $holder) { return $new }
        $Suggested = $new; $TakenBy = $holder
    }
}

function Get-WriteDomainController {
    <#
        A DC that can WRITE this DN: the object's directory (longest DN suffix
        wins, so a child-domain object is not mistaken for its parent's), then
        Resolve-DirectoryDomainController -- the Tenant.json controller if one
        is named, otherwise a writable one discovered in that domain. -Fallback
        (normally OriginatingServer) is used only when the DN matches no
        configured directory; '' means "let the session choose", which is the
        very thing that failed, so callers should say so in the log.
    #>
    param([string]$DistinguishedName, [string]$Fallback)
    try {
        $dir = Resolve-DistinguishedNameDirectory -DistinguishedName $DistinguishedName
        if ($dir) {
            $dc = Resolve-DirectoryDomainController -DirectoryId "$($dir.Id)"
            if ("$dc".Trim()) { return "$dc".Trim() }
        }
    } catch { }
    return "$Fallback".Trim()
}

function Invoke-AttributeStamp {
    <#
        Everything from the report onto the remote mailbox, IDEMPOTENT.

        PINNED TO A WRITABLE DC IN THE OBJECT'S OWN DOMAIN. The session's
        default DC can be in another domain of the forest, and two failures
        taught two halves of the lesson:
          - "couldn't be found on dc01.contoso.com": a write seconds after the
            create had not replicated, and a child-domain DC never holds an
            object in another directory anyway.
          - "Received a referral to contoso.com ... from dc01.contoso.com
            ... wrong server": after a -ViewEntireForest read the object IS
            found -- through the global catalog -- and OriginatingServer
            reports the GC that answered, which is the same child-domain DC.
            A GC can read any domain; it can only WRITE its own. Pinning the
            write to OriginatingServer therefore sent it to the wrong domain.
        So the DC comes from the DistinguishedName: which configured directory
        the DN belongs to, then Resolve-DirectoryDomainController for that
        directory (named in Tenant.json, else discovered). OriginatingServer
        is the fallback only when the DN matches no directory. Same lesson as
        ARCHITECTURE 14f, from the other side.

        Proxies are added only if missing, scalars set to the report's value;
        running it twice changes nothing the second time.
    #>
    param(
        [Parameter(Mandatory)]$Report,
        [Parameter(Mandatory)]$RemoteMailbox,
        [string]$OwnerSmtp,
        [Parameter(Mandatory)][scriptblock]$Log
    )
    $r = $Report; $rm = $RemoteMailbox
    $id = "$($rm.DistinguishedName)"
    $dc = Get-WriteDomainController -DistinguishedName $id -Fallback "$($rm.OriginatingServer)"
    if ($DomainController) { $dc = $DomainController }
    $base = @{ Identity = $id; ErrorAction = 'Stop' }
    if ($dc) { $base.DomainController = $dc; & $Log "Writing via $dc (the object's own domain)." }

    # Proxies: every address the list had, plus its LegacyExchangeDN as
    # X500 so replies to mail sent to the LIST still resolve.
    $want = @(ConvertFrom-CsvList $r.EmailAddresses | ForEach-Object { if ($_ -match '^SMTP:') { 'smtp:' + $_.Substring(5) } else { $_ } })
    if ("$($r.LegacyExchangeDN)") { $want += "X500:$($r.LegacyExchangeDN)" }
    $have = @($rm.EmailAddresses | ForEach-Object { "$_".ToLower() })
    $add = @($want | Where-Object { $have -notcontains "$_".ToLower() } | Select-Object -Unique)

    $set = @{}
    if ($add.Count) { $set.EmailAddresses = @{ Add = $add } }
    foreach ($k in 'HiddenFromAddressListsEnabled', 'RequireSenderAuthenticationEnabled', 'ModerationEnabled') {
        if ("$($r.$k)" -match '^(True|False)$') { $set[$k] = ("$($r.$k)" -eq 'True') }
    }
    foreach ($k in 'AcceptMessagesOnlyFrom', 'AcceptMessagesOnlyFromDLMembers', 'RejectMessagesFrom', 'RejectMessagesFromDLMembers', 'ModeratedBy', 'GrantSendOnBehalfTo') {
        $v = @(ConvertFrom-CsvList $r.$k); if ($v.Count) { $set[$k] = $v }
    }
    if ("$($r.MailTip)") { $set.MailTip = "$($r.MailTip)" }
    for ($n = 1; $n -le 15; $n++) { if ("$($r."CustomAttribute$n")") { $set["CustomAttribute$n"] = "$($r."CustomAttribute$n")" } }
    for ($n = 1; $n -le 5; $n++)  { $v = @(ConvertFrom-CsvList $r."ExtensionCustomAttribute$n"); if ($v.Count) { $set["ExtensionCustomAttribute$n"] = $v } }

    $ok = 0; $fail = 0
    if ($set.Count -eq 0) { & $Log "Stamp: nothing to set (no proxies missing, no attributes in the report)." }
    else {
        try {
            # Learn (Set-RemoteMailbox): -EmailAddresses @{Add=}, the
            # restriction/moderation/custom-attribute parameters as on
            # Set-Mailbox; -DomainController pins the write.
            onprem 'Set-RemoteMailbox' ($base + $set) | Out-Null
            & $Log "Stamped on $dc`: $($add.Count) proxy address(es)$(if ($add.Count) { ' (incl. X500)' }), $(@($set.Keys | Where-Object { $_ -ne 'EmailAddresses' }) -join ', ')"
            $ok = $set.Count
        } catch {
            & $Log "Stamp as one call failed ($($_.Exception.Message)); retrying property by property..."
            foreach ($k in @($set.Keys)) {
                try { onprem 'Set-RemoteMailbox' ($base + @{ $k = $set[$k] }) | Out-Null; & $Log "  ok    $k"; $ok++ }
                catch { & $Log "  FAIL  $k -- $($_.Exception.Message)"; $fail++ }
            }
        }
    }
    # Owner = the account's Manager. That is the field the GAL shows and
    # the closest thing a mailbox has to ManagedBy. Access comes in step 5.
    if ($OwnerSmtp) {
        try {
            onprem 'Set-User' ($base + @{ Manager = $OwnerSmtp }) | Out-Null
            & $Log "Owner recorded as Manager: $OwnerSmtp"
        } catch { & $Log "Manager not set: $($_.Exception.Message)"; $fail++ }
    }
    Write-CompletedLog -Action 'DLtoSMBStamp' -Target (Get-LogTarget $rm.PrimarySmtpAddress) -Result $(if ($fail -and $ok) { 'Partial' } elseif ($fail) { 'Failed' } else { 'Success' }) `
        -Detail "Set-RemoteMailbox on $dc`: $ok set, $fail failed; $($add.Count) proxies added; manager $OwnerSmtp"
    return ($fail -eq 0)
}

function Export-ListReport {
    <#
        Everything about the list, one CSV row, lists ';' separated. The
        columns mirror Projects\Migration step 1 so the two are readable side
        by side. Returns the path.
    #>
    param([Parameter(Mandatory)][string]$Identity)
    $dg = onprem 'Get-DistributionGroup' @{ Identity = $Identity; ErrorAction = 'Stop' }
    $grp = $null; try { $grp = onprem 'Get-Group' @{ Identity = $dg.Identity; ErrorAction = 'Stop' } } catch { }
    $members = @(onprem 'Get-DistributionGroupMember' @{ Identity = $dg.Identity; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
                 ForEach-Object { if ("$($_.PrimarySmtpAddress)") { "$($_.PrimarySmtpAddress)" } else { "$($_.Name)" } })
    $sendAs = @()
    try {
        # On-premises Send As is an AD ACE, reported as DOMAIN\sam.
        foreach ($a in @(onprem 'Get-ADPermission' @{ Identity = $dg.DistinguishedName; ErrorAction = 'Stop' } |
                 Where-Object { -not $_.IsInherited -and $_.ExtendedRights -match 'Send-As' -and "$($_.User)" -notmatch 'NT AUTHORITY\\SELF|^S-1-5' })) {
            $sam = ("$($a.User)" -split '\\')[-1]
            $r = $null; try { $r = onprem 'Get-Recipient' @{ Identity = $sam; ErrorAction = 'Stop' } } catch { }
            $sendAs += $(if ($r) { "$($r.PrimarySmtpAddress)" } else { "$($a.User)" })
        }
    } catch { }

    $row = [ordered]@{
        Name                                 = "$($dg.Name)"
        DisplayName                          = "$($dg.DisplayName)"
        Alias                                = "$($dg.Alias)"
        PrimarySmtpAddress                   = "$($dg.PrimarySmtpAddress)"
        EmailAddresses                       = ConvertTo-CsvList ($dg.EmailAddresses | ForEach-Object { $_.ProxyAddressString })
        LegacyExchangeDN                     = "$($dg.LegacyExchangeDN)"
        GroupType                            = "$($dg.GroupType)"
        SamAccountName                       = "$($dg.SamAccountName)"
        Guid                                 = "$($dg.Guid)"
        DistinguishedName                    = "$($dg.DistinguishedName)"
        OrganizationalUnit                   = "$($dg.OrganizationalUnit)"
        ManagedBy                            = ConvertTo-CsvList (Get-OnPremSmtp $dg.ManagedBy)
        Members                              = ConvertTo-CsvList $members
        MemberCount                          = $members.Count
        SendAs                               = ConvertTo-CsvList $sendAs
        GrantSendOnBehalfTo                  = ConvertTo-CsvList (Get-OnPremSmtp $dg.GrantSendOnBehalfTo)
        RequireSenderAuthenticationEnabled   = "$($dg.RequireSenderAuthenticationEnabled)"
        AcceptMessagesOnlyFrom               = ConvertTo-CsvList (Get-OnPremSmtp $dg.AcceptMessagesOnlyFrom)
        AcceptMessagesOnlyFromDLMembers      = ConvertTo-CsvList (Get-OnPremSmtp $dg.AcceptMessagesOnlyFromDLMembers)
        RejectMessagesFrom                   = ConvertTo-CsvList (Get-OnPremSmtp $dg.RejectMessagesFrom)
        RejectMessagesFromDLMembers          = ConvertTo-CsvList (Get-OnPremSmtp $dg.RejectMessagesFromDLMembers)
        ModerationEnabled                    = "$($dg.ModerationEnabled)"
        ModeratedBy                          = ConvertTo-CsvList (Get-OnPremSmtp $dg.ModeratedBy)
        HiddenFromAddressListsEnabled        = "$($dg.HiddenFromAddressListsEnabled)"
        MailTip                              = "$($dg.MailTip)"
        Notes                                = $(if ($grp) { "$($grp.Notes)" } else { '' })
        WhenCreated                          = "$($dg.WhenCreated)"
    }
    for ($n = 1; $n -le 15; $n++) { $row["CustomAttribute$n"] = "$($dg."CustomAttribute$n")" }
    for ($n = 1; $n -le 5; $n++)  { $row["ExtensionCustomAttribute$n"] = ConvertTo-CsvList $dg."ExtensionCustomAttribute$n" }
    $row['ExportedUtc'] = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')

    $local = ("$($dg.PrimarySmtpAddress)" -split '@')[0]
    $path = Join-Path (Get-ReportPath) "DLtoSMB_${local}_$(Get-Date -Format 'yyyy-MM-dd_HHmm').csv"
    [PSCustomObject]$row | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    return $path
}

function Find-NewestReport {
    param([string]$Smtp)
    $dir = Get-ReportPath
    $pat = 'DLtoSMB_*.csv'
    if ("$Smtp".Trim()) { $pat = "DLtoSMB_$(("$Smtp" -split '@')[0])_*.csv" }
    return (Get-ChildItem -LiteralPath $dir -Filter $pat -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
}
#endregion

#region Form ------------------------------------------------------------------
function Show-ConvertForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'DL to Shared Mailbox'
    $form.ClientSize = New-Object System.Drawing.Size(820, 716)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblConn = New-Object System.Windows.Forms.Label
    $lblConn.Location = New-Object System.Drawing.Point(12, $y); $lblConn.Size = New-Object System.Drawing.Size(796, 18); $lblConn.Tag = 'NoTheme'
    $form.Controls.Add($lblConn)
    $y += 24

    # ---- 1 + 2 ----
    $grp12 = New-Object System.Windows.Forms.GroupBox
    $grp12.Text = '1 - 2   On-premises: export the list, then delete it'
    $grp12.Location = New-Object System.Drawing.Point(12, $y); $grp12.Size = New-Object System.Drawing.Size(796, 58)
    $form.Controls.Add($grp12)
    $btnExport = New-Object System.Windows.Forms.Button; $btnExport.Text = '1  Export DL'; $btnExport.Location = New-Object System.Drawing.Point(12, 22); $btnExport.Size = New-Object System.Drawing.Size(130, 28); $grp12.Controls.Add($btnExport)
    $btnDelete = New-Object System.Windows.Forms.Button; $btnDelete.Text = '2  Delete DL'; $btnDelete.Location = New-Object System.Drawing.Point(150, 22); $btnDelete.Size = New-Object System.Drawing.Size(130, 28); $btnDelete.Enabled = $false; $grp12.Controls.Add($btnDelete)
    $btnCheck = New-Object System.Windows.Forms.Button; $btnCheck.Text = 'Check sync (cloud)'; $btnCheck.Location = New-Object System.Drawing.Point(288, 22); $btnCheck.Size = New-Object System.Drawing.Size(140, 28); $btnCheck.Enabled = $false; $btnCheck.Tag = 'Plain'; $grp12.Controls.Add($btnCheck)
    $lbl12 = New-Object System.Windows.Forms.Label; $lbl12.Location = New-Object System.Drawing.Point(436, 27); $lbl12.Size = New-Object System.Drawing.Size(346, 18); $lbl12.ForeColor = [System.Drawing.Color]::DimGray
    $lbl12.Text = 'Export, delete, wait a sync, Check sync until the list is gone.'
    $grp12.Controls.Add($lbl12)
    $y += 66

    # ---- report ----
    $grpR = New-Object System.Windows.Forms.GroupBox
    $grpR.Text = 'Report  (steps 3 and 4 start from here)'
    $grpR.Location = New-Object System.Drawing.Point(12, $y); $grpR.Size = New-Object System.Drawing.Size(796, 132)
    $form.Controls.Add($grpR)
    $btnLoad = New-Object System.Windows.Forms.Button; $btnLoad.Text = 'Load report...'; $btnLoad.Location = New-Object System.Drawing.Point(12, 22); $btnLoad.Size = New-Object System.Drawing.Size(130, 28); $grpR.Controls.Add($btnLoad)
    $lblRep = New-Object System.Windows.Forms.Label; $lblRep.Location = New-Object System.Drawing.Point(150, 27); $lblRep.Size = New-Object System.Drawing.Size(630, 18); $lblRep.ForeColor = [System.Drawing.Color]::DimGray; $lblRep.Text = '(none loaded)'; $grpR.Controls.Add($lblRep)
    $txtSum = New-Object System.Windows.Forms.TextBox
    $txtSum.Multiline = $true; $txtSum.ReadOnly = $true; $txtSum.ScrollBars = 'Vertical'
    $txtSum.Location = New-Object System.Drawing.Point(12, 56); $txtSum.Size = New-Object System.Drawing.Size(770, 66); $txtSum.Font = New-Object System.Drawing.Font('Consolas', 9)
    $grpR.Controls.Add($txtSum)
    $y += 140

    # ---- 3 ----
    $grp3 = New-Object System.Windows.Forms.GroupBox
    $grp3.Text = '3   On-premises: create the shared mailbox with the list''s address'
    $grp3.Location = New-Object System.Drawing.Point(12, $y); $grp3.Size = New-Object System.Drawing.Size(796, 122)
    $form.Controls.Add($grp3)
    $l = New-Object System.Windows.Forms.Label; $l.Text = 'OU:'; $l.Location = New-Object System.Drawing.Point(12, 26); $l.AutoSize = $true; $grp3.Controls.Add($l)
    $txtOu = New-Object System.Windows.Forms.TextBox; $txtOu.Location = New-Object System.Drawing.Point(110, 22); $txtOu.Size = New-Object System.Drawing.Size(520, 24); $txtOu.Font = New-Object System.Drawing.Font('Consolas', 9)
    try { $txtOu.Text = "$(Get-TenantDirectoryOu -Kind 'SharedMailbox')" } catch { }
    $grp3.Controls.Add($txtOu)
    $l = New-Object System.Windows.Forms.Label; $l.Text = 'Account name:'; $l.Location = New-Object System.Drawing.Point(12, 56); $l.AutoSize = $true; $grp3.Controls.Add($l)
    $txtAcct = New-Object System.Windows.Forms.TextBox; $txtAcct.Location = New-Object System.Drawing.Point(110, 52); $txtAcct.Size = New-Object System.Drawing.Size(200, 24); $txtAcct.MaxLength = 20; $txtAcct.Font = New-Object System.Drawing.Font('Consolas', 9); $grp3.Controls.Add($txtAcct)
    $l = New-Object System.Windows.Forms.Label; $l.Text = '(max 20)'; $l.Location = New-Object System.Drawing.Point(316, 56); $l.AutoSize = $true; $grp3.Controls.Add($l)
    $l = New-Object System.Windows.Forms.Label; $l.Text = 'Owner (required):'; $l.Location = New-Object System.Drawing.Point(12, 86); $l.AutoSize = $true; $grp3.Controls.Add($l)
    $txtOwner = New-Object System.Windows.Forms.TextBox; $txtOwner.Location = New-Object System.Drawing.Point(110, 82); $txtOwner.Size = New-Object System.Drawing.Size(320, 24); $txtOwner.Font = New-Object System.Drawing.Font('Consolas', 9); $grp3.Controls.Add($txtOwner)
    $lblOwnerHint = New-Object System.Windows.Forms.Label; $lblOwnerHint.Location = New-Object System.Drawing.Point(436, 86); $lblOwnerHint.Size = New-Object System.Drawing.Size(200, 18); $lblOwnerHint.ForeColor = [System.Drawing.Color]::DimGray; $lblOwnerHint.Tag = 'NoTheme'; $grp3.Controls.Add($lblOwnerHint)
    $btnCreate = New-Object System.Windows.Forms.Button; $btnCreate.Text = '3  Create mailbox'; $btnCreate.Location = New-Object System.Drawing.Point(642, 80); $btnCreate.Size = New-Object System.Drawing.Size(140, 28); $btnCreate.Enabled = $false; $grp3.Controls.Add($btnCreate)
    $y += 130

    # ---- 4 stamp ----
    $grpS = New-Object System.Windows.Forms.GroupBox
    $grpS.Text = '4   On-premises: stamp attributes from the report (repeatable)'
    $grpS.Location = New-Object System.Drawing.Point(12, $y); $grpS.Size = New-Object System.Drawing.Size(796, 58)
    $form.Controls.Add($grpS)
    $lblS = New-Object System.Windows.Forms.Label; $lblS.Location = New-Object System.Drawing.Point(12, 27); $lblS.Size = New-Object System.Drawing.Size(620, 18); $lblS.ForeColor = [System.Drawing.Color]::DimGray
    $lblS.Text = 'Proxies + X500, hidden/restrictions/moderation, custom attributes, owner as Manager. Adds only what is missing.'
    $grpS.Controls.Add($lblS)
    $btnStamp = New-Object System.Windows.Forms.Button; $btnStamp.Text = '4  Stamp attributes'; $btnStamp.Location = New-Object System.Drawing.Point(642, 20); $btnStamp.Size = New-Object System.Drawing.Size(140, 28); $btnStamp.Enabled = $false; $grpS.Controls.Add($btnStamp)
    $y += 66

    # ---- 5 cloud ----
    $grp4 = New-Object System.Windows.Forms.GroupBox
    $grp4.Text = '5   Cloud: grant the former members access'
    $grp4.Location = New-Object System.Drawing.Point(12, $y); $grp4.Size = New-Object System.Drawing.Size(796, 58)
    $form.Controls.Add($grp4)
    $lbl4 = New-Object System.Windows.Forms.Label; $lbl4.Location = New-Object System.Drawing.Point(12, 27); $lbl4.Size = New-Object System.Drawing.Size(620, 18); $lbl4.ForeColor = [System.Drawing.Color]::DimGray
    $lbl4.Text = 'Members + owners: Full Access (auto-mapped) and Send As. Send on Behalf carried over. Sent-items copy on.'
    $grp4.Controls.Add($lbl4)
    $btnPerms = New-Object System.Windows.Forms.Button; $btnPerms.Text = '5  Apply permissions'; $btnPerms.Location = New-Object System.Drawing.Point(642, 20); $btnPerms.Size = New-Object System.Drawing.Size(140, 28); $btnPerms.Enabled = $false; $grp4.Controls.Add($btnPerms)
    $y += 66

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline = $true; $txtOut.ReadOnly = $true; $txtOut.ScrollBars = 'Vertical'
    $txtOut.Location = New-Object System.Drawing.Point(12, $y); $txtOut.Size = New-Object System.Drawing.Size(796, 150); $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)
    $y += 158

    $btnClose = New-Object System.Windows.Forms.Button; $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(708, $y); $btnClose.Size = New-Object System.Drawing.Size(100, 28); $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    $form | Add-Member -NotePropertyName 'Busy'       -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'Report'     -NotePropertyValue $null  -Force
    $form | Add-Member -NotePropertyName 'ReportPath' -NotePropertyValue ''     -Force
    $form | Add-Member -NotePropertyName 'Exported'   -NotePropertyValue ''     -Force   # address exported in THIS window

    $log = { param([string]$s) $txtOut.AppendText("$s`r`n") }

    $refreshConn = {
        $op = Test-OnPremReady; $cl = Test-CloudReady
        $lblConn.Text = "On-premises: $(if ($op) { 'connected' } else { 'NOT connected -- steps 1-3 need it' })     Cloud: $(if ($cl) { 'connected' } else { 'NOT connected -- step 4 needs it' })"
        $lblConn.ForeColor = $(if ($op -and $cl) { [System.Drawing.Color]::FromArgb(40, 140, 80) } else { [System.Drawing.Color]::FromArgb(200, 120, 0) })
    }

    $showReport = {
        param([string]$Path)
        $rows = @(Import-Csv -LiteralPath $Path)
        if ($rows.Count -ne 1) { throw "Expected one row in $([System.IO.Path]::GetFileName($Path)), found $($rows.Count)." }
        $r = $rows[0]
        $form.Report = $r; $form.ReportPath = $Path
        $lblRep.Text = [System.IO.Path]::GetFileName($Path)
        $owners = @(ConvertFrom-CsvList $r.ManagedBy)
        $members = @(ConvertFrom-CsvList $r.Members)
        $txtSum.Text = ("{0}  <{1}>  alias {2}`r`nOwners : {3}`r`nMembers: {4}  Send As: {5}  Send on behalf: {6}`r`nX500   : {7}" -f
            $r.DisplayName, $r.PrimarySmtpAddress, $r.Alias,
            $(if ($owners.Count) { $owners -join ', ' } else { 'NONE -- you must name one' }),
            $members.Count, @(ConvertFrom-CsvList $r.SendAs).Count, @(ConvertFrom-CsvList $r.GrantSendOnBehalfTo).Count, $r.LegacyExchangeDN)
        if (-not "$($txtAcct.Text)".Trim()) {
            # A security group is KEPT in step 2 and keeps its account name, so
            # the mailbox cannot have it. Suggest <alias>SMB from the start.
            $txtAcct.Text = $(if ("$($r.GroupType)" -match 'SecurityEnabled') { Get-AccountNameFromAlias ($r.Alias + 'SMB') } else { Get-AccountNameFromAlias $r.Alias })
        }
        if ($owners.Count) { $txtOwner.Text = $owners[0]; $lblOwnerHint.ForeColor = [System.Drawing.Color]::DimGray; $lblOwnerHint.Text = "from the list$(if ($owners.Count -gt 1) { " (+$($owners.Count - 1) more get access in step 4)" })" }
        else { $txtOwner.Text = ''; $lblOwnerHint.ForeColor = [System.Drawing.Color]::FromArgb(230, 140, 0); $lblOwnerHint.Text = 'NO OWNER on the list -- name one' }
        # Create stays locked until Check sync says the cloud has forgotten the
        # list; the address cannot be reused while it is still there.
        $btnCreate.Enabled = $false; $btnCheck.Enabled = $true; $btnPerms.Enabled = $true; $btnStamp.Enabled = $true
        $btnDelete.Enabled = ("$($form.Exported)" -ieq "$($r.PrimarySmtpAddress)")
        & $log "Report loaded: $Path"
    }

    # ---- 1 export ----
    $btnExport.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            if (-not (Test-OnPremReady)) { & $log 'On-premises Exchange is not connected in this window.'; return }
            $id = "$script:Target".Trim()
            if (-not $id) { & $log 'Enter the distribution list above and press Look Up.'; return }
            $path = Export-ListReport -Identity $id
            & $log "Exported: $path"
            Write-CompletedLog -Action 'DLtoSMBExport' -Target (Get-LogTarget $script:Target) -Result 'Success' -Detail $path
            & $showReport $path
            $form.Exported = "$($form.Report.PrimarySmtpAddress)"
            $btnDelete.Enabled = $true
            & $log 'Step 1 done. Step 2 (Delete) is unlocked for this list.'
        } catch { & $log "Export failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Export DL (on-premises)' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    # ---- 2 delete ----
    $btnDelete.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $r = $form.Report
            if (-not $r -or "$($form.Exported)" -ine "$($r.PrimarySmtpAddress)") { & $log 'Delete needs the export done in THIS window for THIS list (step 1).'; return }
            if (-not (Test-OnPremReady)) { & $log 'On-premises Exchange is not connected in this window.'; return }
            # THE GROUP'S TYPE DECIDES WHAT "DELETE" MEANS.
            #
            # A distribution group exists only to carry mail: with the address
            # moving to a mailbox it has no reason to exist, so it is deleted.
            # A mail-enabled SECURITY group may also be securing file shares,
            # apps or other permissions, and deleting it would break those
            # silently. It is only MAIL-DISABLED: the address is freed, the AD
            # group and everything it secures stay.
            $isSec = ("$($r.GroupType)" -match 'SecurityEnabled')
            $verb  = if ($isSec) { 'mail-disable' } else { 'delete' }
            $a = Show-ToolDialog(
                ($(if ($isSec) { "MAIL-DISABLE the on-premises security group? (It is a mail-enabled SECURITY group, so it is kept -- only its mail attributes are removed.)" }
                   else        { "DELETE the on-premises distribution group?" }) +
                 "`r`n`r`n  $($r.DisplayName)  <$($r.PrimarySmtpAddress)>`r`n  $($r.GroupType); $($r.MemberCount) member(s)`r`n`r`nThe export is at:`r`n  $($form.ReportPath)`r`n`r`nThe cloud copy goes on the next sync. Then Check sync before step 3."),
                "Confirm $verb", 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            if ($isSec) {
                # Learn (Disable-DistributionGroup): -Identity; removes the
                # Exchange attributes and leaves the AD group. Same manager
                # check as Remove-, same bypass.
                onprem 'Disable-DistributionGroup' @{ Identity = "$($r.Guid)"; BypassSecurityGroupManagerCheck = $true; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                Write-CompletedLog -Action 'DLtoSMBDisable' -Target (Get-LogTarget $r.PrimarySmtpAddress) -Result 'Success' -Detail "Disable-DistributionGroup (mail-enabled security group kept); export $([System.IO.Path]::GetFileName($form.ReportPath))"
                & $log "Mail-disabled on-premises (security group kept): $($r.PrimarySmtpAddress). WAIT for a directory sync (30 min), then Check sync."
            } else {
                # Learn (Remove-DistributionGroup): -Identity; -Confirm:$false;
                # -BypassSecurityGroupManagerCheck -- without it Exchange refuses
                # anyone who is not in the list's ManagedBy, admin or not.
                onprem 'Remove-DistributionGroup' @{ Identity = "$($r.Guid)"; BypassSecurityGroupManagerCheck = $true; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                Write-CompletedLog -Action 'DLtoSMBDelete' -Target (Get-LogTarget $r.PrimarySmtpAddress) -Result 'Success' -Detail "Remove-DistributionGroup; export $([System.IO.Path]::GetFileName($form.ReportPath))"
                & $log "Deleted on-premises: $($r.PrimarySmtpAddress). WAIT for a directory sync (30 min), then Check sync."
            }
            $btnDelete.Enabled = $false
        } catch { & $log "Step 2 failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Remove-/Disable-DistributionGroup'; Write-FailedLog -Action 'DLtoSMBDelete' -Target (Get-LogTarget $script:Target) -ErrorRecord $_ }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    # ---- check sync (cloud) ----
    # One Get-Recipient on the list's address. Three answers: the list is
    # still there (wait), nothing is there (Create unlocks), or the NEW
    # mailbox is there (step 3 has synced; step 4 is ready).
    $btnCheck.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $r = $form.Report
            if (-not $r) { & $log 'Load the report first.'; return }
            if (-not (Test-CloudReady)) { & $log 'Exchange Online is not connected in this window.'; return }
            $smtp = "$($r.PrimarySmtpAddress)"
            $rcp = $null
            try { $rcp = exo 'Get-Recipient' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
            $stamp = Get-Date -Format 'HH:mm:ss'
            if (-not $rcp) {
                & $log "[$stamp] Object has been deleted from the cloud: $smtp is free. Step 3 (Create) is unlocked."
                $btnCreate.Enabled = $true
            } elseif ("$($rcp.RecipientTypeDetails)" -match 'Mailbox') {
                & $log "[$stamp] The NEW mailbox has synced: $smtp is a $($rcp.RecipientTypeDetails). Step 5 (Apply permissions) is ready."
                $btnCreate.Enabled = $false
            } else {
                & $log "[$stamp] Object still exists in the cloud: $smtp is a $($rcp.RecipientTypeDetails). Wait for the sync and check again."
                $btnCreate.Enabled = $false
            }
        } catch { & $log "Check failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Check sync (Get-Recipient, cloud)' }
        finally { $form.Busy = $false }
    })

    # ---- load report ----
    $btnLoad.Add_Click({
        try {
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.InitialDirectory = Get-ReportPath; $dlg.Filter = 'DL to SMB reports (DLtoSMB_*.csv)|DLtoSMB_*.csv|All CSV|*.csv'
            $newest = Find-NewestReport -Smtp $script:Target
            if ($newest) { $dlg.FileName = $newest.Name }
            if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            & $showReport $dlg.FileName
        } catch { & $log "Load failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Load report' }
    })

    # ---- 3 create ----
    $btnCreate.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $r = $form.Report
            if (-not $r) { & $log 'Load the report first.'; return }
            if (-not (Test-OnPremReady)) { & $log 'On-premises Exchange is not connected in this window.'; return }
            $ou = "$($txtOu.Text)".Trim(); $acct = "$($txtAcct.Text)".Trim(); $ownerIn = "$($txtOwner.Text)".Trim()
            if (-not $ou) { & $log 'OU is required (Tenant Manager > Shared mbx OU, or type it).'; return }
            if (-not $acct -or $acct.Length -gt 20 -or $acct.EndsWith('.')) { & $log 'Account name: 1-20 characters, not ending in a period.'; return }
            if (-not $ownerIn) { & $log 'OWNER IS REQUIRED. Name the person responsible for this mailbox.'; $txtOwner.Focus(); return }
            $owner = $null
            try { $owner = onprem 'Get-Recipient' @{ Identity = $ownerIn; ErrorAction = 'Stop' } } catch { }
            if (-not $owner) { & $log "Owner '$ownerIn' does not resolve on-premises."; $txtOwner.Focus(); return }
            $smtp = "$($r.PrimarySmtpAddress)"
            # The address must be free: the list gone, nothing else holding it.
            $holder = $null
            try { $holder = onprem 'Get-Recipient' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
            if ($holder) { & $log "REFUSED: '$smtp' is still held on-premises by $($holder.RecipientTypeDetails) '$($holder.Name)'. Delete the list (step 2) first."; return }
            $holder = Find-AccountNameHolder -Sam $acct
            if ($holder) {
                & $log "Account name '$acct' is used by $holder -- asking for another (alias and address are unaffected)."
                $acct = Read-AccountName -Suggested $acct -TakenBy $holder
                if (-not $acct) { & $log 'Cancelled. Nothing created.'; return }
                $txtAcct.Text = $acct
            }
            $ouObj = @()
            try { $ouObj = @(onprem 'Get-OrganizationalUnit' @{ Identity = $ou; ErrorAction = 'Stop' }) } catch { }
            if ($ouObj.Count -eq 0) { & $log "REFUSED: OU not found: $ou  (canonical form, e.g. domain.com/Accounts/Resource Mailbox)"; return }

            $a = Show-ToolDialog(
                ("Create the on-premises shared mailbox?`r`n`r`n  $($r.DisplayName)  <$smtp>`r`n  account: $acct`r`n  OU: $ou`r`n  owner: $($owner.DisplayName) <$($owner.PrimarySmtpAddress)>`r`n`r`nThen: proxies, X500, restrictions, custom attributes from the report."),
                'Confirm create', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            # Learn (New-RemoteMailbox): Shared set -- -Shared, -Name, -Alias,
            # -DisplayName, -SamAccountName (<=20), -PrimarySmtpAddress,
            # -OnPremisesOrganizationalUnit; Password/UPN not required.
            $p = @{
                Name = "$($r.DisplayName)"; DisplayName = "$($r.DisplayName)"; Alias = "$($r.Alias)"
                SamAccountName = $acct; PrimarySmtpAddress = $smtp; Shared = $true
                OnPremisesOrganizationalUnit = $ou; ErrorAction = 'Stop'
            }
            # Pin the create to a writable DC in the OU's domain. Without this
            # the create is routed by Exchange (which works) but the returned
            # object reports whichever server answered, and a session parked
            # on a child-domain DC then aims the follow-up writes there.
            $ouDc = Get-WriteDomainController -DistinguishedName "$($ouObj[0].DistinguishedName)" -Fallback ''
            if ($ouDc) { $p.DomainController = $ouDc; & $log "Creating via $ouDc (the OU's domain)." }
            if ($DomainController) { $p.DomainController = $DomainController }
            $rm = onprem 'New-RemoteMailbox' $p
            & $log "Created remote shared mailbox: $($rm.PrimarySmtpAddress)  (account $acct)"
            Write-CompletedLog -Action 'DLtoSMBCreate' -Target (Get-LogTarget $smtp) -Result 'Success' -Detail "New-RemoteMailbox -Shared in $ou; account $acct; owner $($owner.PrimarySmtpAddress)"

            # Stamp straight away, pinned to the DC that just created the
            # object -- and step 4 can repeat it if anything failed.
            Invoke-AttributeStamp -Report $r -RemoteMailbox $rm -OwnerSmtp "$($owner.PrimarySmtpAddress)" -Log $log
            $btnStamp.Enabled = $true
            $btnCreate.Enabled = $false
            & $log 'Step 3 done. If any stamp line above says FAIL, press 4 Stamp attributes (safe to repeat). Then WAIT for a sync; Check sync until it reports the new mailbox; then step 5.'
        } catch { & $log "Create failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'New-RemoteMailbox'; Write-FailedLog -Action 'DLtoSMBCreate' -Target (Get-LogTarget $script:Target) -ErrorRecord $_ }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    # ---- 4 stamp (on-premises, repeatable) ----
    $btnStamp.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $r = $form.Report
            if (-not $r) { & $log 'Load the report first.'; return }
            if (-not (Test-OnPremReady)) { & $log 'On-premises Exchange is not connected in this window.'; return }
            $smtp = "$($r.PrimarySmtpAddress)"
            # FIND IT WHEREVER IT IS. The session may be scoped to one domain
            # of the forest (that is how step 3's stamp missed an object it had
            # just created): widen to the whole forest first, then, if the
            # object still does not answer, ask each configured domain
            # controller directly.
            $rm = $null
            try { onprem 'Set-ADServerSettings' @{ ViewEntireForest = $true; ErrorAction = 'Stop' } | Out-Null } catch { }
            try { $rm = onprem 'Get-RemoteMailbox' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
            if (-not $rm) {
                foreach ($dc in @(Get-TenantDomainController)) {
                    try { $rm = onprem 'Get-RemoteMailbox' @{ Identity = $smtp; DomainController = $dc; ErrorAction = 'Stop' } } catch { }
                    if ($rm) { & $log "Found the remote mailbox via $dc."; break }
                }
            }
            if (-not $rm) { & $log "No remote mailbox at '$smtp' on-premises yet -- step 3 not run, or not replicated. Wait a minute and try again."; return }
            $ownerSmtp = ''
            $ownerIn = "$($txtOwner.Text)".Trim()
            if ($ownerIn) { try { $ownerSmtp = "$((onprem 'Get-Recipient' @{ Identity = $ownerIn; ErrorAction = 'Stop' }).PrimarySmtpAddress)" } catch { & $log "Owner '$ownerIn' does not resolve -- Manager will not be set." } }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $clean = Invoke-AttributeStamp -Report $r -RemoteMailbox $rm -OwnerSmtp $ownerSmtp -Log $log
            & $log $(if ($clean) { 'Step 4 done. WAIT for a sync; Check sync until it reports the new mailbox; then step 5.' } else { 'Step 4 finished with failures -- read the lines above; this button is safe to press again.' })
        } catch { & $log "Stamp failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Set-RemoteMailbox' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    # ---- 5 permissions (cloud) ----
    $btnPerms.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $r = $form.Report
            if (-not $r) { & $log 'Load the report first.'; return }
            if (-not (Test-CloudReady)) { & $log 'Exchange Online is not connected in this window.'; return }
            $smtp = "$($r.PrimarySmtpAddress)"
            $mb = $null
            try { $mb = exo 'Get-Mailbox' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
            if (-not $mb) { & $log "No cloud mailbox at '$smtp' yet. Step 3 done? Wait for a sync and try again."; return }
            if ("$($mb.RecipientTypeDetails)" -ne 'SharedMailbox') { & $log "'$smtp' is a $($mb.RecipientTypeDetails), not a SharedMailbox -- stopping." ; return }

            $owners  = @(ConvertFrom-CsvList $r.ManagedBy)
            $ownerIn = "$($txtOwner.Text)".Trim(); if ($ownerIn -and $owners -notcontains $ownerIn) { $owners += $ownerIn }
            $members = @(ConvertFrom-CsvList $r.Members)
            $sendAs  = @(ConvertFrom-CsvList $r.SendAs)
            $sob     = @(ConvertFrom-CsvList $r.GrantSendOnBehalfTo)
            $grant   = @(($members + $owners + $sendAs) | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Select-Object -Unique)
            if ($grant.Count -eq 0) { & $log 'Nothing to grant: the list had no members, owners or Send As.'; return }

            $a = Show-ToolDialog(
                ("Grant on $smtp (cloud):`r`n`r`n  Full Access (auto-mapped) + Send As to $($grant.Count) recipient(s)`r`n  Send on Behalf to $($sob.Count)`r`n  Sent-items copy: Send As and Send on Behalf`r`n`r`nAlready-present grants are left alone."),
                'Confirm permissions', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            $curFA = @(); $curSA = @()
            try { $curFA = @(exo 'Get-MailboxPermission' @{ Identity = $smtp; ErrorAction = 'Stop' } | Where-Object { -not $_.IsInherited -and $_.AccessRights -contains 'FullAccess' } | ForEach-Object { "$($_.User)".ToLower() }) } catch { }
            try { $curSA = @(exo 'Get-RecipientPermission' @{ Identity = $smtp; ErrorAction = 'Stop' } | Where-Object { $_.AccessRights -contains 'SendAs' } | ForEach-Object { "$($_.Trustee)".ToLower() }) } catch { }
            $ok = 0; $skip = 0; $fail = 0
            foreach ($who in $grant) {
                $rcp = $null
                try { $rcp = exo 'Get-Recipient' @{ Identity = $who; ErrorAction = 'Stop' } } catch { }
                if (-not $rcp) { & $log "  skip  $who -- not found in the cloud"; $skip++; continue }
                if ("$($rcp.RecipientTypeDetails)" -match 'DistributionGroup|DynamicDistributionGroup' -and "$($rcp.RecipientTypeDetails)" -notmatch 'Security') {
                    & $log "  skip  $who -- a distribution list cannot hold mailbox permissions (mail-enabled SECURITY groups can)"; $skip++; continue
                }
                $addr = "$($rcp.PrimarySmtpAddress)"
                try {
                    if ($curFA -contains $addr.ToLower() -or $curFA -contains "$($rcp.Name)".ToLower()) { & $log "  has   FullAccess  $addr" }
                    else {
                        # Learn (Add-MailboxPermission): -AccessRights FullAccess,
                        # -InheritanceType All, -AutoMapping.
                        exo 'Add-MailboxPermission' @{ Identity = $smtp; User = $addr; AccessRights = 'FullAccess'; InheritanceType = 'All'; AutoMapping = $true; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        & $log "  add   FullAccess  $addr"
                    }
                    if ($curSA -contains $addr.ToLower() -or $curSA -contains "$($rcp.Name)".ToLower()) { & $log "  has   SendAs      $addr" }
                    else {
                        # Learn (Add-RecipientPermission): -Trustee, -AccessRights SendAs.
                        exo 'Add-RecipientPermission' @{ Identity = $smtp; Trustee = $addr; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        & $log "  add   SendAs      $addr"
                    }
                    $ok++
                } catch { & $log "  FAIL  $addr -- $($_.Exception.Message)"; $fail++ }
            }
            # Send on Behalf + sent-items copies, one Set-Mailbox.
            # Learn (Set-Mailbox): -GrantSendOnBehalfTo @{Add=},
            # -MessageCopyForSentAsEnabled, -MessageCopyForSendOnBehalfEnabled.
            $set = @{ Identity = $smtp; MessageCopyForSentAsEnabled = $true; MessageCopyForSendOnBehalfEnabled = $true; ErrorAction = 'Stop' }
            $sobOk = @($sob | Where-Object { try { [void](exo 'Get-Recipient' @{ Identity = $_; ErrorAction = 'Stop' }); $true } catch { $false } })
            if ($sobOk.Count) { $set.GrantSendOnBehalfTo = @{ Add = $sobOk } }
            try { exo 'Set-Mailbox' $set | Out-Null; & $log "Set: sent-items copy on (Send As + Send on Behalf)$(if ($sobOk.Count) { "; Send on Behalf +$($sobOk.Count)" })" }
            catch { & $log "Set-Mailbox failed: $($_.Exception.Message)"; $fail++ }

            Write-CompletedLog -Action 'DLtoSMBPermissions' -Target (Get-LogTarget $smtp) -Result $(if ($fail -and $ok) { 'Partial' } elseif ($fail) { 'Failed' } else { 'Success' }) `
                -Detail "FullAccess+SendAs for $ok, $skip skipped, $fail failed; SendOnBehalf $($sobOk.Count); sent-items copy on; automap on"
            & $log "Step 5 done: $ok granted, $skip skipped, $fail failed. The conversion is complete."
        } catch { & $log "Permissions failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Cloud permissions'; Write-FailedLog -Action 'DLtoSMBPermissions' -Target (Get-LogTarget $script:Target) -ErrorRecord $_ }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    $bar = Add-ToolTargetBar -Form $form -Label 'Distribution list:' -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        $form.Text = "DL to Shared Mailbox: $script:Target"
        & $refreshConn
        $n = Find-NewestReport -Smtp $script:Target
        if ($n) { try { & $showReport $n.FullName } catch { & $log "Newest report could not be read: $($_.Exception.Message)" } }
        else { & $log "No report yet for $script:Target -- run step 1." }
    })
    $form.Add_Shown({ & $refreshConn; if ("$script:Target".Trim()) { $bar.Button.PerformClick() } })

    Set-ConsoleTheme -Form $form
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL to Shared Mailbox' | Out-Null
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    [void]$form.ShowDialog()
}
#endregion

Show-ConvertForm
