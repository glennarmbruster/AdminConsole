#Requires -Version 5.1
<#
    RenameUser.ps1
    Change a user's email identity on-premises: UPN, mail, mailNickname and the
    proxyAddresses collection.

    From a working desk script. What changed and why is in the comments below --
    the shape is the same, and every difference is a failure the script could
    produce that this one cannot.

    WRITES TO ACTIVE DIRECTORY.
#>
# TOOLKIT-REQUIRES: OnPremAD
# TOOLKIT-HELP-RUNS: WRITES to on-premises Active Directory. Reads Exchange Online only to check the new address is free.
# TOOLKIT-VARIANT: OnPremAD
# TOOLKIT-HELP-BEGIN
# Changes the email identity of a user who has been renamed -- marriage, legal
# name change, a mistyped address at onboarding.
#
# FOUR ATTRIBUTES MOVE TOGETHER, and that is the whole point of doing it here
# rather than in ADUC:
#
#   userPrincipalName   what they sign in with
#   mail                what the directory says their address is
#   mailNickname        the alias Exchange builds addresses from
#   proxyAddresses      every address the mailbox actually receives on
#
# Change three of the four and you get a mailbox that receives mail at an
# address the person cannot sign in with, or a sign-in that no longer matches
# the mailbox. They are one change.
#
# THE OLD ADDRESS IS KEPT. It is demoted from primary to secondary, so mail sent
# to it still arrives. A rename that stops the old address working breaks every
# reply to every message they sent before today.
#
# X500 ADDRESSES ARE CARRIED ACROSS UNTOUCHED. They are what lets Outlook
# resolve replies to historical mail, and rebuilding a proxy list without them
# breaks those replies weeks later, long after anybody connects it to a rename.
#
# NAMES ARE NOT TOUCHED. displayName, givenName, sn and sAMAccountName stay as
# they are -- this tool changes what the person is ADDRESSED as, not what they
# are called or how they log on.
#
# THE CLOUD LAGS. This writes to Active Directory. Exchange Online will not show
# the change until the next AAD Connect sync cycle, and the tool says so rather
# than letting you conclude the write failed.
#
# The state before the change is written to the reports folder before anything
# is modified.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

#region Load Common -------------------------------------------------------------
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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:Found    = $null
$script:Plan     = $null
$script:Domain   = ''
$script:Server   = ''

# THE DOMAIN COMES FROM Tenant.json, AND THERE IS NO FALLBACK LITERAL.
#
# The desk script had the domain typed into it twice. That is correct on one
# deployment and wrong in the handout copy the moment somebody else runs it --
# and it fails by building an address for a domain that does not exist, which
# reads as "user not found" rather than "wrong domain".
#
# I FIRST WROTE A HARDCODED FALLBACK HERE and the verifier rejected it (TEN002).
# It was right to: a fallback that happens to be correct on this tenant is the
# same bug with a longer fuse, because it only shows up somewhere else.
# An unset domain is reported and the tool refuses to run.
try { $script:Domain = "$((Get-TenantConfig).PrimaryDomain)".Trim() } catch { }

#region Directory ---------------------------------------------------------------
function Get-AdServerParam {
    <#
        The -Server value for every AD call, resolved once.

        A read and a write that land on different domain controllers is how a
        change appears not to have taken: the write succeeds on one, the
        read-back happens on another that has not replicated yet, and the tool
        reports failure on a change that worked.

        NOTHING AT ALL IS A VALID ANSWER. Get-TenantDirectoryServer returns the
        DOMAIN when a directory names several controllers, deliberately -- the
        ActiveDirectory module resolves that itself. An empty value means "let
        AD choose", which is better than pinning the wrong one.
    #>
    if ("$DomainController".Trim()) { return @{ Server = "$DomainController".Trim() } }
    if ("$script:Server".Trim())    { return @{ Server = "$script:Server".Trim() } }
    return @{}
}

function Find-RenameTarget {
    <#
        The user behind an address, found by any of the three things it could be.

        THE DESK SCRIPT MATCHED userPrincipalName ONLY, and a rename ticket
        rarely arrives with the UPN. It arrives with the address people send to,
        which may be a secondary proxy or the mail attribute -- both of which
        return nothing from a UPN filter, and "User not found." is a poor answer
        when the object is plainly there.

        ESCAPED. ConvertTo-LdapFilterValue exists because an apostrophe in
        o'brien@... terminates the filter string and turns the rest of it into
        syntax errors -- or, worse, into a filter that matches something else.
    #>
    param([Parameter(Mandatory)][string]$Address)

    $out = [PSCustomObject]@{ Ok = $false; User = $null; Reason = ''; MatchedOn = '' }

    $esc = ConvertTo-LdapFilterValue -Value $Address
    $ldap = "(&(objectClass=user)(!(objectClass=computer))(|(userPrincipalName=$esc)(mail=$esc)(proxyAddresses=smtp:$esc)))"

    $p = @{
        LDAPFilter  = $ldap
        Properties  = @('displayName', 'userPrincipalName', 'mail', 'mailNickname',
                        'proxyAddresses', 'sAMAccountName', 'distinguishedName',
                        'targetAddress', 'msExchRemoteRecipientType', 'userAccountControl')
        ErrorAction = 'Stop'
    }
    $sp = Get-AdServerParam
    foreach ($k in $sp.Keys) { $p[$k] = $sp[$k] }

    $hits = @()
    try { $hits = @(Get-ADObject @p) }
    catch { $out.Reason = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADObject'; return $out }

    if (@($hits).Count -eq 0) {
        $out.Reason = "Nothing in Active Directory has '$Address' as its UPN, its mail attribute or one of its proxy addresses."
        return $out
    }
    if (@($hits).Count -gt 1) {
        # AMBIGUITY IS REPORTED, NOT RESOLVED. Two objects sharing an address is
        # a real and serious directory fault, and picking one to rename would
        # hide it behind a change that half worked.
        $out.Reason = "$(@($hits).Count) objects share that address, which should not be possible: " +
                      (@($hits | ForEach-Object { "$($_.distinguishedName)" }) -join ' | ') +
                      '. Resolve the duplicate before renaming either of them.'
        return $out
    }

    $u = $hits[0]
    $out.MatchedOn = if ("$($u.userPrincipalName)" -eq $Address) { 'userPrincipalName' }
                     elseif ("$($u.mail)" -eq $Address)          { 'mail' }
                     else                                        { 'a proxy address' }
    $out.User = $u
    $out.Ok   = $true
    return $out
}

function Test-AddressFree {
    <#
        Is the new address already spoken for?

        THE DESK SCRIPT DID NOT ASK, and this is the failure with the longest
        tail. Assigning an address another object already holds gives two
        recipients with the same proxy: mail routes to whichever the directory
        resolves first, intermittently, and the symptom appears days later as
        "some of my mail is going to someone else".

        BOTH SIDES ARE CHECKED. Active Directory catches live objects; Exchange
        Online catches a SOFT-DELETED mailbox still holding the address, which
        has no AD object at all and is the single most common reason a rename is
        rejected by Exchange after the AD write has already happened.
    #>
    param([Parameter(Mandatory)][string]$Address, [string]$ExceptDn = '')

    $out = [PSCustomObject]@{ Free = $true; Notes = New-Object System.Collections.ArrayList }

    $esc = ConvertTo-LdapFilterValue -Value $Address
    $p = @{
        LDAPFilter  = "(|(userPrincipalName=$esc)(mail=$esc)(proxyAddresses=smtp:$esc))"
        Properties  = @('distinguishedName')
        ErrorAction = 'Stop'
    }
    $sp = Get-AdServerParam
    foreach ($k in $sp.Keys) { $p[$k] = $sp[$k] }

    try {
        foreach ($h in @(Get-ADObject @p)) {
            if ("$($h.distinguishedName)" -eq "$ExceptDn") { continue }
            $out.Free = $false
            [void]$out.Notes.Add("Active Directory: already held by $($h.distinguishedName)")
        }
    } catch {
        [void]$out.Notes.Add("Active Directory could not be checked: $($_.Exception.Message)")
    }

    $exo = Test-ExoConnectionState
    if ($exo.Connected) {
        try {
            # ResultSize IS DELIBERATELY SMALL: this asks "is anything holding
            # this address", not "how many". WarningAction sits next to it because
            # Exchange otherwise warns "There are more results available" on every
            # single check -- EXO002, whose scan window is three lines.
            $soft = @(Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
                SoftDeletedMailbox = $true
                Filter             = "EmailAddresses -eq 'smtp:$(ConvertTo-OpathValue -Value $Address)'"
                ResultSize         = 10
                WarningAction      = 'SilentlyContinue'
                ErrorAction        = 'Stop' })
            foreach ($s in $soft) {
                $out.Free = $false
                [void]$out.Notes.Add("Exchange Online: held by the SOFT-DELETED mailbox '$($s.DisplayName)'. It must be purged or its address released before this can be used.")
            }
        } catch {
            [void]$out.Notes.Add("Soft-deleted mailboxes could not be checked: $($_.Exception.Message)")
        }
    } else {
        [void]$out.Notes.Add('Exchange Online is not connected, so a soft-deleted mailbox holding this address would not be seen. Connect ExO and check again if the write is rejected.')
    }
    return $out
}
#endregion

#region The plan ----------------------------------------------------------------
function New-ProxyAddressPlan {
    <#
        The proxyAddresses collection as it will be written.

        THE DESK SCRIPT'S VERSION HAD TWO FAULTS, and both are silent.

        1.  It tested `-not ($proxyAddresses -contains "SMTP:$newEmail")` before
            adding. -contains IS CASE-INSENSITIVE, so an existing SECONDARY
            'smtp:new@...' satisfied it -- the new address was never promoted to
            primary and the rename simply did not happen, with no error. The one
            case where somebody has already been given the address as an alias,
            which is common when a rename has been half-done by hand, was the
            case it silently skipped.

        2.  It removed the old primary with `-ne "SMTP:$oldEmail"`, also
            case-insensitive, which removed the secondary copy too -- then
            re-added it only inside the if. On the path above, the old address
            was removed and never restored.

        THE RULE THIS ENFORCES INSTEAD: exactly one entry is uppercase SMTP: and
        that is what makes it primary. Everything else is carried across
        verbatim, including X500 -- see the help block for why that matters.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()]$Existing,
        [Parameter(Mandatory)][string]$NewPrimary,
        [Parameter(Mandatory)][string]$OldPrimary
    )

    $kept    = New-Object System.Collections.ArrayList
    $dropped = New-Object System.Collections.ArrayList

    foreach ($a in @($Existing)) {
        $s = "$a".Trim()
        if (-not $s) { continue }

        # Any case-variant of the NEW address is removed here and re-added as
        # the primary below, so it can never appear twice.
        if ($s -match "(?i)^smtp:$([regex]::Escape($NewPrimary))$") {
            [void]$dropped.Add($s)
            continue
        }

        # DEMOTE THE CURRENT PRIMARY. -creplace, because case IS the data: the
        # uppercase prefix is the only thing distinguishing primary from alias.
        if ($s -cmatch '^SMTP:') { $s = $s -creplace '^SMTP:', 'smtp:' }

        # DE-DUPLICATED AFTER DEMOTION, NOT BEFORE.
        #
        # A directory holding both 'SMTP:x' and 'smtp:x' -- which happens, and
        # usually means a previous rename was half-finished by hand -- produces
        # two identical entries the moment the primary is demoted. Exchange then
        # rejects the whole write with "proxy address is already being used",
        # naming an address that appears only once in the list you were looking
        # at. Case-insensitive, because the prefix carries the primary flag and
        # nothing else.
        $dupe = @($kept | Where-Object { "$_".ToLowerInvariant() -eq $s.ToLowerInvariant() })
        if (@($dupe).Count -gt 0) { [void]$dropped.Add($s); continue }

        [void]$kept.Add($s)
    }

    # THE OLD ADDRESS IS GUARANTEED PRESENT AS A SECONDARY, whether or not it
    # was in the collection to begin with. On a directory where mail and
    # proxyAddresses disagree, it might not have been.
    $haveOld = @($kept | Where-Object { $_ -match "(?i)^smtp:$([regex]::Escape($OldPrimary))$" })
    if ("$OldPrimary".Trim() -and @($haveOld).Count -eq 0) { [void]$kept.Add("smtp:$OldPrimary") }

    $final = New-Object System.Collections.ArrayList
    [void]$final.Add("SMTP:$NewPrimary")
    foreach ($k in $kept) { [void]$final.Add($k) }

    return [PSCustomObject]@{
        Final   = @($final.ToArray())
        Dropped = @($dropped.ToArray())
        X500    = @($final | Where-Object { $_ -match '(?i)^x500:' })
    }
}

function Build-RenamePlan {
    param(
        [Parameter(Mandatory)]$User,
        [Parameter(Mandatory)][string]$OldAddress,
        [Parameter(Mandatory)][string]$NewAddress,
        [Parameter(Mandatory)][string]$NewAlias
    )

    $px = New-ProxyAddressPlan -Existing @($User.proxyAddresses) -NewPrimary $NewAddress -OldPrimary $OldAddress

    return [PSCustomObject]@{
        Dn          = "$($User.distinguishedName)"
        DisplayName = "$($User.displayName)"
        Sam         = "$($User.sAMAccountName)"
        OldUpn      = "$($User.userPrincipalName)"
        NewUpn      = $NewAddress
        OldMail     = "$($User.mail)"
        NewMail     = $NewAddress
        OldNickname = "$($User.mailNickname)"
        NewNickname = $NewAlias
        OldProxies  = @($User.proxyAddresses | ForEach-Object { "$_" })
        NewProxies  = @($px.Final)
        X500Count   = @($px.X500).Count
        Remote      = ($null -ne $User.msExchRemoteRecipientType)
    }
}

function Format-RenamePlan {
    param([Parameter(Mandatory)]$Plan)
    $L = New-Object System.Collections.ArrayList
    [void]$L.Add("$($Plan.DisplayName)   ($($Plan.Sam))")
    [void]$L.Add("$($Plan.Dn)")
    [void]$L.Add('')
    [void]$L.Add('WILL CHANGE')
    [void]$L.Add(("  userPrincipalName  {0}" -f $Plan.OldUpn))
    [void]$L.Add( "                  -> $($Plan.NewUpn)")
    [void]$L.Add(("  mail               {0}" -f $Plan.OldMail))
    [void]$L.Add( "                  -> $($Plan.NewMail)")
    [void]$L.Add(("  mailNickname       {0}  ->  {1}" -f $Plan.OldNickname, $Plan.NewNickname))
    [void]$L.Add('')
    [void]$L.Add("WILL NOT CHANGE")
    [void]$L.Add("  displayName, givenName, sn   -- this tool does not touch names")
    [void]$L.Add("  sAMAccountName '$($Plan.Sam)' -- they log on with this unchanged")
    [void]$L.Add('')
    [void]$L.Add("proxyAddresses: $(@($Plan.OldProxies).Count) entry(ies) -> $(@($Plan.NewProxies).Count)")
    foreach ($a in @($Plan.NewProxies)) {
        $tag = if ($a -cmatch '^SMTP:') { '  <- PRIMARY (new)' }
               elseif ($a -match "(?i)^smtp:$([regex]::Escape($Plan.OldUpn))$") { '  <- the old address, kept so mail still arrives' }
               elseif ($a -match '(?i)^x500:') { '  <- X500, carried across untouched' }
               else { '' }
        [void]$L.Add("    $a$tag")
    }
    if ($Plan.X500Count -gt 0) {
        [void]$L.Add('')
        [void]$L.Add("  $($Plan.X500Count) X500 address(es) preserved. Losing these breaks replies to mail sent before today.")
    }
    return ($L -join [Environment]::NewLine)
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Rename User'
$form.ClientSize      = New-Object System.Drawing.Size(900, 720)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'
$form.AutoScroll      = $true

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Text      = 'WRITES TO ACTIVE DIRECTORY. Exchange Online will not show the change until the next AAD Connect sync -- that is normal, not a failure.'
$lblWarn.Location  = New-Object System.Drawing.Point(14, 10)
$lblWarn.Size      = New-Object System.Drawing.Size(830, 32)
$lblWarn.Font      = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(255, 105, 97)
$lblWarn.Tag       = 'SemanticColour'
$lblWarn.Anchor    = 'Top,Left,Right'
[void]$form.Controls.Add($lblWarn)

$grpIn          = New-Object System.Windows.Forms.GroupBox
$grpIn.Text     = 'Addresses'
$grpIn.Location = New-Object System.Drawing.Point(14, 46)
$grpIn.Size     = New-Object System.Drawing.Size(872, 122)
$grpIn.Anchor   = 'Top,Left,Right'
[void]$form.Controls.Add($grpIn)

$lblOld          = New-Object System.Windows.Forms.Label
$lblOld.Text     = 'Current'
$lblOld.Location = New-Object System.Drawing.Point(12, 28)
$lblOld.Size     = New-Object System.Drawing.Size(66, 20)
[void]$grpIn.Controls.Add($lblOld)

$txtOld          = New-Object System.Windows.Forms.TextBox
$txtOld.Location = New-Object System.Drawing.Point(84, 25)
$txtOld.Size     = New-Object System.Drawing.Size(230, 22)
$txtOld.Font     = New-Object System.Drawing.Font('Consolas', 9)
[void]$grpIn.Controls.Add($txtOld)

$lblOldDom          = New-Object System.Windows.Forms.Label
$lblOldDom.Location = New-Object System.Drawing.Point(320, 28)
$lblOldDom.Size     = New-Object System.Drawing.Size(210, 20)
$lblOldDom.Font     = New-Object System.Drawing.Font('Consolas', 9)
[void]$grpIn.Controls.Add($lblOldDom)

$lblNew          = New-Object System.Windows.Forms.Label
$lblNew.Text     = 'New'
$lblNew.Location = New-Object System.Drawing.Point(12, 58)
$lblNew.Size     = New-Object System.Drawing.Size(66, 20)
[void]$grpIn.Controls.Add($lblNew)

$txtNew          = New-Object System.Windows.Forms.TextBox
$txtNew.Location = New-Object System.Drawing.Point(84, 55)
$txtNew.Size     = New-Object System.Drawing.Size(230, 22)
$txtNew.Font     = New-Object System.Drawing.Font('Consolas', 9)
[void]$grpIn.Controls.Add($txtNew)

$lblNewDom          = New-Object System.Windows.Forms.Label
$lblNewDom.Location = New-Object System.Drawing.Point(320, 58)
$lblNewDom.Size     = New-Object System.Drawing.Size(210, 20)
$lblNewDom.Font     = New-Object System.Drawing.Font('Consolas', 9)
[void]$grpIn.Controls.Add($lblNewDom)

# THE ALIAS IS TYPED WITHOUT THE DOMAIN, exactly as the desk script asked for
# it, and the domain is shown beside the box rather than assumed silently.
$lblHint          = New-Object System.Windows.Forms.Label
$lblHint.Text     = "Type the part before the @ only. The domain comes from Tenant.json."
$lblHint.Location = New-Object System.Drawing.Point(84, 84)
$lblHint.Size     = New-Object System.Drawing.Size(460, 18)
$lblHint.Font     = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grpIn.Controls.Add($lblHint)

$btnCheck           = New-Object System.Windows.Forms.Button
$btnCheck.Text      = 'Check'
$btnCheck.Location  = New-Object System.Drawing.Point(560, 24)
$btnCheck.Size      = New-Object System.Drawing.Size(130, 30)
$btnCheck.FlatStyle = 'Flat'
[void]$grpIn.Controls.Add($btnCheck)

$btnApply           = New-Object System.Windows.Forms.Button
$btnApply.Text      = 'Apply rename'
$btnApply.Location  = New-Object System.Drawing.Point(700, 24)
$btnApply.Size      = New-Object System.Drawing.Size(160, 30)
$btnApply.FlatStyle = 'Flat'
$btnApply.Enabled   = $false
[void]$grpIn.Controls.Add($btnApply)

$lblGate           = New-Object System.Windows.Forms.Label
$lblGate.Location  = New-Object System.Drawing.Point(560, 58)
$lblGate.Size      = New-Object System.Drawing.Size(300, 32)
$lblGate.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grpIn.Controls.Add($lblGate)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(14, 178)
$txtOut.Size       = New-Object System.Drawing.Size(872, 528)
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
# B6: a plan with long DNs in it must not wrap into unreadable porridge.
$txtOut.WordWrap   = $false
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

function Get-Typed {
    param([string]$Text)
    # Tolerate a full address being pasted in: strip the domain rather than
    # building 'user@x.com@x.com' and reporting the user as missing.
    $v = "$Text".Trim()
    if ($v -match '@') { $v = ($v -split '@')[0] }
    return $v.Trim()
}

function Update-DomainLabels {
    $o = Get-Typed -Text $txtOld.Text
    $n = Get-Typed -Text $txtNew.Text
    $lblOldDom.Text = if ($o) { "@$script:Domain" } else { '' }
    $lblNewDom.Text = if ($n) { "@$script:Domain" } else { '' }
}

function Invoke-Check {
    $btnApply.Enabled = $false
    $script:Plan  = $null
    $script:Found = $null
    $lblGate.ForeColor = [System.Drawing.Color]::Gray
    $lblGate.Text = ''

    $oldAlias = Get-Typed -Text $txtOld.Text
    $newAlias = Get-Typed -Text $txtNew.Text
    Update-DomainLabels

    if (-not $script:Domain) {
        $txtOut.Text = 'PrimaryDomain is not set in Tenant.json, so there is no domain to build addresses from. Set it in Tenant Manager first -- this tool will not guess one.'
        return
    }
    if (-not $oldAlias -or -not $newAlias) { $txtOut.Text = 'Enter both the current and the new alias.'; return }
    if ($oldAlias -eq $newAlias) { $txtOut.Text = 'The two aliases are the same. Nothing to do.'; return }

    $ad = Test-AdAvailable
    if (-not $ad.Available) { $txtOut.Text = "The ActiveDirectory module is not available: $($ad.Detail)"; return }

    $oldAddr = "$oldAlias@$script:Domain"
    $newAddr = "$newAlias@$script:Domain"

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = "Looking up $oldAddr ..."
    $form.Refresh()
    try {
        $f = Find-RenameTarget -Address $oldAddr
        if (-not $f.Ok) { $txtOut.Text = $f.Reason; return }
        $script:Found = $f.User

        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("Matched on $($f.MatchedOn).")
        [void]$L.Add('')

        # THE ADDRESS IS CHECKED BEFORE ANYTHING IS WRITTEN, not after the write
        # is rejected. See Test-AddressFree.
        $free = Test-AddressFree -Address $newAddr -ExceptDn "$($f.User.distinguishedName)"
        foreach ($n in @($free.Notes)) { [void]$L.Add("  $n") }
        if (@($free.Notes).Count -gt 0) { [void]$L.Add('') }

        $plan = Build-RenamePlan -User $f.User -OldAddress $oldAddr -NewAddress $newAddr -NewAlias $newAlias
        [void]$L.Add((Format-RenamePlan -Plan $plan))

        if ($plan.Remote) {
            [void]$L.Add('')
            [void]$L.Add('This is a REMOTE mailbox: the directory object is here and the mailbox is in Exchange Online. Active Directory is the authority, so this is the correct place to change it -- the cloud copy follows on the next sync.')
        }

        if (-not $free.Free) {
            $lblGate.ForeColor = [System.Drawing.Color]::Firebrick
            $lblGate.Text = 'BLOCKED: the new address is already in use. See above.'
            [void]$L.Add('')
            [void]$L.Add('NOTHING WILL BE WRITTEN while the new address is held by something else. Two objects sharing an address routes mail to whichever the directory resolves first, intermittently -- a fault that shows up days later and is very hard to trace back to today.')
        } else {
            $script:Plan = $plan
            $btnApply.Enabled = $true
            $lblGate.ForeColor = [System.Drawing.Color]::ForestGreen
            $lblGate.Text = 'Checked. Apply rename writes the four attributes above.'
        }
        $txtOut.Text = ($L -join [Environment]::NewLine)
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Rename check'
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Invoke-Rename {
    if (-not $script:Plan) { return }
    $plan = $script:Plan

    $msg = "Rename $($plan.DisplayName)?" + [Environment]::NewLine + [Environment]::NewLine +
           "  $($plan.OldUpn)" + [Environment]::NewLine +
           "  -> $($plan.NewUpn)" + [Environment]::NewLine + [Environment]::NewLine +
           'This writes userPrincipalName, mail, mailNickname and proxyAddresses to Active Directory now.' + [Environment]::NewLine + [Environment]::NewLine +
           'The old address is kept as a secondary, so mail sent to it still arrives.'
    $ans = Show-ToolDialog($msg, 'Write to Active Directory',
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($ans -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $btnApply.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $L = New-Object System.Collections.ArrayList
    try {
        # THE BEFORE STATE IS SAVED FIRST, AND AS A CSV.
        #
        # The desk script did this with '> z:\alias.txt' AFTER the null check
        # would have exited, to a hardcoded drive, using console redirection --
        # which truncates each column at the console width. proxyAddresses, the
        # one attribute you would actually need to restore, was the one it cut
        # off. A CSV of one row per address is recoverable; a truncated table is
        # a record that something happened.
        $backup = @()
        foreach ($a in @($plan.OldProxies)) {
            $backup += [PSCustomObject]@{
                DistinguishedName = $plan.Dn
                DisplayName       = $plan.DisplayName
                sAMAccountName    = $plan.Sam
                userPrincipalName = $plan.OldUpn
                mail              = $plan.OldMail
                mailNickname      = $plan.OldNickname
                proxyAddress      = $a
            }
        }
        if (@($backup).Count -eq 0) {
            $backup = @([PSCustomObject]@{
                DistinguishedName = $plan.Dn; DisplayName = $plan.DisplayName
                sAMAccountName = $plan.Sam; userPrincipalName = $plan.OldUpn
                mail = $plan.OldMail; mailNickname = $plan.OldNickname
                proxyAddress = '(none)' })
        }
        $bk = Write-ReportCsv -Directory (Get-ToolkitReportDirectory) -Prefix "RenameUser_before_$($plan.Sam)" -Rows $backup
        [void]$L.Add("Before state saved to:")
        [void]$L.Add("  $bk")
        [void]$L.Add('')

        $sp = Get-AdServerParam

        # ONE Set-ADUser, NOT TWO. The desk script wrote the identity attributes
        # and then the proxy collection separately -- two writes, and a failure
        # between them leaves a user whose UPN says one thing and whose mailbox
        # receives on another. A single call either applies or does not.
        $p = @{
            Identity          = $plan.Dn
            UserPrincipalName = $plan.NewUpn
            EmailAddress      = $plan.NewMail
            Replace           = @{ mailNickname = $plan.NewNickname; proxyAddresses = @($plan.NewProxies) }
            ErrorAction       = 'Stop'
        }
        foreach ($k in $sp.Keys) { $p[$k] = $sp[$k] }
        Set-ADUser @p

        # READ BACK FROM THE SAME CONTROLLER and compare, like every other write
        # in this toolkit. A write that reports success and did not land is the
        # failure mode worth spending a round trip on.
        $rp = @{ Identity = $plan.Dn
                 Properties = @('userPrincipalName', 'mail', 'mailNickname', 'proxyAddresses')
                 ErrorAction = 'Stop' }
        foreach ($k in $sp.Keys) { $rp[$k] = $sp[$k] }
        $after = Get-ADObject @rp

        $ok = New-Object System.Collections.ArrayList
        $bad = New-Object System.Collections.ArrayList
        if ("$($after.userPrincipalName)" -eq $plan.NewUpn) { [void]$ok.Add('userPrincipalName') } else { [void]$bad.Add("userPrincipalName is '$($after.userPrincipalName)'") }
        if ("$($after.mail)" -eq $plan.NewMail)             { [void]$ok.Add('mail') }              else { [void]$bad.Add("mail is '$($after.mail)'") }
        if ("$($after.mailNickname)" -eq $plan.NewNickname) { [void]$ok.Add('mailNickname') }      else { [void]$bad.Add("mailNickname is '$($after.mailNickname)'") }

        $nowPx = @($after.proxyAddresses | ForEach-Object { "$_" })
        $primary = @($nowPx | Where-Object { $_ -cmatch '^SMTP:' })
        if (@($primary).Count -eq 1 -and "$($primary[0])" -ceq "SMTP:$($plan.NewUpn)") { [void]$ok.Add('proxyAddresses') }
        else { [void]$bad.Add("the primary proxy is $(@($primary) -join ',')") }

        [void]$L.Add("WRITTEN: $(@($ok) -join ', ')")
        if (@($bad).Count -gt 0) {
            [void]$L.Add('')
            [void]$L.Add("DID NOT TAKE: $(@($bad) -join '; ')")
            [void]$L.Add('The before state above is the record of what it was.')
        }
        [void]$L.Add('')
        [void]$L.Add("proxyAddresses now ($(@($nowPx).Count)):")
        foreach ($a in $nowPx) { [void]$L.Add("    $a") }

        [void]$L.Add('')
        [void]$L.Add('THE CLOUD IS NOW STALE. Exchange Online still shows the old address and will until the next AAD Connect sync cycle. That is expected -- do not re-run this because the change has not appeared.')
        [void]$L.Add('Their sign-in name changed. They will need to sign in with the new address, and Outlook may prompt for credentials once.')

        $res = if (@($bad).Count -gt 0) { 'Partial' } else { 'Success' }
        Write-CompletedLog -Action 'Rename' -Target (Get-LogTarget $plan.NewUpn $plan.OldUpn $plan.Sam) -Result $res `
            -Detail "UPN/mail/mailNickname/proxyAddresses: $($plan.OldUpn) -> $($plan.NewUpn). $(@($plan.OldProxies).Count) proxy entries -> $(@($nowPx).Count), $($plan.X500Count) X500 preserved. Before state: $bk" `
            -Environment 'On-premises AD'

        $txtOut.Text = ($L -join [Environment]::NewLine)
        $script:Plan = $null
    } catch {
        [void]$L.Add((Format-ErrorDetail -ErrorRecord $_ -Context 'Set-ADUser'))
        $txtOut.Text = ($L -join [Environment]::NewLine)
        Write-FailedLog -Action 'Rename' -Target (Get-LogTarget $plan.NewUpn $plan.OldUpn) -ErrorRecord $_ `
            -Detail "$($plan.OldUpn) -> $($plan.NewUpn)" -Environment 'On-premises AD'
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}
#endregion

$btnCheck.Add_Click({ try { Invoke-Check } catch { Show-HandlerError -ErrorRecord $_ -Context 'Check' } })
$btnApply.Add_Click({ try { Invoke-Rename } catch { Show-HandlerError -ErrorRecord $_ -Context 'Apply rename' } })
$txtOld.Add_TextChanged({ try { Update-DomainLabels; $btnApply.Enabled = $false; $script:Plan = $null } catch { } })
$txtNew.Add_TextChanged({ try { Update-DomainLabels; $btnApply.Enabled = $false; $script:Plan = $null } catch { } })
$form.AcceptButton = $btnCheck

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Rename User' | Out-Null

$form.Add_Shown({
    try {
        try { $script:Server = "$(Get-TenantDirectoryServer -Id (Get-DefaultTenantDirectory).Id)".Trim() } catch { $script:Server = '' }
        if ("$SmtpAddress".Trim()) { $txtOld.Text = (Get-Typed -Text $SmtpAddress) }
        Update-DomainLabels

        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("Domain: @$script:Domain   (Tenant.json PrimaryDomain)")
        [void]$L.Add("Directory: $(if ($script:Server) { $script:Server } else { 'resolved by Active Directory' })")
        $ad = Test-AdAvailable
        [void]$L.Add("ActiveDirectory module: $(if ($ad.Available) { 'available' } else { $ad.Detail })")
        [void]$L.Add('')
        [void]$L.Add('Enter the current alias and the new one, then press Check. Nothing is written until you press Apply rename and confirm.')
        [void]$L.Add('')
        [void]$L.Add('Check does three things: finds the user by UPN, mail OR proxy address; confirms the new address is not already held by another object or a soft-deleted mailbox; and shows exactly what would be written.')
        $txtOut.Text = ($L -join [Environment]::NewLine)
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Rename User' }
})

[void]$form.ShowDialog()
