#Requires -Version 5.1
# TOOLKIT-VARIANT: OnPrem
# TOOLKIT-REQUIRES: OwnProcess  -- Get-User, Set-User, Get-Mailbox and
#   Set-Mailbox all exist in BOTH Exchanges and mean a different directory in
#   each. This tool means the ON-PREMISES one. Every call goes through
#   Invoke-OnPremCommand, but sharing a process with a connected Exchange
#   Online session is still how a rename lands in the wrong directory with no
#   error at all.
# ---------------------------------------------------------------------------
# ON-PREMISES, AND NOW THE ONLY VERSION.
#
# This file used to be the cloud twin of DisabledTools\AD-MailAttributes.ps1 --
# one read Exchange Online, the other read raw AD through Get-ADUser/Set-ADUser.
# The split was a mistake for a hybrid tenant and has been collapsed: there is
# one Mail Attributes tool and it reads and writes on-premises Exchange.
#
# WHY ON-PREMISES IS THE ONLY CORRECT ANSWER HERE. In a hybrid tenant the
# authoritative copy of these attributes is the on-premises directory. Directory
# sync pushes them up. Editing them in Exchange Online either fails outright on
# a synchronised object or is silently reverted on the next sync cycle, which is
# worse -- the tool reports success and the change disappears an hour later.
#
# WHY EXCHANGE AND NOT Set-ADUser. Writing proxyAddresses through Set-ADUser
# works and is what the parked AD- variant did, but it bypasses every check
# Exchange makes: address format, uniqueness across the organisation, the
# SMTP:/smtp: casing that decides which address is primary, and the email
# address policy. A malformed proxyAddresses collection is accepted without
# complaint and shows up later as mail that does not arrive.
# ---------------------------------------------------------------------------
<#
    MailAttributes.ps1
    Mail attributes on a user, read and edited on-premises: display name,
    alias, primary SMTP address and the rest of the proxy addresses.

    Author   : Glenn Armbruster
    Revised  : 2026-09-10
#>
# TOOLKIT-HELP-RUNS: reads on open. Every change asks first and says exactly what it will do.
# TOOLKIT-HELP-BEGIN
# View and edit the mail-related attributes on an account -- addresses, display
# name, alias, and how the person appears in the address book.
#
# THIS IS THE RENAME TOOL. Someone marries, changes name, or was created with a
# typo: change the display name, the alias and the primary SMTP address here,
# and keep the old address as a secondary so mail sent to it still arrives.
#
# EVERYTHING HAPPENS ON-PREMISES and syncs to the cloud on the next cycle
# (usually within 30 minutes). Nothing here touches Exchange Online, because on
# a synced object the cloud copy is not the authoritative one -- an edit made
# there is reverted on the next sync.
#
# THE EMAIL ADDRESS POLICY IS THE THING THAT CATCHES PEOPLE OUT. If the account
# is subject to an email address policy -- which is the default -- Exchange
# generates its addresses from that policy and will not let you pin a different
# primary. The tool reads that flag, shows it, and will not pretend to set a
# primary address that the policy is going to regenerate. To change the primary
# on such an account you have to detach it from the policy first, and the tool
# asks you to confirm that explicitly, because from then on that account's
# addresses are yours to maintain by hand.
#
# ALIAS BEHAVES DIFFERENTLY DEPENDING ON THE OBJECT, and this is Microsoft's
# documented behaviour rather than a quirk of the tool:
#   - Remote mailbox (mailbox in the cloud, account on-premises): changing the
#     Alias does NOT change the primary address.
#   - On-premises mailbox subject to a policy: changing the Alias DOES
#     regenerate the primary address.
# The tool works out which kind of object it is looking at and says so.
#
# WHAT IS DELIBERATELY READ-ONLY:
#   UserPrincipalName        sign-in name; changing it affects far more than
#                            mail and belongs in its own confirmation flow.
#   RemoteRoutingAddress     the routing pointer to the cloud mailbox. Getting
#                            it wrong is a mail-flow outage, not a typo.
#   WindowsEmailAddress      Microsoft documents this as having NO effect while
#                            the account is subject to an email address policy.
#                            An edit that silently does nothing is worse than
#                            no edit button.
#   EmailAddresses           edit 'Primary SMTP' to change the primary, or use
#                            'Add Secondary' to add one. Hand-editing the whole
#                            collection is how the SMTP:/smtp: casing gets lost.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. An empty address is fine; the window
# opens and asks for one.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Actions - mail attributes ---------------------------------------------
function onprem { param([string]$n, [hashtable]$p) Invoke-OnPremCommand -Name $n -Parameters $p }

function Test-OnPremReady { try { return [bool](Test-EmsAvailable).Available } catch { return $false } }

<#
    THE ATTRIBUTE SET, and where each one actually lives.

        Shown as                     AD attribute      Read from        Written by
        ---------------------------  ----------------  ---------------  ----------------
        Company (Office)             company           Get-User         Set-User
        DisplayName                  displayName       Get-User         Set-User
        Name                         cn / name         Get-User         Set-User
        FirstName                    givenName         Get-User         Set-User
        Initials                     initials          Get-User         Set-User
        LastName                     sn                Get-User         Set-User
        WindowsEmailAddress          mail              Get-User         (read-only)
        UserPrincipalName            userPrincipalName Get-User         (read-only)
        Alias                        mailNickname      Get-RemoteMailbox / Get-Mailbox
        PrimarySmtpAddress           proxyAddresses    (derived)        via EmailAddresses
        EmailAddresses               proxyAddresses    Get-RemoteMailbox / Get-Mailbox
        RemoteRoutingAddress         targetAddress     Get-RemoteMailbox  (read-only)
        EmailAddressPolicyEnabled    -                 Get-RemoteMailbox / Get-Mailbox

    ArchiveName was dropped rather than carried over from the cloud version.
    Set-RemoteMailbox is not documented as taking it, it has nothing to do with
    a rename, and Outlook desktop ignores the value anyway. A parameter this
    tool cannot verify is a parameter it should not send.
#>
$script:MailAttribsDisplayOrder = @(
    'Company', 'DisplayName', 'Name', 'FirstName', 'Initials', 'LastName',
    'Alias', 'PrimarySmtpAddress', 'EmailAddresses', 'WindowsEmailAddress',
    'RemoteRoutingAddress', 'RecipientTypeDetails', 'EmailAddressPolicyEnabled',
    'UserPrincipalName'
)
$script:MailAttribsDisplayLabels = @{
    'Company'                   = 'Office (company)'
    'DisplayName'               = 'DisplayName (displayName)'
    'Name'                      = 'Name (cn)'
    'FirstName'                 = 'FirstName (givenName)'
    'LastName'                  = 'LastName (sn)'
    'Alias'                     = 'Alias (mailNickname)'
    'PrimarySmtpAddress'        = 'Primary SMTP'
    'EmailAddresses'            = 'EmailAddresses (proxyAddresses)'
    'WindowsEmailAddress'       = 'mail (WindowsEmailAddress)'
    'RemoteRoutingAddress'      = 'RemoteRoutingAddress (targetAddress)'
    'RecipientTypeDetails'      = 'Object type'
    'EmailAddressPolicyEnabled' = 'Email address policy'
}
$script:MailAttribsEditableAttributes = @(
    'Company', 'DisplayName', 'Name', 'FirstName', 'Initials', 'LastName',
    'Alias', 'PrimarySmtpAddress'
)
$script:MailAttribsNeverBlank = @('DisplayName', 'Name', 'Alias', 'PrimarySmtpAddress')

# Set-User and Set-Mailbox both cap Name at 64; DisplayName at 256.
$script:MailAttribsMaxLength = @{ 'Name' = 64; 'DisplayName' = 256 }

function Get-MailAttribsWriteDc {
    <#
        A DC that can WRITE this object: its own domain's, not whatever the
        session happens to be pointed at.

        The same lesson DLtoSharedMB learned the hard way. In a multi-domain
        forest a global catalogue can READ any domain but only WRITE its own,
        and a write sent to the wrong one comes back as
        "Received a referral ... the operation could not be performed on this
        server" -- per attribute, halfway through a rename, leaving the object
        in a state nobody intended.

        -DomainController on the command line wins, because an operator who
        named one is troubleshooting. Otherwise the object's DN decides.
        Returns '' for "let the session choose", which callers simply omit.
    #>
    param([string]$DistinguishedName)
    if ("$DomainController".Trim()) { return "$DomainController".Trim() }
    try {
        $dir = Resolve-DistinguishedNameDirectory -DistinguishedName $DistinguishedName
        if ($dir) {
            $dc = Resolve-DirectoryDomainController -DirectoryId "$($dir.Id)"
            if ("$dc".Trim()) { return "$dc".Trim() }
        }
    } catch { }
    return ''
}

function Add-MailAttribsDc {
    # Adds -DomainController to a parameter hashtable when one was worked out.
    param([Parameter(Mandatory)][hashtable]$Parameters, [string]$Dc)
    if ("$Dc".Trim()) { $Parameters['DomainController'] = "$Dc".Trim() }
    return $Parameters
}

function Get-MailAttribsUser {
    <#
        One record assembled from Get-User plus whichever mailbox cmdlet fits.

        TWO KINDS OF OBJECT, AND THE TOOL HAS TO KNOW WHICH IT HAS. In a hybrid
        tenant the on-premises object for a cloud mailbox is a REMOTE mailbox
        (Get-RemoteMailbox); a mailbox that still lives on-premises is a plain
        mailbox (Get-Mailbox). They take different cmdlets and, per Microsoft's
        documentation, Alias behaves differently on each. Guessing wrong means
        either a cmdlet that errors or a warning that is untrue, so the route is
        established once, here, and carried on the record.

        NEVER CALLED WITH A BLANK IDENTITY. Microsoft documents that
        Get-RemoteMailbox with a $null or non-existent Identity returns ALL
        objects rather than none. In a tool whose next step is an edit, being
        handed the whole directory because the box was empty is not an outcome
        worth risking.
    #>
    param([Parameter(Mandatory)][string]$Smtp)

    if (-not "$Smtp".Trim()) { throw 'No address supplied.' }

    Write-ToolStep 'Reading the directory account ...'
    $u = onprem 'Get-User' @{ Identity = $Smtp; ErrorAction = 'Stop' }

    Write-ToolStep 'Reading the mailbox object ...'
    $mb = $null; $route = ''
    try { $mb = onprem 'Get-RemoteMailbox' @{ Identity = $Smtp; ErrorAction = 'Stop' }; $route = 'RemoteMailbox' } catch { }
    if (-not $mb) {
        try { $mb = onprem 'Get-Mailbox' @{ Identity = $Smtp; ErrorAction = 'Stop' }; $route = 'Mailbox' } catch { }
    }
    if (-not $mb) {
        throw "'$Smtp' resolves to a directory account but not to a mailbox on-premises. Get-RemoteMailbox and Get-Mailbox both failed -- it may be a mail user, a contact, or cloud-only."
    }

    return [PSCustomObject]@{
        Identity                  = "$($mb.Identity)"
        DistinguishedName         = "$($u.DistinguishedName)"
        Route                     = $route
        Company                   = "$($u.Company)"
        DisplayName               = "$($u.DisplayName)"
        Name                      = "$($u.Name)"
        FirstName                 = "$($u.FirstName)"
        Initials                  = "$($u.Initials)"
        LastName                  = "$($u.LastName)"
        WindowsEmailAddress       = "$($u.WindowsEmailAddress)"
        UserPrincipalName         = "$($u.UserPrincipalName)"
        Alias                     = "$($mb.Alias)"
        RemoteRoutingAddress      = "$($mb.RemoteRoutingAddress)"
        RecipientTypeDetails      = "$($mb.RecipientTypeDetails)"
        EmailAddressPolicyEnabled = [bool]$mb.EmailAddressPolicyEnabled
        EmailAddresses            = @($mb.EmailAddresses | ForEach-Object { "$_" })
    }
}

function Get-PrimarySmtpFromProxyAddresses {
    # The primary is the ONE uppercase SMTP: entry. -cmatch, not -match:
    # the case IS the meaning here, and a case-insensitive test would return
    # the first alias it met and call it the primary.
    param($ProxyAddresses)
    foreach ($addr in @($ProxyAddresses)) {
        if ("$addr" -cmatch '^SMTP:(.+)$') { return $Matches[1] }
    }
    return ''
}

function Set-PrimarySmtpInProxyAddresses {
    param(
        $CurrentProxyAddresses,
        [Parameter(Mandatory)][string]$NewPrimary,
        [bool]$KeepOldPrimaryAsAlias = $true
    )
    $kept = New-Object System.Collections.Generic.List[string]
    $oldPrimary = $null
    foreach ($addr in @($CurrentProxyAddresses)) {
        $s = "$addr"
        if ($s -cmatch '^SMTP:(.+)$') { $oldPrimary = $Matches[1]; continue }
        if ($s -match '^smtp:(.+)$' -and $Matches[1] -ieq $NewPrimary) { continue }
        [void]$kept.Add($s)
    }
    if ($KeepOldPrimaryAsAlias -and $oldPrimary -and $oldPrimary -ine $NewPrimary) {
        [void]$kept.Add("smtp:$oldPrimary")
    }
    $result = New-Object System.Collections.Generic.List[string]
    [void]$result.Add("SMTP:$NewPrimary")
    foreach ($k in $kept) { [void]$result.Add($k) }
    return $result.ToArray()
}

function Format-AttributeValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [bool]) { return $(if ($Value) { 'True' } else { 'False' }) }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = New-Object System.Collections.Generic.List[string]
        foreach ($item in $Value) { [void]$items.Add("$item") }
        return ($items -join '; ')
    }
    return "$Value"
}

function Format-MailAttribsTable {
    param($UserObject)
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Attribute', [string])
    [void]$table.Columns.Add('Value', [string])
    [void]$table.Columns.Add('_Editable', [string])
    [void]$table.Columns.Add('_AttributeKey', [string])
    foreach ($attrName in $script:MailAttribsDisplayOrder) {
        # RemoteRoutingAddress only exists on a remote mailbox; showing an
        # empty row for it on an on-premises mailbox invites the question
        # "why is that blank".
        if ($attrName -eq 'RemoteRoutingAddress' -and $UserObject.Route -ne 'RemoteMailbox') { continue }
        $rawValue = if ($attrName -eq 'PrimarySmtpAddress') {
            Get-PrimarySmtpFromProxyAddresses -ProxyAddresses $UserObject.EmailAddresses
        } else {
            $UserObject.$attrName
        }
        if ($attrName -eq 'EmailAddressPolicyEnabled') {
            $rawValue = $(if ($UserObject.EmailAddressPolicyEnabled) { 'ON -- addresses are generated by policy' } else { 'off -- addresses are maintained by hand' })
        }
        $label = if ($script:MailAttribsDisplayLabels.ContainsKey($attrName)) { $script:MailAttribsDisplayLabels[$attrName] } else { $attrName }
        $row = $table.NewRow()
        $row['Attribute']     = $label
        $row['Value']         = Format-AttributeValue -Value $rawValue
        $row['_Editable']     = if ($script:MailAttribsEditableAttributes -contains $attrName) { 'Y' } else { 'N' }
        $row['_AttributeKey'] = $attrName
        [void]$table.Rows.Add($row)
    }
    return ,$table
}

function Set-MailAttribsAttribute {
    <#
        Set-User for the person fields, Set-RemoteMailbox or Set-Mailbox for the
        mail fields, everything pinned to a writable DC in the object's own
        domain.

        THE PRIMARY ADDRESS IS NOT SET WITH -PrimarySmtpAddress, and that is
        deliberate. Microsoft documents two constraints that decide the route:

          - "You can't use the EmailAddresses parameter and the
            PrimarySmtpAddress parameter in the same command."
          - "If you set the EmailAddressPolicyEnabled parameter to $false, you
            can specify the primary address using the PrimarySmtpAddress
            parameter" -- i.e. while the policy is ON, that parameter is not
            the supported way to pin a primary.

        Rewriting the whole EmailAddresses collection with exactly one
        uppercase SMTP: entry does the job in one command, keeps every alias,
        and is the same shape whether the policy is on or off. The caller
        handles the policy itself before getting here.
    #>
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][string]$AttributeName,
        [string]$NewValue,
        $CurrentProxyAddresses,
        [bool]$KeepOldPrimaryAsAlias = $true,
        [Parameter(Mandatory)][string]$Route,
        [string]$Dc
    )
    $isBlank = ("$NewValue".Trim() -eq '')
    $value   = if ($isBlank) { $null } else { $NewValue.Trim() }
    $mbCmd   = if ($Route -eq 'RemoteMailbox') { 'Set-RemoteMailbox' } else { 'Set-Mailbox' }

    switch ($AttributeName) {
        # Learn (Set-User): -Identity plus the person fields; -DomainController
        # is on-premises only, which is exactly where this tool runs.
        { $_ -in @('Company', 'DisplayName', 'Name', 'FirstName', 'Initials', 'LastName') } {
            $p = @{ Identity = $Identity; ErrorAction = 'Stop' }
            $p[$AttributeName] = $value
            onprem 'Set-User' (Add-MailAttribsDc -Parameters $p -Dc $Dc) | Out-Null
        }
        'Alias' {
            if ($isBlank) { throw 'Alias cannot be blank -- an account without one is not addressable.' }
            # Learn (Set-RemoteMailbox / Set-Mailbox): -Alias <String>.
            $p = @{ Identity = $Identity; Alias = $value; ErrorAction = 'Stop' }
            onprem $mbCmd (Add-MailAttribsDc -Parameters $p -Dc $Dc) | Out-Null
        }
        'PrimarySmtpAddress' {
            if ($isBlank) { throw 'PrimarySmtpAddress cannot be blank.' }
            $new = Set-PrimarySmtpInProxyAddresses -CurrentProxyAddresses $CurrentProxyAddresses -NewPrimary $value -KeepOldPrimaryAsAlias $KeepOldPrimaryAsAlias
            # Learn (Set-RemoteMailbox / Set-Mailbox): -EmailAddresses
            # <ProxyAddressCollection>. Passing the collection wholesale
            # "replace[s] all existing proxy email addresses with the values
            # you specify", which is what the rewrite above builds.
            $p = @{ Identity = $Identity; EmailAddresses = $new; ErrorAction = 'Stop' }
            onprem $mbCmd (Add-MailAttribsDc -Parameters $p -Dc $Dc) | Out-Null
        }
        default { throw "Attribute '$AttributeName' is not handled by Set-MailAttribsAttribute." }
    }
}

function Disable-MailAttribsAddressPolicy {
    <#
        Detach the object from its email address policy.

        A SEPARATE FUNCTION BECAUSE IT IS A SEPARATE DECISION. This is not part
        of a rename; it permanently changes who maintains this account's
        addresses, from Exchange to whoever is holding the mouse. It is only
        ever called after the operator has said yes to that in as many words.
    #>
    param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string]$Route, [string]$Dc)
    $cmd = if ($Route -eq 'RemoteMailbox') { 'Set-RemoteMailbox' } else { 'Set-Mailbox' }
    # Learn (Set-RemoteMailbox / Set-Mailbox): -EmailAddressPolicyEnabled
    # <Boolean>; on-premises only on Set-Mailbox. Default is $true.
    $p = @{ Identity = $Identity; EmailAddressPolicyEnabled = $false; ErrorAction = 'Stop' }
    onprem $cmd (Add-MailAttribsDc -Parameters $p -Dc $Dc) | Out-Null
}

function Add-MailAttribsSecondaryProxyAddress {
    <#
        One address added, nothing else touched.

        Uses the documented hashtable form rather than rewriting the whole
        collection: "To add or remove specify proxy addresses without affecting
        other existing values, use the following syntax:
        @{Add="Type:EmailAddress1",...; Remove=...}". Sending the whole
        collection back to add one entry is how the primary's casing gets lost
        in a tool that was only meant to append.
    #>
    param(
        [Parameter(Mandatory)][string]$Identity,
        $CurrentProxyAddresses,
        [Parameter(Mandatory)][string]$NewSecondary,
        [Parameter(Mandatory)][string]$Route,
        [string]$Dc
    )
    foreach ($addr in @($CurrentProxyAddresses)) {
        if ("$addr" -match '^smtp:(.+)$' -and $Matches[1] -ieq $NewSecondary) {
            throw "'$NewSecondary' already exists in EmailAddresses (as $addr)."
        }
    }
    $cmd = if ($Route -eq 'RemoteMailbox') { 'Set-RemoteMailbox' } else { 'Set-Mailbox' }
    $p = @{ Identity = $Identity; EmailAddresses = @{ Add = "smtp:$NewSecondary" }; ErrorAction = 'Stop' }
    onprem $cmd (Add-MailAttribsDc -Parameters $p -Dc $Dc) | Out-Null
}

function Get-MailAttribsWarning {
    <#
        The sentence shown above the edit box, per attribute and per object
        type. The Alias one differs by route and that difference is Microsoft's,
        not this tool's -- see the quotes below.
    #>
    param([string]$AttributeName, [string]$Route, [bool]$PolicyOn)
    switch ($AttributeName) {
        'Alias' {
            if ($Route -eq 'RemoteMailbox') {
                return 'The Alias is the mailbox nickname (mailNickname). On a remote mailbox Microsoft documents that "the Alias parameter never generates or updates the primary email address of a mail contact or a mail user" -- so this will NOT change the primary address. Cannot be blank.'
            }
            if ($PolicyOn) {
                return 'The Alias is the mailbox nickname (mailNickname). This account IS subject to an email address policy, and Microsoft documents that changing the Alias updates the primary email address in that case. Check the Primary SMTP row afterwards. Cannot be blank.'
            }
            return 'The Alias is the mailbox nickname (mailNickname). This account is not subject to an email address policy, so the primary address will not be regenerated. Cannot be blank.'
        }
        'PrimarySmtpAddress' {
            if ($PolicyOn) {
                return 'This account is subject to an EMAIL ADDRESS POLICY, which generates its addresses. Setting a primary by hand requires detaching it from that policy first -- you will be asked to confirm that separately, and it is not reversible from this tool. Cannot be blank.'
            }
            return 'This rewrites the whole EmailAddresses collection: sets this value as the SMTP: (primary) entry. If a different primary already exists you will be asked next whether to keep it as a secondary alias or remove it. Cannot be blank.'
        }
        'Name' {
            return 'Name is the object''s unique name in the directory (cn). Maximum 64 characters -- shorter than DisplayName''s 256, so a long name that fits there may not fit here. Cannot be blank.'
        }
        'DisplayName' { return 'How the person appears in the address book. Maximum 256 characters. Cannot be blank.' }
        default { return '' }
    }
}
#endregion

#region Dialogs ---------------------------------------------------------------
function Show-MailAttribsEditDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$AttributeName,
        [string]$CurrentValue,
        [string]$WarningNote
    )
    $dlgHeight = if ($WarningNote -ne '') { 372 } else { 300 }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit Attribute -- $AttributeName"
    $dlg.Size = New-Object System.Drawing.Size(520, $dlgHeight)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    $lblAttr = New-Object System.Windows.Forms.Label
    $lblAttr.Text = "Attribute: $AttributeName"
    $lblAttr.Location = New-Object System.Drawing.Point(12, 12)
    $lblAttr.AutoSize = $true
    $lblAttr.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblAttr)
    $nextY = 42
    if ($WarningNote -ne '') {
        $lblWarnNote = New-Object System.Windows.Forms.Label
        $lblWarnNote.Text = $WarningNote
        $lblWarnNote.Location = New-Object System.Drawing.Point(12, $nextY)
        $lblWarnNote.Size = New-Object System.Drawing.Size(480, 82)
        $lblWarnNote.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Italic)
        $lblWarnNote.ForeColor = [System.Drawing.Color]::DarkOrange
        $lblWarnNote.Tag = 'NoTheme'
        $dlg.Controls.Add($lblWarnNote)
        $nextY += 88
    }
    $lblCurrent = New-Object System.Windows.Forms.Label
    $lblCurrent.Text = 'Current Value:'
    $lblCurrent.Location = New-Object System.Drawing.Point(12, $nextY)
    $lblCurrent.AutoSize = $true
    $dlg.Controls.Add($lblCurrent)
    $nextY += 18
    $txtCurrent = New-Object System.Windows.Forms.TextBox
    $txtCurrent.Location = New-Object System.Drawing.Point(12, $nextY)
    $txtCurrent.Size = New-Object System.Drawing.Size(480, 40)
    $txtCurrent.Multiline = $true
    $txtCurrent.ReadOnly = $true
    $txtCurrent.ScrollBars = 'Vertical'
    $txtCurrent.Text = $CurrentValue
    $txtCurrent.Font = New-Object System.Drawing.Font('Consolas', 9)
    $dlg.Controls.Add($txtCurrent)
    $nextY += 48
    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = 'New Value:'
    $lblNew.Location = New-Object System.Drawing.Point(12, $nextY)
    $lblNew.AutoSize = $true
    $dlg.Controls.Add($lblNew)
    $nextY += 18
    $txtNew = New-Object System.Windows.Forms.TextBox
    $txtNew.Location = New-Object System.Drawing.Point(12, $nextY)
    $txtNew.Size = New-Object System.Drawing.Size(480, 46)
    $txtNew.Multiline = $true
    $txtNew.ScrollBars = 'Vertical'
    $txtNew.Text = $CurrentValue
    $txtNew.Font = New-Object System.Drawing.Font('Consolas', 9)
    $dlg.Controls.Add($txtNew)
    $nextY += 54
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(320, $nextY)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Location = New-Object System.Drawing.Point(408, $nextY)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    Set-ConsoleTheme -Form $dlg
    $result = $dlg.ShowDialog($Owner)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $txtNew.Text
}

function Show-KeepOrRemoveOldPrimaryDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$OldPrimary,
        [string]$NewPrimary
    )
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Previous Primary SMTP Address'
    $dlg.ClientSize = New-Object System.Drawing.Size(444, 190)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    $dlg.Tag = 'Cancel'
    $lblMsg = New-Object System.Windows.Forms.Label
    $lblMsg.Text = "Current primary SMTP address:`n$OldPrimary`n`nIt is being replaced by:`n$NewPrimary`n`nKeeping it as an alias means mail sent to the old address still arrives. Removing it means that mail bounces."
    $lblMsg.Location = New-Object System.Drawing.Point(12, 12)
    $lblMsg.Size = New-Object System.Drawing.Size(420, 116)
    $dlg.Controls.Add($lblMsg)
    $btnKeep = New-Object System.Windows.Forms.Button
    $btnKeep.Text = 'Keep as Alias'
    $btnKeep.Location = New-Object System.Drawing.Point(20, 144)
    $btnKeep.Size = New-Object System.Drawing.Size(130, 32)
    $btnKeep.Add_Click({ $dlg.Tag = 'Keep'; $dlg.Close() })
    $dlg.Controls.Add($btnKeep)
    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = 'Remove'
    $btnRemove.Location = New-Object System.Drawing.Point(165, 144)
    $btnRemove.Size = New-Object System.Drawing.Size(110, 32)
    $btnRemove.Add_Click({ $dlg.Tag = 'Remove'; $dlg.Close() })
    $dlg.Controls.Add($btnRemove)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(290, 144)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnCancel.Add_Click({ $dlg.Tag = 'Cancel'; $dlg.Close() })
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    Set-ConsoleTheme -Form $dlg
    [void]$dlg.ShowDialog($Owner)
    return "$($dlg.Tag)"
}
#endregion

#region Form ------------------------------------------------------------------
function Show-MailAttribsForm {
    # NO -Smtp PARAMETER. The address can change while the window is open, so
    # it lives in $script:Target rather than a parameter fixed at build time.
    $mailForm = New-Object System.Windows.Forms.Form
    $mailForm.Text = "Mail Attributes: $script:Target"
    $mailForm.ClientSize = New-Object System.Drawing.Size(664, 672)
    $mailForm.StartPosition = 'Manual'
    $mailForm.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads and writes ON-PREMISES Exchange (Get-User + Get-RemoteMailbox / Get-Mailbox). Changes sync to the cloud on the next cycle. Double-click a row, or select it and click 'Edit Attribute', to change it. 'EmailAddresses' is read-only here -- edit 'Primary SMTP' to change the primary (you will be asked whether to keep the old one as an alias), or use 'Add Secondary' below to add another address. Watch the 'Email address policy' row: while it is ON, Exchange generates the addresses and the primary cannot be pinned by hand."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(640, 104)
    $mailForm.Controls.Add($lblNote)

    $grdMail = New-Object System.Windows.Forms.DataGridView
    $grdMail.Location = New-Object System.Drawing.Point(12, 120)
    $grdMail.Size = New-Object System.Drawing.Size(640, 300)
    $grdMail.ReadOnly = $true
    $grdMail.AllowUserToAddRows = $false
    $grdMail.AllowUserToDeleteRows = $false
    $grdMail.AllowUserToResizeRows = $false
    $grdMail.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdMail.MultiSelect = $false
    $grdMail.AutoGenerateColumns = $false
    $grdMail.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $grdMail.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $grdMail.RowHeadersVisible = $false
    $grdMail.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $grdMail.Font = New-Object System.Drawing.Font('Consolas', 9)
    $colMailAttribute = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailAttribute.Name = 'Attribute'; $colMailAttribute.HeaderText = 'Attribute'
    $colMailAttribute.DataPropertyName = 'Attribute'; $colMailAttribute.Width = 240; $colMailAttribute.ReadOnly = $true
    [void]$grdMail.Columns.Add($colMailAttribute)
    $colMailValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailValue.Name = 'Value'; $colMailValue.HeaderText = 'Value'
    $colMailValue.DataPropertyName = 'Value'; $colMailValue.ReadOnly = $true
    $colMailValue.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    [void]$grdMail.Columns.Add($colMailValue)
    $colMailEditableFlag = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailEditableFlag.Name = '_Editable'; $colMailEditableFlag.DataPropertyName = '_Editable'; $colMailEditableFlag.Visible = $false
    [void]$grdMail.Columns.Add($colMailEditableFlag)
    $colMailAttributeKey = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailAttributeKey.Name = '_AttributeKey'; $colMailAttributeKey.DataPropertyName = '_AttributeKey'; $colMailAttributeKey.Visible = $false
    [void]$grdMail.Columns.Add($colMailAttributeKey)
    $mailForm.Controls.Add($grdMail)
    $grdMail.DataSource = New-Object System.Windows.Forms.BindingSource
    $mailBindingSource = $grdMail.DataSource

    $lblAddSmtp = New-Object System.Windows.Forms.Label
    $lblAddSmtp.Text = 'Add secondary SMTP address (appended as smtp: -- does not replace the primary or any other address):'
    $lblAddSmtp.Location = New-Object System.Drawing.Point(12, 428)
    $lblAddSmtp.Size = New-Object System.Drawing.Size(640, 16)
    $mailForm.Controls.Add($lblAddSmtp)
    $txtAddSmtp = New-Object System.Windows.Forms.TextBox
    $txtAddSmtp.Location = New-Object System.Drawing.Point(12, 446)
    $txtAddSmtp.Size = New-Object System.Drawing.Size(430, 22)
    $txtAddSmtp.Font = New-Object System.Drawing.Font('Consolas', 9)
    $mailForm.Controls.Add($txtAddSmtp)
    $btnAddSmtp = New-Object System.Windows.Forms.Button
    $btnAddSmtp.Text = 'Add Secondary'
    $btnAddSmtp.Location = New-Object System.Drawing.Point(452, 445)
    $btnAddSmtp.Size = New-Object System.Drawing.Size(110, 26)
    $mailForm.Controls.Add($btnAddSmtp)

    $lblMailLog = New-Object System.Windows.Forms.Label
    $lblMailLog.Text = 'Log:'
    $lblMailLog.Location = New-Object System.Drawing.Point(12, 480)
    $lblMailLog.AutoSize = $true
    $mailForm.Controls.Add($lblMailLog)
    $txtMailLog = New-Object System.Windows.Forms.TextBox
    $txtMailLog.Multiline = $true
    $txtMailLog.ReadOnly = $true
    $txtMailLog.ScrollBars = 'Vertical'
    $txtMailLog.Location = New-Object System.Drawing.Point(12, 500)
    $txtMailLog.Size = New-Object System.Drawing.Size(640, 120)
    $txtMailLog.Font = New-Object System.Drawing.Font('Consolas', 9)
    $mailForm.Controls.Add($txtMailLog)

    $btnMailRefresh = New-Object System.Windows.Forms.Button
    $btnMailRefresh.Text = 'Refresh'
    $btnMailRefresh.Location = New-Object System.Drawing.Point(12, 632)
    $btnMailRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $mailForm.Controls.Add($btnMailRefresh)
    $btnMailEdit = New-Object System.Windows.Forms.Button
    $btnMailEdit.Text = 'Edit Attribute'
    $btnMailEdit.Location = New-Object System.Drawing.Point(122, 632)
    $btnMailEdit.Size = New-Object System.Drawing.Size(140, 28)
    $btnMailEdit.Enabled = $false
    $btnMailEdit.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $mailForm.Controls.Add($btnMailEdit)
    $btnMailClose = New-Object System.Windows.Forms.Button
    $btnMailClose.Text = 'Close'
    $btnMailClose.Location = New-Object System.Drawing.Point(552, 632)
    $btnMailClose.Size = New-Object System.Drawing.Size(100, 28)
    $mailForm.Controls.Add($btnMailClose)

    $script:MailAttribsCurrentUser = $null

    function Update-MailAttribsGrid {
        param($UserObject)
        $table = Format-MailAttribsTable -UserObject $UserObject
        $mailBindingSource.DataSource = $null
        $mailBindingSource.DataSource = $table
        $grdMail.Refresh()
    }

    function Set-MailAttribsEditEnabled {
        $hasRow = ($grdMail.SelectedRows.Count -gt 0 -and $null -ne $script:MailAttribsCurrentUser)
        $isEditable = $false
        if ($hasRow) { $isEditable = ("$($grdMail.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
        $btnMailEdit.Enabled = ($hasRow -and $isEditable)
    }

    function Update-MailAttribsFromCurrent {
        # OPENED WITH NO ADDRESS IS A NORMAL STATE -- the window appears before
        # anything has been typed, so the refresh must not run empty. It also
        # must not reach Get-RemoteMailbox with a blank Identity; see the note
        # in Get-MailAttribsUser for what that cmdlet does with one.
        if (-not "$script:Target".Trim()) {
            $mailForm.Text = 'Mail Attributes'
            return
        }
        $mailForm.Text = 'Mail Attributes: ' + $script:Target
        $btnMailRefresh.Enabled = $false
        $btnAddSmtp.Enabled = $false
        $txtMailLog.Text = 'Loading mail attributes...'
        $mailForm.Refresh()
        try {
            if (-not (Test-OnPremReady)) { throw 'On-premises Exchange is not connected in this window. Connect it from the console and press Refresh.' }
            $script:MailAttribsCurrentUser = Get-MailAttribsUser -Smtp $script:Target
            Update-MailAttribsGrid -UserObject $script:MailAttribsCurrentUser
            $u = $script:MailAttribsCurrentUser
            $kind = if ($u.Route -eq 'RemoteMailbox') { 'remote mailbox (mailbox in the cloud, account on-premises)' } else { 'on-premises mailbox' }
            $pol  = if ($u.EmailAddressPolicyEnabled) { 'ON -- the primary address cannot be pinned by hand until it is detached' } else { 'off -- addresses are maintained by hand' }
            $txtMailLog.Text = "Loaded '$($u.UserPrincipalName)'.`r`nObject: $kind ($($u.RecipientTypeDetails))`r`nEmail address policy: $pol"
        } catch {
            $script:MailAttribsCurrentUser = $null
            $mailBindingSource.DataSource = $null
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
            Set-MailAttribsEditEnabled
        }
    }

    Update-MailAttribsFromCurrent
    $btnMailRefresh.Add_Click({ Update-MailAttribsFromCurrent })

    function Invoke-MailAttribsEditSelectedRow {
        if ($grdMail.SelectedRows.Count -eq 0 -or $null -eq $script:MailAttribsCurrentUser) { return }
        $row = $grdMail.SelectedRows[0]
        $attrLabel = "$($row.Cells['Attribute'].Value)"
        $attrName  = "$($row.Cells['_AttributeKey'].Value)"
        if ($script:MailAttribsEditableAttributes -notcontains $attrName) { return }
        $btnMailRefresh.Enabled = $false
        $btnMailEdit.Enabled = $false
        $btnAddSmtp.Enabled = $false
        try {
            $u = $script:MailAttribsCurrentUser
            $oldValue = "$($row.Cells['Value'].Value)"
            $warn = Get-MailAttribsWarning -AttributeName $attrName -Route $u.Route -PolicyOn $u.EmailAddressPolicyEnabled
            $newValue = Show-MailAttribsEditDialog -Owner $mailForm -AttributeName $attrLabel -CurrentValue $oldValue -WarningNote $warn
            if ($null -eq $newValue) { $txtMailLog.Text = 'Edit cancelled.'; return }
            $newTrim = $newValue.Trim()

            if ($script:MailAttribsNeverBlank -contains $attrName -and $newTrim -eq '') {
                $null = Show-ToolDialog("'$attrLabel' cannot be blank.", 'Invalid Input', 'OK', 'Warning')
                return
            }
            if ($newTrim -eq '' -and "$oldValue".Trim() -ne '') {
                $null = Show-ToolDialog("'$attrLabel' currently has a value and cannot be saved blank. Enter a replacement value instead.", 'Invalid Input', 'OK', 'Warning')
                return
            }
            # Length caps are Exchange's, and hitting one is a confusing error
            # from the cmdlet. Checked here so the message names the limit.
            if ($script:MailAttribsMaxLength.ContainsKey($attrName) -and $newTrim.Length -gt $script:MailAttribsMaxLength[$attrName]) {
                $null = Show-ToolDialog("'$attrLabel' is limited to $($script:MailAttribsMaxLength[$attrName]) characters. That value is $($newTrim.Length).", 'Too Long', 'OK', 'Warning')
                return
            }
            if ($attrName -eq 'PrimarySmtpAddress' -and $newTrim -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                $null = Show-ToolDialog("'$newTrim' does not look like a valid email address (expected user@domain.tld).", 'Invalid Input', 'OK', 'Warning')
                return
            }
            if ($newTrim -eq "$oldValue".Trim()) { $txtMailLog.Text = 'No change -- the value is the same.'; return }

            # ---- the email address policy gate -------------------------------
            #
            # ASKED BEFORE THE CHANGE IS CONFIRMED, NOT AFTER IT FAILS. While
            # the policy is on, Exchange owns this account's addresses. Setting
            # a primary by hand means detaching it, and that is permanent as far
            # as this tool is concerned -- there is no re-attach button here.
            # Doing it silently as part of a rename would be indefensible.
            $detachPolicy = $false
            if ($attrName -eq 'PrimarySmtpAddress' -and $u.EmailAddressPolicyEnabled) {
                $ans = Show-ToolDialog(
                    ("$($u.UserPrincipalName) is subject to an EMAIL ADDRESS POLICY." +
                     "`r`n`r`nExchange generates this account's addresses from that policy. To set the primary address by hand, the account has to be detached from the policy first." +
                     "`r`n`r`nAFTER THAT, this account's email addresses are maintained BY HAND. Policy changes made for everyone else will not reach it, and this tool has no way to re-attach it." +
                     "`r`n`r`nDetach it from the policy and set the primary address?"),
                    'Email Address Policy', 'YesNo', 'Warning')
                if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
                    $txtMailLog.Text = 'Cancelled -- the account is still attached to its email address policy and the primary address was not changed.'
                    return
                }
                $detachPolicy = $true
            }

            $keepOldPrimaryAsAlias = $true
            if ($attrName -eq 'PrimarySmtpAddress') {
                $oldPrimaryTrim = "$oldValue".Trim()
                if ($oldPrimaryTrim -ne '' -and -not ($oldPrimaryTrim -ieq $newTrim)) {
                    $choice = Show-KeepOrRemoveOldPrimaryDialog -Owner $mailForm -OldPrimary $oldPrimaryTrim -NewPrimary $newTrim
                    if ($choice -ne 'Keep' -and $choice -ne 'Remove') { $txtMailLog.Text = 'Edit cancelled.'; return }
                    $keepOldPrimaryAsAlias = ($choice -eq 'Keep')
                }
            }

            $confirmExtra = switch ($attrName) {
                'PrimarySmtpAddress' {
                    $(if ($keepOldPrimaryAsAlias) { "`r`n`r`nThe EmailAddresses collection is rewritten: this becomes SMTP: (primary) and the previous primary is kept as smtp: (secondary alias)." }
                      else { "`r`n`r`nThe EmailAddresses collection is rewritten: this becomes SMTP: (primary) and the previous primary is REMOVED entirely. Mail sent to it will bounce." }) +
                    $(if ($detachPolicy) { "`r`n`r`nThe account will first be DETACHED from its email address policy." } else { '' })
                }
                default { '' }
            }
            $confirm = Show-ToolDialog(
                "Change '$attrLabel' for $($u.UserPrincipalName) (on-premises):`r`n`r`nFrom: $oldValue`r`nTo:   $newTrim$confirmExtra",
                'Confirm Attribute Edit', 'YesNo', 'Question')
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtMailLog.Text = 'Cancelled.'; return }

            $mailForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $dc = Get-MailAttribsWriteDc -DistinguishedName $u.DistinguishedName
            if ($dc) { Write-ToolStep "Writing on $dc ..." } else { Write-ToolStep 'Writing ...' }

            if ($detachPolicy) {
                Disable-MailAttribsAddressPolicy -Identity $u.Identity -Route $u.Route -Dc $dc
                Write-CompletedLog -Action 'MailAttribsPolicyDetach' -Target (Get-LogTarget $u.UserPrincipalName $u.Identity) -Result 'Success' `
                    -Detail 'EmailAddressPolicyEnabled = $false (required to set the primary address by hand)'
            }

            Set-MailAttribsAttribute -Identity $u.Identity -AttributeName $attrName -NewValue $newTrim `
                -CurrentProxyAddresses $u.EmailAddresses -KeepOldPrimaryAsAlias $keepOldPrimaryAsAlias `
                -Route $u.Route -Dc $dc

            Write-ToolStep 'Reading the object back ...'
            $refreshed = Get-MailAttribsUser -Smtp $script:Target
            $script:MailAttribsCurrentUser = $refreshed
            $afterValue = if ($attrName -eq 'PrimarySmtpAddress') {
                Get-PrimarySmtpFromProxyAddresses -ProxyAddresses $refreshed.EmailAddresses
            } else {
                Format-AttributeValue -Value ($refreshed.$attrName)
            }
            Update-MailAttribsGrid -UserObject $refreshed

            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine('--- BEFORE ---')
            [void]$log.AppendLine("$attrLabel = $oldValue")
            [void]$log.AppendLine('--- AFTER ---')
            [void]$log.AppendLine("$attrLabel = $afterValue")
            if ($attrName -eq 'PrimarySmtpAddress') {
                [void]$log.AppendLine('--- EmailAddresses now ---')
                [void]$log.AppendLine((Format-AttributeValue -Value $refreshed.EmailAddresses))
            }
            [void]$log.AppendLine("Written on: $(if ($dc) { $dc } else { 'the session default DC' })")
            [void]$log.AppendLine('This is an on-premises change. It reaches the cloud on the next directory sync.')
            $txtMailLog.Text = $log.ToString().TrimEnd()

            $detail = "$attrName : '$oldValue' -> '$afterValue'"
            if ($attrName -eq 'PrimarySmtpAddress') {
                $detail += "; old primary $(if ($keepOldPrimaryAsAlias) { 'kept as alias' } else { 'REMOVED' })"
            }
            if ($dc) { $detail += "; on $dc" }
            Write-CompletedLog -Action "MailAttribsSet$attrName" -Target (Get-LogTarget $refreshed.UserPrincipalName $script:Target) -Result 'Success' -Detail $detail
        } catch {
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs edit'
            Show-HandlerError -ErrorRecord $_ -Context 'Mail Attributes edit (on-premises)'
            Write-FailedLog -Action 'MailAttribsSet' -Target (Get-LogTarget $script:Target) -ErrorRecord $_
        } finally {
            $mailForm.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
            Set-MailAttribsEditEnabled
        }
    }

    $grdMail.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $grdMail.ClearSelection()
        $grdMail.Rows[$e.RowIndex].Selected = $true
        Invoke-MailAttribsEditSelectedRow
    })
    $grdMail.Add_SelectionChanged({ Set-MailAttribsEditEnabled })
    $btnMailEdit.Add_Click({ Invoke-MailAttribsEditSelectedRow })

    $btnAddSmtp.Add_Click({
        if ($null -eq $script:MailAttribsCurrentUser) {
            $null = Show-ToolDialog('Load an account first (click Refresh).', 'No User Loaded', 'OK', 'Warning')
            return
        }
        $btnMailRefresh.Enabled = $false
        $btnAddSmtp.Enabled = $false
        try {
            $u = $script:MailAttribsCurrentUser
            $newAddr = $txtAddSmtp.Text.Trim()
            if ($newAddr -eq '') {
                $null = Show-ToolDialog('Enter an address to add.', 'Missing Input', 'OK', 'Warning')
                return
            }
            if ($newAddr -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                $null = Show-ToolDialog("'$newAddr' does not look like a valid email address (expected user@domain.tld).", 'Invalid Input', 'OK', 'Warning')
                return
            }
            $beforeValue = Format-AttributeValue -Value $u.EmailAddresses
            $confirm = Show-ToolDialog(
                ("Add 'smtp:$newAddr' as a secondary address for $($u.UserPrincipalName) (on-premises)?" +
                 "`r`n`r`nThis does not change the primary address or remove any existing address." +
                 $(if ($u.EmailAddressPolicyEnabled) { "`r`n`r`nNote: this account is subject to an email address policy, which may add or remove addresses of its own." } else { '' })),
                'Confirm Add Secondary SMTP Address', 'YesNo', 'Question')
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtMailLog.Text = 'Cancelled.'; return }

            $mailForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $dc = Get-MailAttribsWriteDc -DistinguishedName $u.DistinguishedName
            Write-ToolStep 'Adding the address ...'
            Add-MailAttribsSecondaryProxyAddress -Identity $u.Identity -CurrentProxyAddresses $u.EmailAddresses `
                -NewSecondary $newAddr -Route $u.Route -Dc $dc

            $refreshed = Get-MailAttribsUser -Smtp $script:Target
            $script:MailAttribsCurrentUser = $refreshed
            Update-MailAttribsGrid -UserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine('--- BEFORE (EmailAddresses) ---')
            [void]$log.AppendLine($beforeValue)
            [void]$log.AppendLine('--- AFTER ---')
            [void]$log.AppendLine((Format-AttributeValue -Value $refreshed.EmailAddresses))
            [void]$log.AppendLine('This is an on-premises change. It reaches the cloud on the next directory sync.')
            $txtMailLog.Text = $log.ToString().TrimEnd()
            Write-CompletedLog -Action 'MailAttribsAddSecondary' -Target (Get-LogTarget $refreshed.UserPrincipalName $script:Target) -Result 'Success' `
                -Detail "added smtp:$newAddr$(if ($dc) { "; on $dc" })"
            $txtAddSmtp.Text = ''
        } catch {
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add Secondary SMTP'
            Show-HandlerError -ErrorRecord $_ -Context 'Add secondary SMTP address (on-premises)'
            Write-FailedLog -Action 'MailAttribsAddSecondary' -Target (Get-LogTarget $script:Target) -ErrorRecord $_
        } finally {
            $mailForm.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
            Set-MailAttribsEditEnabled
        }
    })

    $btnMailClose.Add_Click({ $mailForm.Close() })
    $mailForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $mailForm
    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down
    # and grows the form, and the help button places itself from ClientSize at
    # run time.
    $bar = Add-ToolTargetBar -Form $mailForm -Label 'Mailbox or address:' -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-MailAttribsFromCurrent
    })

    Add-ToolHelpButton -Form $mailForm -ScriptPath $PSCommandPath -Title 'Mail Attributes (Rename)' | Out-Null
    Set-FormScreenCorner -Form $mailForm -Corner 'BottomLeft'
    [void]$mailForm.ShowDialog()
}
#endregion

Show-MailAttribsForm
