#Requires -Version 5.1
# TOOLKIT-VARIANT: Dual
#
# ONE WINDOW FOR A PERSON -- the account, not the mailbox.
#
# The Mailbox Editor answers "what is going on with this mailbox". This answers
# "what is going on with this PERSON": is the account enabled, is it locked,
# when does the password expire, what groups are they in, are they licensed,
# and which mail object -- if any -- hangs off them. A mailbox is one facet of
# a person; the account is another. Putting account state into the Mailbox
# Editor is how that window becomes the everything-window, so it lives here.
#
# PURE POWERSHELL, ON PURPOSE. Everything on these tabs comes from Active
# Directory and from what Exchange itself knows about the user. Nothing here
# needs Microsoft Graph, because this ships to a site that does not have it
# yet. The Entra-only facets -- sign-in activity, MFA registration, risk state
# -- are NOT here, and the help says so rather than showing an empty tab.
# An Entra tab was built and taken out again on 2026-09-18 pending Graph being
# available everywhere this ships; it is a separate piece of work.
<#
    PersonEditor.ps1
    A person's account, directory fields, group memberships and licence state,
    in tabs, on the shared tabbed-editor runtime.                  Author : Glenn Armbruster
    Created  : 2026-09-18
#>
#
# TOOLKIT-HELP-RUNS: Reads one person from Active Directory (and what Exchange knows about them) across several tabs. Writes go to Active Directory. Each section's Apply changes only that section, and the risky ones ask first.
# TOOLKIT-HELP-BEGIN
# Everything about one PERSON, in tabs. Type an address, UPN or account name,
# press Load.
#
#   Overview   Who they are: name, UPN, account name, which directory, the
#              job fields (title, department, office, company, manager) with
#              an Apply, and which mail object exists for them -- with a
#              button that opens it in the Mailbox Editor. The job fields are
#              editable either way: Set-ADUser where there is an on-premises
#              account, Set-User through Exchange Online where there is not.
#   Account    The account itself: enabled or disabled, locked out, password
#              last set and when it expires, last logon, created, the OU it
#              sits in. Unlock, enable and disable are here.
#              WITH NO ACTIVE DIRECTORY those three do not apply -- sign-in
#              state and lockout are Entra properties and need Graph -- so
#              they are replaced by the one thing Exchange Online can do:
#              require a password change at the next sign-in.
#   Groups     Every group the person is a member of, from Active Directory,
#              with type, scope and whether it is mail-enabled. Add to a
#              group, remove from one, or open a mail-enabled one in the Group
#              Editor. This is the reverse of the Group Editor's question.
#   Licence    Whether they are licensed, decided the way this tenant decides
#              it -- membership of a licence group -- and which group. Change
#              it on the Groups tab; this tab only reports.
#
# WHERE WRITES GO DEPENDS ON WHERE THE PERSON LIVES. A synced person is
# mastered on-premises, so the Editing switch is set to On-premises and fixed
# there and writes use the ActiveDirectory module. A person with no
# on-premises account -- every person in a cloud-only tenant -- has their job
# fields written straight to Exchange Online with Set-User instead. The window
# is no longer read-only in that case; it was, and that was wrong.
#
# WHAT IS NOT HERE, AND WHY. Sign-in activity, MFA registration and risk
# state live in Entra ID and need Microsoft Graph. This window is built to run
# where Graph is not available, so those are left out rather than shown blank.
# The Sign-in Activity, Auth Methods and Risky Users tools cover them where
# Graph is connected.
#
# WHAT IT WILL NOT DO: reset a password, or run the leaver process. Resetting
# a password is a credential operation with its own window -- Password
# Manager, from the same Unified Lookup row -- which also checks the lockout
# on every domain controller; the leaver process is Term Requests, which does
# the whole job in the right order. Disabling an account here is one step of that process,
# and the confirmation says so.
#
# EACH BLOCK HAS ITS OWN APPLY. There is no global Save.
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

$script:Person        = $null    # the resolved person -- see Resolve-PersonTarget
$script:Busy          = $false
$script:Closing       = $false
$script:Loaded        = @{}
$script:Mode          = 'OnPrem'
$script:TabReadFailed = $false
$script:OnPremWhy     = ''
$script:AdReady       = $null    # Test-AdAvailable, once per window
$script:Licence       = $null

#region Resolve -----------------------------------------------------------------
function Get-PersonAdProperties {
    # The attributes every tab reads, asked for once. lastLogonTimestamp rather
    # than lastLogon because the latter is per-controller and not replicated.
    @('DisplayName', 'GivenName', 'Surname', 'UserPrincipalName', 'SamAccountName', 'mail', 'proxyAddresses',
      'Enabled', 'LockedOut', 'PasswordLastSet', 'PasswordNeverExpires', 'PasswordExpired', 'CannotChangePassword',
      'AccountExpirationDate', 'LastLogonDate', 'lastLogonTimestamp', 'whenCreated', 'whenChanged',
      'Title', 'Department', 'Company', 'Office', 'Manager', 'Description', 'EmployeeID', 'telephoneNumber', 'MobilePhone',
      'DistinguishedName', 'ObjectGUID', 'msExchRecipientTypeDetails', 'msExchRemoteRecipientType', 'memberOf',
      'msDS-UserPasswordExpiryTimeComputed', 'userAccountControl', 'SID', 'primaryGroupID',
      'homeMDB', 'msExchResourceSearchProperties')   # what Resolve-AdObjectType reads to tell a mailbox from a mail user
}

function Resolve-PersonTarget {
    <#
        Find the person, in Active Directory first.

        ONE OBJECT, THREE WAYS IN. An address, a UPN, or a bare account name
        all identify a person and all arrive here, so the LDAP filter accepts
        any of them. Every configured directory is asked in turn, the way
        Resolve-ADAccountBySmtp does, and the SERVER that answered is kept --
        every later read and write goes back to the same controller, so a
        change made here is read back from where it was made rather than from
        a replica that has not heard yet.

        THE CLOUD SHADOW IS OPTIONAL. Get-User in Exchange Online tells us what
        recipient type they are and whether a mailbox exists; if Exchange is
        not connected, the window still works from Active Directory alone.

        NO ON-PREMISES DIRECTORY AT ALL is a real configuration (cloud-only
        tenants). Then the cloud object is all there is, the account tab says
        so, and nothing here can write.
    #>
    param([Parameter(Mandatory)][string]$Identity)
    $r = [PSCustomObject]@{
        Ok = $false; Smtp = ''; Display = ''; Upn = ''; Sam = ''; Kind = ''
        Directory = ''; Server = ''; Dn = ''; Guid = ''
        Ad = $null; Cloud = $null; IsSynced = $false; Reason = ''
    }
    $t = "$Identity".Trim()
    if (-not $t) { $r.Reason = 'Enter an address, a UPN or an account name.'; return $r }

    # ---- Active Directory --------------------------------------------------
    if ($null -eq $script:AdReady) { $script:AdReady = Test-AdAvailable }
    if ($script:AdReady.Available) {
        $esc = ConvertTo-LdapFilterValue -Value $t
        $filter = "(&(objectCategory=person)(objectClass=user)(|(mail=$esc)(userPrincipalName=$esc)(sAMAccountName=$esc)(proxyAddresses=smtp:$esc)))"
        $errors = @()
        foreach ($d in @(Get-TenantDirectory)) {
            $srv = Get-TenantDirectoryServer -Id $d.Id
            if (-not $srv) { continue }
            try {
                $hits = @(Get-ADUser -LDAPFilter $filter -Server $srv -Properties (Get-PersonAdProperties) -ErrorAction Stop)
            } catch { $errors += "$($d.Id): $($_.Exception.Message)"; continue }
            if ($hits.Count -eq 0) { continue }
            if ($hits.Count -gt 1) {
                $r.Reason = "'$t' matches $($hits.Count) accounts in $($d.Id). Use the UPN or the account name."
                return $r
            }
            $u = $hits[0]
            $r.Ad = $u; $r.Directory = "$($d.Id)"; $r.Server = $srv
            $r.Display = "$($u.DisplayName)"; $r.Upn = "$($u.UserPrincipalName)"; $r.Sam = "$($u.SamAccountName)"
            $r.Dn = "$($u.DistinguishedName)"; $r.Guid = "$($u.ObjectGUID)"
            $r.Smtp = "$($u.mail)"
            if (-not $r.Smtp) {
                $p = @($u.proxyAddresses | Where-Object { "$_" -cmatch '^SMTP:' } | Select-Object -First 1)
                if ($p.Count) { $r.Smtp = "$($p[0])".Substring(5) }
            }
            $r.Kind = Resolve-AdObjectType -AdObject $u
            $r.IsSynced = $true
            break
        }
        if (-not $r.Ad -and $errors.Count -gt 0 -and $errors.Count -eq @(Get-TenantDirectory).Count) {
            $r.Reason = "Active Directory could not be asked: $($errors[0])"
            return $r
        }
    }

    # ---- Exchange Online, for the recipient shape ----------------------------
    $lookup = $(if ($r.Smtp) { $r.Smtp } elseif ($r.Upn) { $r.Upn } else { $t })
    try {
        $c = Invoke-EXOCommand -Name 'Get-User' -Parameters @{ Identity = $lookup; ErrorAction = 'Stop' }
        if ($c) {
            $r.Cloud = $c
            if (-not $r.Display) { $r.Display = "$($c.DisplayName)" }
            if (-not $r.Upn)     { $r.Upn = "$($c.UserPrincipalName)" }
            if (-not $r.Smtp)    { $r.Smtp = "$($c.WindowsEmailAddress)" }
            if (-not $r.Kind)    { $r.Kind = "$($c.RecipientTypeDetails)" }
            if (-not $r.Ad -and $c.PSObject.Properties['IsDirSynced']) { $r.IsSynced = [bool]$c.IsDirSynced }
        }
    } catch { }

    if (-not $r.Ad -and -not $r.Cloud) {
        $r.Reason = $(if ($script:AdReady.Available) { "'$t' was not found in Active Directory or Exchange Online." }
                      else { "'$t' was not found in Exchange Online, and there is no on-premises directory to ask ($($script:AdReady.Detail))." })
        return $r
    }
    if (-not $r.Display) { $r.Display = $t }
    if (-not $r.Kind)    { $r.Kind = 'User' }
    $r.Ok = $true
    return $r
}

function Invoke-PersonAd {
    <#
        Every AD cmdlet goes through here so it is pinned to the controller
        the person was found on. A write to one controller read back from
        another is the 'nothing changed' report that is not true.

        -ForDn OVERRIDES THAT PIN with the domain the object actually lives
        in. memberOf routinely crosses domains in a forest, and a group in a
        child domain asked for on the parent's controller is 'cannot find an
        object with identity' -- it is there, it is just not HERE.
    #>
    param([Parameter(Mandatory)][string]$Name, [hashtable]$Parameters = @{}, [string]$ForDn)
    $p = @{} + $Parameters
    $srv = $(if ($script:Person -and $script:Person.Server) { $script:Person.Server } else { '' })
    if ($ForDn) {
        $d = Get-DirectoryFromDn -Dn $ForDn
        if ($d) { $s2 = Get-TenantDirectoryServer -Id $d.Id; if ($s2) { $srv = $s2 } }
    }
    if ($srv) { $p['Server'] = $srv }
    $p['ErrorAction'] = 'Stop'
    return (& $Name @p)
}

function Test-PersonWritable {
    param([string]$What = 'That')
    if (-not $script:Person -or -not $script:Person.Ad) {
        # 'FOUND ONLY IN EXCHANGE ONLINE' READS AS A LOOKUP FAILURE in a tenant
        # that has nowhere else to look. Same refusal, honest reason.
        $why = 'this person was found only in Exchange Online'
        try { if (Test-TenantIsCloudOnly) { $why = 'this tenant has no Active Directory' } } catch { }
        Set-EditorStatus -Text "$What needs an Active Directory account, and $why." -Bad
        return $false
    }
    return $true
}
#endregion

#region Window ------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Person Editor'
$form.ClientSize    = New-Object System.Drawing.Size(1120, 760)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$y = 12
$lblWho          = New-Object System.Windows.Forms.Label
$lblWho.Text     = 'Person:'
$lblWho.Location = New-Object System.Drawing.Point(12, ($y + 3))
$lblWho.Size     = New-Object System.Drawing.Size(48, 18)
$form.Controls.Add($lblWho)

$txtWho          = New-Object System.Windows.Forms.TextBox
$txtWho.Location = New-Object System.Drawing.Point(62, $y)
$txtWho.Size     = New-Object System.Drawing.Size(340, 24)
$txtWho.Font     = New-Object System.Drawing.Font('Consolas', 9)
$txtWho.Text     = "$SmtpAddress".Trim()
$form.Controls.Add($txtWho)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load'
$btnLoad.Location = New-Object System.Drawing.Point(412, ($y - 2))
$btnLoad.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnLoad)

# THE SAME SWITCH THE OTHER EDITORS HAVE, so the window reads the same -- but
# FIXED. A person is mastered in Active Directory; there is nothing to write in
# Exchange Online for a user object here. So it shows On-premises and is
# disabled, and the sync banner explains. Where there is no AD it shows Cloud,
# disabled, and everything is read-only.
$cmbMode = New-EditorModeSwitch -Form $form -X 718 -Y $y -LabelX 660

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
$lblIdent.Text      = 'No person loaded.'
$form.Controls.Add($lblIdent)
$y += 20

$lblSync           = New-Object System.Windows.Forms.Label
$lblSync.Tag       = 'NoTheme'
$lblSync.Location  = New-Object System.Drawing.Point(12, $y)
$lblSync.Size      = New-Object System.Drawing.Size(1096, 18)
$lblSync.ForeColor = (Get-ToolStatusColour -Kind Muted)
$lblSync.Text      = ''
$form.Controls.Add($lblSync)
$y += 22

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, $y)
$tabs.Size     = New-Object System.Drawing.Size(1096, 640)
$form.Controls.Add($tabs)

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(12, ($y + 646))
$lblStatus.Size      = New-Object System.Drawing.Size(1096, 18)
$lblStatus.ForeColor = (Get-ToolStatusColour -Kind Muted)
$lblStatus.Text      = 'Enter an address, UPN or account name and press Load.'
$form.Controls.Add($lblStatus)
$form.ClientSize = New-Object System.Drawing.Size(1120, ($lblStatus.Bottom + 12))

# ---- register with the shared editor runtime ----------------------------------
Initialize-TabbedEditor -Name 'Person' -Noun 'person' -ObjectName 'Person' `
    -Form $form -Tabs $tabs -StatusLabel $lblStatus -IdentityLabel $lblIdent -SyncLabel $lblSync -ModeCombo $cmbMode `
    -UsesOnPremObject $false `
    -TabLoaders @{
        Overview = { Update-OverviewTab }
        Account  = { Update-AccountTab }
        Groups   = { Update-GroupsTab }
        Licence  = { Update-LicenceTab }
    } `
    -Identity {
        param($p)
        $where = $(if ($p.Ad) { "Active Directory ($($p.Directory))" } else { 'Exchange Online only' })
        # THREE DIFFERENT SITUATIONS, NOT TWO. The old wording had one message
        # for 'no on-premises account' and it said the job fields came from
        # Active Directory and could not be changed. In a CLOUD-ONLY tenant
        # both halves of that are false: the fields live in Entra, Exchange
        # surfaces them through Get-User, and Set-User writes them. Reported
        # 2026-09-19 -- a red banner claiming a directory this tenant has never
        # had, over boxes that turn out to be editable after all.
        $cloudTenant = $false
        try { $cloudTenant = Test-TenantIsCloudOnly } catch { $cloudTenant = $false }
        $sync = $(if ($p.Ad) { "Writes go to Active Directory via $($p.Server). Exchange Online sees a change after the next directory sync -- allow about half an hour." }
                  elseif ($cloudTenant) { 'Cloud tenant. The job fields below are written straight to Exchange Online with Set-User and take effect at once. Sign-in state, lockout and licence live in Entra and need Graph, which is not configured here.' }
                  else { 'No on-premises account was found for this person. The job fields below are written to Exchange Online with Set-User; account state, groups and licence are mastered on-premises and are not available for this one.' })
        @{ Text = "$($p.Display)   <$(if ($p.Upn) { $p.Upn } else { $p.Smtp })>   $($p.Kind)   [$where]"; Colour = 'Good'
           Sync = $sync; SyncColour = $(if ($p.Ad) { 'Warn' } elseif ($cloudTenant) { 'Muted' } else { 'Warn' }) }
    }

$tabOverview = New-EditorTab -Key 'Overview' -Text 'Overview'
$tabAccount  = New-EditorTab -Key 'Account'  -Text 'Account'
$tabGroups   = New-EditorTab -Key 'Groups'   -Text 'Groups'
$tabLicence  = New-EditorTab -Key 'Licence'  -Text 'Licence'
#endregion

#region Overview tab ------------------------------------------------------------
$ovY = 8
[void](New-TabLabel $tabOverview 'IDENTITY' 8 $ovY -Bold)
$ovY += 22
[void](New-TabLabel $tabOverview 'Display name:' 8 $ovY)
$ovName = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Kind:' 560 $ovY 135)
$ovKind = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'User ID (UPN):' 8 $ovY)
$ovUpn = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Account name:' 560 $ovY 135)
$ovSam = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Primary address:' 8 $ovY)
$ovSmtp = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Directory:' 560 $ovY 135)
$ovDir = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Object GUID:' 8 $ovY)
$ovGuid = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Employee ID:' 560 $ovY 135)
$ovEmp = New-TabValue $tabOverview 700 $ovY 320
$ovY += 32

# THE JOB FIELDS ARE EDITABLE HERE, because a person editor that cannot change
# a title is a viewer. They are Active Directory attributes and only Active
# Directory; Exchange Online shows them after sync.
[void](New-TabLabel $tabOverview 'DIRECTORY' 8 $ovY -Bold)
$ovDirNote = New-TabLabel $tabOverview '' 175 $ovY 860 -Muted
$ovY += 22
[void](New-TabLabel $tabOverview 'Title:' 8 $ovY)
$txtTitle = New-Object System.Windows.Forms.TextBox; $txtTitle.Location = New-Object System.Drawing.Point(170, $ovY); $txtTitle.Size = New-Object System.Drawing.Size(360, 22); $tabOverview.Controls.Add($txtTitle)
[void](New-TabLabel $tabOverview 'Department:' 560 $ovY 135)
$txtDept = New-Object System.Windows.Forms.TextBox; $txtDept.Location = New-Object System.Drawing.Point(700, $ovY); $txtDept.Size = New-Object System.Drawing.Size(320, 22); $tabOverview.Controls.Add($txtDept)
$ovY += 26
[void](New-TabLabel $tabOverview 'Office:' 8 $ovY)
$txtOffice = New-Object System.Windows.Forms.TextBox; $txtOffice.Location = New-Object System.Drawing.Point(170, $ovY); $txtOffice.Size = New-Object System.Drawing.Size(360, 22); $tabOverview.Controls.Add($txtOffice)
[void](New-TabLabel $tabOverview 'Company:' 560 $ovY 135)
$txtCompany = New-Object System.Windows.Forms.TextBox; $txtCompany.Location = New-Object System.Drawing.Point(700, $ovY); $txtCompany.Size = New-Object System.Drawing.Size(320, 22); $tabOverview.Controls.Add($txtCompany)
$ovY += 26
[void](New-TabLabel $tabOverview 'Manager:' 8 $ovY)
$txtMgr = New-Object System.Windows.Forms.TextBox; $txtMgr.Location = New-Object System.Drawing.Point(170, $ovY); $txtMgr.Size = New-Object System.Drawing.Size(360, 22); $tabOverview.Controls.Add($txtMgr)
$ovMgrNote = New-TabLabel $tabOverview '' 560 ($ovY + 3) 460 -Muted
$ovY += 26
[void](New-TabLabel $tabOverview 'Phone:' 8 $ovY)
$txtPhone = New-Object System.Windows.Forms.TextBox; $txtPhone.Location = New-Object System.Drawing.Point(170, $ovY); $txtPhone.Size = New-Object System.Drawing.Size(360, 22); $tabOverview.Controls.Add($txtPhone)
[void](New-TabLabel $tabOverview 'Mobile:' 560 $ovY 135)
$txtMobile = New-Object System.Windows.Forms.TextBox; $txtMobile.Location = New-Object System.Drawing.Point(700, $ovY); $txtMobile.Size = New-Object System.Drawing.Size(320, 22); $tabOverview.Controls.Add($txtMobile)
$ovY += 30
$btnDirApply = New-TabButton $tabOverview 'Apply directory fields' 170 $ovY 180
$ovY += 40

[void](New-TabLabel $tabOverview 'MAIL OBJECT' 8 $ovY -Bold)
$ovY += 22
$ovMail = New-TabLabel $tabOverview '' 8 ($ovY + 3) 700 -Muted
$btnOpenMailbox = New-TabButton $tabOverview 'Open in Mailbox Editor' 720 ($ovY - 2) 180
$ovY += 38

$ovVerdict = Add-EditorVerdictBox -Tab $tabOverview -Y $ovY -Height 150
$ovY += 182

function Get-PersonManagerText {
    # Manager arrives as a DN. Shown as a display name with the account in
    # brackets so it can be read and, when re-entered, still resolves.
    param([string]$Dn)
    if (-not $Dn) { return '' }
    try {
        $m = Invoke-PersonAd -Name 'Get-ADUser' -Parameters @{ Identity = $Dn; Properties = @('DisplayName') }
        return "$($m.SamAccountName)"
    } catch { return $Dn }
}

function Update-OverviewTab {
    $p = $script:Person
    $ovName.Text = $p.Display; $ovKind.Text = $p.Kind; $ovUpn.Text = $p.Upn; $ovSam.Text = $p.Sam
    $ovSmtp.Text = $(if ($p.Smtp) { $p.Smtp } else { '(none)' })
    $ovDir.Text  = $(if ($p.Ad) { "$($p.Directory) via $($p.Server)" } else { 'Exchange Online (no on-premises account)' })
    $ovGuid.Text = $p.Guid
    $u = $p.Ad
    $cloudOnlyTenant = $false
    try { $cloudOnlyTenant = Test-TenantIsCloudOnly } catch { $cloudOnlyTenant = $false }
    $ovEmp.Text = $(if ($u) { "$($u.EmployeeID)" } else { '' })

    # EDITABLE EITHER WAY. Set-ADUser writes these when there is an
    # on-premises account; Set-User writes the same seven through Exchange
    # Online when there is not -- Title, Department, Office, Company, Manager,
    # Phone and MobilePhone are all documented as writable in Exchange Online.
    # The only thing that turns the boxes off now is having no object at all.
    $can = [bool]($u -or $p.Cloud)
    foreach ($t in @($txtTitle, $txtDept, $txtOffice, $txtCompany, $txtMgr, $txtPhone, $txtMobile)) { $t.Enabled = $can }
    $btnDirApply.Enabled = $can
    if ($u) {
        $txtTitle.Text = "$($u.Title)"; $txtDept.Text = "$($u.Department)"; $txtOffice.Text = "$($u.Office)"
        $txtCompany.Text = "$($u.Company)"; $txtMgr.Text = (Get-PersonManagerText -Dn "$($u.Manager)")
        $txtPhone.Text = "$($u.telephoneNumber)"; $txtMobile.Text = "$($u.MobilePhone)"
        $ovMgrNote.Text = 'Account name of the manager. Blank to clear.'
        $ovDirNote.Text = "Active Directory attributes -- written with Set-ADUser against $($p.Server)."
    } else {
        $c = $p.Cloud
        $txtTitle.Text = "$($c.Title)"; $txtDept.Text = "$($c.Department)"; $txtOffice.Text = "$($c.Office)"
        $txtCompany.Text = "$($c.Company)"; $txtMgr.Text = "$($c.Manager)"
        $txtPhone.Text = "$($c.Phone)"; $txtMobile.Text = "$($c.MobilePhone)"
        $ovMgrNote.Text = 'Display name or address of the manager. Blank to clear.'
        $ovDirNote.Text = 'Exchange Online attributes -- written with Set-User. The change takes effect at once.'
    }
    $ovDirNote.ForeColor = (Get-ToolStatusColour -Kind Muted)

    # ---- which mail object exists, and the door to it -----------------------
    $kind = "$($p.Kind)"
    $hasMailbox = ($kind -match '(?i)Mailbox')
    $ovMail.Text = $(if ($kind -match '(?i)^Remote.*Mailbox') { "$kind -- the mailbox is in Exchange Online; the on-premises object is the stub that routes to it." }
                     elseif ($kind -match '(?i)Mailbox') { "$kind." }
                     elseif ($kind -match '(?i)MailUser') { 'Mail user -- mail-enabled, no mailbox. Mail is delivered elsewhere.' }
                     else { 'No mailbox and no mail-enabled object. Unified Lookup can enable one.' })
    $btnOpenMailbox.Enabled = ($hasMailbox -and [bool]$p.Smtp)

    # ---- verdict, from what this tab and the account already read ----------
    $f = New-Object System.Collections.ArrayList
    if ($u) {
        if (-not $u.Enabled) { [void]$f.Add("The account is DISABLED.$(if ($kind -match '(?i)UserMailbox') { ' It still has a USER mailbox -- a leaver nobody converted to shared, still consuming a licence.' })") }
        if ($u.LockedOut) { [void]$f.Add('The account is LOCKED OUT. See Account to unlock it.') }
        if ($u.PasswordExpired) { [void]$f.Add('The password has EXPIRED. The person cannot sign in until it is changed.') }
        if ($u.AccountExpirationDate -and $u.AccountExpirationDate -lt (Get-Date)) { [void]$f.Add("The account EXPIRED on $($u.AccountExpirationDate.ToString('yyyy-MM-dd')).") }
        if ($u.LastLogonDate -and $u.LastLogonDate -lt (Get-Date).AddDays(-90) -and $u.Enabled) { [void]$f.Add("No logon in over 90 days (last $($u.LastLogonDate.ToString('yyyy-MM-dd'))) on an enabled account.") }
        if (-not $u.Manager -and $kind -match '(?i)User') { [void]$f.Add('No manager set. Approvals and escalations that route to a manager will not route.') }
    }
    $notes = @()
    if (-not $u) {
        $notes += $(if ($cloudOnlyTenant) {
                        'Sign-in state, lockout, password expiry and licence are Entra properties. They need Microsoft Graph, which is not configured for this tenant, so they are not shown.'
                    } else {
                        'Account state, groups and licence need the on-premises account and are not available for this person.'
                    })
    }
    Set-EditorVerdict -Findings @($f) -Notes $notes -Box $ovVerdict `
        -Clean 'Nothing obviously wrong. Enabled, not locked, password current, has a manager, and has logged on recently.'
}

function Invoke-PersonCloudFieldApply {
    <#
        The same seven job fields, written through Exchange Online instead of
        Active Directory.

        WHY THIS EXISTS. A person in a cloud-only tenant has no AD object, and
        until 2026-09-19 that meant the boxes were greyed and a red banner said
        the fields came from Active Directory. Both halves were wrong: the
        fields live in Entra, Get-User surfaces them, and Set-User writes them.
        Title, Department, Office, Company, Manager, Phone and MobilePhone are
        all documented as writable in Exchange Online.

        BLANK CLEARS, and it does so the way Exchange expects: Set-User takes
        $null for a cleared string, not an empty one.

        THE MANAGER IS RESOLVED FIRST, for the same reason the on-premises path
        resolves it first -- a manager that cannot be found should fail before
        six other fields have already been written, not after.

        NO DIRECTORY SYNC TO WAIT FOR. This tenant has nowhere for the value to
        come back from, so the confirmation does not promise one.
    #>
    $p = $script:Person
    $c = $p.Cloud
    if (-not $c) { Set-EditorStatus -Text 'Nothing to write to -- this person was not found in Exchange Online.' -Bad; return }

    $changes = @{}; $shown = @()
    foreach ($pair in @(@('Title', $txtTitle, 'Title'), @('Department', $txtDept, 'Department'), @('Office', $txtOffice, 'Office'),
                        @('Company', $txtCompany, 'Company'), @('Phone', $txtPhone, 'Phone'), @('MobilePhone', $txtMobile, 'MobilePhone'))) {
        $name = $pair[0]; $new = "$($pair[1].Text)".Trim(); $old = "$($c.($pair[2]))"
        if ($new -ne $old) {
            $changes[$name] = $(if ($new) { $new } else { $null })
            $shown += "$name`: '$old' -> '$new'"
        }
    }
    $mgrNew = "$($txtMgr.Text)".Trim(); $mgrOld = "$($c.Manager)"
    if ($mgrNew -ne $mgrOld) {
        if ($mgrNew) {
            try { $m = Invoke-EXOCommand -Name 'Get-User' -Parameters @{ Identity = $mgrNew; ErrorAction = 'Stop' } }
            catch { Set-EditorStatus -Text "Manager '$mgrNew' was not found in Exchange Online." -Bad; return }
            $changes['Manager'] = "$($m.Identity)"
        } else {
            $changes['Manager'] = $null
        }
        $shown += "Manager: '$mgrOld' -> '$mgrNew'"
    }
    if ($shown.Count -eq 0) { Set-EditorStatus -Text 'Nothing changed.'; return }

    $a = Show-ToolDialog(("On $($p.Display) in Exchange Online:`r`n`r`n  " + ($shown -join "`r`n  ") + "`r`n`r`nThe change takes effect at once."),
                         'Apply directory fields', 'OKCancel', 'Question')
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $changes['Identity'] = $(if ($p.Upn) { $p.Upn } elseif ($p.Smtp) { $p.Smtp } else { "$($c.Identity)" })
    $changes['ErrorAction'] = 'Stop'
    Invoke-EXOCommand -Name 'Set-User' -Parameters $changes | Out-Null
    Write-EditorAudit -Action 'PersonDirectoryFields' -Environment 'Cloud' -Detail ($shown -join '; ')
    Invoke-PersonReload
    Set-EditorStatus -Text "Applied: $($shown -join '; ')"
}

$btnDirApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        # NO ON-PREMISES ACCOUNT IS NOT 'NOTHING TO DO' ANY MORE. It routes to
        # Exchange Online instead, which is where the fields actually live for
        # a cloud person.
        if (-not $script:Person.Ad) { Invoke-PersonCloudFieldApply; return }
        $u = $script:Person.Ad
        $changes = @{}; $shown = @()
        # Name = the Set-ADUser parameter; Attr = the property it is read from.
        # They differ for the phones: -OfficePhone writes telephoneNumber.
        foreach ($pair in @(@('Title', $txtTitle, 'Title'), @('Department', $txtDept, 'Department'), @('Office', $txtOffice, 'Office'),
                            @('Company', $txtCompany, 'Company'), @('OfficePhone', $txtPhone, 'telephoneNumber'), @('MobilePhone', $txtMobile, 'MobilePhone'))) {
            $name = $pair[0]; $new = "$($pair[1].Text)".Trim(); $old = "$($u.($pair[2]))"
            if ($new -ne $old) {
                # AN EMPTY BOX CLEARS THE ATTRIBUTE, BY PASSING $null TO THE NAMED
                # PARAMETER. Set-ADUser -Title '' fails, and -Clear wants LDAP
                # display names (Office is physicalDeliveryOfficeName there),
                # so a null through the friendly parameter is the one form that
                # is right for all five fields.
                $changes[$name] = $(if ($new) { $new } else { $null })
                $shown += "$name`: '$old' -> '$new'"
            }
        }
        $mgrNew = "$($txtMgr.Text)".Trim(); $mgrOld = (Get-PersonManagerText -Dn "$($u.Manager)")
        if ($mgrNew -ne $mgrOld) {
            if ($mgrNew) {
                # RESOLVED BEFORE IT IS WRITTEN. Set-ADUser -Manager takes an
                # identity and will happily fail after the other fields have
                # already been written, so the manager is looked up first.
                try { $m = Invoke-PersonAd -Name 'Get-ADUser' -Parameters @{ Identity = $mgrNew } }
                catch { Set-EditorStatus -Text "Manager '$mgrNew' was not found in $($script:Person.Directory)." -Bad; return }
                $changes['Manager'] = "$($m.DistinguishedName)"
            } else {
                $changes['Manager'] = $null
            }
            $shown += "Manager: '$mgrOld' -> '$mgrNew'"
        }
        if ($shown.Count -eq 0) { Set-EditorStatus -Text 'Nothing changed.'; return }
        $a = Show-ToolDialog(("On $($script:Person.Display) in Active Directory:`r`n`r`n  " + ($shown -join "`r`n  ") + "`r`n`r`nExchange Online shows the change after the next directory sync."),
                             'Apply directory fields', 'OKCancel', 'Question')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $changes['Identity'] = $u.DistinguishedName
        Invoke-PersonAd -Name 'Set-ADUser' -Parameters $changes | Out-Null
        Write-EditorAudit -Action 'PersonDirectoryFields' -Environment 'OnPrem' -Detail ($shown -join '; ')
        Invoke-PersonReload
        Set-EditorStatus -Text "Applied: $($shown -join '; ')"
    } 'Apply directory fields'
})

$btnOpenMailbox.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    $path = Join-Path $PSScriptRoot 'MailboxEditor.ps1'
    if (-not (Test-Path -LiteralPath $path)) { Set-EditorStatus -Text 'MailboxEditor.ps1 is not installed beside this tool.' -Bad; return }
    # IN-PROCESS, AS UNIFIED LOOKUP DOES IT. The script gets its own scope and
    # its own copy of the runtime, so its $script:Ed is not this window's.
    Invoke-EditorAction {
        & $path -SmtpAddress $script:Person.Smtp
        Set-EditorStatus -Text 'Mailbox Editor closed.'
    } 'Open the Mailbox Editor'
})
#endregion

#region Account tab -------------------------------------------------------------
<#
    THE ACCOUNT ITSELF. Every line here is an Active Directory attribute, and
    every one of them is something a helpdesk call turns on: 'I can't log in'
    is disabled, locked, expired, or a password that lapsed -- and the answer
    is on this tab in one read.

    THREE BUTTONS, NOT SIX. Unlock is safe and common. Enable and Disable are
    real changes to whether a person can work, and Disable in particular is
    ONE STEP of the leaver process -- Term Requests does the whole thing in
    the right order, and the confirmation here says so. There is no password
    reset: a credential operation with its own audit trail belongs in its own
    tool.
#>
$acY = 8
[void](New-TabLabel $tabAccount 'ACCOUNT STATE' 8 $acY -Bold)
$acY += 22
[void](New-TabLabel $tabAccount 'Enabled:' 8 $acY)
$acEnabled = New-TabValue $tabAccount 170 $acY 360
[void](New-TabLabel $tabAccount 'Locked out:' 560 $acY 135)
$acLocked = New-TabValue $tabAccount 700 $acY 320
$acY += 24
[void](New-TabLabel $tabAccount 'Password last set:' 8 $acY)
$acPwdSet = New-TabValue $tabAccount 170 $acY 360
[void](New-TabLabel $tabAccount 'Password expires:' 560 $acY 135)
$acPwdExp = New-TabValue $tabAccount 700 $acY 320
$acY += 24
[void](New-TabLabel $tabAccount 'Last logon:' 8 $acY)
$acLogon = New-TabValue $tabAccount 170 $acY 360
$acLogonNote = New-TabLabel $tabAccount '' 560 ($acY + 3) 460 -Muted
$acLogonNote.Text = 'From lastLogonTimestamp -- replicated, accurate to about two weeks.'
$acY += 24
[void](New-TabLabel $tabAccount 'Account expires:' 8 $acY)
$acExpires = New-TabValue $tabAccount 170 $acY 360
[void](New-TabLabel $tabAccount 'Created:' 560 $acY 135)
$acCreated = New-TabValue $tabAccount 700 $acY 320
$acY += 24
[void](New-TabLabel $tabAccount 'Container (OU):' 8 $acY)
$acOu = New-TabValue $tabAccount 170 $acY 850
$acY += 24
[void](New-TabLabel $tabAccount 'Description:' 8 $acY)
$acDesc = New-TabValue $tabAccount 170 $acY 850
$acY += 34

[void](New-TabLabel $tabAccount 'ACTIONS' 8 $acY -Bold)
$acY += 24
$script:AcActionsY = $acY
$btnUnlock  = New-TabButton $tabAccount 'Unlock account' 8 $acY 150
$btnEnable  = New-TabButton $tabAccount 'Enable account' 170 $acY 150
$btnDisable = New-TabButton $tabAccount 'Disable account' 332 $acY 150
$acY += 34

# THE ONE ACTION A CLOUD TENANT ACTUALLY HAS ON THIS TAB.
#
# Unlock, Enable and Disable all write userAccountControl in Active Directory
# and have no Exchange Online equivalent -- the full Set-User parameter list
# has nothing for sign-in state or lockout, because those are Entra properties
# and need Graph. In a cloud-only tenant those three buttons could never work,
# so they are removed rather than greyed: a control that can never do anything
# is worse than no control, because it reads as broken rather than as absent.
#
# Set-User -ResetPasswordOnNextLogon IS documented as Applicable to Exchange
# Online, so the tab keeps one real action instead of three dead ones. It sits
# in the same place the removed buttons occupied.
# ITS OWN ROW, AND IT MOVES UP WHEN THE OTHER THREE GO.
#
# Declaring it on top of the three AD buttons would have been honest about the
# intent -- one slot, two mutually exclusive sets -- but Test-Toolkit's layout
# check reads static coordinates and cannot see Visible, so it would have
# reported a real overlap on every run from now on. A gate that cries wolf
# once is a gate somebody stops reading, so the button gets its own row and
# Update-AccountTab lifts it into the vacated one when the three are hidden.
$btnMustChange = New-TabButton $tabAccount 'Must change password at next sign-in' 8 $acY 260
$btnMustChange.Visible = $false
$acY += 34
$acNote = New-TabLabel $tabAccount '' 8 $acY 1040 -Muted
$acNote.Size = New-Object System.Drawing.Size(1040, 48)
$acNote.Text = ('Disabling here is ONE STEP of the leaver process. Term Requests does the whole job -- mailbox, licence, groups, forwarding -- in the right order; use it for a leaver.' + "`r`n" +
                'There is no password reset here on purpose. A credential change belongs in a tool with its own audit trail.')
$acY += 56

$btnMustChange.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $p = $script:Person
        $id = $(if ($p.Upn) { $p.Upn } elseif ($p.Smtp) { $p.Smtp } else { $p.Sam })
        $a = Show-ToolDialog(("Require $($p.Display) to change their password at the next sign-in?" + "`r`n`r`n" +
                              "Their current password keeps working until they do. Anything signed in with it -- phones, mapped drives -- carries on until the password actually changes."),
                             'Must change at next sign-in', 'OKCancel', 'Question')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-EXOCommand -Name 'Set-User' -Parameters @{ Identity = $id; ResetPasswordOnNextLogon = $true; ErrorAction = 'Stop' } | Out-Null
        Write-EditorAudit -Action 'PersonMustChangePassword' -Environment 'Cloud' -Detail 'ResetPasswordOnNextLogon set via Set-User.'
        Invoke-PersonReload
        Set-EditorStatus -Text 'They will be asked to change their password at the next sign-in.'
    } 'Must change at next sign-in'
})

function Update-AccountTab {
    $p = $script:Person; $u = $p.Ad
    $vals = @($acEnabled, $acLocked, $acPwdSet, $acPwdExp, $acLogon, $acExpires, $acCreated, $acOu, $acDesc)
    $cloudOnlyTenant = $false
    try { $cloudOnlyTenant = Test-TenantIsCloudOnly } catch { $cloudOnlyTenant = $false }

    if (-not $u) {
        foreach ($v in $vals) { $v.Text = '' }

        # THE THREE AD BUTTONS GO, AND ONE THAT WORKS TAKES THEIR PLACE.
        # Which way round depends on WHY there is no AD object: in a cloud
        # tenant there is no on-premises half to be missing, and Set-User can
        # still require a password change; in a hybrid tenant this person is
        # simply cloud-only and the same is true of them.
        foreach ($b in @($btnUnlock, $btnEnable, $btnDisable)) { $b.Visible = $false; $b.Enabled = $false }
        $btnMustChange.Location = New-Object System.Drawing.Point(8, $script:AcActionsY)
        $btnMustChange.Visible = $true
        $btnMustChange.Enabled = [bool]$p.Cloud

        $acEnabled.Text = $(if ($cloudOnlyTenant) { '(Entra -- needs Graph)' } else { '(no on-premises account)' })
        $acNote.Text = $(if ($cloudOnlyTenant) {
                            'Sign-in state, lockout, password age and account expiry are Entra properties. Exchange Online cannot read or write them, so they are blank here until Microsoft Graph is configured for this tenant.' + "`r`n" +
                            'Requiring a password change at next sign-in is the one thing Exchange Online can do, and it is the button above.'
                        } else {
                            'This person has no on-premises account, so the Active Directory fields above are blank and unlock, enable and disable do not apply.' + "`r`n" +
                            'Requiring a password change at next sign-in goes through Exchange Online and is the button above.'
                        })
        return
    }

    # An on-premises account: the AD buttons are the right ones, and the
    # Exchange-only action is not.
    foreach ($b in @($btnUnlock, $btnEnable, $btnDisable)) { $b.Visible = $true }
    $btnMustChange.Visible = $false
    $acNote.Text = ('Disabling here is ONE STEP of the leaver process. Term Requests does the whole job -- mailbox, licence, groups, forwarding -- in the right order; use it for a leaver.' + "`r`n" +
                    'There is no password reset here on purpose. A credential change belongs in a tool with its own audit trail -- Password Manager.')
    $acEnabled.Text = $(if ($u.Enabled) { 'Yes' } else { 'NO -- DISABLED' })
    $acLocked.Text  = $(if ($u.LockedOut) { 'YES -- LOCKED' } else { 'no' })
    $acPwdSet.Text  = $(if ($u.PasswordLastSet) { $u.PasswordLastSet.ToString('yyyy-MM-dd HH:mm') } else { '(never)' })
    # msDS-UserPasswordExpiryTimeComputed IS THE DOMAIN'S OWN ANSWER, fine-
    # grained policy included. Computing it from pwdLastSet plus the default
    # max age is wrong for anyone under a PSO.
    $exp = ''
    try {
        $raw = $u.'msDS-UserPasswordExpiryTimeComputed'
        if ($u.PasswordNeverExpires) { $exp = 'never (PasswordNeverExpires is set)' }
        elseif ($raw -and [int64]$raw -gt 0 -and [int64]$raw -lt [int64]::MaxValue) { $exp = [datetime]::FromFileTime([int64]$raw).ToString('yyyy-MM-dd HH:mm') }
        elseif ($u.PasswordExpired) { $exp = 'EXPIRED' }
        else { $exp = '(not reported)' }
    } catch { $exp = '(not reported)' }
    if ($u.PasswordExpired -and $exp -notmatch 'EXPIRED') { $exp = "EXPIRED ($exp)" }
    $acPwdExp.Text  = $exp
    $acLogon.Text   = $(if ($u.LastLogonDate) { $u.LastLogonDate.ToString('yyyy-MM-dd HH:mm') } else { '(never, or not yet replicated)' })
    $acExpires.Text = $(if ($u.AccountExpirationDate) { $u.AccountExpirationDate.ToString('yyyy-MM-dd') } else { 'never' })
    $acCreated.Text = $(if ($u.whenCreated) { ([datetime]$u.whenCreated).ToString('yyyy-MM-dd') } else { '' })
    $acOu.Text      = ("$($u.DistinguishedName)" -replace '^CN=[^,]+,', '')
    $acDesc.Text    = "$($u.Description)"
    $btnUnlock.Enabled  = [bool]$u.LockedOut
    $btnEnable.Enabled  = -not $u.Enabled
    $btnDisable.Enabled = [bool]$u.Enabled
}

$btnUnlock.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-PersonWritable -What 'Unlocking')) { return }
        $a = Show-ToolDialog("Unlock $($script:Person.Display)'s account?`r`n`r`nIf it locks again within minutes, something is still trying an old password -- a phone, a mapped drive, a scheduled task.",
                             'Unlock account', 'OKCancel', 'Question')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-PersonAd -Name 'Unlock-ADAccount' -Parameters @{ Identity = $script:Person.Ad.DistinguishedName } | Out-Null
        Write-EditorAudit -Action 'PersonUnlock' -Environment 'OnPrem' -Detail 'Account unlocked.'
        Invoke-PersonReload
        Set-EditorStatus -Text 'Account unlocked.'
    } 'Unlock account'
})

$btnEnable.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-PersonWritable -What 'Enabling')) { return }
        $a = Show-ToolDialog("Enable $($script:Person.Display)'s account?`r`n`r`nThey will be able to sign in again immediately on-premises, and in the cloud after the next directory sync.",
                             'Enable account', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-PersonAd -Name 'Enable-ADAccount' -Parameters @{ Identity = $script:Person.Ad.DistinguishedName } | Out-Null
        Write-EditorAudit -Action 'PersonEnable' -Environment 'OnPrem' -Detail 'Account enabled.'
        Invoke-PersonReload
        Set-EditorStatus -Text 'Account enabled.'
    } 'Enable account'
})

$btnDisable.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-PersonWritable -What 'Disabling')) { return }
        $a = Show-ToolDialog(("Disable $($script:Person.Display)'s account?" + "`r`n`r`n" +
                              "THIS IS ONE STEP OF A LEAVER PROCESS, NOT THE PROCESS. It stops sign-in. It does not convert the mailbox, remove the licence, clear group memberships or set forwarding -- Term Requests does all of that in order." + "`r`n`r`n" +
                              "Continue only if you mean to disable sign-in and nothing else."),
                             'Disable account', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-PersonAd -Name 'Disable-ADAccount' -Parameters @{ Identity = $script:Person.Ad.DistinguishedName } | Out-Null
        Write-EditorAudit -Action 'PersonDisable' -Environment 'OnPrem' -Detail 'Account disabled (sign-in only; not the leaver process).'
        Invoke-PersonReload
        Set-EditorStatus -Text 'Account disabled.'
    } 'Disable account'
})
#endregion

#region Groups tab --------------------------------------------------------------
<#
    THE REVERSE OF THE GROUP EDITOR'S QUESTION. That window asks 'who is in
    this group'; this tab asks 'which groups is this person in'. It reads
    memberOf from Active Directory -- the attribute is authoritative for a
    synced person -- and resolves each DN to a name, a category and a scope,
    so a security group and a distribution list read differently.

    THE PRIMARY GROUP IS NOT IN memberOf. Domain Users (or whatever the
    primary group is) is held on primaryGroupID instead, so it is added to
    the list by hand and marked, because 'this person is in no groups' is
    never true and would read as a bug.
#>
$gpY = 8
[void](New-TabLabel $tabGroups 'GROUP MEMBERSHIPS (ACTIVE DIRECTORY)' 8 $gpY -Bold)
$gpY += 22
$gpNote = New-TabLabel $tabGroups '' 8 $gpY 1050 -Muted
$gpY += 22
$gridGroups = New-TabGrid $tabGroups @(@('Name', 300), @('Type', 110), @('Scope', 100), @('Mail', 200), @('Domain', 180)) 8 $gpY 900 400
Add-GridPickColumn -Grid $gridGroups
$gpBtnY = $gpY
$btnGrpAdd    = New-TabButton $tabGroups 'Add to group...' 920 $gpBtnY 130
$gpBtnY += 34
$btnGrpRemove = New-TabButton $tabGroups 'Remove ticked' 920 $gpBtnY 130
$gpBtnY += 34
$btnGrpOpen   = New-TabButton $tabGroups 'Open in Group Editor' 920 $gpBtnY 130
$gpBtnY += 44
$btnGrpAll    = New-TabButton $tabGroups 'Tick all' 920 $gpBtnY 62
$btnGrpNone   = New-TabButton $tabGroups 'None' 988 $gpBtnY 62
$gpY += 412
$gpWarn = New-TabLabel $tabGroups '' 8 $gpY 1040 -Muted
$gpWarn.Size = New-Object System.Drawing.Size(1040, 34)

function Update-GroupsTab {
    $cloudOnlyTenant = $false
    try { $cloudOnlyTenant = Test-TenantIsCloudOnly } catch { $cloudOnlyTenant = $false }
    $gridGroups.Rows.Clear()
    $p = $script:Person; $u = $p.Ad
    if (-not $u) {
        # NOT 'READ FROM ACTIVE DIRECTORY' WHERE THERE IS NONE. In a cloud
        # tenant the memberships exist -- they live in Entra, and the
        # mail-enabled ones are reachable through Exchange Online -- this tab
        # just does not read them yet. Saying they come from a directory the
        # tenant has never had is the kind of wrong that sends somebody
        # looking for a sync problem.
        $gpNote.Text = $(if ($cloudOnlyTenant) {
                             'Not shown yet. This tab reads Active Directory, and this tenant has none. Entra group membership needs Microsoft Graph; the mail-enabled groups are visible in the Group Editor in the meantime.'
                         } else {
                             'Group membership is read from Active Directory, and this person has no on-premises account.'
                         })
        $btnGrpAdd.Enabled = $false; $btnGrpRemove.Enabled = $false; $btnGrpOpen.Enabled = $false
        return
    }
    $dns = @($u.memberOf | ForEach-Object { "$_" } | Where-Object { $_ })
    $rows = @()
    foreach ($dn in $dns) {
        $g = $null
        try { $g = Invoke-PersonAd -Name 'Get-ADGroup' -ForDn $dn -Parameters @{ Identity = $dn; Properties = @('mail', 'GroupCategory', 'GroupScope') } } catch { }
        $dom = ''
        if ($dn -match '(DC=.+)$') { $dom = ($Matches[1] -replace 'DC=', '' -replace ',', '.') }
        $rows += [PSCustomObject]@{
            Name = $(if ($g) { "$($g.Name)" } else { ($dn -replace '^CN=([^,]+),.*$', '$1') })
            Type = $(if ($g) { "$($g.GroupCategory)" } else { '(unreadable)' })
            Scope = $(if ($g) { "$($g.GroupScope)" } else { '' })
            Mail = $(if ($g -and $g.mail) { "$($g.mail)" } else { '' })
            Domain = $dom; Dn = $dn; Primary = $false
        }
    }
    # THE PRIMARY GROUP, ADDED BY HAND -- see the region note.
    try {
        $full = Invoke-PersonAd -Name 'Get-ADUser' -Parameters @{ Identity = $u.DistinguishedName; Properties = @('primaryGroupID') }
        $rid  = [int]$full.primaryGroupID
        # The primary group's SID is the domain SID plus the RID held on the
        # user, so the domain SID comes from the user's own SID.
        $domSid = ("$($u.SID)" -replace '-\d+$', '')
        $pg = Invoke-PersonAd -Name 'Get-ADGroup' -Parameters @{ Identity = "$domSid-$rid"; Properties = @('mail') }
        if ($pg) {
            $rows += [PSCustomObject]@{ Name = "$($pg.Name)"; Type = "$($pg.GroupCategory) (primary)"; Scope = "$($pg.GroupScope)"
                                        Mail = ''; Domain = ''; Dn = "$($pg.DistinguishedName)"; Primary = $true }
        }
    } catch { }
    foreach ($r in ($rows | Sort-Object Primary, Name)) {
        $i = $gridGroups.Rows.Add($false, $r.Name, $r.Type, $r.Scope, $r.Mail, $r.Domain)
        $gridGroups.Rows[$i].Tag = $r
        if ($r.Primary) { $gridGroups.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Muted) }
    }
    $gpNote.Text = "$($rows.Count) group(s). Mail-enabled groups show an address and can be opened in the Group Editor. The primary group cannot be removed from here."
    $btnGrpAdd.Enabled = $true; $btnGrpRemove.Enabled = $true; $btnGrpOpen.Enabled = $true
    $gpWarn.Text = "Changes are written with Add-ADGroupMember / Remove-ADGroupMember against $($p.Server). Group membership reaches Exchange Online after the next directory sync; a licence group change takes effect when the licence provisioning runs after that."
    $gpWarn.ForeColor = (Get-ToolStatusColour -Kind Muted)
}

$btnGrpAll.Add_Click({  Set-GridPicked -Grid $gridGroups -On $true })
$btnGrpNone.Add_Click({ Set-GridPicked -Grid $gridGroups -On $false })

$btnGrpAdd.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-PersonWritable -What 'Adding to a group')) { return }
        $typed = Show-InputDialog -Title 'Add to group' -Width 620 `
                    -Prompt "Group name(s) or address(es) to add $($script:Person.Display) to, separated by semicolons:"
        $names = @(ConvertFrom-DialogList -Text "$typed")
        if ($names.Count -eq 0) { return }
        # RESOLVED BEFORE ANYTHING IS WRITTEN, so a typo in the third name
        # does not leave the first two done and the operator unsure.
        # EVERY CONFIGURED DIRECTORY IS ASKED, the person's own first. A group
        # named by hand may live in any domain of the forest, and a name that
        # resolves in exactly one of them is the group meant.
        $servers = @($script:Person.Server) + @(Get-TenantDirectory | ForEach-Object { Get-TenantDirectoryServer -Id $_.Id } | Where-Object { $_ -and $_ -ne $script:Person.Server })
        $groups = @(); $bad = @()
        foreach ($n in $names) {
            $g = $null
            foreach ($srv in $servers) {
                try { $g = Get-ADGroup -Identity $n -Server $srv -ErrorAction Stop } catch { $g = $null }
                if (-not $g) {
                    try { $g = @(Get-ADGroup -Filter "mail -eq '$($n.Replace("'", "''"))'" -Server $srv -ErrorAction Stop)[0] } catch { $g = $null }
                }
                if ($g) { break }
            }
            if ($g) { $groups += $g } else { $bad += $n }
        }
        if ($bad.Count) { Set-EditorStatus -Text "Not found in any configured directory: $($bad -join ', ')" -Bad; return }
        $a = Show-ToolDialog(("Add $($script:Person.Display) to:`r`n`r`n  " + (@($groups | ForEach-Object { $_.Name }) -join "`r`n  ")),
                             'Add to group', 'OKCancel', 'Question')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $lines = @(); $ok = 0; $failed = 0
        foreach ($g in $groups) {
            try {
                Invoke-PersonAd -Name 'Add-ADGroupMember' -ForDn $g.DistinguishedName -Parameters @{ Identity = $g.DistinguishedName; Members = $script:Person.Ad.DistinguishedName } | Out-Null
                $lines += "OK      $($g.Name)"; $ok++
            } catch { $lines += "FAILED  $($g.Name): $($_.Exception.Message)"; $failed++ }
        }
        Write-EditorAudit -Action 'PersonGroupAdd' -Environment 'OnPrem' -Result $(if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }) `
            -Detail (@($groups | ForEach-Object { $_.Name }) -join '; ')
        if ($failed) { Show-ResultListDialog -Title "Add to group -- $ok done, $failed failed" -Lines $lines }
        Invoke-PersonReload
        Set-EditorStatus -Text "Added to $ok group(s), failed $failed." -Bad:($failed -gt 0)
    } 'Add to group'
})

$btnGrpRemove.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-PersonWritable -What 'Removing from a group')) { return }
        $picked = @(Get-GridPicked -Grid $gridGroups | ForEach-Object { $_.Tag })
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick the groups to remove.' -Bad; return }
        $prim = @($picked | Where-Object { $_.Primary })
        if ($prim.Count) { Set-EditorStatus -Text "The primary group ($($prim[0].Name)) cannot be removed from here." -Bad; return }
        $a = Show-ToolDialog(("Remove $($script:Person.Display) from:`r`n`r`n  " + (@($picked | ForEach-Object { $_.Name }) -join "`r`n  ") +
                              "`r`n`r`nIf one of these is a LICENCE group, the licence is removed at the next provisioning run and the mailbox enters its grace period."),
                             'Remove from group', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $lines = @(); $ok = 0; $failed = 0
        foreach ($g in $picked) {
            try {
                Invoke-PersonAd -Name 'Remove-ADGroupMember' -ForDn $g.Dn -Parameters @{ Identity = $g.Dn; Members = $script:Person.Ad.DistinguishedName; Confirm = $false } | Out-Null
                $lines += "OK      $($g.Name)"; $ok++
            } catch { $lines += "FAILED  $($g.Name): $($_.Exception.Message)"; $failed++ }
        }
        Write-EditorAudit -Action 'PersonGroupRemove' -Environment 'OnPrem' -Result $(if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }) `
            -Detail (@($picked | ForEach-Object { $_.Name }) -join '; ')
        if ($failed) { Show-ResultListDialog -Title "Remove from group -- $ok done, $failed failed" -Lines $lines }
        Invoke-PersonReload
        Set-EditorStatus -Text "Removed from $ok group(s), failed $failed." -Bad:($failed -gt 0)
    } 'Remove from group'
})

$btnGrpOpen.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    if ($gridGroups.SelectedRows.Count -eq 0) { Set-EditorStatus -Text 'Select the group to open.' -Bad; return }
    $g = $gridGroups.SelectedRows[0].Tag
    if (-not $g.Mail) { Set-EditorStatus -Text "'$($g.Name)' is not mail-enabled; the Group Editor edits mail-enabled groups." -Bad; return }
    $path = Join-Path $PSScriptRoot 'Group-Editor.ps1'
    if (-not (Test-Path -LiteralPath $path)) { Set-EditorStatus -Text 'Group-Editor.ps1 is not installed beside this tool.' -Bad; return }
    # THE GROUP EDITOR CARRIES TOOLKIT-REQUIRES: OwnProcess for a synced
    # group. A group found in Active Directory is by definition on-premises,
    # so it goes to its own process the way Unified Lookup sends it.
    Invoke-EditorAction {
        Start-ToolInOnPremProcess -ScriptPath $path -SmtpAddress $g.Mail -Extra @('-OnPrem')
        Set-EditorStatus -Text "Group Editor opening for $($g.Name) in its own window -- it connects on-premises first, so give it a moment."
    } 'Open the Group Editor'
})
#endregion

#region Licence tab -------------------------------------------------------------
<#
    READ-ONLY, BY DESIGN. This tenant licenses by GROUP MEMBERSHIP, so the
    question 'is this person licensed' is 'are they in a licence group', and
    Get-UserLicenseState already answers it the way every other tool does. The
    way to CHANGE it is the Groups tab -- one place to edit membership, not
    two that can disagree.
#>
$lcY = 8
[void](New-TabLabel $tabLicence 'LICENCE' 8 $lcY -Bold)
$lcY += 22
[void](New-TabLabel $tabLicence 'Licensed:' 8 $lcY)
$lcState = New-TabValue $tabLicence 170 $lcY 360
[void](New-TabLabel $tabLicence 'Matched via:' 560 $lcY 135)
$lcVia = New-TabValue $tabLicence 700 $lcY 320
$lcY += 30
[void](New-TabLabel $tabLicence 'LICENCE GROUPS THIS PERSON IS IN' 8 $lcY -Bold)
$lcY += 22
$gridLic = New-TabGrid $tabLicence @(@('Group', 420), @('What it grants', 400)) 8 $lcY 830 220
$lcY += 232
$lcNote = New-TabLabel $tabLicence '' 8 $lcY 1040 -Muted
$lcNote.Size = New-Object System.Drawing.Size(1040, 60)
$lcNote.Text = ('Licensing here is by group membership. To license or unlicense somebody, add them to or remove them from the licence group on the Groups tab; the change takes effect when licence provisioning next runs after directory sync.' + "`r`n" +
                'Which groups count as licence groups is configured in Tenant Manager.')

function Update-LicenceTab {
    $gridLic.Rows.Clear()
    $p = $script:Person
    $lcState.Text = '(reading...)'; $lcVia.Text = ''

    # LICENSING HERE MEANS LICENCE-GROUP MEMBERSHIP, and that is read out of a
    # directory. A cloud-only tenant has none, so the answer is not 'no
    # licence' -- it is 'this tool cannot see licences here', and the two must
    # not look the same. Saying 'NO' would have somebody assigning a licence
    # that is already there.
    try {
        if (Test-TenantIsCloudOnly) {
            $lcState.Text = 'not available'
            $lcVia.Text   = 'Licence assignment lives in Entra and needs Microsoft Graph, which is not configured for this tenant.'
            return
        }
    } catch { }

    try {
        $script:Licence = $(if ($p.Smtp) { Get-UserLicenseState -Smtp $p.Smtp } elseif ($p.Sam) { Get-UserLicenseState -SamAccountName $p.Sam } else { $null })
    } catch { $script:Licence = $null; $lcState.Text = "could not be read -- $($_.Exception.Message)"; return }
    $s = $script:Licence
    if (-not $s) { $lcState.Text = '(no address or account name to look up)'; return }
    if (-not $s.Known) { $lcState.Text = "unknown -- $($s.Reason)"; return }
    $lcState.Text = $(if ($s.Licensed) { 'Yes' } else { 'NO' })
    $lcVia.Text   = "$($s.MatchedVia)"
    foreach ($g in @($s.Groups)) {
        $name = $(if ($g.PSObject.Properties['Name']) { "$($g.Name)" } else { "$g" })
        $what = $(if ($g.PSObject.Properties['Skus']) { (@($g.Skus) -join ', ') } elseif ($s.Skus) { (@($s.Skus) -join ', ') } else { '' })
        [void]$gridLic.Rows.Add($name, $what)
    }
}
#endregion

#region Load and tab switching --------------------------------------------------
function Invoke-PersonReload {
    # Re-read from the SAME controller and repaint what is on screen.
    $t = $script:Person.Upn; if (-not $t) { $t = $script:Person.Smtp }; if (-not $t) { $t = $script:Person.Sam }
    $script:Person = Resolve-PersonTarget -Identity $t
    $script:Loaded = @{}
    Update-EditorIdentity
    $script:TabReadFailed = $false
    Update-OverviewTab
    if (-not $script:TabReadFailed) { $script:Loaded['Overview'] = $true }
    Invoke-EditorTabLoad -Key "$($tabs.SelectedTab.Name)"
}

function Invoke-PersonLoad {
    $t = "$($txtWho.Text)".Trim()
    if (-not $t) { Set-EditorStatus -Text 'Enter an address, UPN or account name.' -Bad; return }
    Set-EditorStatus -Text "Resolving $t..."
    $script:Person = $null; $script:Loaded = @{}; $script:Licence = $null
    $p = Resolve-PersonTarget -Identity $t
    if (-not $p.Ok) {
        Update-EditorIdentity
        Set-EditorStatus -Text $p.Reason -Bad
        return
    }
    $script:Person = $p
    # THE MODE FOLLOWS WHERE THE PERSON IS. Found in AD: On-premises, and
    # fixed. Not found in AD: Cloud, read-only, and fixed.
    $script:Mode = $(if ($p.Ad) { 'OnPrem' } else { 'Cloud' })
    Sync-EditorModeCombo -Combo $cmbMode -Mode $script:Mode
    $cmbMode.Enabled = $false
    Update-EditorIdentity
    $script:TabReadFailed = $false
    Update-OverviewTab
    if (-not $script:TabReadFailed) { $script:Loaded['Overview'] = $true }
    Invoke-EditorTabLoad -Key "$($tabs.SelectedTab.Name)"
    Set-EditorStatus -Text "Loaded $($p.Display)."
}

$btnLoad.Add_Click({ Invoke-EditorAction { Invoke-PersonLoad } 'Load person' })
$txtWho.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $btnLoad.PerformClick(); $_.SuppressKeyPress = $true } })

# TAB CHANGE AND WINDOW CLOSE ARE WIRED BY THE SHARED RUNTIME, guards included.
Register-EditorTabHandler
#endregion

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Person Editor' | Out-Null

$form.Add_Shown({
    # THE SWITCH IS FIXED, AND THE REASON IS ON THE BANNER. A person is written
    # in Active Directory or not at all, so the drop-down is a label with a
    # border here. It stays so the four editors look the same.
    if ($null -eq $script:AdReady) { $script:AdReady = Test-AdAvailable }
    $script:Mode = $(if ($script:AdReady.Available) { 'OnPrem' } else { 'Cloud' })
    Sync-EditorModeCombo -Combo $cmbMode -Mode $script:Mode
    $cmbMode.Enabled = $false
    # 'READ-ONLY' WAS THE OLD ANSWER AND IT IS NO LONGER TRUE. Without Active
    # Directory the job fields are still editable -- through Exchange Online --
    # so the opening line says what this window CAN do rather than warning
    # about a directory that is not coming.
    $cloudTenant = $false
    try { $cloudTenant = Test-TenantIsCloudOnly } catch { $cloudTenant = $false }
    $lblSync.Text = $(if ($script:AdReady.Available) { "A person is edited in Active Directory. $($script:AdReady.Detail)" }
                      elseif ($cloudTenant) { 'Cloud tenant. Job fields are written to Exchange Online; sign-in state, groups and licence need Microsoft Graph.' }
                      else { "No Active Directory here: $($script:AdReady.Detail) Job fields still write to Exchange Online." })
    $lblSync.ForeColor = (Get-ToolStatusColour -Kind $(if ($script:AdReady.Available -or $cloudTenant) { 'Muted' } else { 'Warn' }))
    if ("$($txtWho.Text)".Trim()) { $btnLoad.PerformClick() }
})

[void]$form.ShowDialog()
try { $form.Dispose() } catch { }
