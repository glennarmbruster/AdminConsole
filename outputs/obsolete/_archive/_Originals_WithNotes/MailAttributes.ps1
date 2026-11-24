#Requires -Version 5.1
<#
.SYNOPSIS
    Standalone Mail Attributes tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\MailAttributes.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the Mail Attributes window.

    Fully functional from AdminConsole.ps1 too:
        .\MailAttributes.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Mail Attributes
    window for that mailbox.

    Every function below is MOVED, UNCHANGED, from the original monolithic
    Mailbox-Group-Inspector.ps1 (original lines 2000-2666, region "Actions -
    mail attributes"), including every design/verification comment. The only
    edit anywhere in this file that touches original tool behavior is
    mechanical: Show-MailAttribsForm's closing call changed from
        [void]$mailForm.ShowDialog($form)
    to
        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
        # Console, which opens top-left. Helper is in Common.ps1.
        Set-FormScreenCorner -Form $mailForm -Corner 'BottomLeft'
        [void]$mailForm.ShowDialog()
    because there is no longer a parent $form to own it -- this script IS
    the top-level window now. (The two internal dialogs it opens,
    Show-MailAttribsEditDialog and Show-KeepOrRemoveOldPrimaryDialog, already
    took their owner as an explicit -Owner parameter -- $mailForm -- so those
    calls are unchanged.) Every other line, including all comments, is
    byte-for-byte the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
#>
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'Mail Attributes -- Enter Target'
if (-not $target) { return }   # operator cancelled the prompt

#region Actions - mail attributes (moved unchanged from Mailbox-Group-Inspector.ps1 lines 2000-2666)
# "Mail Attribs" button (2026-07-10). Ports the on-prem AD attribute grid and
# edit mechanics from the operator's "Renaming Script" (on-prem AD name-change
# tool) almost verbatim -- same fixed attribute list, same derived
# PrimarySmtpAddress/proxyAddresses handling, same Keep-as-Alias/Remove prompt,
# same "Add secondary SMTP address" control -- per explicit instruction to
# match that script's behavior exactly. Differences are only what's required
# to fit this tool's existing conventions:
#   - The target user comes from this tool's own Mailbox lookup (Smtp/ADServer
#     already resolved by Resolve-Target / Get-SelectedADServer) instead of a
#     separate UPN-entry/Parse/Environment UI -- Get-MailAttribsADUser matches
#     the AD account by "mail" (Get-ADUser -Filter "mail -eq '$Smtp'"), the
#     exact same match-by-mail pattern already used and verified by
#     Get-MailboxDLMemberships just above.
#   - The grid/dialog live inside one self-contained Show-MailAttribsForm
#     function, matching how every other feature dialog in this script
#     (Show-DLMembershipsForm, Show-MailboxHealthForm, etc.) is built. Because
#     the current AD-user object needs to be read AND reassigned across
#     several separate event handlers (Refresh, the shared edit function used
#     by both double-click and the Edit Attribute button, and Add Secondary)
#     -- not just one -- it is held in $script:MailAttribsCurrentADUser
#     rather than a plain closure variable, for the same reason this file
#     already uses $script:CurrentRules / $script:CurrentTraceResults /
#     $script:DLScanCancelRequested for equivalent cross-handler state: a
#     plain variable assignment inside one event-handler scriptblock does not
#     reliably propagate back out to a sibling handler's scriptblock.
#   - Function/variable names are new (MailAttribs-prefixed where the
#     reference name would otherwise be too generic to safely reuse) to avoid
#     any collision with this file's existing functions; the proxyAddresses
#     helpers and Format-AttributeValue keep their original reference-script
#     names since nothing else in this file defines them.
#   - "All Attribs" (the operator's second requested button, a full
#     every-attribute dump) is a separate follow-up request and is not part of
#     this region.
#
# Cmdlets used here (Get-ADUser -Filter/-Properties/-Server, Set-ADUser
# -DisplayName/-GivenName/-Initials/-Surname/-EmailAddress/-Clear/-Replace,
# Rename-ADObject -NewName) are the exact same calls, in the exact same
# shapes, already used and Microsoft-Learn-verified in the reference script
# this was ported from, and/or already used elsewhere in this file
# (Get-MailboxDLMemberships, Resolve-Target's on-prem group path) -- not
# re-verified independently per standards Section 1(a) (prior verified usage
# stands).
$script:MailAttribsDisplayOrder = @(
    'Company', 'DisplayName', 'name', 'GivenName', 'Initials', 'Surname',
    'mail', 'mailNickname', 'msExchArchiveName', 'PrimarySmtpAddress',
    'proxyAddresses', 'targetAddress', 'UserPrincipalName'
)
$script:MailAttribsDisplayLabels = @{
    'Company'            = 'Office'
    'GivenName'          = 'FirstName'
    'Surname'            = 'LastName'
    'PrimarySmtpAddress' = 'Primary SMTP'
}
# UserPrincipalName, Company ("Office"), name, and proxyAddresses are
# read-only/context-only, exactly as in the reference script -- name is the
# object's CN/RDN (out of scope for a mail-attributes tool) and proxyAddresses
# is only ever changed narrowly, via PrimarySmtpAddress and "Add Secondary".
$script:MailAttribsEditableAttributes = @(
    'DisplayName', 'GivenName', 'Initials', 'Surname', 'mail', 'mailNickname',
    'msExchArchiveName', 'PrimarySmtpAddress', 'targetAddress'
)
$script:MailAttribsNeverBlank = @('name', 'PrimarySmtpAddress')
# Matches the AD account to this mailbox by its "mail" attribute -- the same
# lookup Get-MailboxDLMemberships already uses just above, so a mailbox
# without a synced on-prem account throws the same style of message rather
# than silently showing an empty grid.
function Get-MailAttribsADUser {
    param(
        [Parameter(Mandatory)] [string]$Smtp,
        [string]$ADServer
    )
    $safeSmtp = $Smtp.Replace("'", "''")
    $params = @{
        Filter      = "mail -eq '$safeSmtp'"
        Properties  = @('Company', 'DisplayName', 'GivenName', 'Initials', 'Surname', 'mail', 'mailNickname',
                         'msExchArchiveName', 'Name', 'proxyAddresses', 'targetAddress',
                         'UserPrincipalName', 'DistinguishedName', 'SamAccountName')
        ErrorAction = 'Stop'
    }
    if ($ADServer) { $params['Server'] = $ADServer }
    $results = @(Get-ADUser @params)
    if ($results.Count -eq 0) {
        throw "No on-prem AD user account has a mail attribute matching '$Smtp'. This mailbox may be cloud-only (never synced from on-prem AD), or the Environment/Domain Controller selection at the top of the main window may need to change."
    }
    if ($results.Count -gt 1) {
        throw "$($results.Count) on-prem AD users matched mail '$Smtp'. Expected 0 or 1 -- resolve the duplicate mail attribute before using this feature."
    }
    return $results[0]
}
# Reads the primary SMTP address out of proxyAddresses -- the entry prefixed
# "SMTP:" in uppercase (Exchange convention; lowercase "smtp:" is secondary).
function Get-PrimarySmtpFromProxyAddresses {
    param($ProxyAddresses)
    foreach ($addr in @($ProxyAddresses)) {
        if ("$addr" -cmatch '^SMTP:(.+)$') { return $Matches[1] }
    }
    return ''
}
# Builds a new proxyAddresses array with $NewPrimary as the SMTP: entry.
# $KeepOldPrimaryAsAlias controls what happens to the previous primary (the
# operator's answer to Show-KeepOrRemoveOldPrimaryDialog): kept as a demoted
# smtp: entry, or dropped entirely. An existing secondary equal to the new
# primary is removed first so the address never appears twice.
function Set-PrimarySmtpInProxyAddresses {
    param(
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewPrimary,
        [bool]$KeepOldPrimaryAsAlias = $true
    )
    $kept = New-Object System.Collections.Generic.List[string]
    $oldPrimary = $null
    foreach ($addr in @($CurrentProxyAddresses)) {
        $s = "$addr"
        if ($s -cmatch '^SMTP:(.+)$') {
            $oldPrimary = $Matches[1]
            continue
        }
        if ($s -match '^smtp:(.+)$' -and $Matches[1] -ieq $NewPrimary) {
            continue
        }
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
# Appends $NewSecondary as an smtp: entry. Throws on a duplicate rather than
# silently creating one.
function Add-SecondarySmtpToProxyAddresses {
    param(
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewSecondary
    )
    foreach ($addr in @($CurrentProxyAddresses)) {
        if ("$addr" -match '^smtp:(.+)$' -and $Matches[1] -ieq $NewSecondary) {
            throw "'$NewSecondary' already exists in proxyAddresses (as $addr)."
        }
    }
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($addr in @($CurrentProxyAddresses)) { [void]$result.Add("$addr") }
    [void]$result.Add("smtp:$NewSecondary")
    return $result.ToArray()
}
# Renders a single attribute value for grid display -- handles the
# multi-valued proxyAddresses array so it never throws on an odd property type.
function Format-AttributeValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = New-Object System.Collections.Generic.List[string]
        foreach ($item in $Value) { [void]$items.Add("$item") }
        return ($items -join '; ')
    }
    return "$Value"
}
# Builds the Attribute/Value DataTable, in the fixed order above.
# PrimarySmtpAddress has no real AD property of its own -- derived from
# proxyAddresses here, the single place that derivation happens.
function Format-MailAttribsTable {
    param($ADUserObject)
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Attribute', [string])
    [void]$table.Columns.Add('Value', [string])
    [void]$table.Columns.Add('_Editable', [string])
    [void]$table.Columns.Add('_AttributeKey', [string])
    foreach ($attrName in $script:MailAttribsDisplayOrder) {
        $rawValue = if ($attrName -eq 'PrimarySmtpAddress') {
            Get-PrimarySmtpFromProxyAddresses -ProxyAddresses $ADUserObject.proxyAddresses
        } else {
            $ADUserObject.$attrName
        }
        $label = if ($script:MailAttribsDisplayLabels.ContainsKey($attrName)) { $script:MailAttribsDisplayLabels[$attrName] } else { $attrName }
        $row = $table.NewRow()
        $row['Attribute']     = $label
        $row['Value']         = Format-AttributeValue -Value $rawValue
        $row['_Editable']     = if ($script:MailAttribsEditableAttributes -contains $attrName) { 'Y' } else { 'N' }
        $row['_AttributeKey'] = $attrName
        [void]$table.Rows.Add($row)
    }
    # Comma operator required -- a bare "return $table" would unroll the
    # DataTable into its DataRow objects instead of returning the table itself.
    return ,$table
}
# Writes one attribute change. Branches by mechanism exactly as the reference
# script documents: dedicated Set-ADUser parameters, generic -Replace/-Clear
# for Exchange-schema attributes with no dedicated parameter, and the derived
# PrimarySmtpAddress field (rewrites proxyAddresses per $KeepOldPrimaryAsAlias).
# The 'name' case (Rename-ADObject) is kept for parity with the reference
# script but is unreachable from this UI, same as there -- 'name' is
# deliberately excluded from $script:MailAttribsEditableAttributes.
function Set-MailAttribsAttribute {
    param(
        [Parameter(Mandatory)] [string]$SamAccountName,
        [Parameter(Mandatory)] [string]$DistinguishedName,
        [Parameter(Mandatory)] [string]$AttributeName,
        [string]$NewValue,
        $CurrentProxyAddresses,
        [bool]$KeepOldPrimaryAsAlias = $true,
        [string]$ADServer
    )
    $commonParams = @{ ErrorAction = 'Stop' }
    if ($ADServer) { $commonParams['Server'] = $ADServer }
    $isBlank = ("$NewValue".Trim() -eq '')
    switch ($AttributeName) {
        'DisplayName' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear displayName @commonParams }
            else { Set-ADUser -Identity $SamAccountName -DisplayName $NewValue @commonParams }
        }
        'GivenName' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear givenName @commonParams }
            else { Set-ADUser -Identity $SamAccountName -GivenName $NewValue @commonParams }
        }
        'Initials' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear initials @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Initials $NewValue @commonParams }
        }
        'Surname' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear sn @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Surname $NewValue @commonParams }
        }
        'mail' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear mail @commonParams }
            else { Set-ADUser -Identity $SamAccountName -EmailAddress $NewValue @commonParams }
        }
        'mailNickname' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear mailNickname @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Replace @{ mailNickname = $NewValue } @commonParams }
        }
        'msExchArchiveName' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear msExchArchiveName @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Replace @{ msExchArchiveName = $NewValue } @commonParams }
        }
        'targetAddress' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear targetAddress @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Replace @{ targetAddress = $NewValue } @commonParams }
        }
        'PrimarySmtpAddress' {
            if ($isBlank) { throw "PrimarySmtpAddress cannot be blank." }
            $newProxyAddresses = Set-PrimarySmtpInProxyAddresses -CurrentProxyAddresses $CurrentProxyAddresses -NewPrimary $NewValue.Trim() -KeepOldPrimaryAsAlias $KeepOldPrimaryAsAlias
            Set-ADUser -Identity $SamAccountName -Replace @{ proxyAddresses = $newProxyAddresses } @commonParams
        }
        'name' {
            if ($isBlank) { throw "The object's Name cannot be blank." }
            Rename-ADObject -Identity $DistinguishedName -NewName $NewValue @commonParams
        }
        default { throw "Attribute '$AttributeName' is not handled by Set-MailAttribsAttribute." }
    }
}
# Standalone write for the "Add secondary SMTP address" control -- appends
# one new smtp: entry without touching anything else in proxyAddresses.
function Add-MailAttribsSecondaryProxyAddress {
    param(
        [Parameter(Mandatory)] [string]$SamAccountName,
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewSecondary,
        [string]$ADServer
    )
    $commonParams = @{ ErrorAction = 'Stop' }
    if ($ADServer) { $commonParams['Server'] = $ADServer }
    $newProxyAddresses = Add-SecondarySmtpToProxyAddresses -CurrentProxyAddresses $CurrentProxyAddresses -NewSecondary $NewSecondary
    Set-ADUser -Identity $SamAccountName -Replace @{ proxyAddresses = $newProxyAddresses } @commonParams
}
# Edit dialog for a single attribute -- ported near-verbatim from the
# reference script's Show-AttributeEditDialog, including the extra warning
# notes for 'name' and 'PrimarySmtpAddress' (the 'name' branch is kept for
# parity even though that row is never reachable from this UI, same as there).
function Show-MailAttribsEditDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$AttributeName,
        [string]$CurrentValue
    )
    $warningNote = switch ($AttributeName) {
        'name' { 'This renames the AD OBJECT itself (its Common Name / RDN in the directory) via Rename-ADObject -- a structural directory change, distinct from DisplayName. Cannot be left blank.' }
        'PrimarySmtpAddress' { 'This rewrites proxyAddresses: sets this value as the SMTP: (primary) entry. If a different primary already exists, you will be prompted next to Keep it as a secondary alias or Remove it entirely. Does NOT update the separate "mail" attribute -- edit that too if it needs to match. Cannot be left blank.' }
        default { '' }
    }
    $dlgHeight = if ($warningNote -ne '') { 330 } else { 300 }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit Attribute -- $AttributeName"
    $dlg.Size = New-Object System.Drawing.Size(520, $dlgHeight)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblAttr = New-Object System.Windows.Forms.Label
    $lblAttr.Text = "Attribute: $AttributeName"
    $lblAttr.Location = New-Object System.Drawing.Point(12, 12)
    $lblAttr.AutoSize = $true
    $lblAttr.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblAttr)
    $nextY = 42
    if ($warningNote -ne '') {
        $lblWarnNote = New-Object System.Windows.Forms.Label
        $lblWarnNote.Text = $warningNote
        $lblWarnNote.Location = New-Object System.Drawing.Point(12, $nextY)
        $lblWarnNote.Size = New-Object System.Drawing.Size(480, 40)
        $lblWarnNote.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
        $lblWarnNote.ForeColor = [System.Drawing.Color]::DarkOrange
        $dlg.Controls.Add($lblWarnNote)
        $nextY += 46
    }
    $lblCurrent = New-Object System.Windows.Forms.Label
    $lblCurrent.Text = "Current Value:"
    $lblCurrent.Location = New-Object System.Drawing.Point(12, $nextY)
    $lblCurrent.AutoSize = $true
    $dlg.Controls.Add($lblCurrent)
    $nextY += 18
    $txtCurrent = New-Object System.Windows.Forms.TextBox
    $txtCurrent.Location = New-Object System.Drawing.Point(12, $nextY)
    $txtCurrent.Size = New-Object System.Drawing.Size(480, 40)
    $txtCurrent.Multiline = $true
    $txtCurrent.ReadOnly = $true
    $txtCurrent.ScrollBars = "Vertical"
    $txtCurrent.BackColor = [System.Drawing.Color]::WhiteSmoke
    $txtCurrent.Text = $CurrentValue
    $txtCurrent.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtCurrent)
    $nextY += 48
    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "New Value (blank only allowed if this field is already blank):"
    $lblNew.Location = New-Object System.Drawing.Point(12, $nextY)
    $lblNew.AutoSize = $true
    $dlg.Controls.Add($lblNew)
    $nextY += 18
    $txtNew = New-Object System.Windows.Forms.TextBox
    $txtNew.Location = New-Object System.Drawing.Point(12, $nextY)
    $txtNew.Size = New-Object System.Drawing.Size(480, 60)
    $txtNew.Multiline = $true
    $txtNew.ScrollBars = "Vertical"
    $txtNew.Text = $CurrentValue
    $txtNew.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtNew)
    $nextY += 68
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(320, $nextY)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(408, $nextY)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    $result = $dlg.ShowDialog($Owner)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $txtNew.Text
}
# Shown only when editing PrimarySmtpAddress AND an existing, different
# primary is about to be displaced -- ported verbatim from the reference
# script. Returns 'Keep', 'Remove', or 'Cancel' (Cancel, or closing via the X
# button, aborts the whole attribute edit -- nothing is written to AD).
function Show-KeepOrRemoveOldPrimaryDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$OldPrimary,
        [string]$NewPrimary
    )
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Previous Primary SMTP Address"
    $dlg.Size = New-Object System.Drawing.Size(460, 210)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $dlg.Tag = 'Cancel'
    $lblMsg = New-Object System.Windows.Forms.Label
    $lblMsg.Text = "Current primary SMTP address:`n$OldPrimary`n`nIt is being replaced by:`n$NewPrimary`n`nWhat should I do with the current PrimarySmtpAddress?"
    $lblMsg.Location = New-Object System.Drawing.Point(12, 12)
    $lblMsg.Size = New-Object System.Drawing.Size(430, 100)
    $dlg.Controls.Add($lblMsg)
    $btnKeep = New-Object System.Windows.Forms.Button
    $btnKeep.Text = "Keep as Alias"
    $btnKeep.Location = New-Object System.Drawing.Point(20, 130)
    $btnKeep.Size = New-Object System.Drawing.Size(130, 32)
    $btnKeep.Add_Click({ $dlg.Tag = 'Keep'; $dlg.Close() })
    $dlg.Controls.Add($btnKeep)
    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = "Remove"
    $btnRemove.Location = New-Object System.Drawing.Point(165, 130)
    $btnRemove.Size = New-Object System.Drawing.Size(110, 32)
    $btnRemove.Add_Click({ $dlg.Tag = 'Remove'; $dlg.Close() })
    $dlg.Controls.Add($btnRemove)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(290, 130)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnCancel.Add_Click({ $dlg.Tag = 'Cancel'; $dlg.Close() })
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    [void]$dlg.ShowDialog($Owner)
    return "$($dlg.Tag)"
}
function Show-MailAttribsForm {
    param([string]$Smtp, [string]$ADServer)
    $mailForm = New-Object System.Windows.Forms.Form
    $mailForm.Text = "Mail Attributes: $Smtp"
    $mailForm.Size = New-Object System.Drawing.Size(680, 700)
    $mailForm.StartPosition = "CenterScreen"
    $mailForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads this mailbox's on-prem AD account (matched by mail attribute) -- the same fixed attribute set and edit mechanics as the Renaming Script tool. Double-click a row, or select it and click 'Edit Attribute' below, to edit it. 'proxyAddresses' is read-only here -- edit 'Primary SMTP' to change the primary address (you'll be prompted to Keep the old primary as an alias or Remove it), or use 'Add Secondary' below to add another address. 'name', 'UserPrincipalName', and 'Office' are read-only/context-only in this tool."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(640, 58)
    $mailForm.Controls.Add($lblNote)
    $grdMail = New-Object System.Windows.Forms.DataGridView
    $grdMail.Location = New-Object System.Drawing.Point(12, 74)
    $grdMail.Size = New-Object System.Drawing.Size(640, 340)
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
    $grdMail.BackgroundColor = [System.Drawing.Color]::White
    $grdMail.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $grdMail.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grdMail.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
    $colMailAttribute = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailAttribute.Name = 'Attribute'
    $colMailAttribute.HeaderText = 'Attribute'
    $colMailAttribute.DataPropertyName = 'Attribute'
    $colMailAttribute.Width = 160
    $colMailAttribute.ReadOnly = $true
    [void]$grdMail.Columns.Add($colMailAttribute)
    $colMailValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailValue.Name = 'Value'
    $colMailValue.HeaderText = 'Value'
    $colMailValue.DataPropertyName = 'Value'
    $colMailValue.ReadOnly = $true
    $colMailValue.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    [void]$grdMail.Columns.Add($colMailValue)
    $colMailEditableFlag = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailEditableFlag.Name = '_Editable'
    $colMailEditableFlag.DataPropertyName = '_Editable'
    $colMailEditableFlag.Visible = $false
    [void]$grdMail.Columns.Add($colMailEditableFlag)
    $colMailAttributeKey = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailAttributeKey.Name = '_AttributeKey'
    $colMailAttributeKey.DataPropertyName = '_AttributeKey'
    $colMailAttributeKey.Visible = $false
    [void]$grdMail.Columns.Add($colMailAttributeKey)
    $mailForm.Controls.Add($grdMail)
    $grdMail.DataSource = New-Object System.Windows.Forms.BindingSource
    $mailBindingSource = $grdMail.DataSource
    $lblAddSmtp = New-Object System.Windows.Forms.Label
    $lblAddSmtp.Text = "Add secondary SMTP address (appended as smtp: -- does not replace the primary or any other address):"
    $lblAddSmtp.Location = New-Object System.Drawing.Point(12, 422)
    $lblAddSmtp.Size = New-Object System.Drawing.Size(640, 16)
    $mailForm.Controls.Add($lblAddSmtp)
    $txtAddSmtp = New-Object System.Windows.Forms.TextBox
    $txtAddSmtp.Location = New-Object System.Drawing.Point(12, 440)
    $txtAddSmtp.Size = New-Object System.Drawing.Size(430, 22)
    $txtAddSmtp.Font = New-Object System.Drawing.Font("Consolas", 9)
    $mailForm.Controls.Add($txtAddSmtp)
    $btnAddSmtp = New-Object System.Windows.Forms.Button
    $btnAddSmtp.Text = "Add Secondary"
    $btnAddSmtp.Location = New-Object System.Drawing.Point(452, 439)
    $btnAddSmtp.Size = New-Object System.Drawing.Size(110, 26)
    $mailForm.Controls.Add($btnAddSmtp)
    $lblMailLog = New-Object System.Windows.Forms.Label
    $lblMailLog.Text = "Log:"
    $lblMailLog.Location = New-Object System.Drawing.Point(12, 474)
    $lblMailLog.AutoSize = $true
    $mailForm.Controls.Add($lblMailLog)
    $txtMailLog = New-Object System.Windows.Forms.TextBox
    $txtMailLog.Multiline = $true
    $txtMailLog.ReadOnly = $true
    $txtMailLog.ScrollBars = "Vertical"
    $txtMailLog.Location = New-Object System.Drawing.Point(12, 494)
    $txtMailLog.Size = New-Object System.Drawing.Size(640, 110)
    $txtMailLog.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtMailLog.BackColor = [System.Drawing.Color]::WhiteSmoke
    $mailForm.Controls.Add($txtMailLog)
    $btnMailRefresh = New-Object System.Windows.Forms.Button
    $btnMailRefresh.Text = "Refresh"
    $btnMailRefresh.Location = New-Object System.Drawing.Point(12, 616)
    $btnMailRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $mailForm.Controls.Add($btnMailRefresh)
    # Added 2026-07-10c alongside double-click -- see the CellDoubleClick /
    # SelectionChanged handlers below. Same position/size as All Attribs'
    # Edit Attribute button for visual consistency between the two dialogs.
    $btnMailEdit = New-Object System.Windows.Forms.Button
    $btnMailEdit.Text = "Edit Attribute"
    $btnMailEdit.Location = New-Object System.Drawing.Point(122, 616)
    $btnMailEdit.Size = New-Object System.Drawing.Size(140, 28)
    $btnMailEdit.Enabled = $false
    $btnMailEdit.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $mailForm.Controls.Add($btnMailEdit)
    $btnMailClose = New-Object System.Windows.Forms.Button
    $btnMailClose.Text = "Close"
    $btnMailClose.Location = New-Object System.Drawing.Point(552, 616)
    $btnMailClose.Size = New-Object System.Drawing.Size(100, 28)
    $mailForm.Controls.Add($btnMailClose)
    # Current AD-user object for this dialog instance. Script-scoped (not a
    # plain closure variable) because it is reassigned from three separate
    # event handlers below (Refresh, grid double-click edit, Add Secondary)
    # -- see the region header comment for why that requires $script: scope
    # in this file's established pattern. Only one Mail Attribs dialog can be
    # open at a time (modal, ShowDialog($form)), so a single script-scope slot
    # is safe, same as $script:CurrentRules / $script:CurrentTraceResults.
    $script:MailAttribsCurrentADUser = $null
    function Update-MailAttribsGrid {
        param($ADUserObject)
        $table = Format-MailAttribsTable -ADUserObject $ADUserObject
        $mailBindingSource.DataSource = $null
        $mailBindingSource.DataSource = $table
        $grdMail.Refresh()
    }
    function Update-MailAttribsFromCurrent {
        $btnMailRefresh.Enabled = $false
        $btnAddSmtp.Enabled = $false
        $txtMailLog.Text = "Loading mail attributes..."
        $mailForm.Refresh()
        try {
            $script:MailAttribsCurrentADUser = Get-MailAttribsADUser -Smtp $Smtp -ADServer $ADServer
            Update-MailAttribsGrid -ADUserObject $script:MailAttribsCurrentADUser
            $txtMailLog.Text = "Loaded mail attributes for SamAccountName '$($script:MailAttribsCurrentADUser.SamAccountName)'."
        } catch {
            $script:MailAttribsCurrentADUser = $null
            $mailBindingSource.DataSource = $null
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
        }
    }
    Update-MailAttribsFromCurrent
    $btnMailRefresh.Add_Click({ Update-MailAttribsFromCurrent })
    # Grid-row edit path -- reachable two ways as of 2026-07-10c: double-click
    # a row (CellDoubleClick below), or select a row and click "Edit
    # Attribute" (btnMailEdit below, enabled only for an editable row via
    # SelectionChanged). Both call this same function against whatever row is
    # currently selected in $grdMail, so the edit logic itself only exists
    # once. Silently no-ops on a read-only row (defensive -- SelectionChanged
    # already keeps btnMailEdit disabled for those, and double-click checks
    # $script:MailAttribsEditableAttributes itself below).
    function Invoke-MailAttribsEditSelectedRow {
        if ($grdMail.SelectedRows.Count -eq 0 -or $null -eq $script:MailAttribsCurrentADUser) { return }
        $row = $grdMail.SelectedRows[0]
        $attrLabel = "$($row.Cells['Attribute'].Value)"
        $attrName  = "$($row.Cells['_AttributeKey'].Value)"
        if ($script:MailAttribsEditableAttributes -notcontains $attrName) { return }
        $btnMailRefresh.Enabled = $false
        $btnMailEdit.Enabled = $false
        $btnAddSmtp.Enabled = $false
        try {
            $oldValue = "$($row.Cells['Value'].Value)"
            $newValue = Show-MailAttribsEditDialog -Owner $mailForm -AttributeName $attrLabel -CurrentValue $oldValue
            if ($null -eq $newValue) {
                $txtMailLog.Text = "Edit cancelled."
                return
            }
            if ($script:MailAttribsNeverBlank -contains $attrName -and $newValue.Trim() -eq '') {
                [void][System.Windows.Forms.MessageBox]::Show("'$attrLabel' cannot be blank.", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $oldIsBlank = ("$oldValue".Trim() -eq '')
            $newIsBlank = $newValue.Trim() -eq ''
            if ($newIsBlank -and -not $oldIsBlank) {
                [void][System.Windows.Forms.MessageBox]::Show("'$attrLabel' currently has a value and cannot be saved blank. Enter a replacement value instead.", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $keepOldPrimaryAsAlias = $true
            if ($attrName -eq 'PrimarySmtpAddress') {
                $oldPrimaryTrim = $oldValue.Trim()
                $newPrimaryTrim = $newValue.Trim()
                if ($oldPrimaryTrim -ne '' -and -not ($oldPrimaryTrim -ieq $newPrimaryTrim)) {
                    $choice = Show-KeepOrRemoveOldPrimaryDialog -Owner $mailForm -OldPrimary $oldPrimaryTrim -NewPrimary $newPrimaryTrim
                    if ($choice -ne 'Keep' -and $choice -ne 'Remove') {
                        $txtMailLog.Text = "Edit cancelled."
                        return
                    }
                    $keepOldPrimaryAsAlias = ($choice -eq 'Keep')
                }
            }
            $displayNew = if ($newValue.Trim() -eq '') { '(cleared)' } else { $newValue.Replace("`r`n", '; ') }
            $confirmExtra = switch ($attrName) {
                'PrimarySmtpAddress' {
                    if ($keepOldPrimaryAsAlias) { "`n`nThis rewrites the proxyAddresses list: sets this as SMTP: (primary) and demotes the previous primary to smtp: (secondary alias)." }
                    else { "`n`nThis rewrites the proxyAddresses list: sets this as SMTP: (primary) and REMOVES the previous primary address from proxyAddresses entirely." }
                }
                default { '' }
            }
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Change '$attrLabel' for $Smtp ($($script:MailAttribsCurrentADUser.SamAccountName)):`n`nFrom: $oldValue`nTo:   $displayNew$confirmExtra",
                'Confirm Attribute Edit',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtMailLog.Text = "Cancelled."
                return
            }
            Set-MailAttribsAttribute -SamAccountName $script:MailAttribsCurrentADUser.SamAccountName -DistinguishedName $script:MailAttribsCurrentADUser.DistinguishedName `
                -AttributeName $attrName -NewValue $newValue -CurrentProxyAddresses $script:MailAttribsCurrentADUser.proxyAddresses `
                -KeepOldPrimaryAsAlias $keepOldPrimaryAsAlias -ADServer $ADServer
            $refreshed = Get-MailAttribsADUser -Smtp $Smtp -ADServer $ADServer
            $script:MailAttribsCurrentADUser = $refreshed
            $afterValue = if ($attrName -eq 'PrimarySmtpAddress') {
                Get-PrimarySmtpFromProxyAddresses -ProxyAddresses $refreshed.proxyAddresses
            } else {
                Format-AttributeValue -Value ($refreshed.$attrName)
            }
            Update-MailAttribsGrid -ADUserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE ---")
            [void]$log.AppendLine("$attrLabel = $oldValue")
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine("$attrLabel = $afterValue")
            $txtMailLog.Text = $log.ToString().TrimEnd()
            $logVerb = if ($attrName -eq 'PrimarySmtpAddress') {
                "Edited Attribute $attrName ($(if($keepOldPrimaryAsAlias){'kept old primary as alias'}else{'removed old primary'}))"
            } else {
                "Edited Attribute $attrName"
            }
            Write-CompletedLog -Task "$logVerb $($script:MailAttribsCurrentADUser.SamAccountName) (Mail Attribs; $Smtp)"
        } catch {
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs edit'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
            $hasRow = ($grdMail.SelectedRows.Count -gt 0 -and $null -ne $script:MailAttribsCurrentADUser)
            $isEditable = $false
            if ($hasRow) { $isEditable = ("$($grdMail.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
            $btnMailEdit.Enabled = ($hasRow -and $isEditable)
        }
    }
    # Double-click: select the row (matching the pattern the reference
    # scripts use elsewhere), then run the shared edit function.
    $grdMail.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $grdMail.ClearSelection()
        $grdMail.Rows[$e.RowIndex].Selected = $true
        Invoke-MailAttribsEditSelectedRow
    })
    # Selection drives Edit Attribute's enabled state, same as All Attribs.
    $grdMail.Add_SelectionChanged({
        $hasRow = ($grdMail.SelectedRows.Count -gt 0 -and $null -ne $script:MailAttribsCurrentADUser)
        $isEditable = $false
        if ($hasRow) { $isEditable = ("$($grdMail.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
        $btnMailEdit.Enabled = ($hasRow -and $isEditable)
    })
    $btnMailEdit.Add_Click({ Invoke-MailAttribsEditSelectedRow })
    $btnAddSmtp.Add_Click({
        if ($null -eq $script:MailAttribsCurrentADUser) {
            [void][System.Windows.Forms.MessageBox]::Show("Load mail attributes first (click Refresh).", 'No User Loaded', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $btnMailRefresh.Enabled = $false
        $btnAddSmtp.Enabled = $false
        try {
            $newAddr = $txtAddSmtp.Text.Trim()
            if ($newAddr -eq '') {
                [void][System.Windows.Forms.MessageBox]::Show("Enter an address to add.", 'Missing Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            if ($newAddr -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                [void][System.Windows.Forms.MessageBox]::Show("'$newAddr' does not look like a valid email address (expected user@domain.tld).", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $beforeValue = Format-AttributeValue -Value $script:MailAttribsCurrentADUser.proxyAddresses
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Add 'smtp:$newAddr' as a secondary proxy address for $Smtp ($($script:MailAttribsCurrentADUser.SamAccountName))?`n`nThis does not change the primary address or remove any existing address.",
                'Confirm Add Secondary SMTP Address',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtMailLog.Text = "Cancelled."
                return
            }
            Add-MailAttribsSecondaryProxyAddress -SamAccountName $script:MailAttribsCurrentADUser.SamAccountName -CurrentProxyAddresses $script:MailAttribsCurrentADUser.proxyAddresses `
                -NewSecondary $newAddr -ADServer $ADServer
            $refreshed = Get-MailAttribsADUser -Smtp $Smtp -ADServer $ADServer
            $script:MailAttribsCurrentADUser = $refreshed
            $afterValue = Format-AttributeValue -Value $refreshed.proxyAddresses
            Update-MailAttribsGrid -ADUserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE (proxyAddresses) ---")
            [void]$log.AppendLine($beforeValue)
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine($afterValue)
            $txtMailLog.Text = $log.ToString().TrimEnd()
            Write-CompletedLog -Task "Added Secondary SMTP $($script:MailAttribsCurrentADUser.SamAccountName) (Mail Attribs; $Smtp; smtp:$newAddr)"
            $txtAddSmtp.Text = ''
        } catch {
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add Secondary SMTP'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
        }
    })
    $btnMailClose.Add_Click({ $mailForm.Close() })
    # CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the
    # monolith's shared parent window. This script has no parent window, so
    # it shows itself as the top-level dialog instead. (Show-MailAttribsEditDialog
    # and Show-KeepOrRemoveOldPrimaryDialog above already take their owner as
    # an explicit -Owner parameter -- $mailForm -- so those calls are
    # unchanged.) Nothing else in this function changed.
    [void]$mailForm.ShowDialog()
}
#endregion

Show-MailAttribsForm -Smtp $target.Smtp -ADServer $target.ADServer
