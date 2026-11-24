#Requires -Version 5.1
<#
.SYNOPSIS
    Standalone All Attributes tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\AllAttributes.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the All Attributes window.

    Fully functional from AdminConsole.ps1 too:
        .\AllAttributes.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the All Attributes
    window for that mailbox.

    Every function below is MOVED, UNCHANGED, from the original monolithic
    Mailbox-Group-Inspector.ps1 (original lines 2714-3080, region "Actions -
    all attributes"), including every design/verification comment. The only
    edit anywhere in this file that touches original tool behavior is
    mechanical: Show-AllAttribsForm's closing call changed from
        [void]$allForm.ShowDialog($form)
    to
        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
        # Console, which opens top-left. Helper is in Common.ps1.
        Set-FormScreenCorner -Form $allForm -Corner 'BottomLeft'
        [void]$allForm.ShowDialog()
    because there is no longer a parent $form to own it -- this script IS
    the top-level window now. (Show-AllAttribsEditDialog already takes its
    owner as an explicit -Owner parameter -- $allForm -- so that call is
    unchanged.) Every other line, including all comments, is byte-for-byte
    the original.
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

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'All Attributes -- Enter Target'
if (-not $target) { return }   # operator cancelled the prompt

#region Actions - all attributes (moved unchanged from Mailbox-Group-Inspector.ps1 lines 2714-3080)
# "All Attribs" button (2026-07-10). Ports the operator's separate
# AD-UserAttributeEditor reference script (full Get-ADUser -Properties *
# viewer/editor with a curated editable subset) into this tool, following the
# same adaptation approach used for "Mail Attribs" just above:
#   - Target/AD-server come from this tool's own Mailbox lookup, matched by
#     "mail" (Get-ADUser -Filter "mail -eq '$Smtp'") -- same pattern as
#     Get-MailAttribsADUser / Get-MailboxDLMemberships -- instead of the
#     reference's separate UPN-entry/Parse/Environment UI.
#   - Interaction model initially matched the REFERENCE SCRIPT exactly, per
#     instruction ("same as last -- if they are editable here, they are
#     editable in the final version"): select a row, then click the "Edit
#     Attribute" button, enabled only when the selected row's hidden
#     _Editable flag is 'Y' (grdAll's SelectionChanged handler). 2026-07-10c
#     added double-click as a second, equivalent way to trigger the same
#     edit (Invoke-AllAttribsEditSelectedRow), for parity with Mail Attribs
#     -- both dialogs now support both triggers.
#   - The editable-attribute list is copied verbatim from the reference
#     (Title, Department, Description, Company, Division, Office,
#     OfficePhone, MobilePhone, Fax, HomePhone, StreetAddress, City, State,
#     PostalCode, POBox, EmployeeID, EmployeeNumber, Organization, Initials,
#     HomePage) -- each already documented there as a verified, dedicated
#     Set-ADUser string parameter (Microsoft Learn), never a generic -Replace
#     guess. Everything else Get-ADUser -Properties * returns (SamAccountName,
#     UserPrincipalName, DistinguishedName, ObjectGUID, ObjectSID,
#     UserAccountControl, Manager, GivenName, Surname, DisplayName, mail, the
#     whole proxyAddresses/Exchange-schema set already covered by Mail
#     Attribs, etc.) is view-only here, same as the reference -- not
#     re-derived independently per standards Section 1(a).
#   - Get-ADUser -Properties * returns both ADUser's friendly property names
#     (e.g. "Office") and, when populated, the matching raw LDAP attribute
#     name (e.g. "physicalDeliveryOfficeName") as SEPARATE rows with the same
#     value -- normal Get-ADUser behavior per the reference script's own
#     header, not a bug introduced here. Only the friendly-named row is ever
#     in $script:AllAttribsEditableAttributes, so only it is ever editable;
#     the note label below the grid says so.
#   - Function/variable names are AllAttribs-prefixed to avoid colliding with
#     Mail Attribs' identically-purposed but differently-scoped functions
#     (e.g. Format-AllAttribsValue here additionally handles byte[]/SID/
#     GUID/DateTime -- types the curated mail-attribute list never produces
#     but a full -Properties * dump routinely does -- so it is intentionally
#     NOT the same function as Mail Attribs' simpler Format-AttributeValue).
#   - Same $script:-scope reasoning as Mail Attribs for the current AD-user
#     object ($script:AllAttribsCurrentADUser), since it's read/reassigned
#     from three separate event handlers (Refresh, Edit, and the
#     SelectionChanged enable/disable check).
$script:AllAttribsEditableAttributes = @(
    'Title', 'Department', 'Description', 'Company', 'Division',
    'Office', 'OfficePhone', 'MobilePhone', 'Fax', 'HomePhone',
    'StreetAddress', 'City', 'State', 'PostalCode', 'POBox',
    'EmployeeID', 'EmployeeNumber', 'Organization', 'Initials', 'HomePage'
)
# Attribute-object bookkeeping properties added by the ADUser type itself --
# never real directory attributes, always excluded from the display grid.
$script:AllAttribsExcludedProperties = @(
    'PropertyNames', 'AddedProperties', 'RemovedProperties',
    'ModifiedProperties', 'PropertyCount', 'nTSecurityDescriptor'
)
# Matches the AD account to this mailbox by its "mail" attribute -- same
# pattern as Get-MailAttribsADUser -- but requests every populated attribute
# (-Properties '*') rather than a fixed list.
function Get-AllAttribsADUser {
    param(
        [Parameter(Mandatory)] [string]$Smtp,
        [string]$ADServer
    )
    $safeSmtp = $Smtp.Replace("'", "''")
    $params = @{
        Filter      = "mail -eq '$safeSmtp'"
        Properties  = '*'
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
# Renders a single attribute value for grid display. Handles arrays, binary,
# GUID/SID, and DateTime -- ported verbatim from the reference script, since
# a full -Properties * dump routinely includes all of these (objectGUID,
# objectSid, whenCreated, lastLogon, etc.), unlike Mail Attribs' curated
# always-string attribute list.
function Format-AllAttribsValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [byte[]]) { return '<binary data>' }
    if ($Value -is [System.Security.Principal.SecurityIdentifier]) { return $Value.Value }
    if ($Value -is [guid]) { return $Value.ToString() }
    if ($Value -is [datetime]) { return ('{0:yyyy-MM-dd HH:mm:ss}' -f $Value) }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = New-Object System.Collections.Generic.List[string]
        foreach ($item in $Value) {
            if ($item -is [byte]) { return '<binary data>' }
            [void]$items.Add("$item")
        }
        return ($items -join '; ')
    }
    return "$Value"
}
# Builds the Attribute/Value/_Editable DataTable bound to the grid -- every
# populated property on the ADUser object (minus the bookkeeping properties
# above), sorted by name. Ported from the reference script's
# Format-UserAttributeTable.
function Format-AllAttribsTable {
    param($ADUserObject)
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Attribute', [string])
    [void]$table.Columns.Add('Value', [string])
    [void]$table.Columns.Add('_Editable', [string])
    $props = $ADUserObject.PSObject.Properties |
        Where-Object { $script:AllAttribsExcludedProperties -notcontains $_.Name } |
        Sort-Object Name
    foreach ($p in $props) {
        $row = $table.NewRow()
        $row['Attribute'] = $p.Name
        $row['Value']     = Format-AllAttribsValue -Value $p.Value
        $row['_Editable'] = if ($script:AllAttribsEditableAttributes -contains $p.Name) { 'Y' } else { 'N' }
        [void]$table.Rows.Add($row)
    }
    # Comma operator required -- a bare "return $table" would unroll the
    # DataTable into its DataRow objects instead of returning the table itself.
    return ,$table
}
# Set-ADUser via the dedicated named parameter matching $AttributeName (must
# be a member of $script:AllAttribsEditableAttributes -- enforced by the
# caller). Blank new value clears the attribute instead of writing an empty
# string. Ported from the reference script's Set-OnPremADUserAttribute.
function Set-AllAttribsAttribute {
    param(
        [Parameter(Mandatory)] [string]$SamAccountName,
        [Parameter(Mandatory)] [string]$AttributeName,
        [string]$NewValue,
        [string]$ADServer
    )
    $params = @{
        Identity    = $SamAccountName
        ErrorAction = 'Stop'
    }
    if ($ADServer) { $params['Server'] = $ADServer }
    if ("$NewValue".Trim() -eq '') {
        $params['Clear'] = $AttributeName
    } else {
        $params[$AttributeName] = $NewValue
    }
    Set-ADUser @params
}
# Edit dialog for a single attribute -- ported near-verbatim from the
# reference script's Show-AttributeEditDialog (the simpler version, with no
# extra warning notes, since none of the curated attributes here have the
# special-case semantics PrimarySmtpAddress/name have in Mail Attribs).
function Show-AllAttribsEditDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$AttributeName,
        [string]$CurrentValue
    )
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit Attribute -- $AttributeName"
    $dlg.Size = New-Object System.Drawing.Size(480, 240)
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
    $lblCurrent = New-Object System.Windows.Forms.Label
    $lblCurrent.Text = "Current Value:"
    $lblCurrent.Location = New-Object System.Drawing.Point(12, 42)
    $lblCurrent.AutoSize = $true
    $dlg.Controls.Add($lblCurrent)
    $txtCurrent = New-Object System.Windows.Forms.TextBox
    $txtCurrent.Location = New-Object System.Drawing.Point(12, 60)
    $txtCurrent.Size = New-Object System.Drawing.Size(440, 22)
    $txtCurrent.ReadOnly = $true
    $txtCurrent.BackColor = [System.Drawing.Color]::WhiteSmoke
    $txtCurrent.Text = $CurrentValue
    $txtCurrent.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtCurrent)
    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "New Value (leave blank to clear the attribute):"
    $lblNew.Location = New-Object System.Drawing.Point(12, 92)
    $lblNew.AutoSize = $true
    $dlg.Controls.Add($lblNew)
    $txtNew = New-Object System.Windows.Forms.TextBox
    $txtNew.Location = New-Object System.Drawing.Point(12, 110)
    $txtNew.Size = New-Object System.Drawing.Size(440, 22)
    $txtNew.Text = $CurrentValue
    $txtNew.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtNew)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(280, 150)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(368, 150)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    $result = $dlg.ShowDialog($Owner)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $txtNew.Text
}
function Show-AllAttribsForm {
    param([string]$Smtp, [string]$ADServer)
    $allForm = New-Object System.Windows.Forms.Form
    $allForm.Text = "All Attributes: $Smtp"
    $allForm.Size = New-Object System.Drawing.Size(680, 750)
    $allForm.StartPosition = "CenterScreen"
    $allForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads every populated attribute (Get-ADUser -Properties *) on this mailbox's on-prem AD account (matched by mail attribute). Click a column header to sort. Double-click a row, or select it and click 'Edit Attribute' (enabled only for the curated editable set below), to edit it -- everything else is view-only."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(640, 40)
    $allForm.Controls.Add($lblNote)
    $grdAll = New-Object System.Windows.Forms.DataGridView
    $grdAll.Location = New-Object System.Drawing.Point(12, 58)
    $grdAll.Size = New-Object System.Drawing.Size(640, 440)
    $grdAll.ReadOnly = $true
    $grdAll.AllowUserToAddRows = $false
    $grdAll.AllowUserToDeleteRows = $false
    $grdAll.AllowUserToResizeRows = $false
    $grdAll.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdAll.MultiSelect = $false
    $grdAll.AutoGenerateColumns = $false
    $grdAll.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $grdAll.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $grdAll.RowHeadersVisible = $false
    $grdAll.BackgroundColor = [System.Drawing.Color]::White
    $grdAll.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $grdAll.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grdAll.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
    $colAllAttribute = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAllAttribute.Name = 'Attribute'
    $colAllAttribute.HeaderText = 'Attribute'
    $colAllAttribute.DataPropertyName = 'Attribute'
    $colAllAttribute.Width = 260
    $colAllAttribute.ReadOnly = $true
    [void]$grdAll.Columns.Add($colAllAttribute)
    $colAllValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAllValue.Name = 'Value'
    $colAllValue.HeaderText = 'Value'
    $colAllValue.DataPropertyName = 'Value'
    $colAllValue.ReadOnly = $true
    $colAllValue.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    [void]$grdAll.Columns.Add($colAllValue)
    $colAllEditableFlag = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAllEditableFlag.Name = '_Editable'
    $colAllEditableFlag.DataPropertyName = '_Editable'
    $colAllEditableFlag.Visible = $false
    [void]$grdAll.Columns.Add($colAllEditableFlag)
    $allForm.Controls.Add($grdAll)
    # BindingSource is the Microsoft-recommended intermediary for binding a
    # DataGridView to a DataTable -- same as Mail Attribs' grid. Column-header
    # click-to-sort works natively once bound through it (SortMode defaults to
    # Automatic), no custom sort handler needed -- matches the reference script.
    $grdAll.DataSource = New-Object System.Windows.Forms.BindingSource
    $allBindingSource = $grdAll.DataSource
    $lblDupeNote = New-Object System.Windows.Forms.Label
    $lblDupeNote.Text = "Note: some values appear twice -- once under a friendly name (e.g. Office) and once under the raw LDAP name (e.g. physicalDeliveryOfficeName). Only the friendly row is editable."
    $lblDupeNote.Location = New-Object System.Drawing.Point(12, 502)
    $lblDupeNote.Size = New-Object System.Drawing.Size(640, 16)
    $lblDupeNote.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 7.5, [System.Drawing.FontStyle]::Italic)
    $lblDupeNote.ForeColor = [System.Drawing.Color]::DimGray
    $allForm.Controls.Add($lblDupeNote)
    $btnAllRefresh = New-Object System.Windows.Forms.Button
    $btnAllRefresh.Text = "Refresh"
    $btnAllRefresh.Location = New-Object System.Drawing.Point(12, 522)
    $btnAllRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $allForm.Controls.Add($btnAllRefresh)
    $btnAllEdit = New-Object System.Windows.Forms.Button
    $btnAllEdit.Text = "Edit Attribute"
    $btnAllEdit.Location = New-Object System.Drawing.Point(122, 522)
    $btnAllEdit.Size = New-Object System.Drawing.Size(140, 28)
    $btnAllEdit.Enabled = $false
    $btnAllEdit.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $allForm.Controls.Add($btnAllEdit)
    $btnAllClose = New-Object System.Windows.Forms.Button
    $btnAllClose.Text = "Close"
    $btnAllClose.Location = New-Object System.Drawing.Point(552, 522)
    $btnAllClose.Size = New-Object System.Drawing.Size(100, 28)
    $allForm.Controls.Add($btnAllClose)
    $lblAllLog = New-Object System.Windows.Forms.Label
    $lblAllLog.Text = "Log:"
    $lblAllLog.Location = New-Object System.Drawing.Point(12, 558)
    $lblAllLog.AutoSize = $true
    $allForm.Controls.Add($lblAllLog)
    $txtAllLog = New-Object System.Windows.Forms.TextBox
    $txtAllLog.Multiline = $true
    $txtAllLog.ReadOnly = $true
    $txtAllLog.ScrollBars = "Vertical"
    $txtAllLog.Location = New-Object System.Drawing.Point(12, 578)
    $txtAllLog.Size = New-Object System.Drawing.Size(640, 110)
    $txtAllLog.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtAllLog.BackColor = [System.Drawing.Color]::WhiteSmoke
    $allForm.Controls.Add($txtAllLog)
    # Current AD-user object for this dialog instance -- script-scoped for the
    # same cross-handler reason as $script:MailAttribsCurrentADUser (read and
    # reassigned from Refresh, Edit Attribute, and the SelectionChanged
    # enable/disable check). Only one All Attribs dialog can be open at a
    # time (modal, ShowDialog($form)), so a single script-scope slot is safe.
    $script:AllAttribsCurrentADUser = $null
    function Update-AllAttribsGrid {
        param($ADUserObject)
        $table = Format-AllAttribsTable -ADUserObject $ADUserObject
        $allBindingSource.DataSource = $null
        $allBindingSource.DataSource = $table
        $grdAll.Refresh()
    }
    function Update-AllAttribsFromCurrent {
        $btnAllRefresh.Enabled = $false
        $btnAllEdit.Enabled = $false
        $txtAllLog.Text = "Loading all attributes..."
        $allForm.Refresh()
        try {
            $script:AllAttribsCurrentADUser = Get-AllAttribsADUser -Smtp $Smtp -ADServer $ADServer
            Update-AllAttribsGrid -ADUserObject $script:AllAttribsCurrentADUser
            $txtAllLog.Text = "Loaded attributes for SamAccountName '$($script:AllAttribsCurrentADUser.SamAccountName)'."
        } catch {
            $script:AllAttribsCurrentADUser = $null
            $allBindingSource.DataSource = $null
            $txtAllLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'All Attribs'
        } finally {
            $btnAllRefresh.Enabled = $true
        }
    }
    Update-AllAttribsFromCurrent
    $btnAllRefresh.Add_Click({ Update-AllAttribsFromCurrent })
    # Selection drives Edit Attribute's enabled state.
    $grdAll.Add_SelectionChanged({
        $hasRow = ($grdAll.SelectedRows.Count -gt 0 -and $null -ne $script:AllAttribsCurrentADUser)
        $isEditable = $false
        if ($hasRow) { $isEditable = ("$($grdAll.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
        $btnAllEdit.Enabled = ($hasRow -and $isEditable)
    })
    # Grid-row edit path -- reachable two ways as of 2026-07-10c: select a row
    # and click "Edit Attribute" (matches the reference script's own model),
    # or double-click the row directly (added for parity with Mail Attribs).
    # Both call this same function against whatever row is currently selected
    # in $grdAll, so the edit logic itself only exists once.
    function Invoke-AllAttribsEditSelectedRow {
        if ($grdAll.SelectedRows.Count -eq 0 -or $null -eq $script:AllAttribsCurrentADUser) { return }
        $row = $grdAll.SelectedRows[0]
        $attrName = "$($row.Cells['Attribute'].Value)"
        if ($script:AllAttribsEditableAttributes -notcontains $attrName) { return }
        $oldValue = "$($row.Cells['Value'].Value)"
        $btnAllEdit.Enabled = $false
        $btnAllRefresh.Enabled = $false
        try {
            $newValue = Show-AllAttribsEditDialog -Owner $allForm -AttributeName $attrName -CurrentValue $oldValue
            if ($null -eq $newValue) {
                $txtAllLog.Text = "Edit cancelled."
                return
            }
            $displayNew = if ($newValue.Trim() -eq '') { '(cleared)' } else { $newValue }
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Change '$attrName' for $Smtp ($($script:AllAttribsCurrentADUser.SamAccountName)):`n`nFrom: $oldValue`nTo:   $displayNew",
                'Confirm Attribute Edit',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtAllLog.Text = "Cancelled."
                return
            }
            Set-AllAttribsAttribute -SamAccountName $script:AllAttribsCurrentADUser.SamAccountName -AttributeName $attrName -NewValue $newValue -ADServer $ADServer
            $refreshed = Get-AllAttribsADUser -Smtp $Smtp -ADServer $ADServer
            $script:AllAttribsCurrentADUser = $refreshed
            $afterValue = Format-AllAttribsValue -Value ($refreshed.PSObject.Properties[$attrName].Value)
            Update-AllAttribsGrid -ADUserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE ---")
            [void]$log.AppendLine("$attrName = $oldValue")
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine("$attrName = $afterValue")
            $txtAllLog.Text = $log.ToString().TrimEnd()
            Write-CompletedLog -Task "Edited Attribute $attrName $($script:AllAttribsCurrentADUser.SamAccountName) (All Attribs; $Smtp)"
        } catch {
            $txtAllLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'All Attribs edit'
        } finally {
            $btnAllRefresh.Enabled = $true
            $hasRow = ($grdAll.SelectedRows.Count -gt 0 -and $null -ne $script:AllAttribsCurrentADUser)
            $isEditable = $false
            if ($hasRow) { $isEditable = ("$($grdAll.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
            $btnAllEdit.Enabled = ($hasRow -and $isEditable)
        }
    }
    $btnAllEdit.Add_Click({ Invoke-AllAttribsEditSelectedRow })
    # Double-click: select the row, then run the shared edit function. A
    # double-click on a non-editable row is a silent no-op (the function
    # itself checks $script:AllAttribsEditableAttributes).
    $grdAll.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $grdAll.ClearSelection()
        $grdAll.Rows[$e.RowIndex].Selected = $true
        Invoke-AllAttribsEditSelectedRow
    })
    $btnAllClose.Add_Click({ $allForm.Close() })
    # CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the
    # monolith's shared parent window. This script has no parent window, so
    # it shows itself as the top-level dialog instead. (Show-AllAttribsEditDialog
    # above already takes its owner as an explicit -Owner parameter --
    # $allForm -- so that call is unchanged.) Nothing else in this function
    # changed.
    [void]$allForm.ShowDialog()
}
#endregion

Show-AllAttribsForm -Smtp $target.Smtp -ADServer $target.ADServer
