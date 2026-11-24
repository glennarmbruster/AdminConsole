#Requires -Version 5.1
<#
    AllAttributes.ps1
    Standalone All Attributes tool -- Layer 2 tool script.

    Cloud-only build. Dumps everything Exchange Online exposes about the
    recipient -- Get-User for the person, Get-Mailbox for the mailbox -- and
    allows editing of the contact/organisation fields Set-User accepts.

    Author   : Glenn Armbruster
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# A wider view of the directory attributes on an account, beyond the mail ones.
#
# Use it when something is set that should not be, or missing when it should be
# there, and you need to see the raw value rather than what a UI reports.
#
# Editing here writes directly to the attribute with no validation beyond what
# the directory itself enforces. Read Mail Attributes first if the field is a
# mail one -- it understands addresses properly.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -ToolTitle 'All Attributes -- Enter Target'
if (-not $target) { return }

#region Actions - all attributes ----------------------------------------------
<#
    The editable set is exactly the Set-User parameters that describe a person
    -- contact details and organisation placement. Nothing structural is here
    on purpose: Name, UserPrincipalName, SamAccountName, DisplayName,
    FirstName, LastName and Manager all change how the object is identified or
    routed, and a grid of two hundred rows is the wrong place to do that by
    accident.

    Five attributes from the on-premises version have no Set-User equivalent
    and are therefore read-only here rather than silently missing:
    Description, Division, EmployeeID, EmployeeNumber and Organization. Four
    others simply changed name -- State became StateOrProvince, POBox became
    PostOfficeBox, OfficePhone became Phone, HomePage became WebPage.
#>
$script:AllAttribsEditableAttributes = @(
    'Title', 'Department', 'Company', 'Office',
    'Phone', 'MobilePhone', 'Fax', 'HomePhone', 'Pager',
    'StreetAddress', 'City', 'StateOrProvince', 'PostalCode',
    'PostOfficeBox', 'CountryOrRegion',
    'Initials', 'WebPage', 'Notes', 'SimpleDisplayName',
    'AssistantName', 'TelephoneAssistant'
)
$script:AllAttribsExcludedProperties = @(
    'PSComputerName', 'PSShowComputerName', 'RunspaceId',
    'PropertyNames', 'AddedProperties', 'RemovedProperties',
    'ModifiedProperties', 'PropertyCount', 'ObjectState',
    'IsValid', 'ExchangeVersion', 'OriginatingServer'
)
function Get-AllAttribsUser {
    <#
        Two cmdlets, one flat list.

        Get-User and Get-Mailbox overlap on a handful of names and disagree on
        none that matter, but a merged list with no provenance would leave the
        operator guessing which cmdlet to use to change something. Mailbox
        properties are therefore prefixed "Mailbox." and are always read-only:
        this tool edits the user object, and Set-Mailbox work belongs in
        MailAttributes.ps1 where the consequences are spelled out.
    #>
    param([Parameter(Mandatory)] [string]$Smtp)

    $u = Invoke-EXOCommand -Name 'Get-User' -Parameters @{
        Identity = $Smtp; ErrorAction = 'Stop' }

    $mb = $null
    try {
        $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop' }
    } catch {
        $mb = $null
    }

    return [PSCustomObject]@{
        User     = $u
        Mailbox  = $mb
        Identity = "$($u.Identity)"
        Upn      = "$($u.UserPrincipalName)"
    }
}
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
function Format-AllAttribsTable {
    param($UserObject)
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Attribute', [string])
    [void]$table.Columns.Add('Value', [string])
    [void]$table.Columns.Add('_Editable', [string])

    $userProps = $UserObject.User.PSObject.Properties |
        Where-Object { $script:AllAttribsExcludedProperties -notcontains $_.Name } |
        Sort-Object Name
    foreach ($p in $userProps) {
        $row = $table.NewRow()
        $row['Attribute'] = $p.Name
        $row['Value']     = Format-AllAttribsValue -Value $p.Value
        $row['_Editable'] = if ($script:AllAttribsEditableAttributes -contains $p.Name) { 'Y' } else { 'N' }
        [void]$table.Rows.Add($row)
    }

    if ($null -ne $UserObject.Mailbox) {
        $mbProps = $UserObject.Mailbox.PSObject.Properties |
            Where-Object { $script:AllAttribsExcludedProperties -notcontains $_.Name } |
            Sort-Object Name
        foreach ($p in $mbProps) {
            $row = $table.NewRow()
            $row['Attribute'] = "Mailbox.$($p.Name)"
            $row['Value']     = Format-AllAttribsValue -Value $p.Value
            $row['_Editable'] = 'N'
            [void]$table.Rows.Add($row)
        }
    }
    return ,$table
}
function Set-AllAttribsAttribute {
    <#
        Set-User has no -Clear. A blank is written as $null, which is how the
        cloud empties a single-valued property.
    #>
    param(
        [Parameter(Mandatory)] [string]$Identity,
        [Parameter(Mandatory)] [string]$AttributeName,
        [string]$NewValue
    )
    if ($script:AllAttribsEditableAttributes -notcontains $AttributeName) {
        throw "'$AttributeName' is not in the editable set for this tool."
    }
    $params = @{ Identity = $Identity; ErrorAction = 'Stop' }
    $params[$AttributeName] = if ("$NewValue".Trim() -eq '') { $null } else { $NewValue.Trim() }
    Invoke-EXOCommand -Name 'Set-User' -Parameters $params | Out-Null
}
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
    param([Parameter(Mandatory)][string]$Smtp)
    $allForm = New-Object System.Windows.Forms.Form
    $allForm.Text = "All Attributes: $Smtp"
    $allForm.Size = New-Object System.Drawing.Size(680, 750)
    $allForm.StartPosition = 'Manual'
    $allForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads every property Exchange Online exposes: Get-User first, then Get-Mailbox with each of its properties prefixed 'Mailbox.'. Click a column header to sort. Double-click a row, or select it and click 'Edit Attribute' (enabled only for the curated Set-User set), to edit it -- everything else, including every Mailbox. row, is view-only."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(640, 46)
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
    $script:AllAttribsCurrentUser = $null
    function Update-AllAttribsGrid {
        param($UserObject)
        $table = Format-AllAttribsTable -UserObject $UserObject
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
            $script:AllAttribsCurrentUser = Get-AllAttribsUser -Smtp $Smtp
            Update-AllAttribsGrid -UserObject $script:AllAttribsCurrentUser
            $txtAllLog.Text = "Loaded attributes for '$($script:AllAttribsCurrentUser.Upn)'."
        } catch {
            $script:AllAttribsCurrentUser = $null
            $allBindingSource.DataSource = $null
            $txtAllLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'All Attribs'
        } finally {
            $btnAllRefresh.Enabled = $true
        }
    }
    Update-AllAttribsFromCurrent
    $btnAllRefresh.Add_Click({ Update-AllAttribsFromCurrent })
    $grdAll.Add_SelectionChanged({
        $hasRow = ($grdAll.SelectedRows.Count -gt 0 -and $null -ne $script:AllAttribsCurrentUser)
        $isEditable = $false
        if ($hasRow) { $isEditable = ("$($grdAll.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
        $btnAllEdit.Enabled = ($hasRow -and $isEditable)
    })
    function Invoke-AllAttribsEditSelectedRow {
        if ($grdAll.SelectedRows.Count -eq 0 -or $null -eq $script:AllAttribsCurrentUser) { return }
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
                "Change '$attrName' for $Smtp ($($script:AllAttribsCurrentUser.Upn)):`n`nFrom: $oldValue`nTo:   $displayNew",
                'Confirm Attribute Edit',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtAllLog.Text = "Cancelled."
                return
            }
            Set-AllAttribsAttribute -Identity $script:AllAttribsCurrentUser.Identity -AttributeName $attrName -NewValue $newValue
            $refreshed = Get-AllAttribsUser -Smtp $Smtp
            $script:AllAttribsCurrentUser = $refreshed
            $afterValue = Format-AllAttribsValue -Value ($refreshed.User.PSObject.Properties[$attrName].Value)
            Update-AllAttribsGrid -UserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE ---")
            [void]$log.AppendLine("$attrName = $oldValue")
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine("$attrName = $afterValue")
            $txtAllLog.Text = $log.ToString().TrimEnd()
            Write-CompletedLog -Task "Edited Attribute $attrName $($script:AllAttribsCurrentUser.Upn) (All Attribs; $Smtp)"
        } catch {
            $txtAllLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'All Attribs edit'
        } finally {
            $btnAllRefresh.Enabled = $true
            $hasRow = ($grdAll.SelectedRows.Count -gt 0 -and $null -ne $script:AllAttribsCurrentUser)
            $isEditable = $false
            if ($hasRow) { $isEditable = ("$($grdAll.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
            $btnAllEdit.Enabled = ($hasRow -and $isEditable)
        }
    }
    $btnAllEdit.Add_Click({ Invoke-AllAttribsEditSelectedRow })
    $grdAll.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $grdAll.ClearSelection()
        $grdAll.Rows[$e.RowIndex].Selected = $true
        Invoke-AllAttribsEditSelectedRow
    })
    $btnAllClose.Add_Click({ $allForm.Close() })
    $allForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $allForm
    Add-ToolHelpButton -Form $allForm -ScriptPath $PSCommandPath -Title 'All Attributes' | Out-Null
    Set-FormScreenCorner -Form $allForm -Corner 'BottomLeft'
    [void]$allForm.ShowDialog()
}
#endregion

Show-AllAttribsForm -Smtp $target.Smtp
