#Requires -Version 5.1
# TOOLKIT-VARIANT: OnPremAD
# ---------------------------------------------------------------------------
# AD- PREFIX: this is the ACTIVE DIRECTORY version of AllAttributes.
#
# It reads and writes raw AD attributes -- mailNickname, proxyAddresses,
# targetAddress -- through Get-ADUser/Get-ADGroup, and needs a reachable domain
# controller. AllAttributes.ps1 beside it does the same job against Exchange Online
# properties (Alias, EmailAddresses) and needs no directory at all.
#
# TWO FILES RATHER THAN TWO CODE PATHS. The alternative was one tool branching
# on Test-TenantCapability, which would have given the six most attribute-heavy
# tools in the toolkit a second path each, both needing testing. Prefixing the
# names keeps each file doing one thing, and Publish-Toolkit lists which set to
# strip for a given tenant.
#
# Ship AD-* where Tenant.json has OnPremAD true. Ship the unprefixed ones
# everywhere. Shipping both is not harmful -- the wrong one greys itself -- but
# it is two buttons doing one job.
# ---------------------------------------------------------------------------
<#
    AllAttributes.ps1
    Standalone All Attributes tool -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
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

# THE ADDRESS BOX IS ON THE FORM -- see Add-ToolTargetBar in Common.ps1.
#
# Resolve-ToolTarget is KEPT here, unlike in the cloud tools. This one reads and
# writes Active Directory, so it needs to know which directory holds the account
# and which controller to use, and an address it cannot place is an address this
# tool cannot serve. It runs per lookup rather than once at startup -- see
# Resolve-ToolTargetOnDemand in Common.ps1.
$script:Target   = ''
$script:ADServer = $null

function Set-AdTarget {
    # Both values are cleared FIRST. A failed lookup must not leave the previous
    # account's controller behind for the next write to land on.
    param([string]$Text)
    $script:Target   = ''
    $script:ADServer = $null
    $t = Resolve-ToolTargetOnDemand -Text $Text -DomainController $DomainController -ToolTitle 'All Attributes -- Enter Target'
    if (-not $t) { return }
    $script:Target   = $t.Smtp
    $script:ADServer = $t.ADServer
}

Set-AdTarget -Text $SmtpAddress

#region Actions - all attributes (moved unchanged from Mailbox-Group-Inspector.ps1 lines 2714-3080)
$script:AllAttribsEditableAttributes = @(
    'Title', 'Department', 'Description', 'Company', 'Division',
    'Office', 'OfficePhone', 'MobilePhone', 'Fax', 'HomePhone',
    'StreetAddress', 'City', 'State', 'PostalCode', 'POBox',
    'EmployeeID', 'EmployeeNumber', 'Organization', 'Initials', 'HomePage'
)
$script:AllAttribsExcludedProperties = @(
    'PropertyNames', 'AddedProperties', 'RemovedProperties',
    'ModifiedProperties', 'PropertyCount', 'nTSecurityDescriptor'
)
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
    return ,$table
}
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
function Show-AllAttribsEditDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$AttributeName,
        [string]$CurrentValue
    )
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit Attribute -- $AttributeName"
    $dlg.ClientSize = New-Object System.Drawing.Size(464, 190)
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
    # NO PARAMETERS. The account can change while the window is open, so the
    # address and its domain controller live in $script:Target and
    # $script:ADServer rather than in parameters fixed at build time.
    $allForm = New-Object System.Windows.Forms.Form
    $allForm.Text = 'All Attributes'
    $allForm.ClientSize = New-Object System.Drawing.Size(664, 700)
    $allForm.StartPosition = 'Manual'
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
    $script:AllAttribsCurrentADUser = $null
    function Update-AllAttribsGrid {
        param($ADUserObject)
        $table = Format-AllAttribsTable -ADUserObject $ADUserObject
        $allBindingSource.DataSource = $null
        $allBindingSource.DataSource = $table
        $grdAll.Refresh()
    }
    function Update-AllAttribsFromCurrent {
        if (-not "$script:Target".Trim()) {
            $allForm.Text = 'All Attributes'
            $script:AllAttribsCurrentADUser = $null
            $allBindingSource.DataSource = $null
            $txtAllLog.Text = 'Enter a mailbox address above and press Look Up.'
            $btnAllEdit.Enabled = $false
            return
        }
        $allForm.Text = "All Attributes: $script:Target"
        $btnAllRefresh.Enabled = $false
        $btnAllEdit.Enabled = $false
        $txtAllLog.Text = "Loading all attributes..."
        $allForm.Refresh()
        try {
            $script:AllAttribsCurrentADUser = Get-AllAttribsADUser -Smtp $script:Target -ADServer $script:ADServer
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
    $grdAll.Add_SelectionChanged({
        $hasRow = ($grdAll.SelectedRows.Count -gt 0 -and $null -ne $script:AllAttribsCurrentADUser)
        $isEditable = $false
        if ($hasRow) { $isEditable = ("$($grdAll.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
        $btnAllEdit.Enabled = ($hasRow -and $isEditable)
    })
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
                "Change '$attrName' for $script:Target ($($script:AllAttribsCurrentADUser.SamAccountName)):`n`nFrom: $oldValue`nTo:   $displayNew",
                'Confirm Attribute Edit',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtAllLog.Text = "Cancelled."
                return
            }
            Set-AllAttribsAttribute -SamAccountName $script:AllAttribsCurrentADUser.SamAccountName -AttributeName $attrName -NewValue $newValue -ADServer $script:ADServer
            $refreshed = Get-AllAttribsADUser -Smtp $script:Target -ADServer $script:ADServer
            $script:AllAttribsCurrentADUser = $refreshed
            $afterValue = Format-AllAttribsValue -Value ($refreshed.PSObject.Properties[$attrName].Value)
            Update-AllAttribsGrid -ADUserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE ---")
            [void]$log.AppendLine("$attrName = $oldValue")
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine("$attrName = $afterValue")
            $txtAllLog.Text = $log.ToString().TrimEnd()
            Write-CompletedLog -Task "Edited Attribute $attrName $($script:AllAttribsCurrentADUser.SamAccountName) (All Attribs; $script:Target)"
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
    $grdAll.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $grdAll.ClearSelection()
        $grdAll.Rows[$e.RowIndex].Selected = $true
        Invoke-AllAttribsEditSelectedRow
    })
    $btnAllClose.Add_Click({ $allForm.Close() })

    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down and
    # grows the form, and the help button places itself from ClientSize at run
    # time. Before Set-FormScreenCorner too, which positions from the size.
    $bar = Add-ToolTargetBar -Form $allForm -InitialValue $script:Target

    # RESOLVES, THEN LOADS. Set-AdTarget has to run first: the refresh below
    # needs the domain controller it works out, not just the address.
    $bar.Button.Add_Click({
        Set-AdTarget -Text $bar.TextBox.Text
        Update-AllAttribsFromCurrent
    })

    $allForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $allForm
    Set-FormScreenCorner -Form $allForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $allForm -ScriptPath $PSCommandPath -Title 'All Attributes' | Out-Null

    [void]$allForm.ShowDialog()
}
#endregion

Show-AllAttribsForm
