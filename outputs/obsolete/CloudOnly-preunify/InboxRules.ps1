#Requires -Version 5.1
<#
    InboxRules.ps1
    Standalone Inbox Rules tool -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# View, disable, enable and delete the inbox rules on someone's mailbox.
#
# The usual reasons to look: mail vanishing into a folder nobody remembers
# creating, mail being forwarded somewhere it should not be, or a rule left
# behind by a compromise.
#
# Deleting a rule cannot be undone -- disabling is reversible, so prefer that
# while you are still working out whether a rule is wanted.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -ToolTitle 'Inbox Rules -- Enter Target'
if (-not $target) { return }

#region Actions - inbox rules (moved unchanged from Mailbox-Group-Inspector.ps1 lines 807-982; Get-InboxRulesForMailbox is in Common.ps1)
function Enable-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Enable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Disable-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Disable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Remove-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Update-RulesGrid {
    param([string]$Smtp)
    $grdRules.Rows.Clear()
    $result = Get-InboxRulesForMailbox -Smtp $Smtp
    if ($result.Error) {
        $txtOut.Text = $result.Error
        $script:CurrentRules = @()
        return
    }
    $script:CurrentRules = $result.Rules
    foreach ($r in $result.Rules) {
        $desc = if ($r.Description) { ($r.Description -replace '\s+', ' ').Trim() } else { '' }
        [void]$grdRules.Rows.Add($r.Name, $desc, $r.Enabled, $r.Identity.ToString())
    }
}
function Show-RuleDetailForm {
    param($Rule, [string]$Smtp)
    $ruleForm = New-Object System.Windows.Forms.Form
    $ruleForm.Text = "Inbox Rule: $($Rule.Name)"
    $ruleForm.Size = New-Object System.Drawing.Size(560, 430)
    $ruleForm.StartPosition = "CenterParent"
    $ruleForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = "Rule Name : $($Rule.Name)"
    $lblName.Location = New-Object System.Drawing.Point(12, 12)
    $lblName.AutoSize = $true
    $ruleForm.Controls.Add($lblName)
    $lblMailbox = New-Object System.Windows.Forms.Label
    $lblMailbox.Text = "Mailbox   : $Smtp"
    $lblMailbox.Location = New-Object System.Drawing.Point(12, 32)
    $lblMailbox.AutoSize = $true
    $ruleForm.Controls.Add($lblMailbox)
    $lblEnabled = New-Object System.Windows.Forms.Label
    $lblEnabled.Text = "Enabled   : $($Rule.Enabled)"
    $lblEnabled.Location = New-Object System.Drawing.Point(12, 52)
    $lblEnabled.AutoSize = $true
    $lblEnabled.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $ruleForm.Controls.Add($lblEnabled)
    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = "Full Description:"
    $lblDesc.Location = New-Object System.Drawing.Point(12, 78)
    $lblDesc.AutoSize = $true
    $ruleForm.Controls.Add($lblDesc)
    $txtDesc = New-Object System.Windows.Forms.TextBox
    $txtDesc.Multiline = $true
    $txtDesc.ReadOnly = $true
    $txtDesc.ScrollBars = "Vertical"
    $txtDesc.Location = New-Object System.Drawing.Point(12, 98)
    $txtDesc.Size = New-Object System.Drawing.Size(520, 200)
    $txtDesc.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtDesc.Text = $Rule.Description
    $ruleForm.Controls.Add($txtDesc)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.Location = New-Object System.Drawing.Point(12, 304)
    $lblStatus.Size = New-Object System.Drawing.Size(520, 34)
    $ruleForm.Controls.Add($lblStatus)
    $btnEnable = New-Object System.Windows.Forms.Button
    $btnEnable.Text = "Enable"
    $btnEnable.Location = New-Object System.Drawing.Point(12, 344)
    $btnEnable.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnEnable)
    $btnDisable = New-Object System.Windows.Forms.Button
    $btnDisable.Text = "Disable"
    $btnDisable.Location = New-Object System.Drawing.Point(120, 344)
    $btnDisable.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnDisable)
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "Delete"
    $btnDelete.Location = New-Object System.Drawing.Point(228, 344)
    $btnDelete.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnDelete)
    $btnRuleClose = New-Object System.Windows.Forms.Button
    $btnRuleClose.Text = "Close"
    $btnRuleClose.Location = New-Object System.Drawing.Point(432, 344)
    $btnRuleClose.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnRuleClose)
    $ruleIdentity = $Rule.Identity.ToString()
    $btnEnable.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Enable Inbox rule '$($Rule.Name)' on mailbox '$Smtp'?",
            "Confirm Enable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Enable-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Enabled successfully."
            $lblEnabled.Text = "Enabled   : True"
            Write-CompletedLog -Task "Enabled InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Enable-InboxRule'
        } finally {
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnDisable.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Disable Inbox rule '$($Rule.Name)' on mailbox '$Smtp'?",
            "Confirm Disable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Disable-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Disabled successfully."
            $lblEnabled.Text = "Enabled   : False"
            Write-CompletedLog -Task "Disabled InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Disable-InboxRule'
        } finally {
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnDelete.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Permanently delete Inbox rule '$($Rule.Name)' on mailbox '$Smtp'? This cannot be undone.",
            "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Remove-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            Write-CompletedLog -Task "Deleted InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
            [System.Windows.Forms.MessageBox]::Show("Rule deleted.", "Deleted", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            $ruleForm.Close()
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-InboxRule'
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnRuleClose.Add_Click({ $ruleForm.Close() })
    [void]$ruleForm.ShowDialog()
}
#endregion

#region NEW - wrapper window for the grid (see .DESCRIPTION above)
function Show-InboxRulesForm {
    param([string]$Smtp)
    $rulesForm = New-Object System.Windows.Forms.Form
    $rulesForm.Text = "Inbox Rules: $Smtp"
    $rulesForm.Size = New-Object System.Drawing.Size(780, 480)
    $rulesForm.StartPosition = "CenterScreen"
    $rulesForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

    $lblRules = New-Object System.Windows.Forms.Label
    $lblRules.Text = "Inbox Rules (double-click a row to open):"
    $lblRules.Location = New-Object System.Drawing.Point(12, 12)
    $lblRules.AutoSize = $true
    $rulesForm.Controls.Add($lblRules)

    $grdRules = New-Object System.Windows.Forms.DataGridView
    $grdRules.Location = New-Object System.Drawing.Point(12, 32)
    $grdRules.Size = New-Object System.Drawing.Size(740, 340)
    $grdRules.ReadOnly = $true
    $grdRules.AllowUserToAddRows = $false
    $grdRules.AllowUserToDeleteRows = $false
    $grdRules.AllowUserToResizeRows = $false
    $grdRules.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdRules.MultiSelect = $false
    $grdRules.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdRules.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdRules.Columns.Add('Name', 'Name')
    [void]$grdRules.Columns.Add('Description', 'Description')
    [void]$grdRules.Columns.Add('Enabled', 'Enabled')
    [void]$grdRules.Columns.Add('Identity', 'Identity')
    $grdRules.Columns['Name'].FillWeight        = 25
    $grdRules.Columns['Description'].FillWeight = 45
    $grdRules.Columns['Enabled'].FillWeight     = 10
    $grdRules.Columns['Identity'].Visible       = $false
    foreach ($col in $grdRules.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }
    $rulesForm.Controls.Add($grdRules)

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline = $true
    $txtOut.ReadOnly = $true
    $txtOut.ScrollBars = "Vertical"
    $txtOut.Location = New-Object System.Drawing.Point(12, 380)
    $txtOut.Size = New-Object System.Drawing.Size(740, 44)
    $txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
    $rulesForm.Controls.Add($txtOut)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 428)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $rulesForm.Controls.Add($btnRefresh)
    $btnRulesClose = New-Object System.Windows.Forms.Button
    $btnRulesClose.Text = "Close"
    $btnRulesClose.Location = New-Object System.Drawing.Point(652, 428)
    $btnRulesClose.Size = New-Object System.Drawing.Size(100, 28)
    $rulesForm.Controls.Add($btnRulesClose)

    $script:CurrentRules = @()

    Update-RulesGrid -Smtp $Smtp
    $btnRefresh.Add_Click({ Update-RulesGrid -Smtp $Smtp })
    $grdRules.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $identity = $grdRules.Rows[$e.RowIndex].Cells['Identity'].Value
        $rule = $script:CurrentRules | Where-Object { $_.Identity.ToString() -eq $identity } | Select-Object -First 1
        if (-not $rule) {
            $txtOut.Text = "Could not match that row back to a rule -- click Refresh and try again."
            return
        }
        Show-RuleDetailForm -Rule $rule -Smtp $Smtp
        Update-RulesGrid -Smtp $Smtp
    })
    $btnRulesClose.Add_Click({ $rulesForm.Close() })
    Set-ConsoleTheme -Form $rulesForm
    Add-ToolHelpButton -Form $rulesForm -ScriptPath $PSCommandPath -Title 'Inbox Rules' | Out-Null
    Set-FormScreenCorner -Form $rulesForm -Corner 'BottomLeft'
    [void]$rulesForm.ShowDialog()
}
#endregion

Show-InboxRulesForm -Smtp $target.Smtp
