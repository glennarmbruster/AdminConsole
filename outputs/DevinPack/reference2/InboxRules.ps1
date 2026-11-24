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
# Add-ToolTargetBar in Common.ps1. Resolve-ToolTarget is not used any more:
# its job was to decide which DIRECTORY an address belongs to, and every read
# and write below goes through Invoke-EXOCommand to Exchange Online, which
# does not need to be told.
#
# An empty address is fine. The window opens and asks.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

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
    # STILL TAKES -Smtp. This one is a script-scope function that reaches up
    # into the form's controls, so it is not the place to introduce a second
    # source of truth for the address -- the caller passes $script:Target.
    param([string]$Smtp)
    $grdRules.Rows.Clear()
    if (-not "$Smtp".Trim()) {
        $txtOut.Text = 'Enter a mailbox address above and press Look Up.'
        $script:CurrentRules = @()
        return
    }
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
    $ruleForm.ClientSize = New-Object System.Drawing.Size(544, 384)
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
    # NO -Smtp PARAMETER. The mailbox can change while the window is open, so
    # it lives in $script:Target rather than a parameter fixed at build time.
    $rulesForm = New-Object System.Windows.Forms.Form
    $rulesForm.Text = 'Inbox Rules'
    $rulesForm.ClientSize = New-Object System.Drawing.Size(764, 468)
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

    function Update-RulesFromCurrent {
        $rulesForm.Text = if ("$script:Target".Trim()) { "Inbox Rules: $script:Target" } else { 'Inbox Rules' }
        Update-RulesGrid -Smtp $script:Target
    }

    Update-RulesFromCurrent
    $btnRefresh.Add_Click({ Update-RulesFromCurrent })
    $grdRules.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $identity = $grdRules.Rows[$e.RowIndex].Cells['Identity'].Value
        $rule = $script:CurrentRules | Where-Object { $_.Identity.ToString() -eq $identity } | Select-Object -First 1
        if (-not $rule) {
            $txtOut.Text = "Could not match that row back to a rule -- click Refresh and try again."
            return
        }
        Show-RuleDetailForm -Rule $rule -Smtp $script:Target
        Update-RulesFromCurrent
    })
    $btnRulesClose.Add_Click({ $rulesForm.Close() })

    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down and
    # grows the form, and the help button places itself from ClientSize at run
    # time. Before Set-FormScreenCorner too, which positions from the size.
    $bar = Add-ToolTargetBar -Form $rulesForm -InitialValue $script:Target

    # WIRED HERE, NOT PASSED INTO THE HELPER. Update-RulesFromCurrent is nested
    # in this function, so the handler has to be built in this scope to see it.
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-RulesFromCurrent
    })

    Set-ConsoleTheme -Form $rulesForm
    Set-FormScreenCorner -Form $rulesForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $rulesForm -ScriptPath $PSCommandPath -Title 'Inbox Rules' | Out-Null

    [void]$rulesForm.ShowDialog()
}
#endregion

Show-InboxRulesForm
