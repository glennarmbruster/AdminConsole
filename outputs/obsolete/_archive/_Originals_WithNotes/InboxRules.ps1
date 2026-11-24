#Requires -Version 5.1
<#
.SYNOPSIS
    Standalone Inbox Rules tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\InboxRules.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the Inbox Rules window.

    Fully functional from AdminConsole.ps1 too:
        .\InboxRules.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Inbox Rules window
    for that mailbox.

    Enable-TargetInboxRule, Disable-TargetInboxRule, Remove-TargetInboxRule,
    Update-RulesGrid, and Show-RuleDetailForm below are MOVED, UNCHANGED,
    from the original monolithic Mailbox-Group-Inspector.ps1 (original lines
    807-982, region "Actions - inbox rules"; Get-InboxRulesForMailbox from
    that same region now lives in Common.ps1 -- see its header for why).

    UNLIKE the other tool scripts extracted so far, the grid this tool
    displays (grdRules) never had its own Show-XForm in the monolith -- it
    lived directly on the monolith's shared main form, and Update-RulesGrid
    referenced $grdRules/$txtOut as outer-scope variables belonging to that
    shared form (see Update-RulesGrid's own original comment: "Referencing
    $grdRules/$txtOut here without $script: works the same way
    Invoke-MailboxEdit/Invoke-GroupEdit already do -- they're script-scope
    (top-level) variables by the time any button handler calls into this
    function"). So this file adds ONE new function, Show-InboxRulesForm,
    that did not exist in the monolith -- it is a small new wrapper window
    built to hold exactly the same grid layout (same columns, same
    double-click-for-detail interaction) the monolith drew directly on its
    main form, so that Update-RulesGrid and Show-RuleDetailForm can be moved
    here completely unchanged. This is the same treatment
    MailboxPermissions.ps1 and GroupEditor.ps1 need for their own
    main-form-embedded controls -- see REFACTOR-DESIGN.md Section 7.

    Mechanical-only changes to the moved code:
      1. Show-RuleDetailForm's closing call changed from
             [void]$ruleForm.ShowDialog($form)
         to
             [void]$ruleForm.ShowDialog()
         because there is no longer a parent $form to own it.
      2. Update-RulesGrid's $grdRules/$txtOut references now resolve to the
         controls built in the new Show-InboxRulesForm wrapper instead of
         the monolith's main form -- same "script-scope variables by the
         time any handler calls into this function" mechanism the original
         comment already describes, just with a different (also top-level,
         relative to this process) form supplying them.
    Every other line, including all comments, is byte-for-byte the original.
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

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'Inbox Rules -- Enter Target'
if (-not $target) { return }   # operator cancelled the prompt

#region Actions - inbox rules (moved unchanged from Mailbox-Group-Inspector.ps1 lines 807-982; Get-InboxRulesForMailbox is in Common.ps1)
function Enable-TargetInboxRule {
    # BUGFIX (2026-07-09): do NOT also pass -Mailbox here. $RuleIdentity comes
    # from Get-InboxRule -Mailbox $Smtp (see Update-RulesGrid), and Exchange
    # returns that Identity already scoped to the mailbox it was queried
    # against. Passing -Mailbox in addition to that already-scoped -Identity
    # throws "You can't use both the Identity and Mailbox parameters to
    # specify a mailbox" -- confirmed against the Enable-InboxRule syntax on
    # Microsoft Learn, where -Mailbox exists only to disambiguate a bare rule
    # NAME (per the Disable-InboxRule/Remove-InboxRule examples on the same
    # page), which is not what's happening here. $Smtp is still accepted as a
    # parameter and used by the caller for logging (Write-CompletedLog).
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Enable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Disable-TargetInboxRule {
    # BUGFIX (2026-07-09): see the comment above Enable-TargetInboxRule --
    # same fix, same reason (this is the exact cmdlet/error the operator hit).
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Disable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-TargetInboxRule {
    # BUGFIX (2026-07-09): see the comment above Enable-TargetInboxRule --
    # same latent bug (not yet hit in the field, but identical cause).
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Remove-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Update-RulesGrid {
    # Repopulates the grdRules grid (defined in the WinForms UI region) for
    # the given mailbox. Referencing $grdRules/$txtOut here without $script:
    # works the same way Invoke-MailboxEdit/Invoke-GroupEdit already do --
    # they're script-scope (top-level) variables by the time any button
    # handler calls into this function.
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
    # CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the
    # monolith's shared parent window. This script has no parent window, so
    # it shows itself as the top-level dialog instead. Nothing else in this
    # function changed.
    [void]$ruleForm.ShowDialog()
}
#endregion

#region NEW - wrapper window for the grid (see .DESCRIPTION above)
# This function did not exist in the monolith. It exists only to give
# Update-RulesGrid the $grdRules/$txtOut controls it already expected to
# find in outer scope, using the exact same grid layout/columns the
# monolith drew directly on its shared main form.
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
    # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
    # Console, which opens top-left. Helper is in Common.ps1.
    Set-FormScreenCorner -Form $rulesForm -Corner 'BottomLeft'
    [void]$rulesForm.ShowDialog()
}
#endregion

Show-InboxRulesForm -Smtp $target.Smtp
