#Requires -Version 5.1
<#
    CalendarPermissions.ps1
    Standalone Calendar Permissions tool -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# See and change who can view or edit someone's calendar.
#
# Shows the current permission for each person who has one, and lets you add,
# change or remove access at the usual levels -- from free/busy only through to
# full editing.
#
# The most common ticket this answers is "my assistant can't see my calendar"
# or "someone can see more of my calendar than they should".
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
# Add-ToolTargetBar in Common.ps1. An empty address is fine; the window opens
# and asks for one.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Actions - calendar permissions (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1272-1455)
function Get-CalendarFolderIdentity {
    param([string]$Smtp)
    Write-ToolStep 'Finding the calendar folder ...'
    $folders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{ Identity = $Smtp; FolderScope = 'Calendar'; ErrorAction = 'Stop' })
    if ($folders.Count -eq 0) { throw "No Calendar-scope folder found for $Smtp." }
    $primary = $folders | Where-Object { $_.FolderPath -eq '/Calendar' } | Select-Object -First 1
    if (-not $primary) { $primary = $folders | Select-Object -First 1 }
    return "$($Smtp):\$($primary.Name)"
}
function Get-CalendarPermissions {
    param([string]$Smtp)
    try {
        $calId = Get-CalendarFolderIdentity -Smtp $Smtp
        Write-ToolStep 'Reading calendar permissions ...'
        $perms = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $calId; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ CalendarIdentity = $calId; Permissions = $perms; Error = $null }
    } catch {
        return [PSCustomObject]@{ CalendarIdentity = $null; Permissions = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxFolderPermission (Calendar)') }
    }
}
function Set-CalendarPermission {
    param([string]$CalendarIdentity, [string]$Grantee, [string]$Role)
    $existing = @()
    try {
        $existing = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $CalendarIdentity; User = $Grantee; ErrorAction = 'Stop' })
    } catch {
        $existing = @()
    }
    if ($existing.Count -gt 0) {
        Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Updated'
    } else {
        Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Added'
    }
}
function Remove-CalendarPermission {
    param([string]$CalendarIdentity, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxFolderPermission' -Parameters @{
        Identity = $CalendarIdentity; User = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Show-CalendarPermissionsForm {
    # NO -Smtp PARAMETER. The address can change while the window is open, so it
    # lives in $script:Target rather than a parameter fixed at build time.
    $calForm = New-Object System.Windows.Forms.Form
    $calForm.Text = "Calendar Permissions: $script:Target"
    $calForm.ClientSize = New-Object System.Drawing.Size(604, 440)
    $calForm.StartPosition = 'Manual'
    $calForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblCalId = New-Object System.Windows.Forms.Label
    $lblCalId.Text = "Calendar folder:"
    $lblCalId.Location = New-Object System.Drawing.Point(12, 12)
    $lblCalId.AutoSize = $true
    $calForm.Controls.Add($lblCalId)
    $grdCal = New-Object System.Windows.Forms.DataGridView
    $grdCal.Location = New-Object System.Drawing.Point(12, 34)
    $grdCal.Size = New-Object System.Drawing.Size(580, 280)
    $grdCal.ReadOnly = $true
    $grdCal.AllowUserToAddRows = $false
    $grdCal.AllowUserToDeleteRows = $false
    $grdCal.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdCal.MultiSelect = $false
    $grdCal.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdCal.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdCal.Columns.Add('User', 'User')
    [void]$grdCal.Columns.Add('AccessRights', 'Access Rights')
    foreach ($c in $grdCal.Columns) { $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }
    $calForm.Controls.Add($grdCal)
    $lblGrantee = New-Object System.Windows.Forms.Label
    $lblGrantee.Text = "User (email or logon):"
    $lblGrantee.Location = New-Object System.Drawing.Point(12, 326)
    $lblGrantee.AutoSize = $true
    $calForm.Controls.Add($lblGrantee)
    $txtGrantee = New-Object System.Windows.Forms.TextBox
    $txtGrantee.Location = New-Object System.Drawing.Point(160, 323)
    $txtGrantee.Size = New-Object System.Drawing.Size(220, 20)
    $txtGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
    $calForm.Controls.Add($txtGrantee)
    $lblRole = New-Object System.Windows.Forms.Label
    $lblRole.Text = "Role:"
    $lblRole.Location = New-Object System.Drawing.Point(390, 326)
    $lblRole.AutoSize = $true
    $calForm.Controls.Add($lblRole)
    $cboRole = New-Object System.Windows.Forms.ComboBox
    $cboRole.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Reviewer', 'Editor', 'Owner') | ForEach-Object { [void]$cboRole.Items.Add($_) }
    $cboRole.SelectedIndex = 0
    $cboRole.Location = New-Object System.Drawing.Point(430, 323)
    $cboRole.Size = New-Object System.Drawing.Size(120, 20)
    $calForm.Controls.Add($cboRole)
    $lblCalStatus = New-Object System.Windows.Forms.Label
    $lblCalStatus.Text = ""
    $lblCalStatus.Location = New-Object System.Drawing.Point(12, 356)
    $lblCalStatus.Size = New-Object System.Drawing.Size(580, 40)
    $calForm.Controls.Add($lblCalStatus)
    $btnAddPerm = New-Object System.Windows.Forms.Button
    $btnAddPerm.Text = "Add / Update Permission"
    $btnAddPerm.Location = New-Object System.Drawing.Point(12, 400)
    $btnAddPerm.Size = New-Object System.Drawing.Size(170, 28)
    $calForm.Controls.Add($btnAddPerm)
    $btnRemovePerm = New-Object System.Windows.Forms.Button
    $btnRemovePerm.Text = "Remove Selected"
    $btnRemovePerm.Location = New-Object System.Drawing.Point(190, 400)
    $btnRemovePerm.Size = New-Object System.Drawing.Size(150, 28)
    $calForm.Controls.Add($btnRemovePerm)
    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "Export CSV"
    $btnExport.Location = New-Object System.Drawing.Point(348, 400)
    $btnExport.Size = New-Object System.Drawing.Size(140, 28)
    $calForm.Controls.Add($btnExport)
    $btnCalClose = New-Object System.Windows.Forms.Button
    $btnCalClose.Text = "Close"
    $btnCalClose.Location = New-Object System.Drawing.Point(500, 400)
    $btnCalClose.Size = New-Object System.Drawing.Size(90, 28)
    $calForm.Controls.Add($btnCalClose)
    $script:CalendarIdentityForForm = $null
    function Update-CalGridFromCurrent {
        if (-not "$script:Target".Trim()) {
            $calForm.Text = 'Calendar Permissions'
            return
        }
        $calForm.Text = "Calendar Permissions: $script:Target"
        $grdCal.Rows.Clear()
        $result = Get-CalendarPermissions -Smtp $script:Target
        if ($result.Error) {
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = $result.Error
            $lblCalId.Text = "Calendar folder: (unresolved)"
            return
        }
        $script:CalendarIdentityForForm = $result.CalendarIdentity
        $lblCalId.Text = "Calendar folder: $($result.CalendarIdentity)"
        foreach ($p in $result.Permissions) {
            [void]$grdCal.Rows.Add($p.User.ToString(), ($p.AccessRights -join ', '))
        }
    }
    Update-CalGridFromCurrent

    $btnExport.Add_Click({
        <#
            Write what is on screen to the reports folder.

            EXPORTS THE GRID, NOT A FRESH LOOKUP. What the operator is looking
            at is what they mean to export -- re-querying here could hand them a
            file that disagrees with the window it came from, which is the one
            thing a report must never do.

            THE FILE NAME IS THE MAILBOX, NOT A TIMESTAMP. Asked for as
            <local part>_CalendarPermissions.csv, so 1357VacationCal@... becomes
            1357VacationCal_CalendarPermissions.csv. That means a second export
            of the same mailbox OVERWRITES the first, which is the intended
            behaviour of a fixed name: the file is "the current permissions on
            this calendar", not a history.
        #>
        try {
            $target = "$script:Target".Trim()
            if (-not $target) {
                $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
                $lblCalStatus.Text = 'No mailbox loaded -- nothing to export.'
                return
            }
            if ($grdCal.Rows.Count -eq 0) {
                # An empty CSV is worse than no CSV. A file with only headers
                # looks like a mailbox with no calendar permissions, when what
                # actually happened is that the lookup never returned.
                $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
                $lblCalStatus.Text = 'No permissions listed -- no file was written.'
                return
            }

            # LOCAL PART ONLY, AND THEN SCRUBBED.
            #
            # Split on '@' so a plain name with no domain still works. The
            # scrub is not paranoia: a calendar can be named with a slash or a
            # colon, and Join-Path would then build a path into a directory that
            # does not exist and fail with a message about the path rather than
            # the name.
            $local = ("$target" -split '@')[0]
            foreach ($bad in [System.IO.Path]::GetInvalidFileNameChars()) {
                $local = $local.Replace($bad, '_')
            }
            if (-not $local.Trim()) { $local = 'calendar' }

            $dir = Get-ReportPath
            if (-not (Test-Path -LiteralPath $dir)) {
                New-Item -Path $dir -ItemType Directory -Force | Out-Null
            }
            $csv = Join-Path $dir "${local}_CalendarPermissions.csv"

            # The mailbox and folder ride in every row. A CSV that is only User
            # and AccessRights is unusable the moment two of them are open side
            # by side, or one is pasted into a ticket.
            $rows = foreach ($r in $grdCal.Rows) {
                if ($r.IsNewRow) { continue }
                [PSCustomObject]@{
                    Mailbox        = $target
                    CalendarFolder = "$script:CalendarIdentityForForm"
                    User           = "$($r.Cells['User'].Value)"
                    AccessRights   = "$($r.Cells['AccessRights'].Value)"
                    ExportedUtc    = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
                }
            }

            @($rows) | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

            $lblCalStatus.ForeColor = [System.Drawing.Color]::Black
            $lblCalStatus.Text = "Exported $(@($rows).Count) row(s) to $csv"

            # No -Environment, deliberately, and matching the other three log
            # calls in this file. The script takes an $Environment parameter but
            # nothing here has ever passed it on -- Write-CompletedLog fills the
            # column itself. Reading it from inside a Click handler would also be
            # the scope trap the console's timer comment describes: a handler
            # resolves its variables against whatever scope is current when it
            # FIRES, not where it was written.
            Write-CompletedLog -Action 'ExportCalendarPermissions' `
                -Target (Get-LogTarget $target) -Result 'Success' `
                -Detail "$(@($rows).Count) permission(s) written to $csv"
        } catch {
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Export calendar permissions'
            Write-FailedLog -Action 'ExportCalendarPermissions' -Target "$script:Target" -ErrorRecord $_
        }
    })

    $btnAddPerm.Add_Click({
        $grantee = $txtGrantee.Text.Trim()
        if (-not $grantee) { $lblCalStatus.Text = "Enter a user first."; return }
        if (-not $script:CalendarIdentityForForm) { $lblCalStatus.Text = "Calendar folder not resolved -- see status above."; return }
        $confirm = Show-ToolDialog(
            "Grant '$($cboRole.SelectedItem)' calendar access to '$grantee' on '$script:Target'?",
            "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnAddPerm.Enabled = $false; $btnRemovePerm.Enabled = $false
        try {
            $action = Set-CalendarPermission -CalendarIdentity $script:CalendarIdentityForForm -Grantee $grantee -Role $cboRole.SelectedItem
            Write-CompletedLog -Task "$action CalendarPermission $script:Target (EXO; user: $grantee; role: $($cboRole.SelectedItem))"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Black
            $lblCalStatus.Text = "$action permission for $grantee."
            Update-CalGridFromCurrent
        } catch {
            Write-FailedLog -Action 'GrantCalendarPermission' -Target $script:Target -ErrorRecord $_ -Detail "user: $grantee; role: $($cboRole.SelectedItem)"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Calendar permission grant'
        } finally {
            $btnAddPerm.Enabled = $true; $btnRemovePerm.Enabled = $true
        }
    })
    $btnRemovePerm.Add_Click({
        if ($grdCal.SelectedRows.Count -eq 0) { $lblCalStatus.Text = "Select a row first."; return }
        $grantee = $grdCal.SelectedRows[0].Cells['User'].Value
        $confirm = Show-ToolDialog(
            "Remove all calendar permissions for '$grantee' on '$script:Target'?",
            "Confirm Remove", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnAddPerm.Enabled = $false; $btnRemovePerm.Enabled = $false
        try {
            Remove-CalendarPermission -CalendarIdentity $script:CalendarIdentityForForm -Grantee $grantee
            Write-CompletedLog -Task "Removed CalendarPermission $script:Target (EXO; user: $grantee)"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Black
            $lblCalStatus.Text = "Removed permission for $grantee."
            Update-CalGridFromCurrent
        } catch {
            Write-FailedLog -Action 'RemoveCalendarPermission' -Target $script:Target -ErrorRecord $_ -Detail "user: $grantee"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-MailboxFolderPermission (Calendar)'
        } finally {
            $btnAddPerm.Enabled = $true; $btnRemovePerm.Enabled = $true
        }
    })
    $btnCalClose.Add_Click({ $calForm.Close() })
    $calForm.StartPosition = 'Manual'
    # BEFORE the theme and the help button: it shifts every existing control
    # down, so anything added after would be left at the old offset.
    $bar = Add-ToolTargetBar -Form $calForm -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-CalGridFromCurrent
    })

    Set-ConsoleTheme -Form $calForm
    Set-FormScreenCorner -Form $calForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $calForm -ScriptPath $PSCommandPath -Title 'Calendar Permissions' | Out-Null

    [void]$calForm.ShowDialog()
}
#endregion

Show-CalendarPermissionsForm
