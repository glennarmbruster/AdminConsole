#Requires -Version 5.1
<#
    Set-MailboxAccess.ps1
    WinForms tool to grant or remove mailbox delegation for multiple
    mailboxes and users in a single run.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Config  ---------------------------------------------------------------

$Domain              = "@expresspros.com"
$CompletedTaskFolder = (Join-Path $PSScriptRoot 'Logs')
$CompletedTaskFile   = Join-Path $CompletedTaskFolder "CompletedTasks.csv"

#endregion

#region Helpers  --------------------------------------------------------------

function Get-AddressList {
    param([string]$Raw)

    if ([string]::IsNullOrWhiteSpace($Raw)) { return @() }

    $tokens = $Raw -split "[\r\n,;]+"

    $clean = foreach ($t in $tokens) {
        $name = $t.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($name -notmatch '@') { $name = $name + $Domain }
        $name
    }

    return @($clean | Sort-Object -Unique)
}

function Write-CompletedLog {
    param([string]$Task)
    if (-not (Test-Path -LiteralPath $CompletedTaskFolder)) {
        New-Item -Path $CompletedTaskFolder -ItemType Directory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $CompletedTaskFile)) {
        "Task,DateTime" | Out-File -FilePath $CompletedTaskFile -Encoding utf8
    }
    $dt = Get-Date -Format "yyyy-MM-dd HHmm"
    if ($Task -and $Task.Trim() -ne "") {
        Add-Content -Path $CompletedTaskFile -Value ('"{0}","{1}"' -f $Task.Trim(), $dt)
    }
}

#endregion

#region Parsing  --------------------------------------------------------------

function Get-ParsedInputs {
    param(
        [string]$MailboxRaw,
        [string]$UserRaw,
        [string]$Action
    )

    $mailboxes = Get-AddressList -Raw $MailboxRaw
    $users     = Get-AddressList -Raw $UserRaw

    $users = @($users | Where-Object { $mailboxes -notcontains $_ })

    return [PSCustomObject]@{
        Mailboxes = $mailboxes
        Users     = $users
        Action    = $Action
        OpCount   = ($mailboxes.Count * $users.Count)
    }
}

#endregion

#region Actions  --------------------------------------------------------------

function Invoke-MailboxAccess {
    param(
        [string[]]$Mailboxes,
        [string[]]$Users,
        [string]$Action
    )

    $log = New-Object System.Text.StringBuilder

    $actionText = switch ($Action) {
        "FullSend" { "Full Access + Send-As" }
        "Full"     { "Full Access only" }
        "Remove"   { "Remove All Access" }
    }

    [void]$log.AppendLine("=== Mailbox Access Manager ===")
    [void]$log.AppendLine("Action      : $actionText")
    [void]$log.AppendLine("Mailbox(es) : $($Mailboxes -join ', ')")
    [void]$log.AppendLine("User(s)     : $($Users -join ', ')")
    [void]$log.AppendLine("")

    $notFound = @()

    foreach ($m in $Mailboxes) {
        try   { Get-Mailbox -Identity $m -ErrorAction Stop | Out-Null }
        catch { $notFound += $m }
    }
    foreach ($u in $Users) {
        try   { Get-Mailbox -Identity $u -ErrorAction Stop | Out-Null }
        catch { $notFound += $u }
    }

    if ($notFound.Count -gt 0) {
        $list = ($notFound | Sort-Object -Unique) -join "`r`n    "
        [void]$log.AppendLine("PRE-FLIGHT FAILED � the following could not be resolved:")
        [void]$log.AppendLine("    $list")
        [void]$log.AppendLine("")
        [void]$log.AppendLine("No changes were made. Correct the above entries and try again.")
        return $log.ToString()
    }

    [void]$log.AppendLine("Pre-flight : all identities resolved OK.")
    [void]$log.AppendLine("")

    $succeeded = [System.Collections.Generic.List[PSCustomObject]]::new()
    $failed    = [System.Collections.Generic.List[string]]::new()
    $skipped   = [System.Collections.Generic.List[string]]::new()

    foreach ($m in $Mailboxes) {

        if ($Action -eq "FullSend") {
            try {
                Set-Mailbox -Identity $m -MessageCopyForSentAs $true -ErrorAction Stop
                [void]$log.AppendLine("  [OK]   MessageCopyForSentAs = true on $m")
            } catch {
                [void]$log.AppendLine("  [WARN] Could not set MessageCopyForSentAs on $m : $($_.Exception.Message)")
            }
        }

        foreach ($u in $Users) {

            switch ($Action) {

                "FullSend" {
                    try {
                        Add-MailboxPermission    -Identity $m -User    $u -AccessRights FullAccess -AutoMapping:$true -ErrorAction Stop | Out-Null
                        Add-RecipientPermission  -Identity $m -Trustee $u -AccessRights SendAs     -Confirm:$false    -ErrorAction Stop | Out-Null
                        [void]$log.AppendLine("  [OK]   Granted Full + Send-As : $u -> $m")
                        $succeeded.Add([PSCustomObject]@{ Mailbox=$m; User=$u; Op="Full+SendAs" })
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Full + Send-As : $u -> $m : $($_.Exception.Message)")
                        $failed.Add("Full+SendAs: $u -> $m")
                    }
                }

                "Full" {
                    try {
                        Add-MailboxPermission -Identity $m -User $u -AccessRights FullAccess -AutoMapping:$true -ErrorAction Stop | Out-Null
                        [void]$log.AppendLine("  [OK]   Granted Full Access : $u -> $m")
                        $succeeded.Add([PSCustomObject]@{ Mailbox=$m; User=$u; Op="Full" })
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Full Access : $u -> $m : $($_.Exception.Message)")
                        $failed.Add("Full: $u -> $m")
                    }
                }

                "Remove" {
                    $removedSomething = $false

                    try {
                        $fa = Get-MailboxPermission -Identity $m -User $u -ErrorAction SilentlyContinue |
                              Where-Object { $_.AccessRights -contains "FullAccess" }
                        if ($fa) {
                            Remove-MailboxPermission -Identity $m -User $u -AccessRights FullAccess `
                                -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                            $removedSomething = $true
                        }
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Remove FullAccess : $u -> $m : $($_.Exception.Message)")
                        $failed.Add("RemoveFull: $u -> $m")
                    }

                    try {
                        $sa = Get-RecipientPermission -Identity $m -Trustee $u -ErrorAction SilentlyContinue |
                              Where-Object { $_.AccessRights -contains "SendAs" }
                        if ($sa) {
                            Remove-RecipientPermission -Identity $m -Trustee $u -AccessRights SendAs `
                                -Confirm:$false -ErrorAction Stop | Out-Null
                            $removedSomething = $true
                        }
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Remove Send-As : $u -> $m : $($_.Exception.Message)")
                        $failed.Add("RemoveSendAs: $u -> $m")
                    }

                    if ($removedSomething) {
                        [void]$log.AppendLine("  [OK]   Removed access : $u -> $m")
                        $succeeded.Add([PSCustomObject]@{ Mailbox=$m; User=$u; Op="Remove" })
                    } else {
                        [void]$log.AppendLine("  [SKIP] No access found : $u -> $m")
                        $skipped.Add("$u -> $m")
                    }
                }
            }
        }
    }

    [void]$log.AppendLine("")
    [void]$log.AppendLine(("Summary : {0} succeeded, {1} failed, {2} with nothing to do." -f `
        $succeeded.Count, $failed.Count, $skipped.Count))

    [void]$log.AppendLine("")
    [void]$log.AppendLine("=== Verification ===")

    foreach ($m in $Mailboxes) {
        [void]$log.AppendLine("")
        [void]$log.AppendLine("Full Access on $m :")
        try {
            $faList = Get-MailboxPermission -Identity $m -ErrorAction Stop |
                      Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-5-*" }
            if ($faList) {
                foreach ($p in $faList) {
                    [void]$log.AppendLine("    $($p.User)  :  $($p.AccessRights -join ', ')")
                }
            } else {
                [void]$log.AppendLine("    (none)")
            }
        } catch {
            [void]$log.AppendLine("    (error: $($_.Exception.Message))")
        }

        [void]$log.AppendLine("Send-As on $m :")
        try {
            $saList = Get-RecipientPermission -Identity $m -ErrorAction Stop |
                      Where-Object { $_.Trustee -notlike "NT AUTHORITY\*" }
            if ($saList) {
                foreach ($p in $saList) {
                    [void]$log.AppendLine("    $($p.Trustee)  :  $($p.AccessRights -join ', ')")
                }
            } else {
                [void]$log.AppendLine("    (none)")
            }
        } catch {
            [void]$log.AppendLine("    (error: $($_.Exception.Message))")
        }
    }

    if ($succeeded.Count -gt 0) {
        $verb     = if ($Action -eq "Remove") { "Removed" } else { "Granted" }
        $userList = $Users -join "+"
        foreach ($m in $Mailboxes) {
            $mSuccess = @($succeeded | Where-Object { $_.Mailbox -eq $m })
            if ($mSuccess.Count -gt 0) {
                Write-CompletedLog -Task "$verb MailboxAccess [$actionText] $m (users: $userList)"
            }
        }
    }

    return $log.ToString()
}

#endregion

#region WinForms UI  ----------------------------------------------------------

$script:Parsed     = $null
$script:ParsedFrom = $null

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Mailbox Access Manager"
$form.Size            = New-Object System.Drawing.Size(760, 750)
$form.StartPosition   = 'Manual'
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Text      = "Make sure you are connected to Exchange Online before running."
$lblWarn.Location  = New-Object System.Drawing.Point(12, 10)
$lblWarn.AutoSize  = $true
$lblWarn.Font      = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$lblMailbox          = New-Object System.Windows.Forms.Label
$lblMailbox.Text     = "Mailbox(es) � one per line (enter email address, @expresspros.com optional):"
$lblMailbox.AutoSize = $true
$lblMailbox.Location = New-Object System.Drawing.Point(12, 38)
$form.Controls.Add($lblMailbox)

$txtMailbox               = New-Object System.Windows.Forms.TextBox
$txtMailbox.Multiline     = $true
$txtMailbox.ScrollBars    = "Vertical"
$txtMailbox.Size          = New-Object System.Drawing.Size(722, 100)
$txtMailbox.Location      = New-Object System.Drawing.Point(12, 58)
$txtMailbox.Font          = New-Object System.Drawing.Font("Consolas", 9)
$txtMailbox.AcceptsReturn = $true
$form.Controls.Add($txtMailbox)

$lblUser          = New-Object System.Windows.Forms.Label
$lblUser.Text     = "User(s) to grant/remove access � one per line (enter email address, @expresspros.com optional):"
$lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(12, 170)
$form.Controls.Add($lblUser)

$txtUser               = New-Object System.Windows.Forms.TextBox
$txtUser.Multiline     = $true
$txtUser.ScrollBars    = "Vertical"
$txtUser.Size          = New-Object System.Drawing.Size(722, 100)
$txtUser.Location      = New-Object System.Drawing.Point(12, 190)
$txtUser.Font          = New-Object System.Drawing.Font("Consolas", 9)
$txtUser.AcceptsReturn = $true
$form.Controls.Add($txtUser)

$lblParsed          = New-Object System.Windows.Forms.Label
$lblParsed.Text     = "Parsed preview � review before clicking Run:"
$lblParsed.AutoSize = $true
$lblParsed.Location = New-Object System.Drawing.Point(12, 302)
$form.Controls.Add($lblParsed)

$txtParsed              = New-Object System.Windows.Forms.TextBox
$txtParsed.Multiline    = $true
$txtParsed.ScrollBars   = "Vertical"
$txtParsed.ReadOnly     = $true
$txtParsed.BackColor    = [System.Drawing.Color]::WhiteSmoke
$txtParsed.Size         = New-Object System.Drawing.Size(722, 85)
$txtParsed.Location     = New-Object System.Drawing.Point(12, 322)
$txtParsed.Font         = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtParsed)

$grpAction          = New-Object System.Windows.Forms.GroupBox
$grpAction.Text     = "Action"
$grpAction.Size     = New-Object System.Drawing.Size(722, 105)
$grpAction.Location = New-Object System.Drawing.Point(12, 416)
$form.Controls.Add($grpAction)

$rbFullSend          = New-Object System.Windows.Forms.RadioButton
$rbFullSend.Text     = "Add Full Access && Send-As"
$rbFullSend.AutoSize = $true
$rbFullSend.Location = New-Object System.Drawing.Point(15, 25)
$rbFullSend.Checked  = $true
$grpAction.Controls.Add($rbFullSend)

$rbFull          = New-Object System.Windows.Forms.RadioButton
$rbFull.Text     = "Add Full Access only"
$rbFull.AutoSize = $true
$rbFull.Location = New-Object System.Drawing.Point(15, 52)
$grpAction.Controls.Add($rbFull)

$rbRemove          = New-Object System.Windows.Forms.RadioButton
$rbRemove.Text     = "Remove all access (Full + Send-As)"
$rbRemove.AutoSize = $true
$rbRemove.Location = New-Object System.Drawing.Point(15, 79)
$grpAction.Controls.Add($rbRemove)

$btnParse          = New-Object System.Windows.Forms.Button
$btnParse.Text     = "Parse"
$btnParse.Size     = New-Object System.Drawing.Size(80, 28)
$btnParse.Location = New-Object System.Drawing.Point(490, 532)
$form.Controls.Add($btnParse)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = "Run"
$btnRun.Size     = New-Object System.Drawing.Size(80, 28)
$btnRun.Location = New-Object System.Drawing.Point(578, 532)
$form.Controls.Add($btnRun)

$btnCancel          = New-Object System.Windows.Forms.Button
$btnCancel.Text     = "Close"
$btnCancel.Size     = New-Object System.Drawing.Size(80, 28)
$btnCancel.Location = New-Object System.Drawing.Point(666, 532)
$form.Controls.Add($btnCancel)

$lblResult          = New-Object System.Windows.Forms.Label
$lblResult.Text     = "Result:"
$lblResult.AutoSize = $true
$lblResult.Location = New-Object System.Drawing.Point(12, 572)
$form.Controls.Add($lblResult)

$txtOut           = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.ReadOnly  = $true
$txtOut.Size      = New-Object System.Drawing.Size(722, 130)
$txtOut.Location  = New-Object System.Drawing.Point(12, 592)
$txtOut.Font      = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOut)

function Get-SelectedAction {
    if ($rbFullSend.Checked) { return "FullSend" }
    if ($rbFull.Checked)     { return "Full" }
    if ($rbRemove.Checked)   { return "Remove" }
}

function Update-ParsedPreview {
    param($Parsed)

    $actionText = switch ($Parsed.Action) {
        "FullSend" { "Full Access + Send-As" }
        "Full"     { "Full Access only" }
        "Remove"   { "Remove All Access" }
    }

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("Action : $actionText")
    $lines.Add("Mailboxes ($($Parsed.Mailboxes.Count)) : $($Parsed.Mailboxes -join ', ')")
    $lines.Add("Users     ($($Parsed.Users.Count)) : $($Parsed.Users -join ', ')")
    $lines.Add("Total operations : $($Parsed.OpCount)")

    $txtParsed.Lines = $lines.ToArray()
}

$btnParse.Add_Click({
    $btnParse.Enabled = $false
    $btnRun.Enabled   = $false
    try {
        $action = Get-SelectedAction
        $p      = Get-ParsedInputs -MailboxRaw $txtMailbox.Text -UserRaw $txtUser.Text -Action $action

        if ($p.Mailboxes.Count -lt 1) {
            $txtOut.Text = "Please enter at least one mailbox."
            return
        }
        if ($p.Users.Count -lt 1) {
            $txtOut.Text = "Please enter at least one user. (If users were removed due to mailbox overlap, re-check your entries.)"
            return
        }

        $script:Parsed     = $p
        $script:ParsedFrom = $txtMailbox.Text + "|" + $txtUser.Text + "|" + $action

        Update-ParsedPreview -Parsed $p
        $txtOut.Text = "Parsed OK. Review the preview above, then click Run to proceed."
    } catch {
        $txtOut.Text = "Parse error: $($_.Exception.Message)"
    } finally {
        $btnParse.Enabled = $true
        $btnRun.Enabled   = $true
    }
})

$btnRun.Add_Click({
    $btnRun.Enabled   = $false
    $btnParse.Enabled = $false
    try {
        $action   = Get-SelectedAction
        $stateKey = $txtMailbox.Text + "|" + $txtUser.Text + "|" + $action

        if (-not $script:Parsed -or $stateKey -ne $script:ParsedFrom) {
            $script:Parsed     = Get-ParsedInputs -MailboxRaw $txtMailbox.Text -UserRaw $txtUser.Text -Action $action
            $script:ParsedFrom = $stateKey
            Update-ParsedPreview -Parsed $script:Parsed
        }

        $p = $script:Parsed

        if ($p.Mailboxes.Count -lt 1) {
            $txtOut.Text = "Please enter at least one mailbox."
            return
        }
        if ($p.Users.Count -lt 1) {
            $txtOut.Text = "Please enter at least one user. (If users were removed due to mailbox overlap, re-check your entries.)"
            return
        }

        $actionText = switch ($action) {
            "FullSend" { "Full Access + Send-As" }
            "Full"     { "Full Access only" }
            "Remove"   { "Remove All Access" }
        }

        $mbList  = ($p.Mailboxes | ForEach-Object { "  $_" }) -join "`r`n"
        $usrList = ($p.Users     | ForEach-Object { "  $_" }) -join "`r`n"

        $confirmMsg = ("Action: {0}`r`n`r`nMailbox(es) ({1}):`r`n{2}`r`n`r`nUser(s) ({3}):`r`n{4}`r`n`r`nTotal operations: {5}`r`n`r`nProceed?" -f `
            $actionText, $p.Mailboxes.Count, $mbList, $p.Users.Count, $usrList, $p.OpCount)

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            $confirmMsg,
            "Confirm Action",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $txtOut.Text = "Cancelled."
            return
        }

        $txtOut.Text = "Running � please wait..."
        $form.Refresh()

        $result      = Invoke-MailboxAccess -Mailboxes $p.Mailboxes -Users $p.Users -Action $p.Action
        $txtOut.Text = $result

    } catch {
        $txtOut.Text = "ERROR: $($_.Exception.Message)"
    } finally {
        $btnRun.Enabled   = $true
        $btnParse.Enabled = $true
    }
})

$btnCancel.Add_Click({ $form.Close() })

$form.StartPosition = 'Manual'
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$yy = [Math]::Max($wa.Top, $wa.Bottom - $form.Height)
$form.Location = New-Object System.Drawing.Point($wa.Left, $yy)
# Theme if Common.ps1 happens to be loaded; these tools run standalone
# by design, so this must not become a hard dependency.
if (Get-Command Set-ConsoleTheme -ErrorAction SilentlyContinue) {
    Set-ConsoleTheme -Form $form
}
[void]$form.ShowDialog()

#endregion
