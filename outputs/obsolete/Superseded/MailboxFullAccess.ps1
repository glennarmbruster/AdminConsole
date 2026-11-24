#Requires -Version 5.1
<#
    MailboxFullAccess.ps1
    Full Access delegation only -- Exchange Online, no SharePoint.

    Author   : Glenn Armbruster
    Revised  : 2026-09-21

    WHY THIS EXISTS, given MailboxOneDriveAccess.ps1 already grants Full Access.

    Set-MailboxAccess.ps1 and onedriveaccess.ps1 were merged into
    MailboxOneDriveAccess.ps1 because they shared a shape. The merge was right
    about the shape and wrong about the cost: the combined tool inherited the
    OneDrive half's SharePoint Online dependency, so the commonest ticket in
    this area -- "give this person access to that mailbox" -- started requiring
    a connection it has no use for, and is unavailable in a tenant that does
    not declare SharePoint at all.

    This tool is the mailbox half, standing on its own. It grants and removes
    ONE thing: Full Access. No Send As, no Send on Behalf, no OneDrive.
    Anything wider belongs in MailboxPermissions.ps1 or the Mailbox Editor.
#>
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-RUNS: READS and WRITES Full Access mailbox permissions in Exchange Online. Nothing is written until Run, and Run asks first.
# TOOLKIT-HELP-BEGIN
# Give somebody Full Access to somebody else's mailbox -- and take it away.
#
# FULL ACCESS ONLY. It lets the grantee open and read the mailbox. It does NOT
# let them send as the mailbox or on its behalf; those are separate rights and
# live in Mailbox Permissions. "I can see it but I can't send from it" is this
# permission granted on its own, and is usually the correct outcome.
#
# PASTE THE TICKET. The box at the top reads labelled requests, so a ticket
# that says
#
#     Please make as Custodian : someone@contoso.com
#     with access to User : someoneelse@contoso.com
#
# fills both boxes below on its own. Labels understood, in any order and any
# capitalisation:
#
#   grantee  -- Custodian, Manager, Delegate, Grant to, Grant access to
#   mailbox  -- User, Mailbox, Target, Account
#
# If no labels are found, nothing is guessed: type the addresses into the two
# boxes yourself. Which address is which is the one thing this tool must not
# get wrong, and an unlabelled pair does not say.
#
# EVERY MAILBOX x EVERY GRANTEE. Two mailboxes and three grantees is six
# grants, and the preview says so before anything runs.
#
# AUTOMAPPING is on by default: Outlook adds the mailbox to the grantee's
# profile by itself. Turn it off for a large mailbox, or where the grantee
# should open it deliberately rather than have it appear.
#
# Changes are written to the audit log. Access granted here does not expire --
# remove it when it is no longer needed.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

#region Parsing ------------------------------------------------------------------

# THE LABELS, AND WHY THIS IS NOT POSITIONAL.
#
# The first version of this idea read "first address is the mailbox, second is
# the grantee". That is right exactly as often as the person writing the ticket
# happens to put them in that order, and wrong silently -- the grant succeeds,
# it is just backwards, and the first anyone knows is the wrong person reading
# somebody's mail. Labels are the only part of a free-text request that carries
# the operator's intent, so labels are what this reads.
$script:GranteeLabels = @('custodian', 'manager', 'delegate', 'grant to', 'grant access to', 'grantee')
$script:MailboxLabels = @('user', 'mailbox', 'target', 'account')

function Get-AddressesInText {
    # Every address in a blob, in the order they appear. Deliberately loose on
    # the local part -- ticket text arrives with quotes, angle brackets and
    # trailing punctuation around addresses that are otherwise fine.
    param([string]$Text)
    if (-not "$Text".Trim()) { return @() }
    $m = [regex]::Matches("$Text", '[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}')
    return @($m | ForEach-Object { $_.Value })
}

function Get-LabelledAddresses {
    <#
        Pull labelled addresses out of pasted ticket text.

        ONE LABEL, ONE LINE, and the label has to come BEFORE the address on
        that line. "Custodian : a@x  with access to User : b@x" is two labels
        on one line and is handled, because each label is matched with the
        first address that follows IT rather than the first address on the line.

        Returns @{ Mailboxes = @(); Grantees = @() }.
    #>
    param([string]$Text)

    $mailboxes = New-Object System.Collections.Generic.List[string]
    $grantees  = New-Object System.Collections.Generic.List[string]
    if (-not "$Text".Trim()) { return @{ Mailboxes = @(); Grantees = @() } }

    $all = @()
    foreach ($lbl in $script:GranteeLabels) { $all += [PSCustomObject]@{ Label = $lbl; Kind = 'Grantee' } }
    foreach ($lbl in $script:MailboxLabels) { $all += [PSCustomObject]@{ Label = $lbl; Kind = 'Mailbox' } }

    foreach ($line in ("$Text" -split "`r?`n")) {
        if (-not "$line".Trim()) { continue }
        $hits = New-Object System.Collections.Generic.List[object]
        foreach ($e in $all) {
            # Word boundary in front so 'account' does not fire inside another
            # word, and a colon/dash/equals after it so prose that merely uses
            # the word is not mistaken for a field.
            $rx = '(?i)\b' + [regex]::Escape($e.Label) + '\s*[:=\-]'
            foreach ($m in [regex]::Matches("$line", $rx)) {
                $hits.Add([PSCustomObject]@{ At = $m.Index; Kind = $e.Kind; Len = $e.Label.Length })
            }
        }
        if ($hits.Count -eq 0) { continue }

        # LONGEST LABEL WINS AT THE SAME SPOT. 'Grant access to' and 'Grant to'
        # can both match at one index; taking the longer one stops the shorter
        # alias claiming a field it only partly named.
        $ordered = @($hits | Sort-Object At, @{ Expression = 'Len'; Descending = $true })
        $taken = @()
        foreach ($h in $ordered) {
            if ($taken -contains $h.At) { continue }
            $taken += $h.At
            $rest = "$line".Substring($h.At)
            $addr = @(Get-AddressesInText -Text $rest)
            if ($addr.Count -eq 0) { continue }
            if ($h.Kind -eq 'Grantee') { $grantees.Add($addr[0]) } else { $mailboxes.Add($addr[0]) }
        }
    }

    return @{
        Mailboxes = @($mailboxes | Select-Object -Unique)
        Grantees  = @($grantees  | Select-Object -Unique)
    }
}

function Get-AddressList {
    # One box's worth of addresses. A bare name gets the tenant's mail domain
    # appended -- Get-MailDomain reads it from Tenant.json, so this tool works
    # in every tenant without a hardcoded suffix.
    param([string]$Raw)
    if (-not "$Raw".Trim()) { return @() }
    $out = foreach ($tok in ("$Raw" -split "[\r\n,;]+")) {
        $t = "$tok".Trim().Trim('<', '>', '"', "'", ',', '.')
        if (-not $t) { continue }
        if ($t -notmatch '@') { $t = "$t@$(Get-MailDomain)" }
        $t
    }
    return @($out | Sort-Object -Unique)
}

function Get-PlannedOperations {
    <#
        The cross-product, minus the pairs that make no sense.

        A MAILBOX IS NEVER ITS OWN GRANTEE. Exchange rejects it, but rejecting
        it here means the preview count is the count of things that will
        actually be attempted rather than a number with a failure hidden in it.
    #>
    param([string[]]$Mailboxes, [string[]]$Grantees)
    $ops = New-Object System.Collections.Generic.List[object]
    foreach ($mbx in $Mailboxes) {
        foreach ($usr in $Grantees) {
            if ("$mbx".ToLower() -eq "$usr".ToLower()) { continue }
            $ops.Add([PSCustomObject]@{ Mailbox = $mbx; Grantee = $usr })
        }
    }
    return @($ops)
}
#endregion

#region Exchange -----------------------------------------------------------------
function Test-FullAccessPresent {
    # Does the grantee ALREADY hold Full Access? Read before the write so the
    # results line can say 'already had it' instead of reporting a success that
    # changed nothing -- the two look identical in the audit log otherwise.
    param([string]$Mailbox, [string]$Grantee)
    try {
        $p = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{
                    Identity = $Mailbox; User = $Grantee; ErrorAction = 'Stop' })
        foreach ($e in $p) {
            if ("$($e.AccessRights)" -match 'FullAccess') { return $true }
        }
        return $false
    } catch {
        # An unreadable permission list is not proof of absence. Say unknown and
        # let the write decide.
        return $null
    }
}

function Grant-FullAccessTo {
    param([string]$Mailbox, [string]$Grantee, [bool]$AutoMapping)
    # Confirm = $false ON EVERY WRITE. A ShouldProcess prompt has no console to
    # render on in a WinForms tool, so the call would wait for an answer that
    # can never arrive and the window would hang.
    Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
        Identity     = $Mailbox
        User         = $Grantee
        AccessRights = 'FullAccess'
        AutoMapping  = $AutoMapping
        Confirm      = $false
        ErrorAction  = 'Stop'
    } | Out-Null
}

function Revoke-FullAccessFrom {
    param([string]$Mailbox, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{
        Identity     = $Mailbox
        User         = $Grantee
        AccessRights = 'FullAccess'
        Confirm      = $false
        ErrorAction  = 'Stop'
    } | Out-Null
}
#endregion

#region Form ---------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Full Access Manager'
$form.ClientSize = New-Object System.Drawing.Size(812, 648)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = 'Full Access only -- this does NOT grant Send As or Send on Behalf. Requires an Exchange Online connection.'
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(788, 18)
$lblWarn.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblWarn)

$lblPaste = New-Object System.Windows.Forms.Label
$lblPaste.Text = 'Paste the ticket here and press Parse (or fill the two boxes below by hand):'
$lblPaste.Location = New-Object System.Drawing.Point(12, 34)
$lblPaste.AutoSize = $true
$form.Controls.Add($lblPaste)

$txtPaste = New-Object System.Windows.Forms.TextBox
$txtPaste.Multiline = $true
$txtPaste.ScrollBars = 'Vertical'
$txtPaste.Location = New-Object System.Drawing.Point(12, 54)
$txtPaste.Size = New-Object System.Drawing.Size(788, 96)
$txtPaste.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtPaste)

$btnParse = New-Object System.Windows.Forms.Button
$btnParse.Text = 'Parse'
$btnParse.Location = New-Object System.Drawing.Point(12, 156)
$btnParse.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnParse)

$lblParseNote = New-Object System.Windows.Forms.Label
$lblParseNote.Text = ''
$lblParseNote.Location = New-Object System.Drawing.Point(110, 162)
$lblParseNote.Size = New-Object System.Drawing.Size(690, 18)
$form.Controls.Add($lblParseNote)

$lblMbx = New-Object System.Windows.Forms.Label
$lblMbx.Text = 'Mailbox(es) -- the account being opened:'
$lblMbx.Location = New-Object System.Drawing.Point(12, 192)
$lblMbx.AutoSize = $true
$form.Controls.Add($lblMbx)

$txtMbx = New-Object System.Windows.Forms.TextBox
$txtMbx.Multiline = $true
$txtMbx.ScrollBars = 'Vertical'
$txtMbx.Location = New-Object System.Drawing.Point(12, 212)
$txtMbx.Size = New-Object System.Drawing.Size(388, 84)
$txtMbx.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtMbx)

$lblUsr = New-Object System.Windows.Forms.Label
$lblUsr.Text = 'Grant access to -- one per line:'
$lblUsr.Location = New-Object System.Drawing.Point(412, 192)
$lblUsr.AutoSize = $true
$form.Controls.Add($lblUsr)

$txtUsr = New-Object System.Windows.Forms.TextBox
$txtUsr.Multiline = $true
$txtUsr.ScrollBars = 'Vertical'
$txtUsr.Location = New-Object System.Drawing.Point(412, 212)
$txtUsr.Size = New-Object System.Drawing.Size(388, 84)
$txtUsr.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtUsr)

$grpMode = New-Object System.Windows.Forms.GroupBox
$grpMode.Text = 'Action'
$grpMode.Location = New-Object System.Drawing.Point(12, 304)
$grpMode.Size = New-Object System.Drawing.Size(240, 52)
$form.Controls.Add($grpMode)

$rbGrant = New-Object System.Windows.Forms.RadioButton
$rbGrant.Text = 'Grant'
$rbGrant.Location = New-Object System.Drawing.Point(12, 20)
$rbGrant.Size = New-Object System.Drawing.Size(70, 22)
$rbGrant.Checked = $true
$grpMode.Controls.Add($rbGrant)

$rbRemove = New-Object System.Windows.Forms.RadioButton
$rbRemove.Text = 'Remove'
$rbRemove.Location = New-Object System.Drawing.Point(92, 20)
$rbRemove.Size = New-Object System.Drawing.Size(90, 22)
$grpMode.Controls.Add($rbRemove)

$chkAuto = New-Object System.Windows.Forms.CheckBox
$chkAuto.Text = 'AutoMapping (Outlook adds the mailbox by itself)'
$chkAuto.Location = New-Object System.Drawing.Point(264, 322)
$chkAuto.Size = New-Object System.Drawing.Size(330, 22)
$chkAuto.Checked = $true
$form.Controls.Add($chkAuto)

$lblPrev = New-Object System.Windows.Forms.Label
$lblPrev.Text = 'Preview -- read this before pressing Run:'
$lblPrev.Location = New-Object System.Drawing.Point(12, 364)
$lblPrev.AutoSize = $true
$form.Controls.Add($lblPrev)

$txtPrev = New-Object System.Windows.Forms.TextBox
$txtPrev.Multiline = $true
$txtPrev.ScrollBars = 'Vertical'
$txtPrev.ReadOnly = $true
$txtPrev.Location = New-Object System.Drawing.Point(12, 384)
$txtPrev.Size = New-Object System.Drawing.Size(788, 76)
$txtPrev.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtPrev)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = 'Run'
$btnRun.Location = New-Object System.Drawing.Point(12, 468)
$btnRun.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnRun)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(110, 468)
$btnClear.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClear)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(710, 468)
$btnClose.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)

$lblRes = New-Object System.Windows.Forms.Label
$lblRes.Text = 'Result:'
$lblRes.Location = New-Object System.Drawing.Point(12, 504)
$lblRes.AutoSize = $true
$form.Controls.Add($lblRes)

$txtRes = New-Object System.Windows.Forms.TextBox
$txtRes.Multiline = $true
$txtRes.ScrollBars = 'Vertical'
$txtRes.ReadOnly = $true
$txtRes.Location = New-Object System.Drawing.Point(12, 524)
$txtRes.Size = New-Object System.Drawing.Size(788, 104)
$txtRes.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtRes)
#endregion

#region Wiring -------------------------------------------------------------------
function Update-Preview {
    $mbx = @(Get-AddressList -Raw $txtMbx.Text)
    $usr = @(Get-AddressList -Raw $txtUsr.Text)
    $ops = @(Get-PlannedOperations -Mailboxes $mbx -Grantees $usr)
    $verb = $(if ($rbRemove.Checked) { 'REMOVE Full Access' } else { 'GRANT Full Access' })

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Action  : $verb")
    if (-not $rbRemove.Checked) {
        $lines.Add("AutoMap : $(if ($chkAuto.Checked) { 'on' } else { 'off' })")
    }
    $lines.Add("Mailboxes ($($mbx.Count)) : $($mbx -join ', ')")
    $lines.Add("Grantees  ($($usr.Count)) : $($usr -join ', ')")
    $lines.Add("Operations : $($ops.Count)")
    if ($mbx.Count -and $usr.Count -and $ops.Count -eq 0) {
        $lines.Add('Nothing to do -- the only pair was a mailbox granting access to itself.')
    }
    $txtPrev.Lines = $lines.ToArray()
    $btnRun.Enabled = ($ops.Count -gt 0)
}

$btnParse.Add_Click({
    $found = Get-LabelledAddresses -Text $txtPaste.Text
    $m = @($found.Mailboxes); $g = @($found.Grantees)

    if ($m.Count -eq 0 -and $g.Count -eq 0) {
        # NOTHING IS GUESSED. See the note on the label list: an unlabelled pair
        # does not say which address is the mailbox, and getting that backwards
        # is the one failure this tool must not produce quietly.
        $all = @(Get-AddressesInText -Text $txtPaste.Text)
        $lblParseNote.Text = $(if ($all.Count) {
                "No labels found. $($all.Count) address(es) are in the text -- type them into the two boxes so it is clear which is which."
            } else {
                'No addresses found in the pasted text.'
            })
        $lblParseNote.ForeColor = (Get-ToolStatusColour -Kind Warn)
        Update-Preview
        return
    }

    if ($m.Count) { $txtMbx.Text = ($m -join "`r`n") }
    if ($g.Count) { $txtUsr.Text = ($g -join "`r`n") }

    $missing = @()
    if ($m.Count -eq 0) { $missing += 'no mailbox label' }
    if ($g.Count -eq 0) { $missing += 'no grantee label' }
    if ($missing.Count) {
        $lblParseNote.Text = "Partly parsed ($($missing -join ', ')) -- fill the empty box by hand."
        $lblParseNote.ForeColor = (Get-ToolStatusColour -Kind Warn)
    } else {
        $lblParseNote.Text = "Parsed: $($m.Count) mailbox(es), $($g.Count) grantee(s). Check the preview."
        $lblParseNote.ForeColor = (Get-ToolStatusColour -Kind Good)
    }
    Update-Preview
})

$txtMbx.Add_TextChanged({ Update-Preview })
$txtUsr.Add_TextChanged({ Update-Preview })
$rbGrant.Add_CheckedChanged({ Update-Preview })
$rbRemove.Add_CheckedChanged({ Update-Preview })
$chkAuto.Add_CheckedChanged({ Update-Preview })

$btnClear.Add_Click({
    $txtPaste.Clear(); $txtMbx.Clear(); $txtUsr.Clear(); $txtRes.Clear()
    $lblParseNote.Text = ''
    Update-Preview
})
$btnClose.Add_Click({ $form.Close() })

$btnRun.Add_Click({
    $mbx = @(Get-AddressList -Raw $txtMbx.Text)
    $usr = @(Get-AddressList -Raw $txtUsr.Text)
    $ops = @(Get-PlannedOperations -Mailboxes $mbx -Grantees $usr)
    if ($ops.Count -eq 0) { return }

    $removing = $rbRemove.Checked
    $verb = $(if ($removing) { 'REMOVE Full Access' } else { 'GRANT Full Access' })
    $answer = Show-ToolDialog("$verb -- $($ops.Count) operation(s).`r`n`r`nMailboxes: $($mbx -join ', ')`r`nGrantees: $($usr -join ', ')`r`n`r`nProceed?",
                              'Confirm', 'OKCancel', 'Question')
    if ($answer -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $btnRun.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $out = New-Object System.Collections.Generic.List[string]
    $ok = 0; $failed = 0; $skipped = 0
    try {
        $i = 0
        foreach ($op in $ops) {
            $i++
            Set-ToolStatus -Text "$verb $i of $($ops.Count): $($op.Grantee) on $($op.Mailbox)..."
            try {
                $has = Test-FullAccessPresent -Mailbox $op.Mailbox -Grantee $op.Grantee
                if (-not $removing -and $has -eq $true) {
                    $out.Add("SKIP    $($op.Grantee) already has Full Access on $($op.Mailbox)")
                    $skipped++
                    continue
                }
                if ($removing -and $has -eq $false) {
                    $out.Add("SKIP    $($op.Grantee) has no Full Access on $($op.Mailbox)")
                    $skipped++
                    continue
                }
                if ($removing) {
                    Revoke-FullAccessFrom -Mailbox $op.Mailbox -Grantee $op.Grantee
                    $out.Add("REMOVED $($op.Grantee) from $($op.Mailbox)")
                } else {
                    Grant-FullAccessTo -Mailbox $op.Mailbox -Grantee $op.Grantee -AutoMapping $chkAuto.Checked
                    $out.Add("GRANTED $($op.Grantee) on $($op.Mailbox)")
                }
                $ok++
            } catch {
                $out.Add("FAILED  $($op.Grantee) on $($op.Mailbox): $($_.Exception.Message)")
                $failed++
                Write-FailedLog -Action $(if ($removing) { 'RemoveFullAccess' } else { 'GrantFullAccess' }) `
                                -Target "$($op.Mailbox)" -Detail "grantee: $($op.Grantee)" -ErrorRecord $_
            }
        }
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnRun.Enabled = $true
    }

    $out.Insert(0, "$verb -- $ok done, $skipped skipped, $failed failed.")
    $txtRes.Lines = $out.ToArray()
    Set-ToolStatus -Text "$ok done, $skipped skipped, $failed failed."

    # ONE SUMMARY ROW, NOT ONE PER PAIR. The per-pair detail is on screen and
    # the failures are already in the failed log with their own rows; a hundred
    # success rows for one ticket buries the log it is supposed to serve.
    $result = $(if ($failed -eq 0) { 'Success' } elseif ($ok -gt 0) { 'Partial' } else { 'Failed' })
    Write-CompletedLog -Action $(if ($removing) { 'RemoveFullAccess' } else { 'GrantFullAccess' }) `
                       -Target ($mbx -join '; ') -Result $result `
                       -Task "$verb for $($usr.Count) grantee(s) on $($mbx.Count) mailbox(es): $ok done, $skipped skipped, $failed failed"
})
#endregion

# Seed from the console's selected object, if it launched us with one.
if ($SmtpAddress) { $txtMbx.Text = (Get-TargetSmtp -Text $SmtpAddress) }
Update-Preview

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Added after the theme so it is not restyled. Also wires the status bar.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Full Access Manager' | Out-Null

[void]$form.ShowDialog()
