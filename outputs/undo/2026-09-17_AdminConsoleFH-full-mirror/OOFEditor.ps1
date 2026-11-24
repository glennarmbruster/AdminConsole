#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    OOFEditor.ps1
    Out of Office -- rich-text editor with templates and preview.

    Author   : Glenn Armbruster
    Revised  : 2026-09-09  (rich editor, templates, preview, date pickers)
#>
# TOOLKIT-HELP-RUNS: reads the auto-reply configuration. Save changes it and asks first.
# TOOLKIT-HELP-BEGIN
# Read and set someone's out-of-office message without needing their mailbox
# open. Both messages -- internal, and the one external senders receive --
# and the state, audience and schedule.
#
# THE EDITOR IS A REAL ONE. Exchange stores auto-reply text as HTML, so the
# message boxes are HTML editors, not plain text: font, size, colour,
# highlight, bold/italic/underline, bullets, alignment, links -- the same
# things a person would do in Outlook. Ctrl+B / Ctrl+I / Ctrl+U work. What
# you see is what the sender gets.
#
# TEMPLATES. Three standard messages (back on a date, extended leave, left
# the company). Pick one, press Insert, fill in the two or three blanks --
# return date comes from the End picker if it is set, the name from the
# mailbox -- and edit from there. A template REPLACES the tab you are on,
# and asks first if there is something in it.
#
# PREVIEW shows each message exactly as it will render, with the HTML
# source alongside for the times something looks wrong and you need to see
# why. Copy -> External copies the internal message over the external one.
#
# Scheduled needs both a start and an end; the pickers are only live in that
# state. External audience: None (external senders get nothing), Known
# (contacts only), All.
#
# Exchange Online only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

$script:Target      = Get-TargetSmtp -Text $SmtpAddress
$script:DisplayName = ''

#region Data ------------------------------------------------------------------
function Get-OutOfOfficeConfig {
    param([string]$Smtp)
    # Learn 2026-09-09 (Get-MailboxAutoReplyConfiguration): AutoReplyState,
    # ExternalAudience, StartTime, EndTime, InternalMessage, ExternalMessage.
    Write-ToolStep 'Reading automatic replies ...'
    Invoke-EXOCommand -Name 'Get-MailboxAutoReplyConfiguration' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
}
function Set-OutOfOfficeConfig {
    param(
        [string]$Smtp, [string]$AutoReplyState, [string]$ExternalAudience,
        [string]$InternalMessage, [string]$ExternalMessage,
        [Nullable[DateTime]]$StartTime, [Nullable[DateTime]]$EndTime
    )
    # Learn 2026-09-09 (Set-MailboxAutoReplyConfiguration): -AutoReplyState
    # Disabled|Enabled|Scheduled; -ExternalAudience None|Known|All;
    # -InternalMessage / -ExternalMessage accept HTML; -StartTime/-EndTime
    # apply to Scheduled.
    $p = @{
        Identity = $Smtp; AutoReplyState = $AutoReplyState; ExternalAudience = $ExternalAudience
        InternalMessage = $InternalMessage; ExternalMessage = $ExternalMessage
        Confirm = $false; ErrorAction = 'Stop'
    }
    if ($AutoReplyState -eq 'Scheduled') {
        if ($StartTime) { $p.StartTime = $StartTime }
        if ($EndTime)   { $p.EndTime   = $EndTime }
    }
    Invoke-EXOCommand -Name 'Set-MailboxAutoReplyConfiguration' -Parameters $p | Out-Null
}

function Get-BodyInnerHtml {
    # Exchange hands back anything from a bare string to a full document. The
    # editor wants what goes INSIDE its body.
    param([string]$Html)
    $h = "$Html"
    if ($h -match '(?is)<body[^>]*>(.*)</body>') { return $Matches[1] }
    return $h
}

# Templates. Placeholders in braces are filled at insert time: DisplayName
# from the mailbox, ReturnDate from the End picker, the rest asked for.
$script:Templates = @(
    @{ Name = 'Out of office -- back on a date'; Html = @'
<p>Thank you for your email. I am out of the office and will return on <b>{ReturnDate}</b>, with limited access to email.</p>
<p>For anything urgent, please contact <a href="mailto:{BackupEmail}">{BackupName}</a>.</p>
<p>Kind regards,<br>{DisplayName}</p>
'@ },
    @{ Name = 'Extended leave -- not checking email'; Html = @'
<p>Thank you for your message. I am on leave until <b>{ReturnDate}</b> and am not checking email during this time.</p>
<p>Your message will not be forwarded. Please contact <a href="mailto:{BackupEmail}">{BackupName}</a> for assistance, or resend your message after my return.</p>
<p>{DisplayName}</p>
'@ },
    @{ Name = 'No longer with the company'; Html = @'
<p>Thank you for your email. {DisplayName} is no longer with {Company}.</p>
<p>Please direct all correspondence to <a href="mailto:{BackupEmail}">{BackupName}</a>, who will be glad to assist you.</p>
<p>{Company}</p>
'@ }
)
#endregion

#region Editor helpers --------------------------------------------------------
# The editor is the IE WebBrowser control with contenteditable on the body.
# It is the only rich HTML editor WinForms has without a third-party
# dependency, and the HTML it produces is exactly the kind Exchange stores.

$script:EditorShell = @'
<html><head><meta charset="utf-8">
<style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt;margin:8px;} p{margin:0 0 8px 0;}</style>
</head><body contenteditable="true"></body></html>
'@

function New-HtmlEditor {
    $wb = New-Object System.Windows.Forms.WebBrowser
    $wb.Dock = 'Fill'
    $wb.ScriptErrorsSuppressed = $true
    $wb.IsWebBrowserContextMenuEnabled = $false
    $wb.WebBrowserShortcutsEnabled = $true     # Ctrl+B / I / U / Z work natively
    $wb.DocumentText = $script:EditorShell
    return $wb
}

function Wait-Editor {
    # DocumentText loads asynchronously; nothing below the toolbar can run
    # until ReadyState says the body exists.
    param([System.Windows.Forms.WebBrowser]$Wb)
    $n = 0
    while ($Wb.ReadyState -ne [System.Windows.Forms.WebBrowserReadyState]::Complete -and $n -lt 200) {
        [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 25; $n++
    }
}

function Get-EditorHtml  { param([System.Windows.Forms.WebBrowser]$Wb) Wait-Editor $Wb; if ($Wb.Document -and $Wb.Document.Body) { return "$($Wb.Document.Body.InnerHtml)" } return '' }
function Set-EditorHtml  { param([System.Windows.Forms.WebBrowser]$Wb, [string]$Html) Wait-Editor $Wb; if ($Wb.Document -and $Wb.Document.Body) { $Wb.Document.Body.InnerHtml = "$Html" } }
function Get-EditorText  { param([System.Windows.Forms.WebBrowser]$Wb) Wait-Editor $Wb; if ($Wb.Document -and $Wb.Document.Body) { return "$($Wb.Document.Body.InnerText)".Trim() } return '' }

function Invoke-EditorCommand {
    # document.execCommand through the managed wrapper. Value may be $null.
    param([System.Windows.Forms.WebBrowser]$Wb, [string]$Command, [object]$Value = $null)
    Wait-Editor $Wb
    if (-not $Wb.Document) { return }
    try { $Wb.Document.ExecCommand($Command, $false, $Value) | Out-Null } catch { }
    try { $Wb.Document.Focus() } catch { }
}

function Get-EditorSelectionText {
    param([System.Windows.Forms.WebBrowser]$Wb)
    try {
        $sel = $Wb.Document.DomDocument.selection
        if ("$($sel.type)" -eq 'None') { return '' }
        return "$($sel.createRange().text)"
    } catch { return '' }
}

function Add-EditorHtml {
    # IE has no InsertHTML command; the selection range's pasteHTML is the
    # equivalent and replaces whatever is selected.
    param([System.Windows.Forms.WebBrowser]$Wb, [string]$Html)
    try {
        $Wb.Document.Focus()
        $Wb.Document.DomDocument.selection.createRange().pasteHTML($Html)
    } catch {
        # No usable selection (nothing focused yet): append instead.
        $Wb.Document.Body.InnerHtml = "$($Wb.Document.Body.InnerHtml)$Html"
    }
}

function ConvertTo-HtmlColor { param([System.Drawing.Color]$C) return ('#{0:x2}{1:x2}{2:x2}' -f $C.R, $C.G, $C.B) }

function Protect-Html { param([string]$s) return [System.Net.WebUtility]::HtmlEncode("$s") }
#endregion

#region Dialogs ---------------------------------------------------------------
function Show-LinkDialog {
    param([string]$SelectedText)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Insert hyperlink'; $f.ClientSize = New-Object System.Drawing.Size(440, 128)
    $f.FormBorderStyle = 'FixedDialog'; $f.StartPosition = 'CenterParent'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    $l1 = New-Object System.Windows.Forms.Label; $l1.Text = 'Address (URL or mailto:):'; $l1.Location = New-Object System.Drawing.Point(12, 14); $l1.AutoSize = $true; $f.Controls.Add($l1)
    $t1 = New-Object System.Windows.Forms.TextBox; $t1.Location = New-Object System.Drawing.Point(160, 10); $t1.Size = New-Object System.Drawing.Size(268, 24); $t1.Text = 'https://'; $f.Controls.Add($t1)
    $l2 = New-Object System.Windows.Forms.Label; $l2.Text = 'Text to show:'; $l2.Location = New-Object System.Drawing.Point(12, 46); $l2.AutoSize = $true; $f.Controls.Add($l2)
    $t2 = New-Object System.Windows.Forms.TextBox; $t2.Location = New-Object System.Drawing.Point(160, 42); $t2.Size = New-Object System.Drawing.Size(268, 24); $t2.Text = $SelectedText; $f.Controls.Add($t2)
    if ($SelectedText) { $t2.ReadOnly = $true }
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = 'Insert'; $ok.Location = New-Object System.Drawing.Point(232, 86); $ok.Size = New-Object System.Drawing.Size(94, 28); $ok.DialogResult = 'OK'; $f.Controls.Add($ok)
    $no = New-Object System.Windows.Forms.Button; $no.Text = 'Cancel'; $no.Location = New-Object System.Drawing.Point(334, 86); $no.Size = New-Object System.Drawing.Size(94, 28); $no.DialogResult = 'Cancel'; $f.Controls.Add($no)
    $f.AcceptButton = $ok; $f.CancelButton = $no
    Set-ConsoleTheme -Form $f
    if ($f.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    $url = "$($t1.Text)".Trim()
    if (-not $url -or $url -eq 'https://') { return $null }
    if ($url -notmatch '^(https?://|mailto:|tel:)') { $url = 'https://' + $url }
    return @{ Url = $url; Text = $(if ("$($t2.Text)".Trim()) { "$($t2.Text)".Trim() } else { $url }) }
}

function Show-FillDialog {
    # One textbox per blank the template still has after the automatic ones.
    param([string[]]$Names, [hashtable]$Defaults)
    if ($Names.Count -eq 0) { return @{} }
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'Fill in the template'; $f.FormBorderStyle = 'FixedDialog'; $f.StartPosition = 'CenterParent'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    $y = 12; $boxes = @{}
    foreach ($n in $Names) {
        $l = New-Object System.Windows.Forms.Label; $l.Text = ($n -creplace '([a-z])([A-Z])', '$1 $2') + ':'; $l.Location = New-Object System.Drawing.Point(12, ($y + 4)); $l.AutoSize = $true; $f.Controls.Add($l)
        $t = New-Object System.Windows.Forms.TextBox; $t.Location = New-Object System.Drawing.Point(130, $y); $t.Size = New-Object System.Drawing.Size(300, 24)
        if ($Defaults -and $Defaults.ContainsKey($n)) { $t.Text = "$($Defaults[$n])" }
        $f.Controls.Add($t); $boxes[$n] = $t; $y += 32
    }
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = 'OK'; $ok.Location = New-Object System.Drawing.Point(234, ($y + 6)); $ok.Size = New-Object System.Drawing.Size(94, 28); $ok.DialogResult = 'OK'; $f.Controls.Add($ok)
    $no = New-Object System.Windows.Forms.Button; $no.Text = 'Cancel'; $no.Location = New-Object System.Drawing.Point(336, ($y + 6)); $no.Size = New-Object System.Drawing.Size(94, 28); $no.DialogResult = 'Cancel'; $f.Controls.Add($no)
    $f.ClientSize = New-Object System.Drawing.Size(442, ($y + 46)); $f.AcceptButton = $ok; $f.CancelButton = $no
    Set-ConsoleTheme -Form $f
    if ($f.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    $out = @{}; foreach ($n in $Names) { $out[$n] = "$($boxes[$n].Text)".Trim() }
    return $out
}

function Show-PreviewDialog {
    param([string]$Title, [string]$InternalHtml, [string]$ExternalHtml)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Preview -- $Title"; $f.ClientSize = New-Object System.Drawing.Size(760, 560); $f.StartPosition = 'CenterParent'
    $f.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(12, 12); $tabs.Size = New-Object System.Drawing.Size(736, 500)
    foreach ($pair in @(@('Internal -- as recipients see it', $InternalHtml, $false), @('External -- as recipients see it', $ExternalHtml, $false), @('Internal -- HTML source', $InternalHtml, $true), @('External -- HTML source', $ExternalHtml, $true))) {
        $pg = New-Object System.Windows.Forms.TabPage; $pg.Text = $pair[0]
        if ($pair[2]) {
            $tb = New-Object System.Windows.Forms.TextBox; $tb.Multiline = $true; $tb.ReadOnly = $true; $tb.ScrollBars = 'Both'; $tb.WordWrap = $false
            $tb.Dock = 'Fill'; $tb.Font = New-Object System.Drawing.Font('Consolas', 9); $tb.Text = ("$($pair[1])" -replace '(?<!\r)\n', "`r`n")
            $pg.Controls.Add($tb)
        } else {
            $wb = New-Object System.Windows.Forms.WebBrowser; $wb.Dock = 'Fill'; $wb.ScriptErrorsSuppressed = $true; $wb.AllowNavigation = $false
            # String.Replace, not -replace: the message may contain '$' and a
            # regex replacement would read that as a group reference.
            $wb.DocumentText = $script:EditorShell.Replace(' contenteditable="true"', '').Replace('</body>', ("$($pair[1])" + '</body>'))
            $pg.Controls.Add($wb)
        }
        [void]$tabs.TabPages.Add($pg)
    }
    $f.Controls.Add($tabs)
    $btn = New-Object System.Windows.Forms.Button; $btn.Text = 'Close'; $btn.Location = New-Object System.Drawing.Point(648, 520); $btn.Size = New-Object System.Drawing.Size(100, 28); $btn.DialogResult = 'OK'; $f.Controls.Add($btn)
    $f.AcceptButton = $btn
    Set-ConsoleTheme -Form $f
    [void]$f.ShowDialog()
}
#endregion

#region Form ------------------------------------------------------------------
function Show-OutOfOfficeForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Out of Office'
    $form.ClientSize = New-Object System.Drawing.Size(900, 600)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    # ---- row 1: state / audience / schedule ----
    $y = 12
    $lblState = New-Object System.Windows.Forms.Label; $lblState.Text = 'State:'; $lblState.Location = New-Object System.Drawing.Point(12, ($y + 4)); $lblState.AutoSize = $true; $form.Controls.Add($lblState)
    $cboState = New-Object System.Windows.Forms.ComboBox; $cboState.DropDownStyle = 'DropDownList'
    @('Disabled', 'Enabled', 'Scheduled') | ForEach-Object { [void]$cboState.Items.Add($_) }
    $cboState.Location = New-Object System.Drawing.Point(56, $y); $cboState.Size = New-Object System.Drawing.Size(110, 24); $form.Controls.Add($cboState)

    $lblAud = New-Object System.Windows.Forms.Label; $lblAud.Text = 'External audience:'; $lblAud.Location = New-Object System.Drawing.Point(184, ($y + 4)); $lblAud.AutoSize = $true; $form.Controls.Add($lblAud)
    $cboAud = New-Object System.Windows.Forms.ComboBox; $cboAud.DropDownStyle = 'DropDownList'
    @('None', 'Known', 'All') | ForEach-Object { [void]$cboAud.Items.Add($_) }
    $cboAud.Location = New-Object System.Drawing.Point(300, $y); $cboAud.Size = New-Object System.Drawing.Size(90, 24); $form.Controls.Add($cboAud)

    $lblStart = New-Object System.Windows.Forms.Label; $lblStart.Text = 'Start:'; $lblStart.Location = New-Object System.Drawing.Point(410, ($y + 4)); $lblStart.AutoSize = $true; $form.Controls.Add($lblStart)
    $dtStart = New-Object System.Windows.Forms.DateTimePicker; $dtStart.Format = 'Custom'; $dtStart.CustomFormat = 'yyyy-MM-dd HH:mm'; $dtStart.ShowCheckBox = $true; $dtStart.Checked = $false
    $dtStart.Location = New-Object System.Drawing.Point(450, $y); $dtStart.Size = New-Object System.Drawing.Size(170, 24); $form.Controls.Add($dtStart)
    $lblEnd = New-Object System.Windows.Forms.Label; $lblEnd.Text = 'End:'; $lblEnd.Location = New-Object System.Drawing.Point(636, ($y + 4)); $lblEnd.AutoSize = $true; $form.Controls.Add($lblEnd)
    $dtEnd = New-Object System.Windows.Forms.DateTimePicker; $dtEnd.Format = 'Custom'; $dtEnd.CustomFormat = 'yyyy-MM-dd HH:mm'; $dtEnd.ShowCheckBox = $true; $dtEnd.Checked = $false
    $dtEnd.Location = New-Object System.Drawing.Point(672, $y); $dtEnd.Size = New-Object System.Drawing.Size(170, 24); $form.Controls.Add($dtEnd)
    $y += 34

    # ---- row 2: templates / preview / copy ----
    $lblTpl = New-Object System.Windows.Forms.Label; $lblTpl.Text = 'Template:'; $lblTpl.Location = New-Object System.Drawing.Point(12, ($y + 4)); $lblTpl.AutoSize = $true; $form.Controls.Add($lblTpl)
    $cboTpl = New-Object System.Windows.Forms.ComboBox; $cboTpl.DropDownStyle = 'DropDownList'
    foreach ($t in $script:Templates) { [void]$cboTpl.Items.Add($t.Name) }
    $cboTpl.SelectedIndex = 0
    $cboTpl.Location = New-Object System.Drawing.Point(76, $y); $cboTpl.Size = New-Object System.Drawing.Size(280, 24); $form.Controls.Add($cboTpl)
    $btnTpl = New-Object System.Windows.Forms.Button; $btnTpl.Text = 'Insert template'; $btnTpl.Location = New-Object System.Drawing.Point(364, ($y - 2)); $btnTpl.Size = New-Object System.Drawing.Size(120, 28); $form.Controls.Add($btnTpl)
    $btnCopy = New-Object System.Windows.Forms.Button; $btnCopy.Text = 'Copy internal -> external'; $btnCopy.Location = New-Object System.Drawing.Point(560, ($y - 2)); $btnCopy.Size = New-Object System.Drawing.Size(170, 28); $form.Controls.Add($btnCopy)
    $btnPrev = New-Object System.Windows.Forms.Button; $btnPrev.Text = 'Preview'; $btnPrev.Location = New-Object System.Drawing.Point(738, ($y - 2)); $btnPrev.Size = New-Object System.Drawing.Size(104, 28); $form.Controls.Add($btnPrev)
    $y += 34

    # ---- toolbar ----
    # A ToolStrip, not Buttons: ToolStrip items do not take keyboard focus, so
    # the editor's selection is still there when the command lands on it.
    $ts = New-Object System.Windows.Forms.ToolStrip
    $ts.Dock = 'None'; $ts.GripStyle = 'Hidden'; $ts.AutoSize = $false
    $ts.Location = New-Object System.Drawing.Point(12, $y); $ts.Size = New-Object System.Drawing.Size(876, 30)
    $ts.RenderMode = 'System'
    # Fixed light colours and opted out of the theme: a dark theme paints the
    # form's light text onto the ToolStrip's white strip and the toolbar reads
    # as blank. Tag 'NoTheme' is the theme's own opt-out.
    $ts.Tag = 'NoTheme'
    $ts.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $ts.ForeColor = [System.Drawing.Color]::Black
    $cboFont = New-Object System.Windows.Forms.ToolStripComboBox; $cboFont.DropDownStyle = 'DropDownList'; $cboFont.AutoSize = $false; $cboFont.Width = 130
    @('Calibri', 'Arial', 'Segoe UI', 'Verdana', 'Tahoma', 'Georgia', 'Times New Roman', 'Consolas') | ForEach-Object { [void]$cboFont.Items.Add($_) }
    $cboFont.SelectedIndex = 0; $cboFont.ToolTipText = 'Font'; $cboFont.ForeColor = [System.Drawing.Color]::Black
    [void]$ts.Items.Add($cboFont)
    $cboSize = New-Object System.Windows.Forms.ToolStripComboBox; $cboSize.DropDownStyle = 'DropDownList'; $cboSize.AutoSize = $false; $cboSize.Width = 56
    # IE's FontSize command takes 1-7; the label is the point size people expect.
    $script:SizeMap = [ordered]@{ '8' = 1; '10' = 2; '12' = 3; '14' = 4; '18' = 5; '24' = 6; '36' = 7 }
    foreach ($k in $script:SizeMap.Keys) { [void]$cboSize.Items.Add($k) }
    $cboSize.SelectedIndex = 2; $cboSize.ToolTipText = 'Size'; $cboSize.ForeColor = [System.Drawing.Color]::Black
    [void]$ts.Items.Add($cboSize)
    [void]$ts.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    foreach ($b in @(
        @('B', 'Bold', 'Bold (Ctrl+B)'), @('I', 'Italic', 'Italic (Ctrl+I)'), @('U', 'Underline', 'Underline (Ctrl+U)'), @('S', 'StrikeThrough', 'Strikethrough'),
        @('|', '', ''), @('Color', 'ForeColor', 'Font colour'), @('Highlight', 'BackColor', 'Highlight'),
        @('|', '', ''), @('Bullets', 'InsertUnorderedList', 'Bulleted list'), @('1.2.3.', 'InsertOrderedList', 'Numbered list'), @('<-', 'Outdent', 'Decrease indent'), @('->', 'Indent', 'Increase indent'),
        @('|', '', ''), @('Left', 'JustifyLeft', 'Align left'), @('Centre', 'JustifyCenter', 'Centre'), @('Right', 'JustifyRight', 'Align right'),
        @('|', '', ''), @('Link', 'CreateLink', 'Insert hyperlink'), @('Unlink', 'Unlink', 'Remove hyperlink'),
        @('|', '', ''), @('Clear', 'RemoveFormat', 'Clear formatting'), @('Undo', 'Undo', 'Undo'), @('Redo', 'Redo', 'Redo'))) {
        if ($b[0] -eq '|') { [void]$ts.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)); continue }
        $it = New-Object System.Windows.Forms.ToolStripButton
        $it.Text = $b[0]; $it.Tag = $b[1]; $it.ToolTipText = $b[2]; $it.DisplayStyle = 'Text'
        $it.ForeColor = [System.Drawing.Color]::Black
        if ($b[0] -eq 'B') { $it.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold) }
        if ($b[0] -eq 'I') { $it.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Italic) }
        if ($b[0] -eq 'U') { $it.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Underline) }
        if ($b[0] -eq 'S') { $it.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Strikeout) }
        [void]$ts.Items.Add($it)
    }
    $form.Controls.Add($ts)
    $y += 34

    # ---- editors ----
    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(12, $y); $tabs.Size = New-Object System.Drawing.Size(876, 440)
    $pgInt = New-Object System.Windows.Forms.TabPage; $pgInt.Text = 'Internal message'
    $pgExt = New-Object System.Windows.Forms.TabPage; $pgExt.Text = 'External message'
    $wbInt = New-HtmlEditor; $pgInt.Controls.Add($wbInt)
    $wbExt = New-HtmlEditor; $pgExt.Controls.Add($wbExt)
    [void]$tabs.TabPages.Add($pgInt); [void]$tabs.TabPages.Add($pgExt)
    $form.Controls.Add($tabs)
    $y += 450

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, $y); $lblStatus.Size = New-Object System.Drawing.Size(660, 18)
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray; $lblStatus.Text = 'Enter a mailbox and press Look Up.'
    $form.Controls.Add($lblStatus)

    $btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text = 'Save'; $btnSave.Location = New-Object System.Drawing.Point(680, ($y - 4)); $btnSave.Size = New-Object System.Drawing.Size(100, 28); $btnSave.Enabled = $false; $form.Controls.Add($btnSave)
    $btnClose = New-Object System.Windows.Forms.Button; $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(788, ($y - 4)); $btnClose.Size = New-Object System.Drawing.Size(100, 28); $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    $form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force

    # ---- helpers that need the controls ----
    $activeEditor = { if ($tabs.SelectedIndex -eq 1) { $wbExt } else { $wbInt } }

    $syncSchedule = {
        $sched = ("$($cboState.SelectedItem)" -eq 'Scheduled')
        $dtStart.Enabled = $sched; $dtEnd.Enabled = $sched
    }
    $cboState.Add_SelectedIndexChanged($syncSchedule)

    $load = {
        if (-not "$script:Target".Trim()) { $form.Text = 'Out of Office'; $btnSave.Enabled = $false; return }
        $form.Text = "Out of Office: $script:Target"
        $cfg = Get-OutOfOfficeConfig -Smtp $script:Target
        $cboState.SelectedItem = "$($cfg.AutoReplyState)"
        $cboAud.SelectedItem   = "$($cfg.ExternalAudience)"
        if ($cfg.StartTime) { $dtStart.Value = [datetime]$cfg.StartTime; $dtStart.Checked = $true } else { $dtStart.Checked = $false }
        if ($cfg.EndTime)   { $dtEnd.Value   = [datetime]$cfg.EndTime;   $dtEnd.Checked   = $true } else { $dtEnd.Checked   = $false }
        Set-EditorHtml $wbInt (Get-BodyInnerHtml $cfg.InternalMessage)
        Set-EditorHtml $wbExt (Get-BodyInnerHtml $cfg.ExternalMessage)
        try { $script:DisplayName = "$((Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $script:Target; ErrorAction = 'Stop' }).DisplayName)" } catch { $script:DisplayName = '' }
        & $syncSchedule
        $btnSave.Enabled = $true
        $lblStatus.Text = "Loaded. State: $($cfg.AutoReplyState); external audience: $($cfg.ExternalAudience)."
    }

    $doLoad = {
        if ($form.Busy) { return }
        $form.Busy = $true
        try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; & $load }
        catch { $lblStatus.Text = 'Load failed -- see the error dialog.'; Show-HandlerError -ErrorRecord $_ -Context 'Get-MailboxAutoReplyConfiguration' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    }

    # ---- toolbar wiring ----
    $ts.Add_ItemClicked({
        param($sender, $e)
        $cmd = "$($e.ClickedItem.Tag)"
        if (-not $cmd) { return }
        $wb = & $activeEditor
        switch ($cmd) {
            'ForeColor' {
                $cd = New-Object System.Windows.Forms.ColorDialog; $cd.FullOpen = $true
                if ($cd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Invoke-EditorCommand $wb 'ForeColor' (ConvertTo-HtmlColor $cd.Color) }
            }
            'BackColor' {
                $cd = New-Object System.Windows.Forms.ColorDialog; $cd.FullOpen = $true; $cd.Color = [System.Drawing.Color]::Yellow
                if ($cd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Invoke-EditorCommand $wb 'BackColor' (ConvertTo-HtmlColor $cd.Color) }
            }
            'CreateLink' {
                $selText = Get-EditorSelectionText $wb
                $r = Show-LinkDialog -SelectedText $selText
                if (-not $r) { return }
                if ($selText) { Invoke-EditorCommand $wb 'CreateLink' $r.Url }
                else { Add-EditorHtml $wb ('<a href="' + (Protect-Html $r.Url) + '">' + (Protect-Html $r.Text) + '</a>') }
            }
            default { Invoke-EditorCommand $wb $cmd }
        }
    })
    $cboFont.Add_SelectedIndexChanged({ Invoke-EditorCommand (& $activeEditor) 'FontName' "$($cboFont.SelectedItem)" })
    $cboSize.Add_SelectedIndexChanged({ Invoke-EditorCommand (& $activeEditor) 'FontSize' ([int]$script:SizeMap["$($cboSize.SelectedItem)"]) })

    # ---- templates ----
    $btnTpl.Add_Click({
        if ($form.Busy) { return }
        $wb = & $activeEditor
        $t = $script:Templates[$cboTpl.SelectedIndex]
        if ((Get-EditorText $wb)) {
            $a = Show-ToolDialog("Replace the $($tabs.SelectedTab.Text.ToLower()) with the template '$($t.Name)'?", 'Insert template', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        }
        $html = "$($t.Html)"
        $auto = @{}
        $auto['DisplayName'] = $(if ($script:DisplayName) { $script:DisplayName } else { '' })
        $auto['ReturnDate']  = $(if ($dtEnd.Checked) { $dtEnd.Value.ToString('dddd d MMMM yyyy') } else { '' })
        $cfgName = ''
        try { $cfgName = "$((Get-TenantConfig).TenantName)" } catch { }
        $auto['Company'] = $cfgName
        # Whatever is still blank gets asked for.
        $ask = @()
        foreach ($m in ([regex]::Matches($html, '\{(\w+)\}'))) { $n = $m.Groups[1].Value; if (-not "$($auto[$n])" -and $ask -notcontains $n) { $ask += $n } }
        $answers = Show-FillDialog -Names $ask -Defaults @{}
        if ($null -eq $answers) { return }
        foreach ($k in $answers.Keys) { $auto[$k] = $answers[$k] }
        foreach ($k in @($auto.Keys)) { $html = $html.Replace('{' + $k + '}', (Protect-Html $auto[$k])) }
        Set-EditorHtml $wb $html
        $lblStatus.Text = "Template inserted into the $($tabs.SelectedTab.Text.ToLower()). Edit as needed, then Save."
    })

    $btnCopy.Add_Click({
        if ((Get-EditorText $wbExt)) {
            $a = Show-ToolDialog('Replace the external message with a copy of the internal one?', 'Copy', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        }
        Set-EditorHtml $wbExt (Get-EditorHtml $wbInt)
        $tabs.SelectedIndex = 1
        $lblStatus.Text = 'External message replaced with the internal one.'
    })

    $btnPrev.Add_Click({ Show-PreviewDialog -Title $script:Target -InternalHtml (Get-EditorHtml $wbInt) -ExternalHtml (Get-EditorHtml $wbExt) })

    # ---- save ----
    $btnSave.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $state = "$($cboState.SelectedItem)"
            $st = $null; $et = $null
            if ($state -eq 'Scheduled') {
                if (-not $dtStart.Checked -or -not $dtEnd.Checked) { $lblStatus.Text = 'Scheduled needs both a start and an end -- tick both pickers.'; return }
                $st = $dtStart.Value; $et = $dtEnd.Value
                if ($et -le $st) { $lblStatus.Text = 'End must be after start.'; return }
            }
            $intHtml = Get-EditorHtml $wbInt
            $extHtml = Get-EditorHtml $wbExt
            if ($state -ne 'Disabled' -and -not (Get-EditorText $wbInt)) {
                $a = Show-ToolDialog('The internal message is empty. Save anyway?', 'Empty message', 'OKCancel', 'Warning')
                if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
            }
            $a = Show-ToolDialog(
                ("Save Out of Office for:`r`n  $script:Target`r`n`r`nState: $state" + $(if ($state -eq 'Scheduled') { "`r`nFrom:  $($st.ToString('yyyy-MM-dd HH:mm'))`r`nTo:    $($et.ToString('yyyy-MM-dd HH:mm'))" }) + "`r`nExternal audience: $($cboAud.SelectedItem)"),
                'Confirm', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Set-OutOfOfficeConfig -Smtp $script:Target -AutoReplyState $state -ExternalAudience "$($cboAud.SelectedItem)" -InternalMessage $intHtml -ExternalMessage $extHtml -StartTime $st -EndTime $et
            Write-CompletedLog -Action 'SetOutOfOffice' -Target (Get-LogTarget $script:Target) -Result 'Success' `
                -Detail "State $state; audience $($cboAud.SelectedItem)$(if ($state -eq 'Scheduled') { "; $($st.ToString('yyyy-MM-dd HH:mm')) to $($et.ToString('yyyy-MM-dd HH:mm'))" }); internal $($intHtml.Length) chars, external $($extHtml.Length) chars"
            & $load
            $lblStatus.Text = 'Saved.'
        } catch {
            $lblStatus.Text = 'Save failed -- see the error dialog.'
            Show-HandlerError -ErrorRecord $_ -Context 'Set-MailboxAutoReplyConfiguration'
            Write-FailedLog -Action 'SetOutOfOffice' -Target (Get-LogTarget $script:Target) -ErrorRecord $_
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    # The editors need the message loop before their documents exist, so the
    # first load waits for Shown rather than running at build time.
    # The ActiveX behind a WebBrowser is only created once the control has
    # been visible, and the External tab starts hidden -- so show it for one
    # message-loop tick before waiting on either document.
    $form.Add_Shown({
        $tabs.SelectedIndex = 1; [System.Windows.Forms.Application]::DoEvents()
        $tabs.SelectedIndex = 0; [System.Windows.Forms.Application]::DoEvents()
        Wait-Editor $wbInt; Wait-Editor $wbExt
        & $doLoad
    })

    $bar = Add-ToolTargetBar -Form $form -Label 'Mailbox or address:' -InitialValue $script:Target
    $bar.Button.Add_Click({ $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text; & $doLoad })

    # Enter belongs to the HTML editor (new paragraph), not to Look Up.
    $form.Tag = 'NoAcceptButton'
    Set-ConsoleTheme -Form $form
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Out of Office' | Out-Null
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    [void]$form.ShowDialog()
}
#endregion

Show-OutOfOfficeForm
