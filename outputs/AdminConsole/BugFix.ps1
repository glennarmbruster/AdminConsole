#Requires -Version 5.1
<#
    BugFix.ps1
    Lets an operator report something that is broken without leaving the console.

    The twin of FeatureRequest.ps1 -- same relay, same recipient, same one-send
    BCC arrangement. Different questions, because "it does not work" and "it
    should also do X" need different things from the person reporting them.

    Reasoning: ScriptNotes/BugFix.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-20
#>
# TOOLKIT-HELP-BEGIN
# Report something that is broken.
#
# A copy is sent to you, so you have a record of exactly what you reported.
#
# THE FIELD THAT MATTERS MOST IS 'Steps to reproduce'. A bug that can be
# reproduced is usually fixed the same day; one that cannot may never be. If you
# only fill in one optional box, make it that one -- even "it happens on shared
# mailboxes but not user ones" is enough to halve the search.
#
# WHAT MAKES A REPORT EASY TO ACT ON:
#   - The exact error text, pasted rather than described. "It said something
#     about a domain controller" and the actual 0x52 message are different
#     amounts of information.
#   - What you EXPECTED. Half of all bug reports turn out to be the tool doing
#     what it was built to do, and knowing what you expected instead is what
#     turns that into a real answer rather than "working as designed".
#   - Whether it happens every time. An intermittent fault and a constant one
#     are chased completely differently.
#
# IF IT IS BLOCKING YOU RIGHT NOW, tick the box. That is what gets looked at
# first, and it is the only thing on this form that jumps the queue.
#
# You do not need to know what caused it. Report it anyway.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: Sends an email. Makes no changes to any mailbox, account, or group.

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Configuration --------------------------------------------------------
$script:SmtpServer        = "smtp.$(Get-MailDomain)"
$script:DoNotReplyAddress = "DO_NOT_REPLY@$(Get-MailDomain)"

# The reporter, and the person who gets the copy. Same address, deliberately.
$script:OperatorEmail = Get-OperatorEmail

# THE SAME DESTINATION AS A FEATURE REQUEST, and read from the same place.
#
# A separate BugFix.Recipient key would be a second thing to configure per
# tenant and a second thing to get wrong -- and the two land in the same inbox
# anyway. If they ever need to diverge, that is the moment to add the key, not
# before.
#
# NOT A CONSTANT IN THIS FILE. A mail domain in a .ps1 is the one thing
# Test-Toolkit refuses, and this toolkit ships to six tenants.
$script:Recipient = Get-FeatureRequestRecipient
#endregion

#region Actions --------------------------------------------------------------
function Get-ToolkitStateLine {
    <#
        Version, and whether this copy still matches its manifest.

        WHY A BUG REPORT CARRIES IT. "Three files modified" explains an entire
        class of report on its own -- a hand-patched script, a half-finished
        upgrade, a file someone opened in Notepad. Asking the reporter to go and
        find that out is asking them to do the first step of the diagnosis.

        WRAPPED, AND FAILURE IS NOT FATAL. This is decoration on a bug report.
        An unstamped copy, a locked file, a missing Version.json -- none of them
        is a reason the report should not be sent, so every one of them comes
        back as a sentence rather than an exception.
    #>
    try {
        $r = Test-ToolkitIntegrity
        if (-not $r.HasManifest) { return 'not stamped (no Version.json)' }
        return $r.Summary
    } catch {
        return "could not be read: $($_.Exception.Message)"
    }
}

function New-BugReportBody {
    <#
        The mail body.

        Every field is HTML-encoded, and on this form that matters more than on
        the feature request: the single most useful box asks for an ERROR
        MESSAGE, and error messages are full of < > and &. Pasting one raw into
        an HTML body silently eats the rest of the paragraph -- and the failure
        is invisible to the sender, because their copy is broken identically.
    #>
    param(
        [string]$Tool,
        [string]$What,
        [string]$Expected,
        [string]$Steps,
        [string]$ErrorText,
        [string]$Frequency,
        [bool]$Blocking,
        [string]$From,
        [datetime]$When
    )

    $enc = {
        param([string]$Text)
        $t = "$Text"
        # Ampersand FIRST. Escaping it after the others would re-escape the
        # ampersands this very function just introduced, turning &lt; into
        # &amp;lt; and printing the markup instead of the character.
        $t = $t.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
        return $t.Replace("`r`n", '<br>').Replace("`n", '<br>')
    }

    $tenant = "$((Get-TenantConfig).TenantName)"
    if (-not $tenant) { $tenant = '(tenant name not set)' }

    $blockingCell = if ($Blocking) {
        "<span style='color:#B00020'><b>YES -- reporter cannot complete the task</b></span>"
    } else {
        'No -- there is a way round it'
    }

    return @"
<html>
<body style='font-family:Segoe UI,Calibri,sans-serif;font-size:11pt'>
<p><b>Bug report</b> submitted from the Admin Console.</p>
<table cellpadding='6' cellspacing='0' border='0'>
<tr><td valign='top'><b>From</b></td><td>$(& $enc $From)</td></tr>
<tr><td valign='top'><b>Submitted</b></td><td>$($When.ToString('yyyy-MM-dd HH:mm'))</td></tr>
<tr><td valign='top'><b>Tenant</b></td><td>$(& $enc $tenant)</td></tr>
<tr><td valign='top'><b>Machine</b></td><td>$(& $enc $env:COMPUTERNAME)</td></tr>
<tr><td valign='top'><b>Toolkit</b></td><td>$(& $enc (Get-ToolkitStateLine))</td></tr>
<tr><td valign='top'><b>PowerShell</b></td><td>$(& $enc "$($PSVersionTable.PSVersion)")</td></tr>
<tr><td valign='top'><b>Blocking</b></td><td>$blockingCell</td></tr>
<tr><td valign='top'><b>Happens</b></td><td>$(& $enc $Frequency)</td></tr>
</table>
<h3 style='margin-bottom:2px'>Tool</h3>
<p>$(& $enc $Tool)</p>
<h3 style='margin-bottom:2px'>What happened</h3>
<p>$(& $enc $What)</p>
<h3 style='margin-bottom:2px'>What was expected</h3>
<p>$(& $enc $Expected)</p>
<h3 style='margin-bottom:2px'>Steps to reproduce</h3>
<p>$(if ("$Steps".Trim()) { & $enc $Steps } else { '<i>(none given)</i>' })</p>
<h3 style='margin-bottom:2px'>Error message</h3>
$(if ("$ErrorText".Trim()) {
    "<pre style='font-family:Consolas,monospace;font-size:10pt;background:#F4F4F4;padding:8px;border:1px solid #DDD'>$(& $enc $ErrorText)</pre>"
} else { '<p><i>(none given)</i></p>' })
</body>
</html>
"@
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Bug Fix'
# ClientSize, not Size. The controls below are placed in client coordinates and
# .Size would spend about 39px of this on the title bar and border -- which is
# how three tools ended up with their button row drawn off the bottom.
$form.ClientSize    = New-Object System.Drawing.Size(664, 602)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblIntro          = New-Object System.Windows.Forms.Label
$lblIntro.Text     = 'Report something that is broken. A copy is sent to you.'
$lblIntro.Location = New-Object System.Drawing.Point(12, 12)
$lblIntro.Size     = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($lblIntro)

$lblTool          = New-Object System.Windows.Forms.Label
$lblTool.Text     = 'Which tool?  (required -- or "the console itself")'
$lblTool.Location = New-Object System.Drawing.Point(12, 42)
$lblTool.Size     = New-Object System.Drawing.Size(400, 18)
$form.Controls.Add($lblTool)

$txtTool          = New-Object System.Windows.Forms.TextBox
$txtTool.Location = New-Object System.Drawing.Point(12, 62)
$txtTool.Size     = New-Object System.Drawing.Size(400, 24)
$form.Controls.Add($txtTool)

$lblWhat          = New-Object System.Windows.Forms.Label
$lblWhat.Text     = 'What happened?  (required)'
$lblWhat.Location = New-Object System.Drawing.Point(12, 94)
$lblWhat.Size     = New-Object System.Drawing.Size(400, 18)
$form.Controls.Add($lblWhat)

$txtWhat            = New-Object System.Windows.Forms.TextBox
$txtWhat.Multiline  = $true
$txtWhat.ScrollBars = 'Vertical'
$txtWhat.Location   = New-Object System.Drawing.Point(12, 114)
$txtWhat.Size       = New-Object System.Drawing.Size(640, 80)
$txtWhat.Anchor     = 'Top,Left,Right'
$form.Controls.Add($txtWhat)

# EXPECTED IS REQUIRED, and that is deliberate rather than thorough.
#
# A good share of bug reports describe the tool doing exactly what it was built
# to do. Without this field the answer is "working as designed", which helps
# nobody; with it, the disagreement is visible and is usually the real bug.
$lblExpect          = New-Object System.Windows.Forms.Label
$lblExpect.Text     = 'What did you expect to happen instead?  (required)'
$lblExpect.Location = New-Object System.Drawing.Point(12, 202)
$lblExpect.Size     = New-Object System.Drawing.Size(400, 18)
$form.Controls.Add($lblExpect)

$txtExpect            = New-Object System.Windows.Forms.TextBox
$txtExpect.Multiline  = $true
$txtExpect.ScrollBars = 'Vertical'
$txtExpect.Location   = New-Object System.Drawing.Point(12, 222)
$txtExpect.Size       = New-Object System.Drawing.Size(640, 60)
$txtExpect.Anchor     = 'Top,Left,Right'
$form.Controls.Add($txtExpect)

$lblSteps          = New-Object System.Windows.Forms.Label
$lblSteps.Text     = 'Steps to reproduce  (optional -- but this is the field that gets it fixed)'
$lblSteps.Location = New-Object System.Drawing.Point(12, 290)
$lblSteps.Size     = New-Object System.Drawing.Size(600, 18)
$form.Controls.Add($lblSteps)

$txtSteps            = New-Object System.Windows.Forms.TextBox
$txtSteps.Multiline  = $true
$txtSteps.ScrollBars = 'Vertical'
$txtSteps.Location   = New-Object System.Drawing.Point(12, 310)
$txtSteps.Size       = New-Object System.Drawing.Size(640, 80)
$txtSteps.Anchor     = 'Top,Left,Right'
$form.Controls.Add($txtSteps)

$lblErr          = New-Object System.Windows.Forms.Label
$lblErr.Text     = 'Error message, pasted exactly as it appeared  (optional)'
$lblErr.Location = New-Object System.Drawing.Point(12, 398)
$lblErr.Size     = New-Object System.Drawing.Size(600, 18)
$form.Controls.Add($lblErr)

# Consolas, because this box is for pasted output. A proportional font turns a
# stack trace or an LDAP filter into something the reader has to decode.
$txtErr            = New-Object System.Windows.Forms.TextBox
$txtErr.Multiline  = $true
$txtErr.ScrollBars = 'Vertical'
$txtErr.Location   = New-Object System.Drawing.Point(12, 418)
$txtErr.Size       = New-Object System.Drawing.Size(640, 60)
$txtErr.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtErr.Anchor     = 'Top,Left,Right'
$form.Controls.Add($txtErr)

$lblFreq          = New-Object System.Windows.Forms.Label
$lblFreq.Text     = 'How often does it happen?  (required)'
$lblFreq.Location = New-Object System.Drawing.Point(12, 490)
$lblFreq.Size     = New-Object System.Drawing.Size(240, 18)
$form.Controls.Add($lblFreq)

# A LIST, NOT A NUMBER, and not the same question the feature request asks.
#
# "How many times a week" ranks a feature request. For a bug the useful split is
# reproducible versus not: "every time" and "now and then" are chased in
# completely different ways, and a count does not say which you have.
#
# STARTS BLANK ON PURPOSE, like the frequency box on the feature request. A
# pre-selected 'Every time' would let somebody tab past and have the form answer
# for them, which is worse than an empty field -- nobody notices a wrong answer
# that was never given.
$cboFreq               = New-Object System.Windows.Forms.ComboBox
$cboFreq.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
@('Every time', 'Most times', 'Now and then', 'Happened once') |
    ForEach-Object { [void]$cboFreq.Items.Add($_) }
$cboFreq.Location = New-Object System.Drawing.Point(256, 487)
$cboFreq.Size     = New-Object System.Drawing.Size(150, 24)
$form.Controls.Add($cboFreq)

$chkBlocking          = New-Object System.Windows.Forms.CheckBox
$chkBlocking.Text     = 'I cannot complete the task at all'
$chkBlocking.Location = New-Object System.Drawing.Point(424, 489)
$chkBlocking.Size     = New-Object System.Drawing.Size(228, 20)
$form.Controls.Add($chkBlocking)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(12, 516)
$lblStatus.Size     = New-Object System.Drawing.Size(640, 40)
$lblStatus.Anchor   = 'Top,Left,Right'
$form.Controls.Add($lblStatus)

$btnSend          = New-Object System.Windows.Forms.Button
$btnSend.Text     = 'Send'
$btnSend.Location = New-Object System.Drawing.Point(12, 560)
$btnSend.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnSend)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(130, 560)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnClose)

# PREREQUISITES ARE CHECKED BEFORE THE OPERATOR TYPES, not after.
#
# Both are configuration faults, and neither is anything the person in front of
# the form can fix. Discovering them on Send means they have already written the
# whole thing, and the failure arrives at the exact moment their work is lost --
# which on a bug report is doubly annoying, because they were already having a
# bad time when they opened it.
$configProblem = ''
if (-not $script:Recipient) {
    $configProblem = "No recipient is configured. Set FeatureRequest.Recipient in Config\Tenant.json."
} elseif (-not $script:OperatorEmail) {
    $configProblem = "Your own email address is not set, so you could not be sent a copy. Set 'Your email' in Settings."
}
if ($configProblem) {
    $btnSend.Enabled     = $false
    $lblStatus.Text      = $configProblem
    $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
}

$btnSend.Add_Click({
    $tool     = $txtTool.Text.Trim()
    $what     = $txtWhat.Text.Trim()
    $expected = $txtExpect.Text.Trim()
    $steps    = $txtSteps.Text.Trim()
    $errText  = $txtErr.Text.Trim()
    $freq     = "$($cboFreq.SelectedItem)"

    $missing = @()
    if (-not $tool)     { $missing += 'which tool' }
    if (-not $what)     { $missing += 'what happened' }
    if (-not $expected) { $missing += 'what you expected instead' }
    if (-not $freq)     { $missing += 'how often it happens' }

    if ($missing.Count -gt 0) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = "Still needed: $($missing -join '; ')."
        return
    }

    $btnSend.Enabled = $false
    $form.Cursor     = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $when = Get-Date
        $body = New-BugReportBody -Tool $tool -What $what -Expected $expected -Steps $steps `
                    -ErrorText $errText -Frequency $freq -Blocking $chkBlocking.Checked `
                    -From $script:OperatorEmail -When $when

        # BLOCKING IS IN THE SUBJECT, not only in the body.
        #
        # The subject line is what gets read in a list view, and "I cannot do my
        # job" losing to "the label is misaligned" because both look identical
        # from outside is the failure this prevents. The date is there for the
        # same reason as on a feature request: two reports on one topic are
        # otherwise indistinguishable, and Sent is the date the mail moved
        # rather than the date the fault happened.
        $flag    = if ($chkBlocking.Checked) { '[BLOCKING] ' } else { '' }
        $subject = "${flag}Bug Report - $($when.ToString('yyyy-MM-dd HH:mm'))"

        # ONE SEND. The operator is the To:, and the real destination is BCC.
        #
        # Sending twice would put the recipient's address in the operator's copy
        # -- the thing this is meant not to do -- and would give two independent
        # chances to fail on one button press.
        #
        # -Body $body -BodyAsHtml, spelled out. '-BodyAsHtml $body' also works,
        # but only because BodyAsHtml is a SWITCH and $body then binds
        # positionally to -Body. That is an accident the reader has to
        # reconstruct, and it stops working the moment anyone converts the call
        # to splatting -- which is exactly how the OneDrive notification in
        # Disable-User broke.
        Send-MailMessage -To $script:OperatorEmail -Bcc $script:Recipient `
            -Subject $subject -Body $body -BodyAsHtml `
            -From $script:DoNotReplyAddress -SmtpServer $script:SmtpServer -ErrorAction Stop

        # STILL DOES NOT CLAIM DELIVERY. Send-MailMessage returns when the RELAY
        # ACCEPTS the message, which proves handoff and nothing further -- the
        # mail can still be dropped, quarantined or junked after that. "Handed
        # to" is the strongest true statement available and it stays.
        #
        # The address and the junk-folder sentence were removed 2026-08-20, for
        # the reasons in FeatureRequest.ps1. These two lines stay identical on
        # purpose: they are the same message about the same send.
        $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen
        $lblStatus.Text      = "Handed to $($script:SmtpServer) -- subject '$subject'."

        # TARGET IS NOT THE OPERATOR'S ADDRESS. It used to be, and that was
        # wrong twice over.
        #
        # It DUPLICATED the Operator column, which already identifies who ran
        # the tool -- as a domain logon rather than a mail address, which is the
        # form the rest of the log uses. Two columns naming the same person, in
        # two different notations, is not a record; it is noise.
        #
        # And Target means "the thing that was acted on". Nothing was: this tool
        # sends an email and changes no mailbox, account or group. Here there IS a real one --
        # the tool the bug is about -- so the column earns its place: the log
        # becomes searchable by which tool is being complained about.
        Write-CompletedLog -Action 'BugReport' -Target $tool -Result 'Success' `
            -Detail "Bug reported against '$tool' ($freq$(if ($chkBlocking.Checked) { '; BLOCKING' })). Relay $($script:SmtpServer)."

        # CLEARED ONLY AFTER A CONFIRMED SEND. If the send throws, everything
        # typed is still on screen and Send can be pressed again -- a form that
        # empties itself on failure loses the report outright, and this one may
        # hold a pasted error message the operator can no longer reproduce.
        $txtTool.Clear(); $txtWhat.Clear(); $txtExpect.Clear()
        $txtSteps.Clear(); $txtErr.Clear()
        $cboFreq.SelectedIndex = -1
        $chkBlocking.Checked   = $false
    }
    catch {
        # REDACTED ONCE, THEN USED TWICE. An SMTP rejection quotes the address
        # it could not deliver to, and this message goes to two places that must
        # not carry it: the status label in front of the operator, and the
        # Detail column of the audit log.
        #
        # -Detail, NOT -ErrorRecord. Write-FailedLog appends
        # ErrorRecord.Exception.Message itself, so passing the record would put
        # the raw text back in the log underneath the cleaned copy.
        $safeMsg = Remove-RecipientAddress -Text $_.Exception.Message

        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = "Not sent -- $safeMsg  Your text has been kept; press Send to try again."
        Write-FailedLog -Action 'BugReport' -Target $tool `
            -Detail "Bug report mail failed -- $safeMsg"
    }
    finally {
        $btnSend.Enabled = $true
        $form.Cursor     = [System.Windows.Forms.Cursors]::Default
    }
})

$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form
Set-FormResizable -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Bug Fix' | Out-Null

[void]$form.ShowDialog()
#endregion
