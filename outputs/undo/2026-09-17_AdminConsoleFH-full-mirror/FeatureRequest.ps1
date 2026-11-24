#Requires -Version 5.1
<#
    FeatureRequest.ps1
    Lets an operator send a console feature request without leaving the console.

    Reasoning: ScriptNotes/FeatureRequest.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-19
#>
# TOOLKIT-HELP-BEGIN
# Ask for something the console does not do yet.
#
# Fill in the three boxes and press Send. The request is mailed off and a copy
# lands in your own inbox, so you have a record of exactly what you asked for.
#
# ANSWER THE FREQUENCY QUESTION HONESTLY. It is the field that decides what gets
# built first. "Twice a day" and "once a quarter" are both fine answers -- a
# small annoyance that happens constantly is usually worth more than a large one
# that almost never does, and there is no way to know which yours is unless you
# say. Guessing high does not help you; it just moves the queue around.
#
# WHAT MAKES A REQUEST EASY TO ACT ON:
#   - Describe the PROBLEM, not the solution you have in mind. "I have to open
#     three tools to answer one ticket" leads somewhere better than "add a
#     button here", because the button might not be the best fix.
#   - Name the tool and the exact step, if it is about one tool.
#   - Say what you do today instead. The workaround is often most of the design.
#
# You do not need to know whether it is possible. Ask anyway.
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

# The requester, and the person who gets the copy. Same address, deliberately.
$script:OperatorEmail = Get-OperatorEmail

# NOT A CONSTANT IN THIS FILE. See Get-FeatureRequestRecipient -- a mail domain
# in a .ps1 is the one thing Test-Toolkit refuses, and this toolkit ships to six
# tenants.
$script:Recipient = Get-FeatureRequestRecipient
#endregion

#region Actions --------------------------------------------------------------
function New-FeatureRequestBody {
    <#
        The mail body.

        Every field is HTML-encoded. An operator describing a problem with the
        console has every reason to type "<" or "&" -- "the < button does
        nothing" is a perfectly ordinary sentence -- and pasting that raw into
        an HTML body silently eats the rest of the paragraph. The failure is
        invisible to the sender, because their copy is broken in exactly the
        same way.
    #>
    param(
        [string]$Request,
        [string]$Problem,
        [string]$Frequency,
        [string]$Details,
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

    return @"
<html>
<body style='font-family:Segoe UI,Calibri,sans-serif;font-size:11pt'>
<p><b>Feature request</b> submitted from the Admin Console.</p>
<table cellpadding='6' cellspacing='0' border='0'>
<tr><td valign='top'><b>From</b></td><td>$(& $enc $From)</td></tr>
<tr><td valign='top'><b>Submitted</b></td><td>$($When.ToString('yyyy-MM-dd HH:mm'))</td></tr>
<tr><td valign='top'><b>Tenant</b></td><td>$(& $enc $tenant)</td></tr>
<tr><td valign='top'><b>Machine</b></td><td>$(& $enc $env:COMPUTERNAME)</td></tr>
</table>
<h3 style='margin-bottom:2px'>The request</h3>
<p>$(& $enc $Request)</p>
<h3 style='margin-bottom:2px'>Problem it solves</h3>
<p>$(& $enc $Problem)</p>
<h3 style='margin-bottom:2px'>How often this comes up</h3>
<p>$(& $enc $Frequency) time(s) per week</p>
<h3 style='margin-bottom:2px'>Additional details / justification</h3>
<p>$(if ("$Details".Trim()) { & $enc $Details } else { '<i>(none given)</i>' })</p>
</body>
</html>
"@
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Feature Request'
# ClientSize, not Size: the controls are placed in client coordinates, and
# .Size would spend about 39px of this on the title bar and border.
#
# 532 = the Send/Close row at 490 plus its 30px height plus a 12px margin. It
# was 640, which left 81px of empty window under the buttons -- not clipped, so
# LAY001/LAY002 never said a word about it, and nobody noticed until it was
# looked at next to the Bug Fix form.
$form.ClientSize    = New-Object System.Drawing.Size(664, 544)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblIntro           = New-Object System.Windows.Forms.Label
$lblIntro.Text      = 'Ask for something the console does not do yet. A copy is sent to you.'
$lblIntro.Location  = New-Object System.Drawing.Point(12, 12)
$lblIntro.Size      = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($lblIntro)

$lblReq          = New-Object System.Windows.Forms.Label
$lblReq.Text     = 'What would you like the console to do?  (required)'
$lblReq.Location = New-Object System.Drawing.Point(12, 42)
$lblReq.Size     = New-Object System.Drawing.Size(400, 18)
$form.Controls.Add($lblReq)

$txtReq            = New-Object System.Windows.Forms.TextBox
$txtReq.Multiline  = $true
$txtReq.ScrollBars = 'Vertical'
$txtReq.Location   = New-Object System.Drawing.Point(12, 62)
$txtReq.Size       = New-Object System.Drawing.Size(640, 80)
$txtReq.Anchor     = 'Top,Left,Right'
$form.Controls.Add($txtReq)

$lblProb          = New-Object System.Windows.Forms.Label
$lblProb.Text     = 'What problem are you trying to solve?  (required)'
$lblProb.Location = New-Object System.Drawing.Point(12, 150)
$lblProb.Size     = New-Object System.Drawing.Size(400, 18)
$form.Controls.Add($lblProb)

$txtProb            = New-Object System.Windows.Forms.TextBox
$txtProb.Multiline  = $true
$txtProb.ScrollBars = 'Vertical'
$txtProb.Location   = New-Object System.Drawing.Point(12, 170)
$txtProb.Size       = New-Object System.Drawing.Size(640, 100)
$txtProb.Anchor     = 'Top,Left,Right'
$form.Controls.Add($txtProb)

$lblFreq          = New-Object System.Windows.Forms.Label
$lblFreq.Text     = 'How many times a week does this come up?  (required)'
$lblFreq.Location = New-Object System.Drawing.Point(12, 280)
$lblFreq.Size     = New-Object System.Drawing.Size(340, 18)
$form.Controls.Add($lblFreq)

# A NumericUpDown, not a text box, because "how often" invites prose -- "loads",
# "all the time", "3-4ish" -- and prose cannot be sorted. The whole point of
# this field is to rank one request against another.
$numFreq          = New-Object System.Windows.Forms.NumericUpDown
$numFreq.Location = New-Object System.Drawing.Point(356, 278)
$numFreq.Size     = New-Object System.Drawing.Size(80, 24)
$numFreq.Minimum  = 0
$numFreq.Maximum  = 500
# ZERO IS THE STARTING VALUE, and it fails validation on purpose. A NumericUpDown
# defaulting to 1 would let somebody tab past the field and have it answer for
# them, which is worse than an empty box: nobody notices a wrong number that was
# never typed.
$numFreq.Value    = 0
$form.Controls.Add($numFreq)

$lblFreqHint           = New-Object System.Windows.Forms.Label
$lblFreqHint.Text      = 'Best estimate is fine. This is what decides priority.'
$lblFreqHint.Location  = New-Object System.Drawing.Point(444, 281)
# TWO LINES. At 8pt italic in 208px this text wraps, and 18px showed the
# first line and half of the word 'priority' -- which is the word the sentence
# exists for. A wrapping Label never reports that it ran out of room.
$lblFreqHint.Size      = New-Object System.Drawing.Size(208, 32)
$lblFreqHint.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Italic)
$lblFreqHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblFreqHint)

$lblDet          = New-Object System.Windows.Forms.Label
$lblDet.Text     = 'Additional details / justification  (optional)'
$lblDet.Location = New-Object System.Drawing.Point(12, 322)
$lblDet.Size     = New-Object System.Drawing.Size(400, 18)
$form.Controls.Add($lblDet)

$txtDet            = New-Object System.Windows.Forms.TextBox
$txtDet.Multiline  = $true
$txtDet.ScrollBars = 'Vertical'
$txtDet.Location   = New-Object System.Drawing.Point(12, 342)
$txtDet.Size       = New-Object System.Drawing.Size(640, 100)
$txtDet.Anchor     = 'Top,Left,Right'
$form.Controls.Add($txtDet)

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(12, 452)
$lblStatus.Size      = New-Object System.Drawing.Size(640, 40)
$lblStatus.Anchor    = 'Top,Left,Right'
$form.Controls.Add($lblStatus)

$btnSend          = New-Object System.Windows.Forms.Button
$btnSend.Text     = 'Send'
$btnSend.Location = New-Object System.Drawing.Point(12, 502)
$btnSend.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnSend)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(130, 502)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnClose)

# PREREQUISITES ARE CHECKED BEFORE THE OPERATOR TYPES, not after.
#
# Both of these are configuration faults, and neither is anything the person in
# front of the form can fix. Discovering them on Send means they have already
# written three paragraphs, and the failure message arrives at the exact moment
# their work is about to be lost.
$configProblem = ''
if (-not $script:Recipient) {
    $configProblem = "No Feature Request recipient is configured. Set FeatureRequest.Recipient in Config\Tenant.json."
} elseif (-not $script:OperatorEmail) {
    $configProblem = "Your own email address is not set, so you could not be sent a copy. Set 'Your email' in Tool Manager > Settings."
}
if ($configProblem) {
    $btnSend.Enabled     = $false
    $lblStatus.Text      = $configProblem
    $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
}

$btnSend.Add_Click({
    $request = $txtReq.Text.Trim()
    $problem = $txtProb.Text.Trim()
    $details = $txtDet.Text.Trim()
    $freq    = [int]$numFreq.Value

    $missing = @()
    if (-not $request)  { $missing += 'what you would like the console to do' }
    if (-not $problem)  { $missing += 'the problem it solves' }
    if ($freq -le 0)    { $missing += 'how many times a week it comes up' }

    if ($missing.Count -gt 0) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = "Still needed: $($missing -join '; ')."
        return
    }

    $btnSend.Enabled = $false
    $form.Cursor     = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $when = Get-Date
        $body = New-FeatureRequestBody -Request $request -Problem $problem -Frequency "$freq" `
                    -Details $details -From $script:OperatorEmail -When $when

        # The date and time are in the SUBJECT, not just the headers.
        #
        # Two requests from the same person on the same topic are otherwise
        # indistinguishable in a mail list, and the Sent date is the date the
        # mail moved rather than the date the request was made. This also stops
        # a mail client threading unrelated requests together on subject alone.
        $subject = "Feature Request - $($when.ToString('yyyy-MM-dd HH:mm'))"

        # ONE SEND. The operator is the To:, and the real destination is BCC.
        #
        # Sending twice would put the recipient's address in the operator's copy
        # -- which is the thing this is meant not to do -- and would give two
        # independent chances to fail on one button press.
        # -Body $body -BodyAsHtml, spelled out.
        #
        # '-BodyAsHtml $body' also works, but only because BodyAsHtml is a
        # SWITCH and $body then binds positionally to -Body. That is an accident
        # the reader has to reconstruct, and it stops working the moment anyone
        # converts the call to splatting -- which is exactly how the OneDrive
        # notification in Disable-User broke.
        Send-MailMessage -To $script:OperatorEmail -Bcc $script:Recipient `
            -Subject $subject -Body $body -BodyAsHtml `
            -From $script:DoNotReplyAddress -SmtpServer $script:SmtpServer -ErrorAction Stop

        # STILL DOES NOT CLAIM DELIVERY. Send-MailMessage returns when the RELAY
        # ACCEPTS the message, which proves handoff and nothing further -- the
        # mail can still be dropped, quarantined or junked after that. "Handed
        # to" is the strongest true statement available and it stays.
        #
        # THE ADDRESS AND THE JUNK-FOLDER SENTENCE WERE REMOVED, 2026-08-20.
        # They existed to turn "it says sent but nothing arrived" into a
        # one-line answer, which is a real cost being paid here -- but the line
        # ran to two wrapped rows for something read once and never again, and
        # the subject is enough to find the mail. The audit log still records
        # the operator and the relay if a send ever has to be traced.
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
        # sends an email and changes no mailbox, account or group. Saying so plainly is more honest than
        # naming a bystander.
        Write-CompletedLog -Action 'FeatureRequest' -Target '(mail only -- nothing changed)' -Result 'Success' `
            -Detail "Feature request submitted ($freq/week). Relay $($script:SmtpServer), from $($script:DoNotReplyAddress)."

        # CLEARED ONLY AFTER A CONFIRMED SEND. If the send throws, everything
        # the operator typed is still on screen and Send can be pressed again --
        # a form that empties itself on failure loses the request outright.
        $txtReq.Clear(); $txtProb.Clear(); $txtDet.Clear()
        $numFreq.Value = 0
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
        Write-FailedLog -Action 'FeatureRequest' -Target '(mail only -- nothing changed)' `
            -Detail "Feature request mail failed -- $safeMsg"
    }
    finally {
        $btnSend.Enabled = $true
        $form.Cursor     = [System.Windows.Forms.Cursors]::Default
    }
})

$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form
Set-FormResizable -Form $form

# BOTTOM-LEFT, like the rest of the toolkit. Missed when this was written --
# same omission as Group-Manager-Editor, found by the same sweep.
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Feature Request' | Out-Null

[void]$form.ShowDialog()
#endregion
