#Requires -Version 5.1
<#
    OnPrem-MailTest.ps1
    Send a test message through each hub server and report what happened.

    BASED ON A HAND-RUN SCRIPT, with one bug corrected: that script built the
    subject from $Server but then passed SmtpServer = $servers -- the whole
    ARRAY rather than the current server. Send-MailMessage takes the first
    entry, so every iteration went through the FIRST server while the subject
    claimed otherwise. Four tests that all said different things and all took
    the same path.

    -Environment IS ACCEPTED AND USED, and passed to every Write-*Log call.
#>
# TOOLKIT-HELP-RUNS: SENDS real email through each on-premises hub server. Delivers mail.
# TOOLKIT-VARIANT: OnPrem
# TOOLKIT-HELP-BEGIN
# Send a test message through each hub server, one at a time, and report which
# ones accepted it.
#
# THIS SENDS REAL EMAIL. It is not a simulation. Anything you put in the
# recipient boxes will receive a message from you.
#
# WHY TWO RECIPIENT BOXES. Internal and external delivery are different tests
# and they fail independently:
#
#   Internal   proves the server accepts mail and delivers inside the org.
#   External   proves outbound relay, the send connector, and whatever sits in
#              front of it. This is the one that breaks after a firewall or
#              connector change, and an internal-only test will not notice.
#
# Either box on its own is enough. Fill in both and each server sends two
# separate messages, so a server that delivers internally but cannot relay out
# shows as one pass and one fail rather than a single ambiguous failure.
#
# EACH SERVER IS TESTED SEPARATELY and the server name goes in the subject, so
# the messages that arrive tell you which route worked without reading this
# grid. Doing them one at a time is the point -- if three deliver and one does
# not, that names the broken server.
#
# ACCEPTED IS NOT DELIVERED. A green row means the server took the message, not
# that it arrived. Check the mailbox as well; a message can be accepted and then
# dropped by a rule, a filter, or the receiving side.
#
# ATTACHMENTS. 'Add files...' picks one or several; go again to add more from a
# different folder. Double-click a row, or select rows and press Remove, to drop
# them. Every test message carries the same files.
#
# WATCH THE SIZE LINE, because the number that gets a message rejected is not
# the number on disk. Attachments travel base64-encoded, which is about 1.37x
# bigger, and that encoded size is what a connector limit is measured against --
# a 20MB file will not pass a 25MB limit. The panel shows both, and the total
# across every message queued, because one 10MB file to two recipients on four
# servers is 80MB of real traffic.
#
# IF A SEND FAILS ON SIZE the summary says so rather than leaving it looking
# like a routing fault, and tells you to retry with nothing attached to prove
# the route itself is fine. That distinction is the whole reason this box is
# here: "mail is broken" and "mail is broken for a 30MB file" are different
# problems with different owners.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

#region Load Common -------------------------------------------------------------
$root = $PSScriptRoot
while ($root -and -not (Test-Path -LiteralPath (Join-Path $root 'Common.ps1'))) {
    $parent = Split-Path -Parent $root
    if (-not $parent -or $parent -eq $root) { $root = ''; break }
    $root = $parent
}
if (-not $root) { throw "Common.ps1 not found above '$PSScriptRoot'." }
. (Join-Path $root 'Common.ps1')
#endregion

if (-not (Test-ConsoleLaunch)) { return }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# From Tenant.json, not from here -- see Get-TenantExchangeServerList.
$script:Servers = @()
try { $script:Servers = @(Get-TenantExchangeServerList) } catch { $script:Servers = @() }
$script:Results = New-Object System.Collections.ArrayList

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'On-Premises Mail Test'
# 824 RATHER THAN 720. The attachment row added 104px to the Message box and
# everything below it is placed absolutely, so the window grew by the same
# amount instead of squeezing the output box -- which is where the diagnosis
# gets read.
$form.ClientSize      = New-Object System.Drawing.Size(1000, 824)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$grp          = New-Object System.Windows.Forms.GroupBox
$grp.Text     = 'Message'
$grp.Location = New-Object System.Drawing.Point(14, 46)
$grp.Size     = New-Object System.Drawing.Size(972, 314)
$grp.Anchor   = 'Top,Left,Right'
[void]$form.Controls.Add($grp)

function New-Row {
    param([string]$Label, [int]$Y, [string]$Hint)
    $l          = New-Object System.Windows.Forms.Label
    $l.Text     = $Label
    $l.Location = New-Object System.Drawing.Point(14, ($Y + 3))
    $l.Size     = New-Object System.Drawing.Size(130, 20)
    [void]$grp.Controls.Add($l)

    $t          = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point(150, $Y)
    $t.Size     = New-Object System.Drawing.Size(360, 22)
    [void]$grp.Controls.Add($t)

    if ($Hint) {
        $h           = New-Object System.Windows.Forms.Label
        $h.Text      = $Hint
        $h.Location  = New-Object System.Drawing.Point(520, ($Y + 3))
        $h.Size      = New-Object System.Drawing.Size(438, 20)
        $h.ForeColor = [System.Drawing.Color]::Gray
        $h.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
        [void]$grp.Controls.Add($h)
    }
    return $t
}

$txtFrom     = New-Row -Label 'From'               -Y 26  -Hint 'The sender. Must be an address these servers will accept as internal.'
$txtInternal = New-Row -Label 'Internal recipient' -Y 56  -Hint 'Proves the server accepts and delivers inside the org.'
$txtExternal = New-Row -Label 'External recipient' -Y 86  -Hint 'Proves outbound relay. This is the one that breaks after a connector change.'
$txtSubject  = New-Row -Label 'Subject'            -Y 116 -Hint 'The server name is appended, so each message says which route it took.'

$lblBody          = New-Object System.Windows.Forms.Label
$lblBody.Text     = 'Body'
$lblBody.Location = New-Object System.Drawing.Point(14, 149)
$lblBody.Size     = New-Object System.Drawing.Size(130, 20)
[void]$grp.Controls.Add($lblBody)

$txtBody            = New-Object System.Windows.Forms.TextBox
$txtBody.Multiline  = $true
$txtBody.ScrollBars = 'Vertical'
$txtBody.Location   = New-Object System.Drawing.Point(150, 146)
$txtBody.Size       = New-Object System.Drawing.Size(360, 52)
[void]$grp.Controls.Add($txtBody)

# NO SSL CHECKBOX, AND NO -UseSsl. STARTTLS is not used between these servers
# and M365, so a toggle for it would offer a mode that is never correct.
#
# SECURE IS A DIFFERENT THING ENTIRELY. It does not change how the message is
# transmitted -- it puts [secure] in the subject, which is what the encryption
# transport rule matches on. Nothing here encrypts anything; the rule does, and
# only if it is still configured to look for that keyword. That is precisely
# what this tests.
$chkSecure          = New-Object System.Windows.Forms.CheckBox
$chkSecure.Text     = 'Secure'
$chkSecure.Location = New-Object System.Drawing.Point(520, 148)
$chkSecure.Size     = New-Object System.Drawing.Size(90, 20)
$chkSecure.Checked  = $false
[void]$grp.Controls.Add($chkSecure)

$lblSecHint           = New-Object System.Windows.Forms.Label
$lblSecHint.Text      = 'Prefixes the subject with [secure] so the encryption rule acts on it. Tests the rule, not the transport.'
$lblSecHint.Location  = New-Object System.Drawing.Point(520, 170)
$lblSecHint.Size      = New-Object System.Drawing.Size(438, 30)
$lblSecHint.ForeColor = [System.Drawing.Color]::Gray
$lblSecHint.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grp.Controls.Add($lblSecHint)

# --- attachments ----------------------------------------------------------------
#
# FULL PATHS ARE HELD HERE, not in the ListBox. The list shows leaf names and a
# size, because a column of UNC paths is unreadable at 360px and the operator
# picked the files so they know where they came from. The two stay index-aligned
# because the list is only ever rebuilt from this collection, in order.
$script:Attachments   = New-Object System.Collections.ArrayList
$script:LastAttachDir = $null

$lblAtt          = New-Object System.Windows.Forms.Label
$lblAtt.Text     = 'Attachments'
$lblAtt.Location = New-Object System.Drawing.Point(14, 209)
$lblAtt.Size     = New-Object System.Drawing.Size(130, 20)
[void]$grp.Controls.Add($lblAtt)

# MultiExtended so several can be removed in one go -- picking six files and
# then removing them one at a time is a worse experience than not having the
# feature.
$lstAtt                = New-Object System.Windows.Forms.ListBox
$lstAtt.Location       = New-Object System.Drawing.Point(150, 206)
$lstAtt.Size           = New-Object System.Drawing.Size(360, 95)
$lstAtt.SelectionMode  = 'MultiExtended'
$lstAtt.IntegralHeight = $false
$lstAtt.HorizontalScrollbar = $true
[void]$grp.Controls.Add($lstAtt)

$btnAttAdd           = New-Object System.Windows.Forms.Button
$btnAttAdd.Text      = 'Add files...'
$btnAttAdd.Location  = New-Object System.Drawing.Point(520, 206)
$btnAttAdd.Size      = New-Object System.Drawing.Size(120, 26)
$btnAttAdd.FlatStyle = 'Flat'
[void]$grp.Controls.Add($btnAttAdd)

$btnAttRemove           = New-Object System.Windows.Forms.Button
$btnAttRemove.Text      = 'Remove'
$btnAttRemove.Location  = New-Object System.Drawing.Point(520, 238)
$btnAttRemove.Size      = New-Object System.Drawing.Size(120, 26)
$btnAttRemove.FlatStyle = 'Flat'
[void]$grp.Controls.Add($btnAttRemove)

$btnAttClear           = New-Object System.Windows.Forms.Button
$btnAttClear.Text      = 'Clear all'
$btnAttClear.Location  = New-Object System.Drawing.Point(520, 270)
$btnAttClear.Size      = New-Object System.Drawing.Size(120, 26)
$btnAttClear.FlatStyle = 'Flat'
[void]$grp.Controls.Add($btnAttClear)

# THE SIZE LINE IS THE POINT OF THIS LABEL, not decoration. See
# Update-AttachmentList: the number that gets a message rejected is the
# MIME-encoded size, not the size on disk, and it is multiplied by the number
# of test messages.
$lblAttHint           = New-Object System.Windows.Forms.Label
$lblAttHint.Location  = New-Object System.Drawing.Point(650, 206)
$lblAttHint.Size      = New-Object System.Drawing.Size(308, 95)
$lblAttHint.ForeColor = [System.Drawing.Color]::Gray
$lblAttHint.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grp.Controls.Add($lblAttHint)

$btnSend           = New-Object System.Windows.Forms.Button
$btnSend.Text      = 'Send tests'
$btnSend.Location  = New-Object System.Drawing.Point(14, 12)
$btnSend.Size      = New-Object System.Drawing.Size(130, 28)
$btnSend.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnSend)

$lblProg           = New-Object System.Windows.Forms.Label
$lblProg.Location  = New-Object System.Drawing.Point(158, 18)
$lblProg.Size      = New-Object System.Drawing.Size(600, 20)
$lblProg.ForeColor = [System.Drawing.Color]::Gray
[void]$form.Controls.Add($lblProg)

# EVERYTHING FROM HERE DOWN MOVED 104px, which is what the attachment row added
# to the Message box above. These are absolute coordinates with no layout
# manager, so the shift is manual and has to be exact -- the Message box now
# ends at y=360 and anything still sitting at its old y is underneath it.
$pb = Add-ReportProgressBar -Form $form -X 14 -Y 366 -Width 972 -Height 12

$grid                     = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(14, 386)
$grid.Size                = New-Object System.Drawing.Size(972, 200)
$grid.ReadOnly            = $true
$grid.AllowUserToAddRows  = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode       = 'FullRowSelect'
$grid.Anchor              = 'Top,Left,Right'
[void]$form.Controls.Add($grid)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(14, 596)
$txtOut.Size       = New-Object System.Drawing.Size(972, 178)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(866, 784)
$btnClose.Size      = New-Object System.Drawing.Size(120, 28)
$btnClose.FlatStyle = 'Flat'
$btnClose.Anchor    = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
[void]$form.Controls.Add($btnClose)
#endregion

#region Attachments -------------------------------------------------------------
function Format-AttachSize {
    <#
        Bytes as something readable.

        Local rather than in Common.ps1: nothing else in the toolkit formats a
        byte count, and adding a function to the shared layer for one caller
        means regenerating Common.ps1 (rule 4) for a cosmetic helper.
    #>
    param([Parameter(Mandatory)][double]$Bytes)
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N2} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return ('{0:N0} bytes' -f $Bytes)
}

function Get-AttachmentTotals {
    <#
        Raw size, MIME-encoded estimate, and anything that has gone missing.

        THE ENCODED SIZE IS THE ONE THAT GETS REJECTED. Attachments travel
        base64, which is 4 bytes out for every 3 in -- about 1.37x once line
        breaks are counted. A 20MB file against a 25MB connector limit fails,
        and the size on disk gives no hint why. This tool exists to explain
        rejections, so it does the arithmetic the operator would otherwise do
        after the failure.
    #>
    $raw = 0.0
    $missing = New-Object System.Collections.ArrayList
    foreach ($p in @($script:Attachments)) {
        if (Test-Path -LiteralPath $p) {
            try { $raw += [double](Get-Item -LiteralPath $p -ErrorAction Stop).Length } catch { }
        } else {
            [void]$missing.Add($p)
        }
    }
    return [PSCustomObject]@{
        Count   = @($script:Attachments).Count
        Raw     = $raw
        Encoded = [double]($raw * 1.37)
        Missing = @($missing.ToArray())
    }
}

function Update-AttachmentHint {
    <#
        The size line only.

        SEPARATE FROM THE LIST because the message count depends on the
        recipient boxes, so this is called on every keystroke in either of them.
        Rebuilding the ListBox that often would flicker and waste the work.
    #>
    $tot = Get-AttachmentTotals
    if ($tot.Count -eq 0) {
        $lblAttHint.Text = 'No attachments. Add files to test whether size or file type is what a ' +
            'message is being rejected for.'
        return
    }

    # Messages are per recipient per server, so the attachment is sent once for
    # each of them. Worth saying before somebody pushes 400MB through four hub
    # servers to test one connector.
    $perMsg = $tot.Encoded
    $msgs   = @($script:Servers).Count * @(@("$($txtInternal.Text)".Trim(), "$($txtExternal.Text)".Trim()) | Where-Object { $_ }).Count
    if ($msgs -lt 1) { $msgs = 1 }

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("$($tot.Count) file(s), $(Format-AttachSize -Bytes $tot.Raw) on disk.")
    [void]$lines.Add("About $(Format-AttachSize -Bytes $perMsg) per message once base64-encoded -- that is the")
    [void]$lines.Add("size the connector limit applies to, not the size on disk.")
    if ($msgs -gt 1) {
        [void]$lines.Add("$msgs messages queued = roughly $(Format-AttachSize -Bytes ($perMsg * $msgs)) in total.")
    }
    if (@($tot.Missing).Count -gt 0) {
        [void]$lines.Add("$(@($tot.Missing).Count) file(s) are MISSING and will block the send.")
    }
    $lblAttHint.Text = ($lines -join [Environment]::NewLine)
}

function Update-AttachmentList {
    <#
        Redraw the list, then the size line.

        Rebuilt from $script:Attachments every time rather than patched in
        place, so the rows and the paths behind them cannot drift out of index
        alignment -- removal is the operation that would otherwise get it wrong.
    #>
    $lstAtt.BeginUpdate()
    try {
        $lstAtt.Items.Clear()
        foreach ($p in @($script:Attachments)) {
            $leaf = Split-Path -Leaf $p
            if (Test-Path -LiteralPath $p) {
                $len = 0.0
                try { $len = [double](Get-Item -LiteralPath $p -ErrorAction Stop).Length } catch { }
                [void]$lstAtt.Items.Add("$leaf   ($(Format-AttachSize -Bytes $len))")
            } else {
                # SAID PLAINLY IN THE LIST, not discovered at send time. The
                # file was there when it was picked; between then and sending it
                # can be moved, and eight identical "could not find file" rows
                # is the worst way to learn that.
                [void]$lstAtt.Items.Add("$leaf   (MISSING)")
            }
        }
    } finally {
        $lstAtt.EndUpdate()
    }
    Update-AttachmentHint
}

function Add-Attachment {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title       = 'Select the file(s) to attach to the test messages'
    $dlg.Filter      = 'All files (*.*)|*.*'
    $dlg.Multiselect = $true
    if ($script:LastAttachDir -and (Test-Path -LiteralPath $script:LastAttachDir)) {
        $dlg.InitialDirectory = $script:LastAttachDir
    }
    if ($dlg.ShowDialog() -ne 'OK') { return }

    # ADDED, NOT REPLACED, so several trips to different folders build one list.
    # Duplicates are skipped: the same file twice is never the intent and
    # doubles the message size without saying so.
    $added = 0; $dupe = 0
    foreach ($file in @($dlg.FileNames)) {
        if (@($script:Attachments) -contains $file) { $dupe++; continue }
        [void]$script:Attachments.Add($file)
        $added++
    }
    if (@($dlg.FileNames).Count -gt 0) {
        $script:LastAttachDir = Split-Path -Parent @($dlg.FileNames)[0]
    }
    Update-AttachmentList
    $note = "Added $added file(s)."
    if ($dupe -gt 0) { $note += " $dupe already on the list and skipped." }
    $lblProg.Text = $note
}

function Remove-SelectedAttachment {
    $idx = @($lstAtt.SelectedIndices)
    if (@($idx).Count -eq 0) {
        $lblProg.Text = 'Select a file in the list first.'
        return
    }
    # DESCENDING, because removing index 2 renumbers everything above it and
    # ascending removal then deletes the wrong rows.
    foreach ($i in @($idx | Sort-Object -Descending)) {
        if ($i -ge 0 -and $i -lt @($script:Attachments).Count) { $script:Attachments.RemoveAt($i) }
    }
    Update-AttachmentList
    $lblProg.Text = "Removed $(@($idx).Count) file(s)."
}

function Clear-Attachment {
    if (@($script:Attachments).Count -eq 0) { return }
    $n = @($script:Attachments).Count
    $script:Attachments.Clear()
    Update-AttachmentList
    $lblProg.Text = "Removed all $n attachment(s)."
}
#endregion

#region Send --------------------------------------------------------------------
function Invoke-MailTest {
    $from     = "$($txtFrom.Text)".Trim()
    $internal = "$($txtInternal.Text)".Trim()
    $external = "$($txtExternal.Text)".Trim()
    $subject  = "$($txtSubject.Text)".Trim()
    $body     = "$($txtBody.Text)"

    if (-not $from) { $txtOut.Text = 'Enter a From address.'; return }
    # EITHER RECIPIENT WILL DO, and the message says which two boxes count --
    # "enter a recipient" with two empty boxes is not an instruction.
    if (-not $internal -and -not $external) {
        $txtOut.Text = 'Enter an internal recipient, an external recipient, or both. Either one on its own is a valid test.'
        return
    }
    foreach ($a in @($from, $internal, $external)) {
        if ($a -and ($a -notlike '*@*' -or $a -like '* *')) { $txtOut.Text = "'$a' is not an email address."; return }
    }
    if (-not $subject) { $subject = 'Mail flow test' }
    if (-not $body)    { $body    = "Test message sent $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." }

    # ATTACHMENTS ARE CHECKED ONCE, HERE, BEFORE ANYTHING IS SENT.
    #
    # Send-MailMessage opens the files itself, so a file that has moved since it
    # was picked fails inside the loop -- and it would fail for every target,
    # producing eight identical "could not find file" rows that look like a mail
    # problem. One check up front, and nothing is sent.
    $attach = @()
    $attTot = Get-AttachmentTotals
    if ($attTot.Count -gt 0) {
        if (@($attTot.Missing).Count -gt 0) {
            $txtOut.Text = 'These attachments are no longer where they were when you picked them:' +
                [Environment]::NewLine + [Environment]::NewLine +
                (@($attTot.Missing) -join [Environment]::NewLine) +
                [Environment]::NewLine + [Environment]::NewLine +
                'Put the files back, or select them in the list and press Remove. Nothing has been sent.'
            return
        }
        $attach = @($script:Attachments)
    }

    # ONE TARGET PER RECIPIENT PER SERVER, not one message with two recipients.
    # A single message to both would give one result for two different routes,
    # and internal delivery succeeding would mask outbound relay failing --
    # which is the failure this tool exists to find.
    $targets = New-Object System.Collections.ArrayList
    foreach ($srv in $script:Servers) {
        if ($internal) { [void]$targets.Add([PSCustomObject]@{ Server = $srv; Kind = 'Internal'; To = $internal }) }
        if ($external) { [void]$targets.Add([PSCustomObject]@{ Server = $srv; Kind = 'External'; To = $external }) }
    }

    $ans = Show-ToolDialog(
        "Send $(@($targets).Count) real test message(s)?" + [Environment]::NewLine + [Environment]::NewLine +
        "Servers  : $($script:Servers -join ', ')" + [Environment]::NewLine +
        "From     : $from" + [Environment]::NewLine +
        $(if ($internal) { "Internal : $internal" + [Environment]::NewLine } else { '' }) +
        $(if ($external) { "External : $external" + [Environment]::NewLine } else { '' }) +
        $(if ($chkSecure.Checked) { '[secure] : yes -- subject is prefixed, the encryption rule should act on it' + [Environment]::NewLine } else { '' }) +
        $(if (@($attach).Count -gt 0) {
            "Attached : $(@($attach).Count) file(s), $(Format-AttachSize -Bytes $attTot.Raw) on disk" + [Environment]::NewLine +
            "           ~$(Format-AttachSize -Bytes $attTot.Encoded) per message encoded, ~$(Format-AttachSize -Bytes ($attTot.Encoded * @($targets).Count)) in total" + [Environment]::NewLine
        } else { '' }) +
        [Environment]::NewLine + 'These are real emails and will arrive in real mailboxes.',
        'Send test mail', [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = 'Cancelled. Nothing was sent.'; return }

    $script:Results.Clear()
    $pb.Maximum = [Math]::Max(1, @($targets).Count)
    $pb.Value   = 0
    $pb.Visible = $true
    $btnSend.Enabled = $false

    try {
        foreach ($t in $targets) {
            $lblProg.Text = "Sending via $($t.Server) to $($t.To) ..."
            $form.Refresh()

            # THE SUBJECT CARRIES THE SERVER, so the messages that arrive say
            # which route delivered them without cross-referencing this grid.
            # [secure] GOES FIRST, before anything this tool appends. The rule
            # matches on the subject and a prefix survives any client that
            # truncates or prepends "RE:" -- burying the keyword after the
            # server name is how a rule quietly stops matching.
            $subj = '{0} [{1}] via {2}' -f $subject, $t.Kind, $t.Server
            if ($chkSecure.Checked) { $subj = "[secure] $subj" }
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $ok = $false; $detail = ''
            try {
                $params = @{
                    To         = $t.To
                    From       = $from
                    Subject    = $subj
                    Body       = $body
                    # ONE SERVER, THE CURRENT ONE. The script this came from
                    # passed the whole array here, so every send went through
                    # the first server while the subject named a different one.
                    SmtpServer = $t.Server
                    ErrorAction = 'Stop'
                }

                # ADDED ONLY WHEN THERE ARE SOME. -Attachments @() is not the
                # same as leaving it out: an empty array still binds the
                # parameter and Send-MailMessage rejects it.
                if (@($attach).Count -gt 0) { $params['Attachments'] = @($attach) }

                # Send-MailMessage is marked obsolete by Microsoft and is still
                # the right call here: it is present in Windows PowerShell 5.1
                # with no module to install, and this is a connectivity test
                # rather than a mail-sending feature.
                Send-MailMessage @params
                $ok = $true
                $detail = 'Accepted by the server.'
            } catch {
                $detail = "$($_.Exception.Message)"
                if ($_.Exception.InnerException) { $detail += " -- $($_.Exception.InnerException.Message)" }
            }
            $sw.Stop()

            [void]$script:Results.Add([PSCustomObject]@{
                Server = $t.Server; Kind = $t.Kind; To = $t.To
                Result = $(if ($ok) { 'Accepted' } else { 'FAILED' })
                Ms     = [int]$sw.ElapsedMilliseconds
                Detail = $detail
            })
            $pb.Value = [Math]::Min($pb.Maximum, $pb.Value + 1)
        }

        Update-ResultGrid
        Write-Summary
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail test'
        Write-FailedLog -Action 'Report' -Target (Get-LogTarget 'On-Premises Mail Test') `
            -Detail "$($_.Exception.Message)" -Environment $Environment
    } finally {
        $pb.Visible = $false
        $btnSend.Enabled = $true
        $lblProg.Text = ''
    }
}

function Update-ResultGrid {
    $t = New-Object System.Data.DataTable
    foreach ($c in @('Server', 'Kind', 'To', 'Result')) { [void]$t.Columns.Add($c) }
    # Typed, so it sorts as a number rather than as text.
    [void]$t.Columns.Add('Ms', [int])
    [void]$t.Columns.Add('Detail')
    foreach ($r in $script:Results) {
        [void]$t.Rows.Add($r.Server, $r.Kind, $r.To, $r.Result, [int]$r.Ms, $r.Detail)
    }
    $grid.DataSource = $t
    if ($grid.Columns.Contains('Ms')) {
        $grid.Columns['Ms'].DefaultCellStyle.Format = 'N0'
        $grid.Columns['Ms'].DefaultCellStyle.Alignment = 'MiddleRight'
    }
    foreach ($gr in $grid.Rows) {
        if ("$($gr.Cells['Result'].Value)" -eq 'FAILED') { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
        else { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::ForestGreen }
    }
}

function Write-Summary {
    $ok   = @($script:Results | Where-Object { $_.Result -eq 'Accepted' })
    $bad  = @($script:Results | Where-Object { $_.Result -ne 'Accepted' })

    $L = New-Object System.Collections.ArrayList
    [void]$L.Add("$(@($ok).Count) accepted, $(@($bad).Count) failed, at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss').")
    if (@($script:Attachments).Count -gt 0) {
        $at = Get-AttachmentTotals
        [void]$L.Add("Each message carried $($at.Count) attachment(s), $(Format-AttachSize -Bytes $at.Raw) on disk (~$(Format-AttachSize -Bytes $at.Encoded) encoded).")
    }
    [void]$L.Add('')

    if (@($bad).Count -gt 0) {
        [void]$L.Add('FAILURES')
        foreach ($b in $bad) { [void]$L.Add("   $($b.Server) [$($b.Kind)] -> $($b.To)"); [void]$L.Add("      $($b.Detail)") }
        [void]$L.Add('')

        # A PATTERN IN THE FAILURES IS THE ANSWER, and it is easy to miss by
        # reading rows one at a time.
        $extBad = @($bad | Where-Object { $_.Kind -eq 'External' })
        $intBad = @($bad | Where-Object { $_.Kind -eq 'Internal' })
        if (@($extBad).Count -gt 0 -and @($intBad).Count -eq 0) {
            [void]$L.Add('EVERY INTERNAL TEST PASSED AND EXTERNAL FAILED. The servers are healthy and')
            [void]$L.Add('accepting mail -- the problem is outbound: the send connector, the smart')
            [void]$L.Add('host, or a firewall rule in front of it. Do not go looking at the servers.')
            [void]$L.Add('')
        }
        # A SIZE REJECTION IS NOT A MAIL FLOW FAULT and reads exactly like one.
        # Matched on what the servers actually say: 5.3.4 and SMTP 552 are the
        # size codes, and the wording varies by version and by receiving org.
        if (@($script:Attachments).Count -gt 0) {
            $sizeBad = @($bad | Where-Object {
                "$($_.Detail)" -match '(?i)5\.3\.4|\b552\b|message.{0,12}too (large|big)|size limit|exceeds the (maximum|size)'
            })
            if (@($sizeBad).Count -gt 0) {
                $tot = Get-AttachmentTotals
                [void]$L.Add('THIS LOOKS LIKE A SIZE LIMIT, NOT A ROUTING PROBLEM.')
                [void]$L.Add('')
                [void]$L.Add("   $(@($sizeBad).Count) failure(s) mention a size or a 5.3.4 / 552 code.")
                [void]$L.Add("   Attached: $($tot.Count) file(s), $(Format-AttachSize -Bytes $tot.Raw) on disk,")
                [void]$L.Add("   about $(Format-AttachSize -Bytes $tot.Encoded) once base64-encoded -- and the ENCODED size is")
                [void]$L.Add('   what the limit is measured against.')
                [void]$L.Add('')
                [void]$L.Add('   Check the limit that actually applied: the receive connector on the hub,')
                [void]$L.Add('   the send connector, the organisation transport config, or the receiving')
                [void]$L.Add('   side. Send again with no attachment to prove the route itself is fine.')
                [void]$L.Add('')
            }
        }

        $byServer = @($bad | Group-Object Server)
        if (@($byServer).Count -eq 1 -and @($script:Servers).Count -gt 1) {
            [void]$L.Add("ONLY $($byServer[0].Name) FAILED. The other servers took the same message, so")
            [void]$L.Add('this is that one server rather than mail flow generally. Compare its receive')
            [void]$L.Add('connectors and certificate against a server that worked.')
            [void]$L.Add('')
        }
    }

    [void]$L.Add('ACCEPTED IS NOT DELIVERED.')
    [void]$L.Add('')
    [void]$L.Add('A pass means the server took the message, not that it arrived. Check the')
    [void]$L.Add('mailboxes as well -- a message can be accepted and then dropped by a transport')
    [void]$L.Add('rule, a spam filter, or the receiving organisation.')
    [void]$L.Add('')
    [void]$L.Add('The server name is in each subject line, so what arrives tells you which')
    [void]$L.Add('route worked without coming back to this grid.')
    $txtOut.Text = ($L -join [Environment]::NewLine)

    # THE ATTACHMENT COUNT GOES IN THE AUDIT LINE. Two runs an hour apart with
    # different results are only comparable if the record says what was in the
    # message -- "it worked yesterday" is usually this.
    $attNote = ''
    if (@($script:Attachments).Count -gt 0) {
        $at = Get-AttachmentTotals
        $attNote = " $($at.Count) attachment(s), $(Format-AttachSize -Bytes $at.Raw)."
    }
    Write-CompletedLog -Action 'Report' -Target (Get-LogTarget 'On-Premises Mail Test') `
        -Result $(if (@($bad).Count -gt 0) { 'Partial' } else { 'Success' }) `
        -Detail "Sent $(@($script:Results).Count) test message(s) across $(@($script:Servers).Count) server(s). $(@($ok).Count) accepted, $(@($bad).Count) failed.$attNote" `
        -Environment $Environment
}
#endregion

$btnSend.Add_Click({ try { Invoke-MailTest } catch { Show-HandlerError -ErrorRecord $_ -Context 'Send tests' } })

$btnAttAdd.Add_Click({    try { Add-Attachment }            catch { Show-HandlerError -ErrorRecord $_ -Context 'Add attachments' } })
$btnAttRemove.Add_Click({ try { Remove-SelectedAttachment } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove attachment' } })
$btnAttClear.Add_Click({  try { Clear-Attachment }          catch { Show-HandlerError -ErrorRecord $_ -Context 'Clear attachments' } })

# Double-click a row to drop it, which is what the list invites.
$lstAtt.Add_DoubleClick({ try { Remove-SelectedAttachment } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove attachment' } })

# THE MESSAGE COUNT IN THE SIZE LINE COMES FROM THESE TWO BOXES, so it goes
# stale the moment either is edited. Cheap to keep honest -- the hint is a label
# update, not a list rebuild.
$txtInternal.Add_TextChanged({ try { Update-AttachmentHint } catch { } })
$txtExternal.Add_TextChanged({ try { Update-AttachmentHint } catch { } })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'On-Premises Mail Test' | Out-Null

$form.Add_Shown({
    try {
        if ($SmtpAddress) { $txtFrom.Text = "$SmtpAddress" }
        $txtSubject.Text = "Mail flow test $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        $txtBody.Text    = 'Test message from the Admin Console mail flow test.'
        # Draws the empty-state hint rather than leaving a blank label that
        # looks like a control that failed to load.
        Update-AttachmentList
        $txtOut.Text = "THIS SENDS REAL EMAIL." + [Environment]::NewLine +
            [Environment]::NewLine +
            "Servers: $($script:Servers -join ', ')" + [Environment]::NewLine +
            "Environment: $(if ($Environment) { $Environment } else { '(none passed)' })" + [Environment]::NewLine +
            [Environment]::NewLine +
            'Fill in From and at least one recipient. Either recipient box on its own is' + [Environment]::NewLine +
            'a valid test; filling both sends two separate messages per server so internal' + [Environment]::NewLine +
            'and external routing are reported independently.'
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening On-Premises Mail Test' }
})

[void]$form.ShowDialog()
