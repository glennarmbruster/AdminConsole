#Requires -Version 5.1
<#
    Unified-Trace.ps1
    One message trace across Exchange Online, on-premises, Proofpoint and
    SendGrid.

    THE EXCHANGE HALVES ARE REAL. The Proofpoint and SendGrid halves are wired
    to the credential layer and STOP THERE, because no keys exist yet and their
    API surfaces are not confirmed. That is deliberate: the plumbing is written
    now so the tools are not retrofitted later.

    THE MOCKUP IS Unified-Trace-Mockup.ps1 and is a separate tool. It shows what
    the finished view looks like with invented data. This one shows real data
    and is honest about the half it cannot reach yet.

    THIS TOOL MUST NOT REQUEST ITS OWN PROCESS.

    A tool in its own process gets its own global scope, so the API credential
    cache cannot be shared AND it starts with no Exchange Online session --
    every trace then reports "no command named Get-MessageTraceV2".

    THE DIRECTIVE IS NOT SPELLED OUT HERE ON PURPOSE. Writing it in full, even
    inside a comment saying not to use it, is enough for the console scanner to
    match -- which is exactly what happened, and why this tool opened in a new
    session on its first run.

    -Environment IS ACCEPTED AND USED, and passed to every Write-*Log call.
#>
# TOOLKIT-HELP-RUNS: Reads Exchange Online and on-premises message traces. Prompts for API credentials on first use. Read-only.
# TOOLKIT-VARIANT: Hybrid
# TOOLKIT-HELP-BEGIN
# One search across every system a message crosses.
#
# WHAT WORKS TODAY:
#   Exchange Online   real, via Get-MessageTraceV2
#   On-premises       real, via Get-MessageTrackingLog across the hub servers
#
# WHAT DOES NOT WORK YET:
#   Proofpoint        credential prompt works, the API call is not built
#   SendGrid          credential prompt works, the API call is not built
#
# THAT IS NOT A BUG AND THE TOOL SAYS SO RATHER THAN RETURNING NOTHING. There
# are no API keys yet. When they arrive the credential half is already done and
# only the query needs writing.
#
# CREDENTIALS ARE ASKED FOR ONCE PER SESSION, and only for the source you
# actually select. They are held in memory and never written to disk. Settings
# has a "Forget API credentials" item that clears them.
#
# For a demonstration of the finished view with invented data, use
# Unified Trace (Mockup).
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

$script:Rows = New-Object System.Collections.ArrayList
$script:Notes = New-Object System.Collections.ArrayList

# Exchange Online accepts at most 10 days per query; the on-premises logs have
# no such limit. Same rule as Hybrid Message Trace.
$script:CloudMaxDays = 10

#region Sources -----------------------------------------------------------------
function Add-TraceRow {
    param($When, [string]$Source, [string]$SenderAddress, [string]$Recipient,
          [string]$Subject, [string]$Outcome, [string]$Detail, [long]$SizeBytes = 0)
    [void]$script:Rows.Add([PSCustomObject]@{
        When = $When; Source = $Source; Sender = $SenderAddress; Recipient = $Recipient
        Subject = $Subject; Outcome = $Outcome; Detail = $Detail; SizeBytes = $SizeBytes
    })
}

function Get-ExoRows {
    param([string]$From, [string]$To, [datetime]$Start, [datetime]$End)
    $p = @{ StartDate = $Start; EndDate = $End; ResultSize = 5000
            WarningAction = 'SilentlyContinue'; ErrorAction = 'Stop' }
    if ($From) { $p.SenderAddress = $From }
    if ($To)   { $p.RecipientAddress = $To }
    foreach ($m in @(Invoke-EXOCommand -Name 'Get-MessageTraceV2' -Parameters $p)) {
        Add-TraceRow -When $m.Received -Source 'Exchange Online' -SenderAddress "$($m.SenderAddress)" `
            -Recipient "$($m.RecipientAddress)" -Subject "$($m.Subject)" -Outcome "$($m.Status)" `
            -Detail '' -SizeBytes ([long]0 + $m.Size)
    }
}

function Get-OnPremRows {
    param([string]$From, [string]$To, [datetime]$Start, [datetime]$End)
    # All four hub servers -- see AGENTS.md. Get-MessageTrackingLog has no
    # organisation-wide view, so it is one query per server.
    # From Tenant.json (Get-TenantExchangeServerList), not a list in this file.
    foreach ($srv in @(Get-TenantExchangeServerList)) {
        try {
            $p = @{ Server = $srv; Start = $Start; End = $End
                    ResultSize = 5000; WarningAction = 'SilentlyContinue'; ErrorAction = 'Stop' }
            if ($From) { $p.Sender = $From }
            foreach ($e in @(Invoke-EmsCommand -Name 'Get-MessageTrackingLog' -Parameters $p)) {
                $rcp = (@($e.Recipients) -join ', ')
                if ($To -and $rcp -notlike "*$To*") { continue }
                Add-TraceRow -When $e.Timestamp -Source 'On-premises' -SenderAddress "$($e.Sender)" `
                    -Recipient $rcp -Subject "$($e.MessageSubject)" -Outcome "$($e.EventId)" `
                    -Detail "$($e.ServerHostname)" -SizeBytes ([long]0 + $e.TotalBytes)
            }
        } catch {
            [void]$script:Notes.Add("On-premises $srv did not answer: $($_.Exception.Message)")
        }
    }
}

function Get-PendingApiRows {
    <#
        Acquire the credential, then stop.

        THE CREDENTIAL HALF IS REAL AND THE QUERY HALF IS NOT. That is the whole
        point of this task -- the plumbing exists so that when a key arrives it
        is pasted into a working tool rather than starting a build.

        NO CALL IS MADE TO THE PROOFPOINT POD API. Its authentication mechanism
        is unconfirmed and a support ticket is open; guessing at it here would
        produce code nobody can check. See UNVERIFIED.md.
    #>
    param([Parameter(Mandatory)][string]$Service, [Parameter(Mandatory)][string]$Label)

    $cred = Get-ApiSecret -Service $Service -Purpose 'Unified Trace'
    if (-not $cred.Ok) {
        # CANCEL IS A NORMAL ANSWER, not an error. No dialog, no stack trace,
        # and emphatically not proceeding with an empty credential and letting
        # the API answer 401.
        [void]$script:Notes.Add("$Label -- $($cred.Reason)")
        return
    }

    if ($Service -eq 'ProofpointSer') {
        # THE TOKEN EXCHANGE IS REAL, because it is documented and checkable.
        # What would be done with the token is not built.
        $tok = Get-SerAccessToken
        if (-not $tok.Ok) { [void]$script:Notes.Add("$Label -- $($tok.Reason)"); return }
        [void]$script:Notes.Add("$Label -- credential accepted and a bearer token was obtained (expires $($tok.ExpiresAt.ToString('HH:mm:ss'))). The reporting query itself is not built yet.")
        return
    }

    [void]$script:Notes.Add("$Label -- credential is loaded for this session. The API query is not built yet: there is no confirmed endpoint to write against. Nothing was contacted.")
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Unified Trace'
$form.ClientSize      = New-Object System.Drawing.Size(1280, 800)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$grpFind          = New-Object System.Windows.Forms.GroupBox
$grpFind.Text     = 'Find'
$grpFind.Location = New-Object System.Drawing.Point(14, 46)
$grpFind.Size     = New-Object System.Drawing.Size(1200, 140)
$grpFind.Anchor   = 'Top,Left,Right'
[void]$form.Controls.Add($grpFind)

$lblFrom          = New-Object System.Windows.Forms.Label
$lblFrom.Text     = 'Sender'
$lblFrom.Location = New-Object System.Drawing.Point(12, 29)
$lblFrom.Size     = New-Object System.Drawing.Size(70, 20)
[void]$grpFind.Controls.Add($lblFrom)

$txtFrom          = New-Object System.Windows.Forms.TextBox
$txtFrom.Location = New-Object System.Drawing.Point(88, 26)
$txtFrom.Size     = New-Object System.Drawing.Size(300, 22)
[void]$grpFind.Controls.Add($txtFrom)

$lblTo          = New-Object System.Windows.Forms.Label
$lblTo.Text     = 'Recipient'
$lblTo.Location = New-Object System.Drawing.Point(400, 29)
$lblTo.Size     = New-Object System.Drawing.Size(70, 20)
[void]$grpFind.Controls.Add($lblTo)

$txtTo          = New-Object System.Windows.Forms.TextBox
$txtTo.Location = New-Object System.Drawing.Point(476, 26)
$txtTo.Size     = New-Object System.Drawing.Size(300, 22)
[void]$grpFind.Controls.Add($txtTo)

$lblStart          = New-Object System.Windows.Forms.Label
$lblStart.Text     = 'From'
$lblStart.Location = New-Object System.Drawing.Point(12, 61)
$lblStart.Size     = New-Object System.Drawing.Size(70, 20)
[void]$grpFind.Controls.Add($lblStart)

$dtStart              = New-Object System.Windows.Forms.DateTimePicker
$dtStart.Format       = 'Custom'
$dtStart.CustomFormat = 'yyyy-MM-dd HH:mm'
$dtStart.Location     = New-Object System.Drawing.Point(88, 58)
$dtStart.Size         = New-Object System.Drawing.Size(180, 22)
$dtStart.Value        = (Get-Date).AddDays(-1)
[void]$grpFind.Controls.Add($dtStart)

$lblEnd          = New-Object System.Windows.Forms.Label
$lblEnd.Text     = 'To'
$lblEnd.Location = New-Object System.Drawing.Point(284, 61)
$lblEnd.Size     = New-Object System.Drawing.Size(30, 20)
[void]$grpFind.Controls.Add($lblEnd)

$dtEnd              = New-Object System.Windows.Forms.DateTimePicker
$dtEnd.Format       = 'Custom'
$dtEnd.CustomFormat = 'yyyy-MM-dd HH:mm'
$dtEnd.Location     = New-Object System.Drawing.Point(320, 58)
$dtEnd.Size         = New-Object System.Drawing.Size(180, 22)
$dtEnd.Value        = (Get-Date)
[void]$grpFind.Controls.Add($dtEnd)

$lblSrc          = New-Object System.Windows.Forms.Label
$lblSrc.Text     = 'Sources'
$lblSrc.Location = New-Object System.Drawing.Point(12, 97)
$lblSrc.Size     = New-Object System.Drawing.Size(70, 20)
[void]$grpFind.Controls.Add($lblSrc)

# CHECKBOXES, NOT RADIOS. A unified trace is several sources at once -- that is
# the entire premise -- so the selection is a set, not a choice.
$chkExo          = New-Object System.Windows.Forms.CheckBox
$chkExo.Text     = 'Exchange Online'
$chkExo.Location = New-Object System.Drawing.Point(88, 96)
$chkExo.Size     = New-Object System.Drawing.Size(130, 22)
$chkExo.Checked  = $true
[void]$grpFind.Controls.Add($chkExo)

$chkOnPrem          = New-Object System.Windows.Forms.CheckBox
$chkOnPrem.Text     = 'On-premises'
$chkOnPrem.Location = New-Object System.Drawing.Point(224, 96)
$chkOnPrem.Size     = New-Object System.Drawing.Size(110, 22)
$chkOnPrem.Checked  = $true
[void]$grpFind.Controls.Add($chkOnPrem)

$chkPod          = New-Object System.Windows.Forms.CheckBox
$chkPod.Text     = 'Proofpoint'
$chkPod.Location = New-Object System.Drawing.Point(340, 96)
$chkPod.Size     = New-Object System.Drawing.Size(100, 22)
[void]$grpFind.Controls.Add($chkPod)

$chkSer          = New-Object System.Windows.Forms.CheckBox
$chkSer.Text     = 'SER'
$chkSer.Location = New-Object System.Drawing.Point(446, 96)
$chkSer.Size     = New-Object System.Drawing.Size(60, 22)
[void]$grpFind.Controls.Add($chkSer)

$chkSg          = New-Object System.Windows.Forms.CheckBox
$chkSg.Text     = 'SendGrid'
$chkSg.Location = New-Object System.Drawing.Point(512, 96)
$chkSg.Size     = New-Object System.Drawing.Size(90, 22)
[void]$grpFind.Controls.Add($chkSg)

# The Proofpoint/SER/SendGrid caveat was a label here. Removed -- it is in the
# help and in the result summary, and a third copy on the form was noise.

$btnGo           = New-Object System.Windows.Forms.Button
$btnGo.Text      = 'Trace'
$btnGo.Location  = New-Object System.Drawing.Point(800, 24)
$btnGo.Size      = New-Object System.Drawing.Size(130, 28)
$btnGo.FlatStyle = 'Flat'
[void]$grpFind.Controls.Add($btnGo)

$btnForget           = New-Object System.Windows.Forms.Button
$btnForget.Text      = 'Forget credentials'
$btnForget.Location  = New-Object System.Drawing.Point(940, 24)
$btnForget.Size      = New-Object System.Drawing.Size(140, 28)
$btnForget.FlatStyle = 'Flat'
[void]$grpFind.Controls.Add($btnForget)

$lblProg           = New-Object System.Windows.Forms.Label
$lblProg.Location  = New-Object System.Drawing.Point(800, 58)
$lblProg.Size      = New-Object System.Drawing.Size(388, 20)
$lblProg.ForeColor = [System.Drawing.Color]::Gray
[void]$grpFind.Controls.Add($lblProg)

$pb = Add-ReportProgressBar -Form $form -X 14 -Y 192 -Width 1252 -Height 12

$grid                     = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(14, 212)
$grid.Size                = New-Object System.Drawing.Size(1252, 360)
$grid.ReadOnly            = $true
$grid.AllowUserToAddRows  = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode       = 'FullRowSelect'
$grid.MultiSelect         = $false
$grid.Anchor              = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($grid)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(14, 582)
$txtOut.Size       = New-Object System.Drawing.Size(1252, 168)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(1146, 760)
$btnClose.Size      = New-Object System.Drawing.Size(120, 28)
$btnClose.FlatStyle = 'Flat'
$btnClose.Anchor    = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
[void]$form.Controls.Add($btnClose)
#endregion

#region Run ---------------------------------------------------------------------
function Invoke-UnifiedTrace {
    $from = "$($txtFrom.Text)".Trim()
    $to   = "$($txtTo.Text)".Trim()
    if (-not $from -and -not $to) { $txtOut.Text = 'Give a sender or a recipient.'; return }

    $start = $dtStart.Value
    $end   = $dtEnd.Value
    if ($end -le $start) { $txtOut.Text = 'The end of the range is not after the start.'; return }

    # Exchange Online refuses more than 10 days per query -- a hard cmdlet
    # limit, not throttling, so a wider range is an error rather than a slower
    # answer.
    if ($chkExo.Checked -and ($end - $start).TotalDays -gt $script:CloudMaxDays) {
        $txtOut.Text = "Exchange Online accepts at most $($script:CloudMaxDays) days per search. Narrow the range, or untick Exchange Online."
        return
    }

    $script:Rows.Clear()
    $script:Notes.Clear()
    $btnGo.Enabled = $false
    $pb.Maximum = 5; $pb.Value = 0; $pb.Visible = $true

    try {
        if ($chkExo.Checked) {
            $lblProg.Text = 'Exchange Online ...'; $form.Refresh()
            try { Get-ExoRows -From $from -To $to -Start $start -End $end }
            catch { [void]$script:Notes.Add("Exchange Online failed: $($_.Exception.Message)") }
        }
        $pb.Value = 1

        if ($chkOnPrem.Checked) {
            $lblProg.Text = 'On-premises transport logs ...'; $form.Refresh()
            try { Get-OnPremRows -From $from -To $to -Start $start -End $end }
            catch { [void]$script:Notes.Add("On-premises failed: $($_.Exception.Message)") }
        }
        $pb.Value = 2

        if ($chkPod.Checked) { $lblProg.Text = 'Proofpoint ...'; $form.Refresh()
            Get-PendingApiRows -Service 'ProofpointPod' -Label 'Proofpoint' }
        $pb.Value = 3

        if ($chkSer.Checked) { $lblProg.Text = 'SER ...'; $form.Refresh()
            Get-PendingApiRows -Service 'ProofpointSer' -Label 'SER' }
        $pb.Value = 4

        if ($chkSg.Checked) { $lblProg.Text = 'SendGrid ...'; $form.Refresh()
            Get-PendingApiRows -Service 'SendGrid' -Label 'SendGrid' }
        $pb.Value = 5

        Update-TraceGrid
        Write-TraceSummary -From $from -To $to
    } catch {
        $txtOut.Text = Get-SafeApiError -ErrorRecord $_ -Context 'Unified trace'
        Write-FailedLog -Action 'Report' -Target (Get-LogTarget 'Unified Trace') `
            -Detail "$($_.Exception.Message)" -Environment $Environment
    } finally {
        $pb.Visible = $false
        $btnGo.Enabled = $true
        $lblProg.Text = ''
    }
}

function Update-TraceGrid {
    $t = New-Object System.Data.DataTable
    foreach ($c in @('Time', 'Source', 'Sender', 'Recipient', 'Subject', 'Outcome', 'Detail')) { [void]$t.Columns.Add($c) }
    [void]$t.Columns.Add('Size', [long])
    foreach ($r in @($script:Rows | Sort-Object When -Descending)) {
        $w = ''
        try { $w = ([datetime]$r.When).ToString('yyyy-MM-dd HH:mm:ss') } catch { }
        [void]$t.Rows.Add($w, $r.Source, $r.Sender, $r.Recipient, $r.Subject, $r.Outcome, $r.Detail, [long]$r.SizeBytes)
    }
    $grid.DataSource = $t
    if ($grid.Columns.Contains('Size')) {
        $grid.Columns['Size'].DefaultCellStyle.Format = 'N0'
        $grid.Columns['Size'].DefaultCellStyle.Alignment = 'MiddleRight'
    }
    foreach ($gr in $grid.Rows) {
        $o = "$($gr.Cells['Outcome'].Value)"
        if ($o -match '(?i)fail|reject|quarantin') { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
        elseif ($o -match '(?i)defer|pending')     { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGoldenrod }
    }
}

function Write-TraceSummary {
    param([string]$From, [string]$To)
    $L = New-Object System.Collections.ArrayList
    [void]$L.Add("$(@($script:Rows).Count) row(s) at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss').")
    foreach ($g in @($script:Rows | Group-Object Source | Sort-Object Name)) {
        [void]$L.Add(("   {0,-18} {1,5} row(s)" -f $g.Name, @($g.Group).Count))
    }
    if (@($script:Notes).Count -gt 0) {
        [void]$L.Add('')
        [void]$L.Add('NOTES -- read these before treating the result as complete:')
        foreach ($n in $script:Notes) { [void]$L.Add("   $n") }
    }
    [void]$L.Add('')
    [void]$L.Add('THE PROOFPOINT AND SENDGRID HALVES ARE NOT QUERIED YET. There are no API')
    [void]$L.Add('keys, and their endpoints are not confirmed. Ticking them proves the')
    [void]$L.Add('credential path works; it does not add rows. Unified Trace (Mockup) shows')
    [void]$L.Add('what the finished view looks like.')
    $txtOut.Text = ($L -join [Environment]::NewLine)

    # THE AUDIT ROW RECORDS TOOL, OPERATOR, TARGET AND RESULT. What
    # authenticated any call is not part of a question the audit log answers.
    Write-CompletedLog -Action 'Report' -Target (Get-LogTarget $(if ($To) { $To } else { $From }) 'Unified Trace') -Result 'Success' `
        -Detail "Traced $(@($script:Rows).Count) row(s) across the selected sources." -Environment $Environment
}
#endregion

$btnGo.Add_Click({ try { Invoke-UnifiedTrace } catch { Show-HandlerError -ErrorRecord $_ -Context 'Trace' } })
$btnForget.Add_Click({
    try {
        $n = Clear-ApiSecret
        $txtOut.Text = "$n credential(s) cleared. Any cached bearer token was cleared with them." + [Environment]::NewLine +
            'The next source that needs one will ask again.'
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Forget credentials' }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Unified Trace' | Out-Null

$form.Add_Shown({
    try {
        if ($SmtpAddress) { $txtTo.Text = "$SmtpAddress" }
        $state = @(Get-ApiSecretState | Where-Object { $_.Loaded } | ForEach-Object { $_.Label })
        $txtOut.Text = 'Enter a sender or a recipient, choose the sources, press Trace.' + [Environment]::NewLine +
            [Environment]::NewLine +
            "Environment: $(if ($Environment) { $Environment } else { '(none passed)' })" + [Environment]::NewLine +
            "Credentials loaded this session: $(if (@($state).Count) { $state -join ', ' } else { 'none' })" + [Environment]::NewLine +
            [Environment]::NewLine +
            'Exchange Online and on-premises return real data. Proofpoint, SER and' + [Environment]::NewLine +
            'SendGrid prompt for a credential and stop there -- no keys exist yet and' + [Environment]::NewLine +
            'their endpoints are not confirmed, so nothing is guessed at.'
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Unified Trace' }
})

[void]$form.ShowDialog()
