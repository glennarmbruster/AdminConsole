#Requires -Version 5.1
<#
    Get-Queues.ps1
    Every transport queue on the four hub servers, on one screen.

    -Environment IS ACCEPTED AND USED, and passed to every Write-*Log call.
#>
# TOOLKIT-HELP-RUNS: Reads transport queues from four on-premises Exchange servers. Read-only.
# TOOLKIT-VARIANT: OnPrem
# TOOLKIT-HELP-BEGIN
# Every transport queue on the four hub servers, on one screen.
#
# WHY THIS EXISTS. Get-Queue has no organisation-wide view -- it is one call per
# server, and checking a mail-flow complaint means four logons or four windows.
# By the time the fourth is open the first is out of date, so a queue that is
# draining looks identical to one that is stuck.
#
# WHAT MATTERS, IN ORDER:
#   Status Retry      the queue cannot reach the next hop. This is the one.
#   Status Suspended  somebody stopped it by hand and may have forgotten.
#   A large MessageCount that is NOT falling between refreshes.
#
# A BIG QUEUE IS NOT AUTOMATICALLY A PROBLEM. A queue of several hundred that is
# shrinking is a queue doing its job. The number on its own tells you nothing --
# press Refresh again and compare. That is why the count from the previous run
# is shown beside the current one.
#
# THE SUBMISSION AND SHADOW QUEUES ARE NORMAL. Shadow queues hold copies for
# redundancy and sit at a steady non-zero count on a healthy server. Do not
# chase them.
#
# A SERVER THAT DID NOT ANSWER IS REPORTED, NOT SKIPPED. A silent omission here
# would read as "that server has no queues", which is the opposite of the truth
# when it is the one that is down.
#
# Read-only. It reports and changes nothing -- it cannot retry, suspend or
# delete a queue.
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

# SERVERS COME FROM Tenant.json, NEVER FROM THIS FILE.
#
# This list used to be four hostnames written here as literals -- in this
# tool, in Check Queues and in Exchange Health Check, three copies of the same
# fact. That breaks the rule the whole toolkit rests on (Tenant.json and
# Settings.json are the only files allowed to name anything belonging to an
# organisation), and a toolkit handed to another tenant reached for servers
# that do not exist there, with an admin-share timeout as the only clue.
#
# Get-TenantExchangeServerList reads OnPremExchange.Servers from Tenant.json,
# falls back to asking the directory with Get-ExchangeServer when an
# on-premises session is open, and subtracts OnPremExchange.ExcludedServers
# either way. Empty is a real answer and the window says where to set it,
# because an empty list must not read as "nothing to do".
$script:Servers = @()
try { $script:Servers = @(Get-TenantExchangeServerList) } catch { $script:Servers = @() }

# AN EMPTY SERVER LIST IS AN ANSWER, NOT A QUIET NO-OP. With no Servers in
# Tenant.json and no on-premises session, every scan below would sweep zero
# servers, find zero files and report success -- which reads as "there is
# nothing to clean" rather than "I was never told where to look".
function Test-ServerListPresent {
    if (@($script:Servers).Count -gt 0) { return $true }
    $null = Show-ToolDialog(
        ("No on-premises Exchange servers are configured, so there is nothing to work against." + [Environment]::NewLine + [Environment]::NewLine +
         "Set OnPremExchange.Servers in Config\Tenant.json (Tenant Manager), or connect OnPrem and this tool will ask the directory itself." + [Environment]::NewLine + [Environment]::NewLine +
         "OnPremExchange.ExcludedServers is subtracted from either list."),
        'No servers configured', 'OK', 'Warning')
    return $false
}

$script:Rows      = New-Object System.Collections.ArrayList
$script:Failed    = New-Object System.Collections.ArrayList
# Counts from the previous run, so a queue can be seen to be draining rather
# than merely large. Keyed on server + queue identity.
$script:LastCount = @{}
$script:LastRun   = $null

#region Query -------------------------------------------------------------------
function Get-QueueRows {
    <#
        One server's queues.

        NO -ResultSize. Get-Queue takes a Server and returns that server's
        queues; there is no truncation to be explicit about. See the note in
        CertificateExpiry.ps1 -- a convention is not a cmdlet contract, and
        adding a parameter the cmdlet does not declare is a hard failure at the
        first call rather than a warning.
    #>
    param([Parameter(Mandatory)][string]$Server)

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Rows = @() }
    try {
        $r.Rows = @(Invoke-EmsCommand -Name 'Get-Queue' -Parameters @{
            Server      = $Server
            ErrorAction = 'Stop'
        })
        $r.Ok = $true
    } catch {
        $r.Reason = Format-ErrorDetail -ErrorRecord $_ -Context "Get-Queue -Server $Server"
    }
    return $r
}

function Get-QueueSeverity {
    <#
        How much attention a queue deserves. Higher sorts first.

        STATUS BEATS SIZE. A retrying queue of 3 is a broken route; a delivering
        queue of 3000 is a busy evening. Sorting by MessageCount would put the
        second at the top and bury the first.
    #>
    param($Queue)

    $status = "$($Queue.Status)"
    $count  = [int]0 + $Queue.MessageCount

    if ($status -match '(?i)retry')     { return 100 }
    if ($status -match '(?i)suspend')   { return 90 }
    if ("$($Queue.LastError)".Trim())   { return 80 }
    if ($count -ge 500)                 { return 60 }
    if ($count -ge 100)                 { return 40 }
    if ($status -match '(?i)connecting'){ return 30 }
    return 10
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Check Queues'
$form.ClientSize      = New-Object System.Drawing.Size(1200, 1116)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$btnRun           = New-Object System.Windows.Forms.Button
$btnRun.Text      = 'Refresh'
$btnRun.Location  = New-Object System.Drawing.Point(14, 14)
$btnRun.Size      = New-Object System.Drawing.Size(120, 28)
$btnRun.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnRun)

$chkBusy           = New-Object System.Windows.Forms.CheckBox
$chkBusy.Text      = 'Only queues needing attention'
$chkBusy.Location  = New-Object System.Drawing.Point(148, 19)
$chkBusy.Size      = New-Object System.Drawing.Size(230, 20)
[void]$form.Controls.Add($chkBusy)

$lblProg          = New-Object System.Windows.Forms.Label
$lblProg.Location = New-Object System.Drawing.Point(392, 19)
$lblProg.Size     = New-Object System.Drawing.Size(500, 20)
$lblProg.ForeColor = [System.Drawing.Color]::Gray
[void]$form.Controls.Add($lblProg)

$btnExport           = New-Object System.Windows.Forms.Button
$btnExport.Text      = 'Export CSV'
$btnExport.Location  = New-Object System.Drawing.Point(1000, 14)
$btnExport.Size      = New-Object System.Drawing.Size(120, 28)
$btnExport.FlatStyle = 'Flat'
$btnExport.Anchor    = 'Top,Right'
[void]$form.Controls.Add($btnExport)

$pb = Add-ReportProgressBar -Form $form -X 14 -Y 48 -Width 1172 -Height 12

$grid                     = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(14, 68)
$grid.Size                = New-Object System.Drawing.Size(1172, 380)
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
# B6: a wrapped queue identity or error string is unreadable.
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(14, 458)
$txtOut.Size       = New-Object System.Drawing.Size(1172, 608)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(1066, 1076)
$btnClose.Size      = New-Object System.Drawing.Size(120, 28)
$btnClose.FlatStyle = 'Flat'
$btnClose.Anchor    = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
[void]$form.Controls.Add($btnClose)
#endregion

#region Display -----------------------------------------------------------------
function Update-QueueGrid {
    $shown = @($script:Rows)
    if ($chkBusy.Checked) { $shown = @($shown | Where-Object { $_.Severity -ge 40 }) }
    $shown = @($shown | Sort-Object Severity, MessageCount -Descending)

    $t = New-Object System.Data.DataTable
    foreach ($c in @('Server', 'Queue', 'Type', 'Status')) { [void]$t.Columns.Add($c) }

    # MESSAGES AND WAS ARE TYPED AS INTEGERS, NOT STRINGS.
    #
    # DataTable columns default to String, and the grid sorts a String column
    # as text -- so a reverse sort gave 7, 4, 25, 24, 14, 1, 0. It was ordering
    # on the first character. The counts were formatted with '{0:N0}' BEFORE
    # being added, which turned numbers into strings on the way in.
    #
    # Added raw and formatted at DISPLAY time instead: keeps the thousands
    # separators and sorts properly, which is what the old code only appeared
    # to do.
    [void]$t.Columns.Add('Messages', [int])
    [void]$t.Columns.Add('Was', [int])
    foreach ($c in @('Next hop', 'Last error')) { [void]$t.Columns.Add($c) }

    foreach ($r in $shown) {
        [void]$t.Rows.Add($r.Server, $r.Queue, $r.Type, $r.Status,
            ([int]$r.MessageCount),
            # DBNull, not 0: a queue seen for the first time has no previous
            # count, and 0 would read as "it drained to nothing".
            $(if ($null -eq $r.Previous) { [System.DBNull]::Value } else { [int]$r.Previous }),
            $r.NextHop, $r.LastError)
    }
    $grid.DataSource = $t

    foreach ($c in @('Messages', 'Was')) {
        if (-not $grid.Columns.Contains($c)) { continue }
        $grid.Columns[$c].DefaultCellStyle.Format    = 'N0'
        $grid.Columns[$c].DefaultCellStyle.Alignment = 'MiddleRight'
    }

    foreach ($gr in $grid.Rows) {
        $s = "$($gr.Cells['Status'].Value)"
        if ($s -match '(?i)retry')        { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
        elseif ($s -match '(?i)suspend')  { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGoldenrod }
    }
    return @($shown).Count
}

function Invoke-QueueCheck {
    $script:Rows.Clear()
    $script:Failed.Clear()

    if (-not (Test-ServerListPresent)) { return }
    $pb.Maximum = [Math]::Max(1, @($script:Servers).Count)
    $pb.Value   = 0
    $pb.Visible = $true
    $btnRun.Enabled = $false

    try {
        foreach ($srv in $script:Servers) {
            $lblProg.Text = "Querying $srv ..."
            $form.Refresh()

            $q = Get-QueueRows -Server $srv
            if (-not $q.Ok) {
                [void]$script:Failed.Add([PSCustomObject]@{ Server = $srv; Reason = $q.Reason })
            } else {
                foreach ($item in $q.Rows) {
                    $id   = "$($item.Identity)"
                    $key  = "$srv|$id"
                    $prev = $null
                    if ($script:LastCount.ContainsKey($key)) { $prev = $script:LastCount[$key] }

                    [void]$script:Rows.Add([PSCustomObject]@{
                        Server       = $srv
                        Queue        = $id
                        Type         = "$($item.DeliveryType)"
                        Status       = "$($item.Status)"
                        MessageCount = [int]0 + $item.MessageCount
                        Previous     = $prev
                        NextHop      = "$($item.NextHopDomain)"
                        LastError    = "$($item.LastError)"
                        Severity     = (Get-QueueSeverity -Queue $item)
                    })
                }
            }
            $pb.Value = [Math]::Min($pb.Maximum, $pb.Value + 1)
        }

        # Remember counts for the next refresh, so 'Was' means something.
        $script:LastCount = @{}
        foreach ($r in $script:Rows) { $script:LastCount["$($r.Server)|$($r.Queue)"] = $r.MessageCount }
        $script:LastRun = Get-Date

        $shown = Update-QueueGrid

        $L = New-Object System.Collections.ArrayList
        $total  = (@($script:Rows | Measure-Object MessageCount -Sum).Sum)
        if (-not $total) { $total = 0 }
        $retry  = @($script:Rows | Where-Object { "$($_.Status)" -match '(?i)retry' })
        $susp   = @($script:Rows | Where-Object { "$($_.Status)" -match '(?i)suspend' })

        [void]$L.Add("Checked $(@($script:Servers).Count) server(s) at $($script:LastRun.ToString('yyyy-MM-dd HH:mm:ss')).")
        [void]$L.Add("$(@($script:Rows).Count) queue(s), $('{0:N0}' -f $total) message(s) in total.")
        [void]$L.Add('')

        if (@($retry).Count -gt 0) {
            [void]$L.Add("$(@($retry).Count) QUEUE(S) IN RETRY -- these cannot reach the next hop:")
            foreach ($r in $retry) {
                [void]$L.Add(("   {0,-16} {1,-40} {2,7} msg   {3}" -f $r.Server, $r.Queue, $r.MessageCount, $r.LastError))
            }
            [void]$L.Add('')
        }
        if (@($susp).Count -gt 0) {
            [void]$L.Add("$(@($susp).Count) QUEUE(S) SUSPENDED -- somebody stopped these by hand:")
            foreach ($r in $susp) {
                [void]$L.Add(("   {0,-16} {1,-40} {2,7} msg" -f $r.Server, $r.Queue, $r.MessageCount))
            }
            [void]$L.Add('')
        }
        if (@($retry).Count -eq 0 -and @($susp).Count -eq 0) {
            [void]$L.Add('No queue is retrying and none is suspended.')
            [void]$L.Add('')
            [void]$L.Add('A LARGE COUNT ON ITS OWN IS NOT A FAULT. Press Refresh again and compare')
            [void]$L.Add('the Messages column against Was -- a falling count is a queue draining')
            [void]$L.Add('normally. A count that does not move between refreshes is the one to look')
            [void]$L.Add('at, even if it is small.')
            [void]$L.Add('')
        }

        # ---- the honesty line ----
        if (@($script:Failed).Count -gt 0) {
            [void]$L.Add('THIS IS A PARTIAL ANSWER.')
            [void]$L.Add('')
            [void]$L.Add('The following server(s) did not respond, so their queues are NOT included')
            [void]$L.Add('above. An empty result for a server that is down looks exactly like a')
            [void]$L.Add('server with nothing queued, which is why they are named rather than')
            [void]$L.Add('quietly skipped:')
            foreach ($f in $script:Failed) { [void]$L.Add("   $($f.Server): $($f.Reason)") }
        }

        $txtOut.Text = ($L -join [Environment]::NewLine)
        $lblProg.Text = "$shown queue(s) shown."

        $result = if (@($script:Failed).Count -gt 0) { 'Partial' } else { 'Success' }
        Write-CompletedLog -Action 'Report' -Target (Get-LogTarget 'Check Queues') -Result $result `
            -Detail "Queried $(@($script:Servers).Count) server(s). $(@($script:Rows).Count) queue(s), $total message(s). $(@($script:Failed).Count) server(s) failed." `
            -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Check Queues'
        Write-FailedLog -Action 'Report' -Target (Get-LogTarget 'Check Queues') `
            -Detail "$($_.Exception.Message)" -Environment $Environment
    } finally {
        $pb.Visible = $false
        $btnRun.Enabled = $true
    }
}

function Show-QueueContents {
    <#
        What is actually sitting in one queue.

        THE CLICKED ROW IS READ FROM ITS CELLS, not from an index into the source
        list. The grid sorts itself, so after a click on a column header the
        display order no longer matches the array and an index lookup returns a
        different queue than the one under the cursor.

        CAPPED AT 200. A stuck queue can hold thousands, and the question being
        asked is "what KIND of thing is stuck" -- which the first 200 answers.
        Pulling all of them freezes the window for no extra information, so the
        cap is stated in the output rather than hidden.
    #>
    param([int]$RowIndex)

    if ($RowIndex -lt 0 -or $RowIndex -ge $grid.Rows.Count) { return }
    $row = $grid.Rows[$RowIndex]
    $srv = "$($row.Cells['Server'].Value)"
    $qid = "$($row.Cells['Queue'].Value)"
    if (-not $qid) { return }

    # The cell is formatted N0, so strip separators before parsing.
    $cnt = 0
    [void][int]::TryParse(("$($row.Cells['Messages'].Value)" -replace '[^\d]', ''), [ref]$cnt)

    if ($cnt -le 0) {
        $txtOut.Text = "$qid on $srv is empty -- there is nothing in it to inspect."
        return
    }

    $lblProg.Text = "Reading $qid ..."
    $form.Refresh()

    $L = New-Object System.Collections.ArrayList
    try {
        # Get-Message DOES declare -ResultSize. EXO002: WarningAction sits with
        # it, because a deliberately limited ResultSize makes Exchange warn
        # "There are more results available" on every single run.
        $msgs = @(Invoke-EmsCommand -Name 'Get-Message' -Parameters @{
            Queue         = $qid
            ResultSize    = 200
            WarningAction = 'SilentlyContinue'
            ErrorAction   = 'Stop'
        })

        [void]$L.Add("$qid on $srv  --  $('{0:N0}' -f $cnt) message(s) queued")
        if (@($msgs).Count -ge 200) {
            [void]$L.Add('Showing the first 200. Enough to see WHAT is stuck; not the whole queue.')
        }
        [void]$L.Add('')

        # WHO IS SENDING IT, FIRST. A queue of 4000 is almost always one sender
        # or one recipient domain, and that grouping names the cause faster than
        # reading any individual message does.
        $bySender = @($msgs | Group-Object { "$($_.FromAddress)" } | Sort-Object Count -Descending | Select-Object -First 5)
        if (@($bySender).Count -gt 0) {
            [void]$L.Add('TOP SENDERS in this sample:')
            foreach ($g in $bySender) {
                [void]$L.Add(("   {0,6}  {1}" -f $g.Count, $(if ($g.Name) { $g.Name } else { '(none)' })))
            }
            [void]$L.Add('')
        }

        [void]$L.Add(('{0,-21} {1,-32} {2,-32} {3,-11} {4}' -f 'Queued', 'From', 'To', 'Status', 'Subject'))
        foreach ($m in @($msgs | Sort-Object DateReceived)) {
            $to = (@($m.Recipients | ForEach-Object { "$($_.Address)" }) -join ', ')
            if ($to.Length -gt 31) { $to = $to.Substring(0, 28) + '...' }
            $fr = "$($m.FromAddress)"
            if ($fr.Length -gt 31) { $fr = $fr.Substring(0, 28) + '...' }
            $dt = ''
            if ($m.DateReceived) { $dt = ([datetime]$m.DateReceived).ToString('yyyy-MM-dd HH:mm:ss') }
            [void]$L.Add(('{0,-21} {1,-32} {2,-32} {3,-11} {4}' -f $dt, $fr, $to, "$($m.Status)", "$($m.Subject)"))
        }

        $lastErr = @($msgs | Where-Object { "$($_.LastError)".Trim() } | Select-Object -First 1)
        if (@($lastErr).Count -gt 0) {
            [void]$L.Add('')
            [void]$L.Add('LAST ERROR reported against a message in this queue:')
            [void]$L.Add("   $($lastErr[0].LastError)")
        }
    } catch {
        [void]$L.Add((Format-ErrorDetail -ErrorRecord $_ -Context "Get-Message -Queue $qid"))
    } finally {
        $lblProg.Text = ''
    }
    $txtOut.Text = ($L -join [Environment]::NewLine)
}

function Export-QueueCsv {
    if (@($script:Rows).Count -eq 0) { $txtOut.Text = 'Nothing to export. Press Refresh first.'; return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV (*.csv)|*.csv'
    $dlg.FileName = "Queues-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    @($script:Rows) | Select-Object Server, Queue, Type, Status, MessageCount, Previous, NextHop, LastError |
        Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -Encoding UTF8
    $txtOut.Text = "Exported $(@($script:Rows).Count) row(s) to $($dlg.FileName)"
}
#endregion

$btnRun.Add_Click({ try { Invoke-QueueCheck } catch { Show-HandlerError -ErrorRecord $_ -Context 'Refresh' } })
$chkBusy.Add_CheckedChanged({ try { [void](Update-QueueGrid) } catch { Show-HandlerError -ErrorRecord $_ -Context 'Filter' } })
$btnExport.Add_Click({ try { Export-QueueCsv } catch { Show-HandlerError -ErrorRecord $_ -Context 'Export' } })

# DOUBLE-CLICK A QUEUE TO SEE WHAT IS IN IT. The result goes to the pane at the
# bottom rather than a new window -- the queue list is the context for what you
# are reading, and a dialog on top of it hides the thing you clicked.
$grid.Add_CellDoubleClick({
    param($sender, $e)
    try { if ($e.RowIndex -ge 0) { Show-QueueContents -RowIndex $e.RowIndex } }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Inspect queue' }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Check Queues' | Out-Null

$form.Add_Shown({
    try {
        $ems = Test-EmsAvailable
        if ($ems.Available) {
            $txtOut.Text = "Press Refresh. Servers: $($script:Servers -join ', ')" + [Environment]::NewLine +
                "Environment: $(if ($Environment) { $Environment } else { '(none passed)' })" + [Environment]::NewLine +
                [Environment]::NewLine +
                'Refresh twice a minute apart to tell a draining queue from a stuck one -- the Was column holds the count from the previous run.' + [Environment]::NewLine +
                'DOUBLE-CLICK any queue with messages in it to see what is sitting there.'
        } else {
            $txtOut.Text = "On-premises Exchange is not available: $($ems.Detail)"
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Check Queues' }
})

[void]$form.ShowDialog()
