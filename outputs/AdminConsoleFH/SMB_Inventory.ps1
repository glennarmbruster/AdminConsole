#Requires -Version 5.1
<#
    SMB_Inventory.ps1
    Every shared mailbox in Exchange Online, with the four facts that decide
    whether it is a safe candidate to source-anchor to Microsoft 365.

    WRITTEN FOR A SPECIFIC JOB: picking a pilot group. The report therefore does
    not just list facts, it scores them -- a mailbox with one or two Full Access
    delegates and no delivery restrictions is the easy case, and the report says
    so in a column rather than leaving the operator to work it out across four
    others in Excel at 7am on a Saturday.

    READS ONLY. Nothing here changes a mailbox.

    IT IS SLOW BY DESIGN. Full Access and Send As cannot be read in bulk -- both
    are a call per mailbox -- so a tenant with 800 shared mailboxes is 1,600
    round trips. Sequential, with the mailbox named as it goes, because a
    visible slow answer beats a silent one that looks hung.
#>
# TOOLKIT-HELP-RUNS: READS every shared mailbox in Exchange Online. Writes nothing.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Lists every shared mailbox in Exchange Online with the things that decide
# whether it is safe to source-anchor to Microsoft 365.
#
# For each mailbox:
#
#   Anchored     CLOUD if it was created in Exchange Online, SYNCED if it comes
#                from on-premises Active Directory. A synced mailbox is managed
#                on-premises -- that is the one that needs the anchor moved.
#   FullAccess   how many real people can open it. NT AUTHORITY\SELF and the
#                other built-in entries are not counted -- they are noise and
#                counting them makes every mailbox look busy.
#   SendAs       how many real people can send as it.
#   Restricted   whether anyone has set delivery restrictions -- an explicit
#                accept-only or reject list.
#
# PILOT COLUMN. "Easy" means synced, one or two Full Access delegates, and no
# delivery restrictions. Those are the mailboxes with the fewest people to
# notice if something goes wrong. Sort on that column and start at the top.
#
# TWO THINGS THE PILOT COLUMN DELIBERATELY DOES NOT DO:
#
# It does not count "blocks external senders" as a restriction. That setting is
# on by default for most shared mailboxes, so treating it as a complication
# would flag nearly everything and make the column useless. It is reported in
# its own column instead.
#
# It does not look at how BUSY a mailbox is. A mailbox with one delegate that
# receives four hundred messages a day is not an easy pilot, and this report
# cannot tell you that. Check the top candidates against message volume before
# committing to them.
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

$script:Rows      = @()
$script:ReportDir = Get-ReportPath
$script:Cancel    = $false

#region Facts -------------------------------------------------------------------
function Get-DeliveryRestriction {
    <#
        Are there delivery restrictions on this mailbox?

        THE PROPERTY NAMES ARE VERIFIED, NOT REMEMBERED. Learn 2026-08-23
        (Set-Mailbox): AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers,
        AcceptMessagesOnlyFromSendersOrMembers, and the matching
        RejectMessagesFrom / RejectMessagesFromDLMembers /
        RejectMessagesFromSendersOrMembers.

        Getting one of those names wrong would not throw. Reading a property
        that does not exist on a PSObject returns $null, every mailbox would
        report "no restrictions", and the pilot group would be chosen from data
        that was silently wrong. That is the whole reason this is one function
        with the names written once.

        REQUIRESENDERAUTHENTICATIONENABLED IS REPORTED SEPARATELY AND ON
        PURPOSE. It means "reject mail from outside the organisation" and it is
        the DEFAULT for most shared mailboxes -- folding it in here would mark
        almost every mailbox as restricted and make the pilot column worthless.
    #>
    param([Parameter(Mandatory)]$Mailbox)

    $lists = @(
        'AcceptMessagesOnlyFrom'
        'AcceptMessagesOnlyFromDLMembers'
        'AcceptMessagesOnlyFromSendersOrMembers'
        'RejectMessagesFrom'
        'RejectMessagesFromDLMembers'
        'RejectMessagesFromSendersOrMembers'
    )
    $set = New-Object System.Collections.ArrayList
    foreach ($n in $lists) {
        if (-not $Mailbox.PSObject.Properties[$n]) { continue }
        $v = @($Mailbox.$n | Where-Object { "$_".Trim() })
        if (@($v).Count -gt 0) { [void]$set.Add("$n ($(@($v).Count))") }
    }

    return [PSCustomObject]@{
        Restricted = (@($set).Count -gt 0)
        Detail     = (@($set) -join '; ')
        ExternalBlocked = [bool]$Mailbox.RequireSenderAuthenticationEnabled
    }
}

function Get-DelegateCount {
    <#
        How many REAL people can open it, and how many can send as it.

        Test-IsRealTrustee is what makes these numbers mean anything.
        Get-MailboxPermission returns NT AUTHORITY\SELF for every mailbox in the
        tenant, so an unfiltered count is "at least one" everywhere and the
        pilot filter matches nothing useful.

        INHERITED AND DENY ENTRIES ARE EXCLUDED TOO. An inherited allow comes
        from an admin role rather than from somebody being given access to this
        mailbox, and it is not what "who can open this" means to the person
        planning a migration.
    #>
    param([Parameter(Mandatory)][string]$Smtp)

    $r = [PSCustomObject]@{ FullAccess = 0; SendAs = 0; FullAccessWho = ''; SendAsWho = ''; Failed = ''; FullAccessNote = '' }
    $problems = New-Object System.Collections.ArrayList

    try {
        $rawFa = @(Invoke-BulkExoQuery -Cmdlet 'Get-MailboxPermission' -Mailbox $Smtp)

        # EVERYTHING IS COMPARED AS A STRING, and that is not defensive
        # over-engineering -- it is the bug that made this report say every
        # mailbox had zero delegates on 2026-08-23.
        #
        # Across a remoting session these objects arrive DESERIALIZED.
        # AccessRights is not a string array, it is a rehydrated enum
        # collection, so -contains 'FullAccess' compares an object to a string
        # and is quietly false for every row. IsInherited and Deny come back
        # the same way and are truthy as objects even when they mean False.
        #
        # There is no error. The filter simply matches nothing, and a mailbox
        # with six delegates reports zero -- which reads as a clean pilot
        # candidate and is the most dangerous wrong answer this tool could give.
        $fa = @($rawFa | Where-Object {
            $rights = (@($_.AccessRights | ForEach-Object { "$_" }) -join ',')
            "$($_.IsInherited)" -ne 'True' -and "$($_.Deny)" -ne 'True' -and $rights -match 'FullAccess'
        })
        $who = @($fa | ForEach-Object { "$($_.User)".Trim() } |
                 Where-Object { Test-IsRealTrustee -Trustee $_ } | Sort-Object -Unique)
        $r.FullAccess    = @($who).Count
        $r.FullAccessWho = (@($who) -join '; ')

        # A mailbox that returned rows but none survived filtering is worth
        # distinguishing from one that genuinely returned nothing. Silent zero
        # is how the original defect hid.
        if (@($rawFa).Count -gt 0 -and @($who).Count -eq 0) {
            $r.FullAccessNote = "$(@($rawFa).Count) permission row(s) returned, none were a real user with FullAccess"
        }
    } catch { [void]$problems.Add("FullAccess: $($_.Exception.Message)") }

    try {
        # Same string coercion, same reason.
        $sa = @(Invoke-BulkExoQuery -Cmdlet 'Get-RecipientPermission' -Mailbox $Smtp |
                Where-Object {
                    "$($_.AccessControlType)" -eq 'Allow' -and
                    (@($_.AccessRights | ForEach-Object { "$_" }) -join ',') -match 'SendAs'
                })
        $who = @($sa | ForEach-Object { "$($_.Trustee)".Trim() } |
                 Where-Object { Test-IsRealTrustee -Trustee $_ } | Sort-Object -Unique)
        $r.SendAs    = @($who).Count
        $r.SendAsWho = (@($who) -join '; ')
    } catch { [void]$problems.Add("SendAs: $($_.Exception.Message)") }

    $r.Failed = (@($problems) -join ' | ')
    return $r
}

function Get-PilotVerdict {
    <#
        The column this report exists for.

        EASY means: synced from on-premises (so there is something to move),
        one or two Full Access delegates, and no delivery restrictions.

        A COUNT OF ZERO IS NOT EASY, IT IS SUSPICIOUS. A shared mailbox nobody
        can open is either abandoned or accessed some way this report cannot
        see, and finding out which is not a thing to do on a Saturday. It gets
        its own verdict rather than being lumped in with the good ones.
    #>
    param([int]$FullAccess, [bool]$Restricted, [bool]$IsSynced, [string]$Failed)

    if ($Failed)        { return 'UNKNOWN - permissions could not be read' }
    if (-not $IsSynced) { return 'n/a - already cloud-anchored' }
    if ($Restricted)    { return 'No - has delivery restrictions' }
    if ($FullAccess -eq 0) { return 'Check - nobody has Full Access' }
    if ($FullAccess -le 2) { return 'EASY' }
    return "No - $FullAccess delegates"
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Shared Mailbox Anchor Pilot'
$form.ClientSize      = New-Object System.Drawing.Size(1300, 800)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$lblWhat          = New-Object System.Windows.Forms.Label
$lblWhat.Text     = 'Every shared mailbox in Exchange Online, scored for how safe it is to source-anchor to M365.'
$lblWhat.Location = New-Object System.Drawing.Point(14, 12)
$lblWhat.Size     = New-Object System.Drawing.Size(900, 18)
$lblWhat.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($lblWhat)

$btnRun           = New-Object System.Windows.Forms.Button
$btnRun.Text      = 'Run report'
$btnRun.Location  = New-Object System.Drawing.Point(14, 38)
$btnRun.Size      = New-Object System.Drawing.Size(130, 30)
$btnRun.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnRun)

$btnStop           = New-Object System.Windows.Forms.Button
$btnStop.Text      = 'Stop'
$btnStop.Location  = New-Object System.Drawing.Point(154, 38)
$btnStop.Size      = New-Object System.Drawing.Size(90, 30)
$btnStop.FlatStyle = 'Flat'
$btnStop.Enabled   = $false
[void]$form.Controls.Add($btnStop)

$chkEasyOnly          = New-Object System.Windows.Forms.CheckBox
$chkEasyOnly.Text     = 'Show only EASY candidates'
$chkEasyOnly.Location = New-Object System.Drawing.Point(258, 44)
$chkEasyOnly.Size     = New-Object System.Drawing.Size(210, 22)
[void]$form.Controls.Add($chkEasyOnly)

$btnExport           = New-Object System.Windows.Forms.Button
$btnExport.Text      = 'Export CSV'
$btnExport.Location  = New-Object System.Drawing.Point(478, 38)
$btnExport.Size      = New-Object System.Drawing.Size(120, 30)
$btnExport.FlatStyle = 'Flat'
$btnExport.Enabled   = $false
[void]$form.Controls.Add($btnExport)

$lblCount          = New-Object System.Windows.Forms.Label
$lblCount.Location = New-Object System.Drawing.Point(612, 44)
$lblCount.Size     = New-Object System.Drawing.Size(640, 20)
$lblCount.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($lblCount)

$pb = Add-ReportProgressBar -Form $form -X 14 -Y 76 -Width 1272 -Height 14

$grid                     = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(14, 98)
$grid.Size                = New-Object System.Drawing.Size(1272, 480)
$grid.ReadOnly            = $true
$grid.AllowUserToAddRows  = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode       = 'FullRowSelect'
$grid.MultiSelect         = $false
$grid.Anchor              = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($grid)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(14, 586)
$txtOut.Size       = New-Object System.Drawing.Size(1272, 200)
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
# B6.
$txtOut.WordWrap   = $true
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

function Update-Grid {
    $shown = @($script:Rows)
    if ($chkEasyOnly.Checked) { $shown = @($shown | Where-Object { $_.Pilot -eq 'EASY' }) }

    $t = New-Object System.Data.DataTable
    foreach ($c in @('Pilot', 'DisplayName', 'PrimarySmtpAddress', 'Anchored', 'FullAccess', 'SendAs', 'Restricted', 'BlocksExternal')) {
        [void]$t.Columns.Add($c)
    }
    foreach ($r in $shown) {
        [void]$t.Rows.Add($r.Pilot, $r.DisplayName, $r.PrimarySmtpAddress, $r.Anchored,
                          $r.FullAccess, $r.SendAs, $r.Restricted, $r.BlocksExternal)
    }
    $grid.DataSource = $t
    $lblCount.Text = "$(@($shown).Count) shown of $(@($script:Rows).Count)  --  $(@($script:Rows | Where-Object { $_.Pilot -eq 'EASY' }).Count) EASY"
}

function Invoke-Report {
    $script:Cancel = $false
    $script:Rows = @()
    $btnRun.Enabled = $false
    $btnStop.Enabled = $true
    $btnExport.Enabled = $false
    $grid.DataSource = $null

    try {
        $exo = Test-ExoConnectionState
        if (-not $exo.Connected) {
            $txtOut.Text = "Not connected to Exchange Online: $($exo.Detail)"
            return
        }

        $txtOut.Text = 'Listing shared mailboxes ...'
        $form.Refresh()

        # C3, and one bulk call rather than one per mailbox -- the properties
        # needed for the restriction check come back with the mailbox.
        $mbx = @(Invoke-BulkExoQuery -Cmdlet 'Get-Mailbox' `
                    -RecipientTypeDetails 'SharedMailbox' `
                    -Properties @('DisplayName', 'PrimarySmtpAddress', 'IsDirSynced',
                                  'RecipientTypeDetails', 'RequireSenderAuthenticationEnabled',
                                  'AcceptMessagesOnlyFrom', 'AcceptMessagesOnlyFromDLMembers',
                                  'AcceptMessagesOnlyFromSendersOrMembers',
                                  'RejectMessagesFrom', 'RejectMessagesFromDLMembers',
                                  'RejectMessagesFromSendersOrMembers'))

        $total = @($mbx).Count
        if ($total -eq 0) { $txtOut.Text = 'No shared mailboxes found in Exchange Online.'; return }

        $pb.Maximum = $total
        $pb.Value   = 0
        $rows = New-Object System.Collections.ArrayList
        $started = Get-Date
        $failed = New-Object System.Collections.ArrayList
        $n = 0

        foreach ($m in $mbx) {
            if ($script:Cancel) { break }
            $n++
            $smtp = "$($m.PrimarySmtpAddress)"

            # LIVE PROGRESS, EVERY MAILBOX.
            #
            # This is two network round trips per mailbox and a thousand
            # mailboxes takes minutes. The first version updated a label and
            # left the grid empty until the end, which is indistinguishable
            # from hung -- and an operator who cannot tell the difference kills
            # it at four minutes and reruns it, which is worse than slow.
            #
            # DoEvents on EVERY iteration, not every third. A repaint costs
            # microseconds against a round trip measured in tens of
            # milliseconds, and the saving was never real.
            $pb.Value = [Math]::Min($n, $pb.Maximum)
            $elapsed = (Get-Date) - $started
            $rate = 0.0
            if ($elapsed.TotalSeconds -gt 0) { $rate = $n / $elapsed.TotalSeconds }
            $remain = ''
            if ($rate -gt 0 -and $n -lt $total) {
                $secsLeft = [int](($total - $n) / $rate)
                $remain = ' -- about {0:mm\:ss} left' -f ([TimeSpan]::FromSeconds($secsLeft))
            }
            $lblCount.Text = "$n of $total ($([int](100 * $n / $total))%)$remain   |   $(@($rows | Where-Object { $_.Pilot -eq 'EASY' }).Count) EASY so far"

            # The scrolling detail, so there is always something MOVING. A
            # static percentage still reads as stuck to somebody who has been
            # watching it for two minutes.
            $recent = @($rows | Select-Object -Last 6 | ForEach-Object {
                '    {0,-9} {1,-42} FA={2,-3} SA={3,-3} {4}' -f $_.Pilot.Split(' ')[0], $_.PrimarySmtpAddress, $_.FullAccess, $_.SendAs, $_.Anchored
            })
            $txtOut.Text = ("Reading $smtp" + [Environment]::NewLine +
                "$n of $total, {0:mm\:ss} elapsed$remain" -f $elapsed) +
                [Environment]::NewLine + [Environment]::NewLine +
                'last few:' + [Environment]::NewLine + (@($recent) -join [Environment]::NewLine)
            [System.Windows.Forms.Application]::DoEvents()

            $res = Get-DeliveryRestriction -Mailbox $m
            $del = Get-DelegateCount -Smtp $smtp
            if ($del.Failed) { [void]$failed.Add("$smtp -- $($del.Failed)") }

            # IsDirSynced: TRUE means the object is mastered in on-premises AD
            # and synced up. That is the one whose anchor has to move.
            $synced = [bool]$m.IsDirSynced

            # THE GRID FILLS AS IT GOES, every 25 mailboxes. On a thousand-
            # mailbox tenant the first EASY candidates are visible within
            # seconds, and the operator can start reading them while the rest
            # runs -- rather than waiting for a complete answer to a question
            # they only needed a few rows of.
            if ($n -gt 1 -and ($n % 25) -eq 0) {
                $script:Rows = @($rows.ToArray())
                Update-Grid
                [System.Windows.Forms.Application]::DoEvents()
            }

            [void]$rows.Add([PSCustomObject]@{
                Pilot              = (Get-PilotVerdict -FullAccess $del.FullAccess -Restricted $res.Restricted -IsSynced $synced -Failed $del.Failed)
                DisplayName        = "$($m.DisplayName)"
                PrimarySmtpAddress = $smtp
                Anchored           = $(if ($synced) { 'SYNCED from on-prem' } else { 'CLOUD' })
                FullAccess         = $del.FullAccess
                SendAs             = $del.SendAs
                Restricted         = $(if ($res.Restricted) { 'Yes' } else { 'No' })
                BlocksExternal     = $(if ($res.ExternalBlocked) { 'Yes' } else { 'No' })
                FullAccessWho      = $del.FullAccessWho
                SendAsWho          = $del.SendAsWho
                RestrictionDetail  = $res.Detail
                ReadError          = $del.Failed
            })
        }

        $script:Rows = @($rows.ToArray() | Sort-Object @{ E = { $_.Pilot -ne 'EASY' } }, FullAccess, DisplayName)
        Update-Grid
        $btnExport.Enabled = (@($script:Rows).Count -gt 0)

        $easy   = @($script:Rows | Where-Object { $_.Pilot -eq 'EASY' })
        $synced = @($script:Rows | Where-Object { $_.Anchored -like 'SYNCED*' })
        $L = New-Object System.Collections.ArrayList
        # ARITHMETIC ON A PARTIAL RUN IS A LIE, AND THIS ONE WAS.
        #
        # The first version said "$total - synced = already cloud-anchored".
        # Stopped after 31 of 1,162 it reported "31 synced, 1131 already
        # cloud-anchored" -- presenting 1,131 mailboxes it had never looked at
        # as a finding. The operator caught it; the tool did not.
        #
        # Every number below is now counted from rows ACTUALLY READ, and what
        # was not read is stated as unexamined rather than folded into the
        # other bucket. A count you did not measure is not the complement of
        # one you did.
        $scanned = @($script:Rows).Count
        $cloud   = @($script:Rows | Where-Object { $_.Anchored -eq 'CLOUD' })
        if ($script:Cancel) {
            [void]$L.Add("STOPPED after $scanned of $total. EVERYTHING BELOW COUNTS ONLY THOSE $scanned.")
            [void]$L.Add("$($total - $scanned) mailbox(es) were NOT examined. They are not counted as anything -- not synced, not cloud, not easy.")
            [void]$L.Add('')
        }
        [void]$L.Add("$total shared mailbox(es) exist. $scanned examined.")
        [void]$L.Add("Of those $scanned : $(@($synced).Count) synced from on-premises, $(@($cloud).Count) already cloud-anchored.")
        [void]$L.Add("$(@($easy).Count) look EASY: synced, one or two Full Access delegates, no delivery restrictions.")
        [void]$L.Add('')
        if (@($failed).Count -gt 0) {
            # D6. A permission read that failed is not "no delegates".
            [void]$L.Add("PERMISSIONS COULD NOT BE READ ON $(@($failed).Count) MAILBOX(ES). They are marked UNKNOWN, not zero:")
            foreach ($f in @($failed | Select-Object -First 10)) { [void]$L.Add("    $f") }
            if (@($failed).Count -gt 10) { [void]$L.Add("    ...and $(@($failed).Count - 10) more") }
            [void]$L.Add('')
        }
        [void]$L.Add('BEFORE YOU COMMIT TO THESE:')
        [void]$L.Add('  - This does not know how BUSY a mailbox is. One delegate and four hundred messages a day is not an easy pilot.')
        [void]$L.Add('  - "Blocks external senders" is shown separately and is NOT counted as a restriction. It is the default on most shared mailboxes, so counting it would flag nearly everything.')
        [void]$L.Add('  - Delivery restrictions do not stop a Full Access delegate. Microsoft Learn: a user with Full Access can still update the contents even if restricted. So Restricted=No is about mail flow, not about who can get in.')
        [void]$L.Add('  - Send As is reported but not scored. A mailbox people send AS is more visible when it breaks than one people only read.')
        $txtOut.Text = ($L -join [Environment]::NewLine)

        Write-CompletedLog -Action 'Read' -Target (Get-LogTarget 'Shared Mailbox Pilot') `
            -Result $(if (@($failed).Count -gt 0 -or $script:Cancel) { 'Partial' } else { 'Success' }) `
            -Detail "Scanned $n of $total shared mailboxes. $(@($easy).Count) EASY candidates. $(@($failed).Count) permission read failure(s)." `
            -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Shared mailbox report'
    } finally {
        $btnRun.Enabled = $true
        $btnStop.Enabled = $false
        $pb.Value = 0
    }
}

function Export-Report {
    if (@($script:Rows).Count -eq 0) { return }
    $out = @($script:Rows | Select-Object Pilot, DisplayName, PrimarySmtpAddress, Anchored,
             FullAccess, SendAs, Restricted, BlocksExternal, FullAccessWho, SendAsWho,
             RestrictionDetail, ReadError)
    $p = Write-ReportCsv -Directory $script:ReportDir -Prefix 'SMB_Inventory' -Rows $out
    $txtOut.Text = "Exported $(@($out).Count) row(s) to:" + [Environment]::NewLine + $p + [Environment]::NewLine + [Environment]::NewLine +
        'The CSV carries WHO has access, which the grid leaves out for width. Sort on Pilot, then FullAccess.'
}
#endregion

$btnRun.Add_Click({ try { Invoke-Report } catch { Show-HandlerError -ErrorRecord $_ -Context 'Run report' } })
$btnStop.Add_Click({ try { $script:Cancel = $true } catch { } })
$btnExport.Add_Click({ try { Export-Report } catch { Show-HandlerError -ErrorRecord $_ -Context 'Export' } })
$chkEasyOnly.Add_CheckedChanged({ try { Update-Grid } catch { Show-HandlerError -ErrorRecord $_ -Context 'Filter' } })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Shared Mailbox Anchor Pilot' | Out-Null

$form.Add_Shown({
    try {
        $exo = Test-ExoConnectionState
        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("Exchange Online: $(if ($exo.Connected) { 'connected' } else { $exo.Detail })")
        [void]$L.Add('')
        [void]$L.Add('Press Run report. It reads every shared mailbox, then asks each one who has Full Access and Send As.')
        [void]$L.Add('')
        [void]$L.Add('THAT SECOND PART IS TWO CALLS PER MAILBOX and cannot be done in bulk, so several hundred mailboxes takes several minutes. The mailbox being read is named as it goes, and Stop keeps what it has already found.')
        $txtOut.Text = ($L -join [Environment]::NewLine)
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Shared Mailbox Pilot' }
})

[void]$form.ShowDialog()
