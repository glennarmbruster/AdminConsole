#Requires -Version 5.1
<#
    MBX-Stats.ps1
    Mailbox size and per-folder statistics for one mailbox, with quota use and the biggest folders. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-08-24
#>
# TOOLKIT-HELP-RUNS: reads one mailbox's statistics. Changes nothing.
# TOOLKIT-HELP-BEGIN
# Where the space went. Total size against the quota, then the folders that
# account for it.
#
# The ticket this answers is "my mailbox is full" or "why can I not send", and
# the useful answer is almost never the total -- it is which folder holds it.
# Usually Deleted Items nobody ever emptied, or a Sent Items going back nine
# years, or Recoverable Items quietly holding everything a retention policy
# was asked to keep.
#
# WHAT IT SHOWS:
#   * total size, item count, and percentage of the send quota used
#   * Inbox, Sent, Calendar and Deleted Items, called out first
#   * every other well-known folder
#   * the ten biggest folders anywhere in the mailbox
#   * the archive, if there is one, on the same basis
#
# RECOVERABLE ITEMS IS NOT A FOLDER THE USER CAN SEE, and it is the usual
# reason a mailbox is far bigger than the sum of everything visible. It holds
# deleted mail a hold or a retention policy is keeping. Emptying Deleted Items
# does not shrink it. If that is where the space is, the answer is a retention
# or hold question rather than a housekeeping one -- Mailbox Health shows which
# holds are applied.
#
# Read-only. Exchange Online.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

# An empty address is fine. The window opens and asks.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Sizes ----------------------------------------------------------------
function ConvertTo-SizeBytes {
    <#
        A size from Exchange, as a number of bytes.

        THREE SHAPES COME BACK AND THE CODE SEES ALL THREE. On-premises returns
        a ByteQuantifiedSize object with a ToBytes() method. Exchange Online's
        REST cmdlets return a STRING, usually

            1.234 GB (1,325,000,000 bytes)

        and occasionally just "1.234 GB" with no byte count at all. Formatting
        the object's ToString() and parsing that back is not optional -- it is
        the only thing that works across both.

        Returns $null rather than 0 when it cannot tell. A zero would sort as
        "smallest folder" and read as "empty", and an unknown size is neither.
    #>
    param($Value)

    if ($null -eq $Value) { return $null }

    # On-premises, and anything else carrying the real type.
    try {
        if ($Value.PSObject.Methods['ToBytes']) { return [long]$Value.ToBytes() }
    } catch { }

    $s = "$Value".Trim()
    if (-not $s) { return $null }

    # "1.234 GB (1,325,000,000 bytes)" -- the exact count, when it is offered.
    if ($s -match '\(([\d,\.\s]+)\s*bytes\)') {
        $n = $Matches[1] -replace '[,\s]', ''
        try { return [long][double]$n } catch { return $null }
    }

    # "1.234 GB" on its own. Binary multipliers: Exchange means GiB by GB.
    if ($s -match '^([\d\.]+)\s*(B|KB|MB|GB|TB)$') {
        $mult = @{ 'B' = 1L; 'KB' = 1024L; 'MB' = 1048576L
                   'GB' = 1073741824L; 'TB' = 1099511627776L }
        try { return [long]([double]$Matches[1] * $mult[$Matches[2].ToUpper()]) } catch { return $null }
    }

    return $null
}

function Format-SizeBytes {
    <#
        Bytes as something a person reads at a glance.

        [Math]::Round, NOT [int]. A cast to [int] rounds half to EVEN -- 2.5
        becomes 2 and 3.5 becomes 4 -- which produces sizes that are wrong by a
        unit and look like typing errors (constraint A4).
    #>
    param($Bytes)
    if ($null -eq $Bytes) { return '(unknown)' }
    $b = 0.0
    try { $b = [double]$Bytes } catch { return '(unknown)' }
    if ($b -lt 1024)          { return "$([long]$b) B" }
    if ($b -lt 1048576)       { return ($b / 1024).ToString('N1') + ' KB' }
    if ($b -lt 1073741824)    { return ($b / 1048576).ToString('N1') + ' MB' }
    if ($b -lt 1099511627776) { return ($b / 1073741824).ToString('N2') + ' GB' }
    return ($b / 1099511627776).ToString('N2') + ' TB'
}

function Format-ItemCount {
    <#
        An item count, or '' when there isn't one.

        [long]$null is 0, and "0 items" printed beside an "(unknown)" size reads
        as an empty mailbox rather than as a call that failed (D6).
    #>
    param($Count)
    if ($null -eq $Count) { return '' }
    try { return ([long]$Count).ToString('N0') + ' items' } catch { return '' }
}

function Format-StatLine {
    <#
        One label / value / note line, padded.

        NO COMPOSITE FORMAT STRINGS ANYWHERE IN THIS FILE, DELIBERATELY.

        The -f operator throws at RUN TIME when the argument list does not match
        the placeholders, and it produced two separate defects here: an invalid
        alignment specifier that only fired when a folder was absent, and an
        argument list that collapsed on a real mailbox. Both parse cleanly. Both
        are invisible until the exact data shape that triggers them turns up.

        PadRight and PadLeft cannot fail that way -- there is no argument list to
        get wrong -- and the result is the same fixed-width output.
    #>
    param([string]$Label, [string]$Value, [string]$Note = '')
    $line = '  ' + $Label.PadRight(24) + ' ' + $Value.PadLeft(12)
    if ($Note) { $line += '  ' + $Note }
    return $line
}
#endregion

#region Data -----------------------------------------------------------------
# WELL-KNOWN FOLDERS ARE MATCHED ON FolderType, NEVER ON Name.
#
# Folder names are LOCALISED. A mailbox created in a French locale has
# 'Elements envoyes' where an English one has 'Sent Items', and a user can
# rename most of them anyway. FolderType is the invariant, and it is the same
# reason Calendar Permissions cannot look the calendar up by its name.
#
# Ordered deliberately: the four asked for first, then the rest.
$script:HeadlineFolders = [ordered]@{
    'Inbox'        = 'Inbox'
    'SentItems'    = 'Sent'
    'Calendar'     = 'Calendar'
    'DeletedItems' = 'Deleted Items'
}
$script:OtherFolders = [ordered]@{
    'Drafts'              = 'Drafts'
    'JunkEmail'           = 'Junk Email'
    'Outbox'              = 'Outbox'
    'Contacts'            = 'Contacts'
    'Tasks'               = 'Tasks'
    'Notes'               = 'Notes'
    'Archive'             = 'Archive (in-place folder)'
    'ConversationHistory' = 'Conversation History'
    'SyncIssues'          = 'Sync Issues'
}

function Get-MbxStatsSnapshot {
    <#
        Everything this window shows, for one mailbox.

        Errors are collected rather than thrown. A folder-statistics call that
        fails should cost the folder table and not the total size -- a partial
        answer to "how big is this mailbox" is still an answer.
    #>
    param([Parameter(Mandatory)][string]$Smtp)

    $snap = [ordered]@{
        Primary        = $Smtp
        DisplayName    = $null
        MailboxType    = $null
        TotalBytes     = $null
        TotalDisplay   = $null
        ItemCount      = $null
        DeletedBytes   = $null
        DeletedCount   = $null
        QuotaBytes     = $null
        QuotaDisplay   = $null
        PercentUsed    = $null
        LastLogon      = $null
        ArchiveEnabled = $false
        ArchiveBytes   = $null
        ArchiveCount   = $null
        Folders        = @()
        ArchiveFolders = @()
        Errors         = @()
    }

    # ROUTED THROUGH Invoke-EXOCommand. Get-Mailbox, Get-MailboxStatistics and
    # Get-MailboxFolderStatistics all exist in BOTH Exchange editions, and in a
    # process holding both sessions the last import owns the name (C1). An
    # on-premises answer under a cloud heading is the whole failure mode.
    #
    # NO ResultSize ON ANY OF THESE. They fetch one object by identity rather
    # than enumerating a collection, so the 1000-item default (C3) does not
    # apply -- adding it here would just be cargo.
    $mbx = $null
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop' }
        $snap.DisplayName = "$($mbx.DisplayName)"
        $snap.MailboxType = Get-SimpleMailboxType -RecipientTypeDetails "$($mbx.RecipientTypeDetails)"
        $snap.QuotaBytes  = ConvertTo-SizeBytes $mbx.ProhibitSendQuota
        $snap.QuotaDisplay= "$($mbx.ProhibitSendQuota)"
        $snap.ArchiveEnabled = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
    } catch {
        $snap.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (quota and type)')
    }

    try {
        $st = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop' }
        $snap.TotalBytes   = ConvertTo-SizeBytes $st.TotalItemSize
        $snap.TotalDisplay = "$($st.TotalItemSize)"
        $snap.ItemCount    = $st.ItemCount
        $snap.DeletedBytes = ConvertTo-SizeBytes $st.TotalDeletedItemSize
        $snap.DeletedCount = $st.DeletedItemCount
        $snap.LastLogon    = $st.LastLogonTime
    } catch {
        $snap.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics')
    }

    # PERCENTAGE ONLY WHEN BOTH NUMBERS ARE REAL. An unlimited quota parses to
    # $null, and dividing by it would either throw or invent a number.
    if ($snap.TotalBytes -and $snap.QuotaBytes -and $snap.QuotaBytes -gt 0) {
        $snap.PercentUsed = [Math]::Round(100 * ($snap.TotalBytes / $snap.QuotaBytes), 1)
    }

    try {
        $snap.Folders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop' } |
            ForEach-Object {
                [PSCustomObject]@{
                    Name       = "$($_.Name)"
                    Path       = "$($_.FolderPath)"
                    Type       = "$($_.FolderType)"
                    Items      = $_.ItemsInFolder
                    Bytes      = ConvertTo-SizeBytes $_.FolderSize
                    TotalItems = $_.ItemsInFolderAndSubfolders
                    TotalBytes = ConvertTo-SizeBytes $_.FolderAndSubfolderSize
                }
            })
    } catch {
        $snap.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxFolderStatistics')
    }

    if ($snap.ArchiveEnabled) {
        try {
            $ast = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{
                Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $snap.ArchiveBytes = ConvertTo-SizeBytes $ast.TotalItemSize
            $snap.ArchiveCount = $ast.ItemCount
        } catch {
            $snap.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive')
        }
        try {
            $snap.ArchiveFolders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{
                Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' } |
                ForEach-Object {
                    [PSCustomObject]@{
                        Name  = "$($_.Name)"
                        Type  = "$($_.FolderType)"
                        Items = $_.ItemsInFolder
                        Bytes = ConvertTo-SizeBytes $_.FolderSize
                    }
                })
        } catch {
            $snap.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxFolderStatistics -Archive')
        }
    }

    return [PSCustomObject]$snap
}

function Get-FolderByType {
    param($Folders, [string]$Type)
    return @($Folders | Where-Object { $_.Type -eq $Type } | Select-Object -First 1)[0]
}

function Get-RecoverableBytes {
    <#
        Everything under Recoverable Items, added up.

        Its subfolders -- Deletions, Purges, Versions, DiscoveryHolds -- are
        reported separately and none of them is the total. Summing the ones
        whose type starts 'RecoverableItems' is the honest answer, and it is
        the number that explains a mailbox far larger than its visible folders.
    #>
    param($Folders)
    $rows = @($Folders | Where-Object { $_.Type -like 'RecoverableItems*' -and $null -ne $_.Bytes })
    if ($rows.Count -eq 0) { return $null }
    $sum = 0L
    foreach ($r in $rows) { $sum += [long]$r.Bytes }
    return $sum
}
#endregion

#region Rendering ------------------------------------------------------------
function Format-FolderLine {
    <#
        "NOT PRESENT" AND "COULD NOT BE IDENTIFIED" ARE DIFFERENT FACTS.

        The first is a mailbox without that folder. The second is a mailbox
        whose folder has lost its FolderType -- which is damage, and a finding
        in its own right rather than a gap in this report. Printing the same
        text for both sent a real investigation looking for a bug in the folder
        map when the answer was a corrupt mailbox.

        Naming the type that was searched for is what makes the difference
        diagnosable without reading this source.
    #>
    param([string]$Label, $Folder, [string]$Type)
    if (-not $Folder) {
        $why = if ($Type) { "(no folder reports FolderType '$Type')" } else { '(not present)' }
        return (Format-StatLine -Label $Label -Value '' -Note $why)
    }
    return (Format-StatLine -Label $Label `
                            -Value (Format-SizeBytes $Folder.Bytes) `
                            -Note  (Format-ItemCount $Folder.Items))
}

function Get-UntypedFolders {
    <#
        Folders that came back with no FolderType at all.

        THE LOCALISATION-SAFE CORRUPTION SIGNAL. Matching folders by NAME would
        find a damaged Sent Items on an English mailbox and miss it on every
        other locale, so this counts the problem instead of guessing at it: a
        healthy mailbox types every folder it returns, and one that does not is
        telling you something.

        Root is excluded -- it legitimately reports nothing on some mailboxes
        and would put every mailbox on this list.
    #>
    param($Folders)
    return @($Folders | Where-Object {
        -not "$($_.Type)".Trim() -and "$($_.Path)".Trim() -and "$($_.Path)" -ne '/'
    })
}

function Format-MbxStats {
    param($Snap)
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("Mailbox      : $($Snap.Primary)")
    if ($Snap.DisplayName) { [void]$sb.AppendLine("Name         : $($Snap.DisplayName)") }
    if ($Snap.MailboxType) { [void]$sb.AppendLine("Type         : $($Snap.MailboxType)") }
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('--- Size ---')
    $totalItems = Format-ItemCount $Snap.ItemCount
    if (-not $totalItems) { $totalItems = '(count unavailable)' }
    [void]$sb.AppendLine((Format-StatLine -Label 'Total' -Value (Format-SizeBytes $Snap.TotalBytes) -Note $totalItems))

    if ($Snap.QuotaBytes) {
        $pct = if ($null -ne $Snap.PercentUsed) { "$($Snap.PercentUsed)% of quota" } else { '' }
        [void]$sb.AppendLine((Format-StatLine -Label 'Send quota' -Value (Format-SizeBytes $Snap.QuotaBytes) -Note $pct))
    } else {
        [void]$sb.AppendLine((Format-StatLine -Label 'Send quota' -Value "$($Snap.QuotaDisplay)"))
    }
    [void]$sb.AppendLine((Format-StatLine -Label 'Deleted (recoverable)' `
                            -Value (Format-SizeBytes $Snap.DeletedBytes) `
                            -Note  (Format-ItemCount $Snap.DeletedCount)))
    if ($Snap.LastLogon) { [void]$sb.AppendLine("  Last logon               $($Snap.LastLogon)") }
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('--- Default folders ---')
    foreach ($k in $script:HeadlineFolders.Keys) {
        [void]$sb.AppendLine((Format-FolderLine -Label $script:HeadlineFolders[$k] `
                                -Folder (Get-FolderByType -Folders $Snap.Folders -Type $k) `
                                -Type   $k))
    }

    # THE MAILBOX ITSELF MAY BE THE PROBLEM. Reported here, next to the folders
    # it affects, rather than buried with the field errors -- a folder missing
    # from the list above and an untyped folder in the mailbox are the same
    # event seen from two ends.
    $untyped = @(Get-UntypedFolders -Folders $Snap.Folders)
    if ($untyped.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("  $($untyped.Count) folder(s) returned NO FolderType. That is mailbox damage,")
        [void]$sb.AppendLine('  not a gap in this report, and it is why a folder above may be missing:')
        foreach ($u in @($untyped | Select-Object -First 8)) {
            [void]$sb.AppendLine('    ' + (Format-SizeBytes $u.Bytes).PadLeft(12) + '  ' + "$($u.Path)")
        }
        if ($untyped.Count -gt 8) {
            [void]$sb.AppendLine("    ... and $($untyped.Count - 8) more.")
        }
        [void]$sb.AppendLine('  Exchange Online has no self-service repair -- New-MailboxRepairRequest is')
        [void]$sb.AppendLine('  on-premises only. A genuinely damaged cloud mailbox needs a support case.')
    }
    [void]$sb.AppendLine('')

    $rec = Get-RecoverableBytes -Folders $Snap.Folders
    if ($null -ne $rec) {
        [void]$sb.AppendLine('--- Recoverable Items (not visible to the user) ---')
        [void]$sb.AppendLine((Format-StatLine -Label 'Total held' -Value (Format-SizeBytes $rec)))
        [void]$sb.AppendLine('  Held by a retention policy or a hold. Emptying Deleted Items')
        [void]$sb.AppendLine('  does NOT shrink this -- see Mailbox Health for which holds apply.')
        [void]$sb.AppendLine('')
    }

    [void]$sb.AppendLine('--- Other well-known folders ---')
    $shown = $false
    foreach ($k in $script:OtherFolders.Keys) {
        $f = Get-FolderByType -Folders $Snap.Folders -Type $k
        if ($f) { [void]$sb.AppendLine((Format-FolderLine -Label $script:OtherFolders[$k] -Folder $f -Type $k)); $shown = $true }
    }
    if (-not $shown) { [void]$sb.AppendLine('  (none reported)') }
    [void]$sb.AppendLine('')

    # THE SECTION THAT ACTUALLY ANSWERS THE TICKET. A user with a full mailbox
    # rarely has it in a folder anyone would guess -- it is a nine-year Sent
    # Items, or one subfolder somebody archived into by hand.
    [void]$sb.AppendLine('--- Ten biggest folders ---')
    $top = @($Snap.Folders | Where-Object { $null -ne $_.Bytes -and $_.Bytes -gt 0 } |
             Sort-Object -Property Bytes -Descending | Select-Object -First 10)
    if ($top.Count -eq 0) {
        [void]$sb.AppendLine('  (no folder sizes returned)')
    } else {
        foreach ($f in $top) {
            $size  = (Format-SizeBytes $f.Bytes).PadLeft(12)
            $items = (Format-ItemCount $f.Items).PadLeft(14)
            [void]$sb.AppendLine("  $size  $items  $($f.Path)")
        }
    }
    [void]$sb.AppendLine('')

    if ($Snap.ArchiveEnabled) {
        [void]$sb.AppendLine('--- Archive ---')
        [void]$sb.AppendLine((Format-StatLine -Label 'Total' `
                                -Value (Format-SizeBytes $Snap.ArchiveBytes) `
                                -Note  (Format-ItemCount $Snap.ArchiveCount)))
        $atop = @($Snap.ArchiveFolders | Where-Object { $null -ne $_.Bytes -and $_.Bytes -gt 0 } |
                  Sort-Object -Property Bytes -Descending | Select-Object -First 5)
        foreach ($f in $atop) {
            [void]$sb.AppendLine('  ' + (Format-SizeBytes $f.Bytes).PadLeft(12) + '  ' + "$($f.Name)")
        }
        [void]$sb.AppendLine('')
    } else {
        [void]$sb.AppendLine('--- Archive ---')
        [void]$sb.AppendLine('  Not enabled.')
        [void]$sb.AppendLine('')
    }

    if ($Snap.Errors.Count -gt 0) {
        [void]$sb.AppendLine('--- Field errors (everything above is still valid) ---')
        $Snap.Errors | ForEach-Object { [void]$sb.AppendLine($_) }
    }
    return $sb.ToString()
}
#endregion

#region UI -------------------------------------------------------------------
function Show-MbxStatsForm {
    $form             = New-Object System.Windows.Forms.Form
    $form.Text        = 'Mailbox Statistics'
    $form.ClientSize  = New-Object System.Drawing.Size(700, 770)
    $form.StartPosition = 'Manual'
    $form.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $txtOut            = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline  = $true
    $txtOut.ReadOnly   = $true
    $txtOut.ScrollBars = 'Vertical'
    $txtOut.Location   = New-Object System.Drawing.Point(12, 12)
    $txtOut.Size       = New-Object System.Drawing.Size(676, 710)
    $txtOut.Anchor     = 'Top,Left,Right,Bottom'
    $txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
    # Word wrap on (B6). A diagnosis off the right edge is a diagnosis nobody reads.
    $txtOut.WordWrap   = $true
    $form.Controls.Add($txtOut)

    $btnRefresh          = New-Object System.Windows.Forms.Button
    $btnRefresh.Text     = 'Refresh'
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 730)
    $btnRefresh.Size     = New-Object System.Drawing.Size(100, 28)
    $btnRefresh.Anchor   = 'Left,Bottom'
    $form.Controls.Add($btnRefresh)

    $btnCopy          = New-Object System.Windows.Forms.Button
    $btnCopy.Text     = 'Copy'
    $btnCopy.Location = New-Object System.Drawing.Point(120, 730)
    $btnCopy.Size     = New-Object System.Drawing.Size(100, 28)
    $btnCopy.Anchor   = 'Left,Bottom'
    $form.Controls.Add($btnCopy)

    $btnClose          = New-Object System.Windows.Forms.Button
    $btnClose.Text     = 'Close'
    $btnClose.Location = New-Object System.Drawing.Point(578, 730)
    $btnClose.Size     = New-Object System.Drawing.Size(110, 28)
    $btnClose.Anchor   = 'Right,Bottom'
    $form.Controls.Add($btnClose)

    function Update-StatsFromCurrent {
        # Every handler wrapped (B1). An exception inside a WinForms handler is
        # caught by .NET rather than PowerShell, so the operator would otherwise
        # get a CLR crash dialog with no script name and no line number.
        try {
            if (-not "$script:Target".Trim()) {
                $txtOut.Text = 'Enter an address above and press Look up.'
                return
            }
            $exo = Test-ExoConnectionState
            if (-not $exo.Connected) {
                $txtOut.Text = "Not connected to Exchange Online. Use the Connect button on the console first." +
                               [Environment]::NewLine + [Environment]::NewLine + "$($exo.Detail)"
                return
            }
            $form.Text   = "Mailbox Statistics: $script:Target"
            $txtOut.Text = 'Reading folder statistics -- this takes a few seconds on a large mailbox...'
            $btnRefresh.Enabled = $false
            $txtOut.Refresh()

            $snap = Get-MbxStatsSnapshot -Smtp $script:Target
            $txtOut.Text = Format-MbxStats -Snap $snap

            Write-CompletedLog -Action 'Read' -Target (Get-LogTarget $snap.Primary $script:Target) `
                -Result 'Success' -Detail "Mailbox statistics: $(Format-SizeBytes $snap.TotalBytes)"
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox statistics'
            Write-FailedLog -Action 'Read' -Target "$script:Target" -ErrorRecord $_
        } finally {
            $btnRefresh.Enabled = $true
        }
    }

    $btnRefresh.Add_Click({ Update-StatsFromCurrent })
    $btnClose.Add_Click({ $form.Close() })
    $btnCopy.Add_Click({
        try {
            if ("$($txtOut.Text)".Trim()) { [System.Windows.Forms.Clipboard]::SetText($txtOut.Text) }
        } catch {
            Show-ToolDialog(
                "Copy to clipboard failed: $($_.Exception.Message)",
                'Copy', 'OK', 'Warning') | Out-Null
        }
    })

    # FIRST: it shifts every existing control down, so anything added after it
    # would sit at the old offset.
    $bar = Add-ToolTargetBar -Form $form -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-StatsFromCurrent
    })

    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so it
    # is not restyled, and BringToFront'd inside the helper because WinForms
    # draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Mailbox Statistics' | Out-Null

    Update-StatsFromCurrent
    [void]$form.ShowDialog()
}
#endregion

Show-MbxStatsForm
