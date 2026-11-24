#Requires -Version 5.1
# TOOLKIT-HELP-RUNS: nothing new. It reads what the console already knows and
# the logs already on this machine.
# TOOLKIT-HELP-BEGIN
# The start-of-day page. Six panels answering the questions worth asking before
# you touch anything:
#
#   ENVIRONMENT      Am I connected to what I need, and to the right tenant?
#   TOOLKIT HEALTH   Is my copy complete and current?
#   TODAY            What has run today, and did any of it fail?
#   TOP TOOLS        What actually gets used.
#   FAILURE HOTSPOTS What keeps going wrong, so it can be fixed.
#
# READ-ONLY. Nothing here changes a mailbox, a group or a file. The buttons
# open other tools; they do not act.
#
# The connection and version panels call the SAME functions as the console
# header and footer rather than working it out again -- so the two cannot
# disagree with each other. A second opinion about your own state is worse
# than one.
#
# Cloud only: Exchange Online and SharePoint Online. The capability map also
# probes on-premises Exchange and Active Directory, but this toolkit runs in
# tenants that have neither, and a row that is permanently red is a row people
# learn to skip.
#
# ABOUT "UNKNOWN" RESULTS: rows converted from the old two-column log have no
# Result, because that log never recorded one. They are counted as Unknown, not
# folded into Success. The count shrinks on its own as those rows age out.
# TOOLKIT-HELP-END
<#
    Dashboard.ps1
    Operational status page for the tenant-agnostic toolkit.

    Reasoning: ScriptNotes/Dashboard.ps1.notes.md (reference toolkit)

    Author   : Glenn Armbruster
    Created  : 2026-08-07
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Colours and layout ------------------------------------------------------
# Same values as the console header, for the same reason: on the dark navy
# theme, stock Green and Red are dark-on-dark.
$script:ColGood = [System.Drawing.Color]::FromArgb(80, 220, 130)
$script:ColBad  = [System.Drawing.Color]::FromArgb(255, 105, 97)
$script:ColWarn = [System.Drawing.Color]::FromArgb(230, 170, 60)
$script:ColDim  = [System.Drawing.Color]::DimGray

$script:ColW  = 446
$script:ColL  = 12
$script:ColR  = 470
#endregion

#region Log reading -------------------------------------------------------------
function Get-DashLogRows {
    <#
        Every row of the completed-task log, or an empty set with a reason.

        Never throws. A dashboard that dies because a CSV is mid-write is worse
        than one that says "could not read the log" -- the first hides all five
        other panels, the second hides one.
    #>
    $out = [PSCustomObject]@{ Rows = @(); Problem = '' }

    $folder = $null
    try { $folder = Get-LogPath } catch { $out.Problem = 'Log folder not resolvable.'; return $out }
    if (-not $folder -or -not (Test-Path -LiteralPath $folder)) {
        $out.Problem = 'No Logs folder yet.'
        return $out
    }

    $file = Join-Path $folder 'CompletedTasks.csv'
    if (-not (Test-Path -LiteralPath $file)) {
        $out.Problem = 'No CompletedTasks.csv yet.'
        return $out
    }

    $head = ''
    try { $head = (Get-Content -LiteralPath $file -TotalCount 1 -ErrorAction Stop) } catch { }
    if ("$head".Trim() -ne $script:LogHeaderLine) {
        # The old two-column file. Say so precisely, and say what fixes it --
        # this is the exact state Convert-LogHistory.ps1 exists for.
        $out.Problem = 'Log is still in the OLD format. Run Convert-LogHistory.ps1 once.'
        return $out
    }

    try { $out.Rows = @(Import-Csv -LiteralPath $file -ErrorAction Stop) }
    catch { $out.Problem = "Could not read the log: $($_.Exception.Message)" }
    return $out
}

function Get-RowDate {
    param($Row)
    $d = [datetime]::MinValue
    if ([datetime]::TryParse("$($Row.Timestamp)", [ref]$d)) { return $d }
    return $null
}

function Get-RowToolName {
    <#
        Script name, or a clearly-marked stand-in.

        Rows converted from the old log have no Script -- the old format did not
        record one. They are labelled rather than dropped, because silently
        omitting them would understate every total on this page.
    #>
    param($Row)
    $s = "$($Row.Script)".Trim()
    if ($s) { return ($s -replace '\.ps1$', '') }
    return '(not recorded)'
}
#endregion

#region Section helpers ---------------------------------------------------------
function New-DashSection {
    param([string]$Title, [int]$X, [int]$Y, [int]$W, [int]$H)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text     = $Title
    $g.Location = New-Object System.Drawing.Point($X, $Y)
    $g.Size     = New-Object System.Drawing.Size($W, $H)
    $g.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    return $g
}

function Add-DashLine {
    <#
        One "Label : value" row inside a section. Returns the value label so the
        caller can colour it.
    #>
    param(
        [System.Windows.Forms.GroupBox]$Box,
        [string]$Key, [string]$Value, [int]$Y,
        [System.Drawing.Color]$Colour = [System.Drawing.Color]::Empty,
        [int]$KeyWidth = 130
    )
    $k = New-Object System.Windows.Forms.Label
    $k.Text     = $Key
    $k.Location = New-Object System.Drawing.Point(12, $Y)
    $k.Size     = New-Object System.Drawing.Size($KeyWidth, 18)
    $k.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    $Box.Controls.Add($k)

    $v = New-Object System.Windows.Forms.Label
    $v.Tag      = 'NoTheme'
    $v.Text     = $Value
    $v.Location = New-Object System.Drawing.Point((12 + $KeyWidth), $Y)
    $v.Size     = New-Object System.Drawing.Size(($Box.Width - $KeyWidth - 30), 18)
    $v.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $v.AutoEllipsis = $true
    if ($Colour -ne [System.Drawing.Color]::Empty) { $v.ForeColor = $Colour }
    $Box.Controls.Add($v)
    return $v
}

function New-DashList {
    param([System.Windows.Forms.GroupBox]$Box, [int]$Y, [int]$H, [string[]]$Columns, [int[]]$Widths)
    $lv = New-Object System.Windows.Forms.ListView
    $lv.Tag           = 'NoTheme'
    $lv.View          = [System.Windows.Forms.View]::Details
    $lv.FullRowSelect = $true
    $lv.GridLines     = $false
    $lv.Location      = New-Object System.Drawing.Point(12, $Y)
    $lv.Size          = New-Object System.Drawing.Size(($Box.Width - 24), $H)
    $lv.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    for ($i = 0; $i -lt $Columns.Count; $i++) {
        [void]$lv.Columns.Add($Columns[$i], $Widths[$i])
    }
    $Box.Controls.Add($lv)
    return $lv
}

function Start-ToolDetached {
    <#
        Launch another tool in a NEW process.

        Not with '&' in this one. The dashboard runs inside ShowDialog(), and a
        modal form disables every other window on its thread -- a WinForms tool
        started from here would draw and then refuse input, which is section 14d
        and cost days the first time.

        Every path is quoted. Start-Process does not quote for you, so one space
        in a folder name turns a path into two arguments (PRC001).
    #>
    param([string]$ScriptName)
    $path = Join-Path $PSScriptRoot $ScriptName
    if (-not (Test-Path -LiteralPath $path)) {
        [System.Windows.Forms.MessageBox]::Show("$ScriptName is not in this folder.", 'Not found') | Out-Null
        return
    }
    try {
        Start-Process -FilePath 'powershell.exe' `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$path`"") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Could not launch $ScriptName`r`n$($_.Exception.Message)", 'Launch failed') | Out-Null
    }
}
#endregion

#region Form --------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Admin Dashboard'
$form.Size          = New-Object System.Drawing.Size(946, 860)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

# --- Row 1 ---------------------------------------------------------------
$boxEnv    = New-DashSection -Title 'Environment'    -X $script:ColL -Y 42 -W $script:ColW -H 196
$boxHealth = New-DashSection -Title 'Toolkit Health' -X $script:ColR -Y 42 -W $script:ColW -H 196
$form.Controls.Add($boxEnv)
$form.Controls.Add($boxHealth)

$lblTenant  = Add-DashLine -Box $boxEnv -Key 'Tenant'      -Value '...' -Y 24
$lblExo     = Add-DashLine -Box $boxEnv -Key 'Exchange Online' -Value '...' -Y 44
$lblSpo     = Add-DashLine -Box $boxEnv -Key 'SharePoint'  -Value '...' -Y 64
$lblVer     = Add-DashLine -Box $boxEnv -Key 'Version'     -Value '...' -Y 88

$lblFiles   = Add-DashLine -Box $boxHealth -Key 'Files verified' -Value '...' -Y 24
$lblDrift   = Add-DashLine -Box $boxHealth -Key 'Differences'    -Value '...' -Y 44
# One line, not one row per file. An earlier version listed three filenames by
# hand -- and got one of them wrong, so a file that did not exist reported as
# 'Missing' rather than as a typo. Here the list comes from the folder, so it
# cannot disagree with reality and does not need editing when a config is added.
$lblCfg     = Add-DashLine -Box $boxHealth -Key 'Config files'   -Value '...' -Y 68
$lblRel     = Add-DashLine -Box $boxHealth -Key 'Released'       -Value '...' -Y 92

$btnMgr = New-Object System.Windows.Forms.Button
$btnMgr.Text = 'Tool Manager'
$btnMgr.Location = New-Object System.Drawing.Point(12, 158)
$btnMgr.Size = New-Object System.Drawing.Size(110, 26)
$btnMgr.Add_Click({ Start-ToolDetached -ScriptName 'ToolManager.ps1' })
$boxHealth.Controls.Add($btnMgr)

$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = 'Run Tester'
$btnTest.Location = New-Object System.Drawing.Point(130, 158)
$btnTest.Size = New-Object System.Drawing.Size(110, 26)
$btnTest.Add_Click({ Start-ToolDetached -ScriptName 'Test-Toolkit.ps1' })
$boxHealth.Controls.Add($btnTest)

# --- Row 2 ---------------------------------------------------------------
# No Retry Queue panel. That panel read disable-retry.csv, which belongs to a
# disable pipeline this toolkit does not have. An empty panel headed 'Retry
# Queue' would imply a queue exists and is clear, which is a different and
# false statement.
$boxToday = New-DashSection -Title "Today"                  -X $script:ColL -Y 246 -W $script:ColW -H 196
$boxTop   = New-DashSection -Title 'Top Tools - last 30 days' -X $script:ColR -Y 246 -W $script:ColW -H 196
$form.Controls.Add($boxToday)
$form.Controls.Add($boxTop)

$lblActions = Add-DashLine -Box $boxToday -Key 'Actions'   -Value '...' -Y 24
$lblOk      = Add-DashLine -Box $boxToday -Key 'Success'   -Value '...' -Y 44
$lblPart    = Add-DashLine -Box $boxToday -Key 'Partial'   -Value '...' -Y 64
$lblFail    = Add-DashLine -Box $boxToday -Key 'Failed'    -Value '...' -Y 84
$lblUnk     = Add-DashLine -Box $boxToday -Key 'Unknown'   -Value '...' -Y 104
$lblLast    = Add-DashLine -Box $boxToday -Key 'Last action' -Value '...' -Y 128
$lblLastAt  = Add-DashLine -Box $boxToday -Key 'At'        -Value '...' -Y 148

$btnLog = New-Object System.Windows.Forms.Button
$btnLog.Text = 'Log Viewer'
$btnLog.Location = New-Object System.Drawing.Point(12, 168)
$btnLog.Size = New-Object System.Drawing.Size(110, 24)
$btnLog.Add_Click({ Start-ToolDetached -ScriptName 'LogViewer.ps1' })
$boxToday.Controls.Add($btnLog)


# --- Row 3 ---------------------------------------------------------------
# Failure Hotspots spans the full width now that Retry Queue is gone. The
# alternative was an empty quadrant, which reads as something failed to load.
$boxHot  = New-DashSection -Title 'Failure Hotspots - last 30 days' -X $script:ColL -Y 450 -W 904 -H 226
$form.Controls.Add($boxHot)

$lvTop = New-DashList -Box $boxTop -Y 24 -H 150 -Columns @('Tool', 'Runs') -Widths @(300, 80)
$lvHot = New-DashList -Box $boxHot -Y 24 -H 150 -Columns @('Tool', 'Failed', 'Partial') -Widths @(500, 100, 100)

$lblTopNote = Add-DashLine -Box $boxTop -Key '' -Value '' -Y 180 -KeyWidth 0
$lblTopNote.ForeColor = $script:ColDim
$lblTopNote.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)

$lblHotNote = Add-DashLine -Box $boxHot -Key '' -Value '' -Y 180 -KeyWidth 0
$lblHotNote.ForeColor = $script:ColDim
$lblHotNote.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)

# --- Footer --------------------------------------------------------------
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Tag      = 'NoTheme'
$lblStatus.Location = New-Object System.Drawing.Point(12, 686)
$lblStatus.Size     = New-Object System.Drawing.Size(700, 18)
$lblStatus.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)
$lblStatus.ForeColor = $script:ColDim
$lblStatus.AutoEllipsis = $true
$form.Controls.Add($lblStatus)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = 'Refresh'
$btnRefresh.Location = New-Object System.Drawing.Point(796, 680)
$btnRefresh.Size = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnRefresh)

# Created here, ABOVE every Update- function that calls SetToolTip. The
# console footer had this the other way round today and a null-method
# throw got reported as 'version unavailable'.
$toolTip = New-Object System.Windows.Forms.ToolTip

# 30px of clear space at the top before the first panel. The help '?' is
# placed at ClientSize-40 by Add-ToolHelpButton, which put it at x880 y8 --
# straight through the Toolkit Health group box. Measured, then moved, rather
# than trusting the button's run-time search to find a gap (14l).
$form.ClientSize = New-Object System.Drawing.Size(920, 720)
#endregion

#region Population --------------------------------------------------------------
function Update-EnvironmentPanel {
    param([switch]$Refresh)
    try {
        # The console's own probe, not a second implementation. -SkipSpo on the
        # first pass for the same reason the console skips it: Get-SPOTenant is
        # a network round trip and this page should open immediately.
        $map = if ($Refresh) { Get-CapabilityStateMap -Refresh } else { Get-CapabilityStateMap -SkipSpo }

        # ExO and SPO only. The capability map also probes on-premises Exchange
        # and AD, but this toolkit is cloud-only in the tenants it ships to, so
        # those rows would be permanently 'Not connected' -- and a panel with
        # rows that are always red teaches people to stop reading it.
        foreach ($pair in @(
            @{ Lbl = $lblExo; Key = 'ExO' },
            @{ Lbl = $lblSpo; Key = 'SPO' })) {

            $st = $map[$pair.Key]
            if (-not $st -or $st.Unknown) {
                $pair.Lbl.Text = 'Not checked'; $pair.Lbl.ForeColor = $script:ColDim
            } elseif ($st.Available) {
                $pair.Lbl.Text = 'Connected';   $pair.Lbl.ForeColor = $script:ColGood
            } else {
                $pair.Lbl.Text = 'Not connected'; $pair.Lbl.ForeColor = $script:ColBad
            }
        }
    } catch {
        foreach ($l in @($lblExo, $lblSpo)) {
            $l.Text = 'Unknown'; $l.ForeColor = $script:ColDim
        }
    }

    # Same two calls the console header makes, in the same order. The mismatch
    # test matters more here than there: this page is read at a glance at the
    # start of the day, which is exactly when "connected to the wrong tenant"
    # is cheapest to catch and most expensive to miss.
    try {
        $d = Get-ConnectedTenantDomain
        if ($d) {
            $lblTenant.Text = "$d"; $lblTenant.ForeColor = $script:ColGood
        } else {
            $lblTenant.Text = 'Not connected'; $lblTenant.ForeColor = $script:ColDim
        }
        $mismatch = Test-TenantDomainMismatch
        if ($mismatch) {
            # Amber and non-blocking, matching the console: a tenant with several
            # accepted domains can legitimately disagree with its onmicrosoft
            # prefix, so this is a question, not a fault.
            $lblTenant.Text = "$($lblTenant.Text)  (?)"
            $lblTenant.ForeColor = $script:ColWarn
            $script:TenantNote = $mismatch
        } else {
            $script:TenantNote = ''
        }
        $toolTip.SetToolTip($lblTenant, "$script:TenantNote")
    } catch { $lblTenant.Text = 'Unknown'; $lblTenant.ForeColor = $script:ColDim }

}

function Update-HealthPanel {
    try {
        $v = Test-ToolkitIntegrity
        if (-not $v.HasManifest) {
            $lblFiles.Text = 'Not stamped'; $lblFiles.ForeColor = $script:ColDim
            $lblDrift.Text = 'Run Publish-Toolkit.ps1 on the master'; $lblDrift.ForeColor = $script:ColDim
            $lblVer.Text   = 'unversioned';  $lblVer.ForeColor = $script:ColDim
            $lblRel.Text   = '-'
        } else {
            $total = $v.Matched + $v.Modified.Count + $v.Missing.Count
            $lblFiles.Text = "$($v.Matched) of $total"
            $bad = $v.Modified.Count + $v.Missing.Count + $v.Extra.Count
            if ($bad -eq 0) {
                $lblFiles.ForeColor = $script:ColGood
                $lblDrift.Text = 'None'; $lblDrift.ForeColor = $script:ColGood
            } else {
                $lblFiles.ForeColor = $script:ColWarn
                $bits = @()
                if ($v.Modified.Count) { $bits += "$($v.Modified.Count) modified" }
                if ($v.Missing.Count)  { $bits += "$($v.Missing.Count) missing" }
                if ($v.Extra.Count)    { $bits += "$($v.Extra.Count) extra" }
                $lblDrift.Text = ($bits -join ', ')
                $lblDrift.ForeColor = $script:ColWarn
            }
            $lblVer.Text = "v$($v.Version)"
            $lblVer.ForeColor = $(if ($bad -eq 0) { $script:ColGood } else { $script:ColWarn })
            $lblRel.Text = "$($v.Released)"
        }
    } catch {
        $lblFiles.Text = 'Unknown'; $lblFiles.ForeColor = $script:ColDim
        $lblDrift.Text = 'Could not check'; $lblDrift.ForeColor = $script:ColDim
    }

    # Present and parseable are different questions, and the second is the one
    # that bites -- a truncated JSON is present.
    try {
        $cfgDir = Split-Path -Parent (Get-ConfigFilePath -Name 'Tools.json')
        $cfgs = @(Get-ChildItem -LiteralPath $cfgDir -Filter '*.json' -File -ErrorAction Stop)
        if ($cfgs.Count -eq 0) {
            $lblCfg.Text = 'none found'; $lblCfg.ForeColor = $script:ColWarn
        } else {
            $bad = @()
            foreach ($c in $cfgs) {
                try { $null = Get-Content -LiteralPath $c.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
                catch { $bad += $c.Name }
            }
            if ($bad.Count -eq 0) {
                $lblCfg.Text = "$($cfgs.Count) of $($cfgs.Count) valid"
                $lblCfg.ForeColor = $script:ColGood
            } else {
                $lblCfg.Text = "$($cfgs.Count - $bad.Count) of $($cfgs.Count) valid -- $($bad -join ', ')"
                $lblCfg.ForeColor = $script:ColBad
            }
        }
    } catch {
        $lblCfg.Text = 'Could not read Config'; $lblCfg.ForeColor = $script:ColDim
    }
}

function Update-LogPanels {
    param($Log)

    $rows = @($Log.Rows)
    if ($Log.Problem) {
        foreach ($l in @($lblActions, $lblOk, $lblPart, $lblFail, $lblUnk)) {
            $l.Text = '-'; $l.ForeColor = $script:ColDim
        }
        $lblLast.Text = $Log.Problem; $lblLast.ForeColor = $script:ColWarn
        $lblLastAt.Text = ''
        $lblTopNote.Text = $Log.Problem
        $lblHotNote.Text = $Log.Problem
        return
    }

    # ---- Today ----------------------------------------------------------
    $today = (Get-Date).Date
    $todayRows = @($rows | Where-Object {
        $d = Get-RowDate -Row $_
        $d -and $d.Date -eq $today
    })

    if ($todayRows.Count -eq 0) {
        foreach ($l in @($lblOk, $lblPart, $lblFail, $lblUnk)) { $l.Text = '0'; $l.ForeColor = $script:ColDim }
        $lblActions.Text = '0'; $lblActions.ForeColor = $script:ColDim
        $lblLast.Text = 'No activity logged today.'; $lblLast.ForeColor = $script:ColDim
        $lblLastAt.Text = ''
    } else {
        $ok   = @($todayRows | Where-Object { $_.Result -eq 'Success' }).Count
        $part = @($todayRows | Where-Object { $_.Result -eq 'Partial' }).Count
        $fail = @($todayRows | Where-Object { $_.Result -eq 'Failed'  }).Count
        # Anything not one of the three. Counted, never absorbed into Success.
        $unk  = $todayRows.Count - $ok - $part - $fail

        $lblActions.Text = "$($todayRows.Count)"; $lblActions.ForeColor = [System.Drawing.Color]::Empty
        $lblOk.Text   = "$ok";   $lblOk.ForeColor   = $(if ($ok)   { $script:ColGood } else { $script:ColDim })
        $lblPart.Text = "$part"; $lblPart.ForeColor = $(if ($part) { $script:ColWarn } else { $script:ColDim })
        $lblFail.Text = "$fail"; $lblFail.ForeColor = $(if ($fail) { $script:ColBad }  else { $script:ColDim })
        $lblUnk.Text  = "$unk";  $lblUnk.ForeColor  = $(if ($unk)  { $script:ColDim }  else { $script:ColDim })

        $last = $todayRows | Sort-Object { Get-RowDate -Row $_ } | Select-Object -Last 1
        $desc = @("$($last.Action)", "$($last.Target)") | Where-Object { $_ }
        if (-not $desc) { $desc = @("$($last.Detail)") }
        $lblLast.Text = ((@(Get-RowToolName -Row $last) + $desc) -join ' - ')
        $lblLast.ForeColor = [System.Drawing.Color]::Empty
        $d = Get-RowDate -Row $last
        $lblLastAt.Text = $(if ($d) { $d.ToString('HH:mm') } else { '' })
        if ("$($last.Result)") { $lblLastAt.Text += "   ($($last.Result))" }
    }

    # ---- 30-day windows -------------------------------------------------
    $since = (Get-Date).Date.AddDays(-30)
    $recent = @($rows | Where-Object {
        $d = Get-RowDate -Row $_
        $d -and $d -ge $since
    })

    $lvTop.Items.Clear()
    $lvHot.Items.Clear()

    if ($recent.Count -eq 0) {
        $lblTopNote.Text = 'No tool usage found for the selected period.'
        $lblHotNote.Text = 'No failures found for the selected period.'
        return
    }

    $byTool = $recent | Group-Object { Get-RowToolName -Row $_ } |
              Sort-Object Count -Descending | Select-Object -First 5
    foreach ($g in $byTool) {
        $it = New-Object System.Windows.Forms.ListViewItem($g.Name)
        [void]$it.SubItems.Add("$($g.Count)")
        [void]$lvTop.Items.Add($it)
    }
    $noScript = @($recent | Where-Object { -not "$($_.Script)".Trim() }).Count
    $lblTopNote.Text = $(if ($noScript) {
        "$($recent.Count) actions. $noScript have no script name (converted rows)."
    } else {
        "$($recent.Count) actions in the last 30 days."
    })

    $bad = @($recent | Where-Object { $_.Result -eq 'Failed' -or $_.Result -eq 'Partial' })
    if ($bad.Count -eq 0) {
        $lblHotNote.Text = 'No failures found for the selected period.'
    } else {
        $byBad = $bad | Group-Object { Get-RowToolName -Row $_ } |
                 Sort-Object { @($_.Group | Where-Object { $_.Result -eq 'Failed' }).Count } -Descending |
                 Select-Object -First 5
        foreach ($g in $byBad) {
            $f = @($g.Group | Where-Object { $_.Result -eq 'Failed'  }).Count
            $p = @($g.Group | Where-Object { $_.Result -eq 'Partial' }).Count
            $it = New-Object System.Windows.Forms.ListViewItem($g.Name)
            [void]$it.SubItems.Add("$f")
            [void]$it.SubItems.Add("$p")
            [void]$lvHot.Items.Add($it)
        }
        $lblHotNote.Text = "$($bad.Count) failed or partial in the last 30 days."
    }
}


function Update-Dashboard {
    param([switch]$Refresh)
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # Each panel is independent on purpose. One unreadable file must not
        # take the other five down -- that is the difference between a page
        # that is partly useful and a page that is blank.
        Update-EnvironmentPanel -Refresh:$Refresh
        Update-HealthPanel
        Update-LogPanels -Log (Get-DashLogRows)
        $lblStatus.Text = "Refreshed $(Get-Date -Format 'HH:mm:ss').  Read-only -- nothing on this page changes anything."
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

$btnRefresh.Add_Click({ Update-Dashboard -Refresh })
#endregion

Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Dashboard' | Out-Null

Update-Dashboard

Write-CompletedLog -Action 'View' -Target 'Dashboard' -Result 'Success' -Detail 'Opened the dashboard'

[void]$form.ShowDialog()
