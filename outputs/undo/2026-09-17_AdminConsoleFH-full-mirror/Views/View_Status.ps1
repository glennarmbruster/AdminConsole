#Requires -Version 5.1
# VIEW-ID: Status
# VIEW-NAME: Status board
# VIEW-NOTE: Connections, PIM, integrity, retry queue, recent activity.
# VIEW-FUNCTION: Show-ConsoleStatusView
# VIEW-ORDER: 40
<#
    View_Status.ps1
    Admin Console view: Status board

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    NOT A VIEW OF TOOLS -- a view of the environment the tools run in.
    Which services this window is signed in to, what PIM is active, what
    the last integrity run said, what is sitting in the retry queue, and
    what has been run lately. The answer to 'why did that tool fail' is
    usually on this screen.

    THIS FILE IS NOT RUN. It is DOT-SOURCED by AdminConsole.ps1, which is why
    it defines a function and does nothing else. Running it on its own defines
    Show-ConsoleStatusView into that prompt and exits; nothing appears, and nothing
    is wrong. Open the console instead.

    WHAT IT MAY USE. Everything the console has already defined by the time it
    is sourced: Get-ConsoleToolModel, New-ConsoleViewShell,
    Update-ConsoleViewShell, Add-ConsoleViewFontStepper, Get-ConsoleViewColors,
    Invoke-ConsoleViewTool, New-ConsoleToolTile, the theme helpers in
    Common.ps1, and $form -- the console window, used as the dialog owner.

    ADDING A VIEW. Copy this file, change the five VIEW- lines at the top and
    the function name, and drop it in this folder. The console finds it by
    reading those lines; there is no list to edit. Keep the VIEW-ID short and
    unique -- it is what a per-user default view is stored as.
#>

function Show-ConsoleStatusView {
    <#
        NOT A VIEW OF TOOLS. A view of the ENVIRONMENT the tools run in --
        what is connected, how long PIM has left, whether this copy of the
        toolkit still matches its manifest, what is sitting in the retry queue,
        and what failed recently.

        THE MORNING CHECK. Every fact here is already somewhere in the console
        or a log file; the value is having them on one screen at the point
        where the answer changes what you do next. Read-only by design -- a
        board that can also change things is a control panel, and a control
        panel is what the other views are for.
    #>
    $sh = New-ConsoleViewShell -Title 'Status' -Hint 'Read-only.  Refresh re-runs every probe.  Esc returns to the console.' -ViewId 'Status'
    $v = $sh.Form
    $col = Get-ConsoleViewColors -Form $v

    $btnRef = New-Object System.Windows.Forms.Button
    $btnRef.Text = 'Refresh'
    $sh.Bar.Controls.Add($btnRef)
    [void]$sh.Right.Add($btnRef)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true; $txt.ReadOnly = $true; $txt.ScrollBars = 'Vertical'
    $txt.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txt.WordWrap = $false
    $sh.Body.Controls.Add($txt)

    $build = {
        $v.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $sb = New-Object System.Text.StringBuilder
        $line = { param([string]$s) [void]$sb.AppendLine($s) }
        try {
            & $line 'CONNECTIONS'
            & $line ''
            foreach ($p in @(
                @{ N = 'Exchange Online'; T = { (Test-ExoInProcess).Present } },
                @{ N = 'Exchange on-prem'; T = { (Test-EmsAvailable).Available } },
                @{ N = 'Active Directory'; T = { (Test-AdAvailable).Available } })) {
                $ok = $false
                try { $ok = [bool](& $p.T) } catch { }
                & $line ("  {0,-22} {1}" -f $p.N, $(if ($ok) { 'connected' } else { 'NOT connected' }))
            }
            $tenant = ''
            try { $tenant = "$(Get-ConnectedTenantDomain)" } catch { }
            & $line ("  {0,-22} {1}" -f 'Tenant', $(if ($tenant) { $tenant } else { '(not known)' }))
            $pim = ''
            try { $pim = "$(Get-PimStatusLine)" } catch { }
            & $line ("  {0,-22} {1}" -f 'PIM', $(if ($pim) { $pim } else { 'no activation recorded' }))
            & $line ''

            & $line 'TOOLKIT'
            & $line ''
            try {
                $vv = Test-ToolkitIntegrity
                if (-not $vv.HasManifest) {
                    & $line '  unversioned copy (no Version.json)'
                } else {
                    & $line ("  {0,-22} v{1}   released {2}" -f 'Version', $vv.Version, $vv.Released)
                    & $line ("  {0,-22} {1} verified, {2} modified, {3} missing, {4} extra" -f 'Files', $vv.Matched, @($vv.Modified).Count, @($vv.Missing).Count, @($vv.Extra).Count)
                    foreach ($m in @($vv.Modified | Select-Object -First 8)) { & $line "      modified: $m" }
                    foreach ($m in @($vv.Missing  | Select-Object -First 8)) { & $line "      MISSING : $m" }
                }
            } catch { & $line "  integrity check failed: $($_.Exception.Message)" }
            & $line ''

            & $line 'TOOLS'
            & $line ''
            $model = Get-ConsoleToolModel
            $blocked = @($model | Where-Object { $_.Blocked })
            & $line ("  {0,-22} {1}" -f 'Total', @($model).Count)
            & $line ("  {0,-22} {1}" -f 'Blocked right now', @($blocked).Count)
            foreach ($g in @($blocked | Group-Object Blocked | Sort-Object Count -Descending)) {
                & $line ("      {0,-3} {1}" -f $g.Count, $g.Name)
            }
            $missing = @($model | Where-Object { -not (Test-Path -LiteralPath $_.Path) })
            if ($missing.Count) {
                & $line ("  {0,-22} {1}" -f 'Script file missing', $missing.Count)
                foreach ($m in @($missing | Select-Object -First 10)) { & $line "      $($m.Name)  ->  $(Split-Path -Leaf $m.Path)" }
            }
            & $line ''

            & $line 'RETRY QUEUE'
            & $line ''
            try {
                $q = Get-RetryQueuePath
                $files = @()
                if ($q -and (Test-Path -LiteralPath $q)) { $files = @(Get-ChildItem -LiteralPath $q -Filter '*.csv' -File -ErrorAction SilentlyContinue) }
                if ($files.Count -eq 0) { & $line '  empty' }
                else { foreach ($fl in $files) { & $line ("  {0,-34} {1} row(s)   {2}" -f $fl.Name, (@(Import-Csv -LiteralPath $fl.FullName -ErrorAction SilentlyContinue).Count), $fl.LastWriteTime) } }
            } catch { & $line "  could not be read: $($_.Exception.Message)" }
            & $line ''

            & $line 'RECENT ACTIVITY   (from the audit log)'
            & $line ''
            try {
                $lp = Join-Path (Get-LogPath) 'CompletedTasks.csv'
                if (-not (Test-Path -LiteralPath $lp)) { & $line '  no audit log yet' }
                else {
                    $rows = @(Import-Csv -LiteralPath $lp -ErrorAction Stop)
                    $tail = @($rows | Select-Object -Last 12)
                    [array]::Reverse($tail)
                    foreach ($r in $tail) {
                        & $line ("  {0}  {1,-10} {2,-26} {3,-8} {4}" -f "$($r.Timestamp)", "$($r.Operator)", "$($r.Target)", "$($r.Result)", "$($r.Action)")
                    }
                    $fails = @($rows | Where-Object { "$($_.Result)" -match 'Fail' })
                    & $line ''
                    & $line ("  {0} row(s) in the log, {1} recorded as failed." -f $rows.Count, $fails.Count)
                }
            } catch { & $line "  could not be read: $($_.Exception.Message)" }
        } catch {
            & $line ''
            & $line "Status build failed: $($_.Exception.Message)"
        } finally {
            $v.Cursor = [System.Windows.Forms.Cursors]::Default
        }
        $txt.Text = $sb.ToString()
        $txt.Select(0, 0)
    }

    $btnRef.Add_Click({ & $build })
    Add-ConsoleViewFontStepper -Shell $sh -OnChange {
        $txt.Font = New-Object System.Drawing.Font('Consolas', ([int]$v.FontSize))
    }
    $v.Add_Resize({ try { Update-ConsoleViewShell -Shell $sh } catch { } })

    Set-ConsoleTheme -Form $v
    Update-ConsoleViewShell -Shell $sh
    # Monospaced, and not negotiable: every line above is column-aligned with
    # composite formatting, which a proportional font turns into a mess.
    $txt.Font = New-Object System.Drawing.Font('Consolas', ([int]$v.FontSize))
    & $build
    [void]$v.ShowDialog($form)
}
