#Requires -Version 5.1
# ---------------------------------------------------------------------------
<#
    Deploy.ps1
    Build a handout copy of this toolkit, beside this one, and zip it.

    e:\scripts\m365\adminconsole
        -> e:\scripts\m365\adminconsole3.1_2026-08-13_0930
        -> e:\scripts\m365\adminconsole3.1_2026-08-13_0930.zip

    NO BUTTON, ON PURPOSE. This is not a tool, it is how the toolkit is handed
    out, and it belongs in the hands of whoever holds the master. It is also
    excluded from the copy it makes, so a deployed console cannot deploy
    onward -- there is one source, and any folder can be traced back to it.

    Author   : Glenn Armbruster
    Revised  : 2026-08-13
#>
# TOOLKIT-HELP-BEGIN
# Builds a handout copy of this toolkit and zips it.
#
# One build, everyone gets the same one. It creates a staging folder beside
# this one --
#
#     adminconsole3.1_2026-08-13_0930
#
# -- version and timestamp, so you can see at a glance what you shipped and
# when, and an older build is still there if you need to go back to it.
#
# Then it zips that folder and drops a README.txt inside for whoever receives
# it. Hand over the .zip; they extract it to a folder of their own choosing and
# keep that path.
#
# NO NAME PROMPT. It used to ask who the copy was for and bake the name into
# the folder. That was wrong twice over: the build is identical whoever gets
# it, and a name baked into a folder is stale the moment it lands on their
# machine -- Levi does not want to run from adminconsole_levi_1308 forever.
# The timestamp is a BUILD label, not an install path.
#
# WHAT IS LEFT OUT, AND WHY:
#
#   Logs, Reports, RetryQueue   your history, not theirs. They start empty.
#   DisabledTools               tools deliberately out of service here.
#   Config\Settings.json        YOUR paths and email address. Per-engineer.
#                               Everything else in Config is tenant fact and
#                               does travel.
#   Publish-Toolkit.ps1         stamping is the master's job.
#   Deploy.ps1                  copies do not make copies.
#   *Cache*.json                scan caches. Stale data is worse than none;
#                               their console rebuilds them on first use.
#
# Unblock.ps1 IS included -- it is the first thing they will need, because
# files out of a downloaded zip arrive blocked by Windows.
#
# NOTHING IS EVER OVERWRITTEN. Every run makes a new timestamped folder, so a
# build can be reissued, compared, or gone back to. Delete old ones yourself
# when you are done with them.
#
# The manifest (Version.json) is REBUILT for the copy rather than copied, so
# their self-test does not report the excluded files as missing forever.
# TOOLKIT-HELP-END

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Rules ------------------------------------------------------------------

# Folders skipped wholesale, matched on the FIRST path segment only. A folder
# named Logs nested three deep inside Projects is that project's business.
$script:SkipFolders = @('DisabledTools', 'Logs', 'RetryQueue', 'Reports')

# Files skipped wherever they appear.
$script:SkipFiles = @('Publish-Toolkit.ps1', 'Deploy.ps1', 'Version.json')

# Skipped only at this exact relative path -- Settings.json elsewhere would be
# somebody's data file, not this one.
$script:SkipExactPaths = @('Config\Settings.json')

# Scan caches. Not in the original brief; called out here and named in the
# report so the decision is visible rather than assumed. DLInventoryCache.json
# is a snapshot of this tenant's distribution groups -- handing over a stale
# copy of it is worse than handing over none, and the tool rebuilds it.
$script:SkipPatterns = @('*Cache*.json')

function Test-ShouldSkip {
    <#
        Returns the REASON a relative path is skipped, or '' to copy it.

        A reason rather than a boolean, because the report names why each item
        was left out. "37 files skipped" invites a shrug; "Config\Settings.json
        -- yours, not theirs" does not.
    #>
    param([Parameter(Mandatory)][string]$RelativePath)

    $rel   = $RelativePath.TrimStart('\')
    $first = ($rel -split '\\')[0]
    $leaf  = Split-Path -Leaf $rel

    foreach ($d in $script:SkipFolders) {
        if ($first -ieq $d) { return "$d\ is not carried over" }
    }
    foreach ($p in $script:SkipExactPaths) {
        if ($rel -ieq $p) { return 'yours, not theirs' }
    }
    foreach ($f in $script:SkipFiles) {
        if ($leaf -ieq $f) {
            if ($f -eq 'Version.json') { return 'rebuilt for the copy' }
            return 'master only'
        }
    }
    foreach ($pat in $script:SkipPatterns) {
        if ($leaf -like $pat) { return 'scan cache -- rebuilt on first use' }
    }
    return ''
}

#endregion

#region Copy -------------------------------------------------------------------

function Get-RelativeTo {
    param([string]$Root, [string]$FullPath)
    $r = $Root.TrimEnd('\')
    if ($FullPath.StartsWith($r, [StringComparison]::OrdinalIgnoreCase)) {
        return $FullPath.Substring($r.Length).TrimStart('\')
    }
    return (Split-Path -Leaf $FullPath)
}

function Invoke-Deploy {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Target,
        [scriptblock]$Progress
    )

    $copied  = New-Object System.Collections.ArrayList
    $skipped = New-Object System.Collections.ArrayList

    $all = @(Get-ChildItem -LiteralPath $Source -Recurse -File -ErrorAction Stop)
    $i = 0
    foreach ($f in $all) {
        $i++
        $rel = Get-RelativeTo -Root $Source -FullPath $f.FullName
        $why = Test-ShouldSkip -RelativePath $rel
        if ($why) {
            [void]$skipped.Add([PSCustomObject]@{ Path = $rel; Why = $why })
            continue
        }

        if ($Progress -and ($i % 10 -eq 0)) {
            & $Progress "Copying $rel" ([int](100 * $i / [Math]::Max(1, $all.Count)))
        }

        $dest = Join-Path $Target $rel
        $dir  = Split-Path -Parent $dest
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -Path $dir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
        [void]$copied.Add($rel)
    }

    return [PSCustomObject]@{ Copied = @($copied); Skipped = @($skipped) }
}

function Write-CopyManifest {
    <#
        A fresh Version.json describing what was ACTUALLY copied.

        THIS USED TO BE THE ONLY CORRECT MANIFEST FOR A COPY. The master's
        listed Publish-Toolkit.ps1 and Deploy.ps1, which the copy deliberately
        does not have, so handing it over gave the recipient a MISSING list on
        every self-test -- correct behaviour reported as damage, and a check
        that is always wrong is a check people stop reading.

        Those two are now excluded from the manifest at source
        (Get-ManifestExcludedFile), so the master's and the copy's describe the
        same file list and the master's CAN be handed over. That is what makes
        shipping a single patched script with Version.json workable.

        THE REBUILD IS KEPT ANYWAY, and not out of caution for its own sake: it
        hashes what actually landed in the target. A copy that was interrupted,
        or that hit a locked file, produces a manifest that matches the folder
        it describes rather than the folder it was meant to be -- and the
        recipient's first self-test then tells them so.

        Version and release date are carried across unchanged: this IS that
        version, just a subset of it. Only the file list and hashes are rebuilt.

        Paths are computed against the TARGET root. Get-ToolkitRelativePath in
        Common.ps1 resolves against $PSScriptRoot, which here is the SOURCE --
        using it would silently write correct-looking relative paths measured
        from the wrong folder.
    #>
    param([Parameter(Mandatory)][string]$Target, [string]$SourceVersionPath)

    $version  = '0.0'
    $released = (Get-Date).ToString('yyyy-MM-dd')
    $tenant   = ''
    if ($SourceVersionPath -and (Test-Path -LiteralPath $SourceVersionPath)) {
        try {
            $old = Get-Content -LiteralPath $SourceVersionPath -Raw | ConvertFrom-Json
            if ($old.Version)  { $version  = "$($old.Version)" }
            if ($old.Released) { $released = "$($old.Released)" }
            if ($old.Tenant)   { $tenant   = "$($old.Tenant)" }
        } catch { }
    }

    # Same coverage rule as Get-ToolkitCodeFile: root *.ps1 minus the
    # master-only tools, plus Projects.
    #
    # The target never contains those two -- SkipFiles already kept them out of
    # the copy -- so this filter changes nothing here today. It is applied
    # anyway so the rule READS the same in both places. The one thing that must
    # never happen to these two implementations is that they diverge, and a rule
    # written differently in two files is a rule that eventually is.
    $skip  = Get-ManifestExcludedFile
    $files = New-Object System.Collections.ArrayList
    foreach ($f in (Get-ChildItem -LiteralPath $Target -Filter '*.ps1' -File -ErrorAction SilentlyContinue)) {
        if ($skip -contains $f.Name) { continue }
        [void]$files.Add($f.FullName)
    }
    $proj = Join-Path $Target 'Projects'
    if (Test-Path -LiteralPath $proj) {
        foreach ($f in (Get-ChildItem -LiteralPath $proj -Filter '*.ps1' -File -Recurse -ErrorAction SilentlyContinue)) {
            [void]$files.Add($f.FullName)
        }
    }

    $hashes = @{}
    foreach ($f in @($files.ToArray() | Sort-Object)) {
        $rel = Get-RelativeTo -Root $Target -FullPath $f
        $h   = Get-FileHashHex -Path $f
        if ($h) { $hashes[$rel] = $h }
    }

    $manifest = [PSCustomObject]@{
        Version  = $version
        Released = $released
        Tenant   = $tenant
        Files    = $hashes
    }
    $path = Join-Path $Target 'Version.json'
    Set-Content -LiteralPath $path -Value ($manifest | ConvertTo-Json -Depth 4) -Encoding UTF8
    return $hashes.Count
}

function Write-RecipientReadme {
    <#
        The note that lands in the zip.

        Written for someone who has never seen this toolkit and is about to
        double-click something. Plain sentences, numbered steps, and the two
        facts that otherwise cause a support message on day one: files arrive
        BLOCKED, and an empty Logs folder is correct rather than broken.

        Deliberately .txt. It opens in Notepad on any machine with no rendering,
        no markdown, and no "which app should open this".
    #>
    param([Parameter(Mandatory)][string]$Target, [string]$Version, [string]$BuildName)

    $nl = [Environment]::NewLine
    $sb = New-Object System.Text.StringBuilder

    # NO TENANT NAME HERE. This file ships to every tenant, and this line
    # opened with one until the Tester caught it as TEN002 on 2026-08-13 -- a
    # hardcoded tenant value in a tenant-agnostic script, exactly what that
    # rule exists for. The build name below already says which toolkit this is,
    # and it is derived rather than typed.
    #
    # The offending word is not repeated here either: TEN002 reads content, and
    # a comment explaining a tenant leak is still a tenant leak as far as a
    # scanner is concerned.
    [void]$sb.AppendLine("ADMIN CONSOLE TOOLKIT  --  version $Version")
    [void]$sb.AppendLine("Build: $BuildName")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('===========================================================================')
    [void]$sb.AppendLine('IF THIS IS YOUR FIRST COPY')
    [void]$sb.AppendLine('===========================================================================')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('1. Extract this zip somewhere you control and can write to. For example:')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('       E:\Scripts\M365\AdminConsole')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('   Pick a path you are happy to keep, and NOT Downloads -- the console')
    [void]$sb.AppendLine('   writes its logs and reports beside itself.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('   The folder inside this zip has a date in its name. That is the BUILD')
    [void]$sb.AppendLine('   label, not a name you have to keep. Rename it once extracted --')
    [void]$sb.AppendLine('   AdminConsole is a good choice, and see the upgrade note below for why')
    [void]$sb.AppendLine('   the name you pick is worth sticking to.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('2. Run Unblock.ps1 -- FIRST, BEFORE ANYTHING ELSE.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('   Right-click it and choose "Run with PowerShell".')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('   Windows marks every file that came out of a zip as untrusted. Until')
    [void]$sb.AppendLine('   you do this, scripts either refuse to run or stop and ask you to')
    [void]$sb.AppendLine('   confirm each one -- which looks exactly like the toolkit being')
    [void]$sb.AppendLine('   broken. It is not. This is one step and it is the fix.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('3. Run AdminConsole.ps1. That is the launcher; everything else opens')
    [void]$sb.AppendLine('   from it.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('4. Fill in your own settings. In the console:')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('       Dashboard  ->  Tool Manager  ->  Settings...')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('   This is where your paths and your email address go. The copy arrived')
    [void]$sb.AppendLine('   with none, on purpose -- they are per-engineer, and you should not be')
    [void]$sb.AppendLine('   running with someone else''s.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('===========================================================================')
    [void]$sb.AppendLine('IF YOU ALREADY HAVE A COPY -- UPGRADING')
    [void]$sb.AppendLine('===========================================================================')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Do NOT extract over the top of your existing folder, and do not retype')
    [void]$sb.AppendLine('your settings. Move them across:')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('1. In your CURRENT console, run ConfigBackup.ps1 and take a backup.')
    [void]$sb.AppendLine('   That captures Tools.json, Settings.json, Sections.json,')
    [void]$sb.AppendLine('   QuickLaunch.json, Themes.json and Tenant.json.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('   Backups are written OUTSIDE the console folder, so replacing the')
    [void]$sb.AppendLine('   folder cannot take them with it.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('2. Rename your old console folder out of the way -- AdminConsole_old,')
    [void]$sb.AppendLine('   for instance. Keep it until you are happy with the new one.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('3. Extract this zip and rename the extracted folder to EXACTLY the name')
    [void]$sb.AppendLine('   your old folder had.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('   THIS MATTERS. Backups are filed under the console folder''s name. A')
    [void]$sb.AppendLine('   console in a differently-named folder looks in a different place and')
    [void]$sb.AppendLine('   will show you no backups at all -- they are not lost, it is simply')
    [void]$sb.AppendLine('   not looking where they are.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('4. Run Unblock.ps1 in the new folder.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('5. Run ConfigBackup.ps1 in the new folder, pick your backup from the')
    [void]$sb.AppendLine('   list, and restore it. Your settings and layout come back as they')
    [void]$sb.AppendLine('   were.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('6. Open the console and check it looks right before deleting the old')
    [void]$sb.AppendLine('   folder.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('===========================================================================')
    [void]$sb.AppendLine('WHAT TO EXPECT ON FIRST RUN')
    [void]$sb.AppendLine('===========================================================================')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('  - Logs, Reports and the retry queue are EMPTY. That is correct. They')
    [void]$sb.AppendLine('    are your history, and you have not made any yet.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('  - The console writes its own settings file the first time it starts,')
    [void]$sb.AppendLine('    with your paths and your email address. Nobody else''s settings came')
    [void]$sb.AppendLine('    with this copy.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('  - Buttons you have no connection for are greyed out or hidden. Connect')
    [void]$sb.AppendLine('    first -- the buttons along the top -- and they come to life.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('  - The console does not grant you anything. It runs as you, with the')
    [void]$sb.AppendLine('    permissions you already have. If a task needs a role you have not')
    [void]$sb.AppendLine('    activated, it will fail the same way it would by hand.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('===========================================================================')
    [void]$sb.AppendLine('FINDING YOUR WAY AROUND')
    [void]$sb.AppendLine('===========================================================================')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('  - Every window has a "?" in the top-right corner. It explains what')
    [void]$sb.AppendLine('    that tool does, what it touches, and what will bite you. Read it')
    [void]$sb.AppendLine('    before using a tool for the first time.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('  - Test-Toolkit.ps1 checks this copy against itself and reports')
    [void]$sb.AppendLine('    anything wrong with it. It connects to nothing and changes nothing,')
    [void]$sb.AppendLine('    so it is safe to run at any time. Run it if something seems off.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('  - Every change any tool makes is written to a log beside the console.')
    [void]$sb.AppendLine('    Successes and failures both.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('===========================================================================')
    [void]$sb.AppendLine('IF SOMETHING IS WRONG')
    [void]$sb.AppendLine('===========================================================================')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Send Glenn:')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("  - the build line at the top of this file ($BuildName)")
    [void]$sb.AppendLine('  - what you clicked')
    [void]$sb.AppendLine('  - the message you got, or a screenshot')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Do not edit the .ps1 files to work around something. Say so instead --')
    [void]$sb.AppendLine('a copy that has drifted from the master is very hard to support, and')
    [void]$sb.AppendLine('the next build will overwrite your change anyway.')
    [void]$sb.AppendLine('')

    $path = Join-Path $Target 'README.txt'
    # ASCII, CRLF, no BOM -- this one opens in Notepad on someone else's machine
    # rather than being dot-sourced, so the toolkit's BOM convention does not
    # apply and a BOM would show as three stray characters in some viewers.
    [System.IO.File]::WriteAllText($path, ($sb.ToString() -replace "`r?`n", "`r`n"),
        (New-Object System.Text.ASCIIEncoding))
    return $path
}

function Compress-Build {
    <#
        Zip the staged folder, INCLUDING the folder itself.

        Compress-Archive -Path <folder> keeps the top-level folder inside the
        archive; -Path <folder>\* would not, and the recipient would get a
        hundred and sixty loose files sprayed into whatever directory they
        happened to extract into. That is a bad first thirty seconds.
    #>
    param([Parameter(Mandatory)][string]$Folder)
    $zip = "$Folder.zip"
    if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force -ErrorAction Stop }
    Compress-Archive -Path $Folder -DestinationPath $zip -CompressionLevel Optimal -ErrorAction Stop
    return $zip
}

#endregion

#region Form -------------------------------------------------------------------

$SourceRoot = $PSScriptRoot.TrimEnd('\')
$ParentDir  = Split-Path -Parent $SourceRoot
$SourceLeaf = Split-Path -Leaf $SourceRoot

function Get-SourceVersion {
    <#
        The version this build is OF, read from the master's own manifest.

        Falls back to '' rather than to a made-up number. A folder called
        adminconsole0.0_... would be a lie about a stamp that was never
        applied; adminconsole_... is merely quiet about it, and the report says
        so out loud.
    #>
    $vp = Join-Path $SourceRoot 'Version.json'
    if (-not (Test-Path -LiteralPath $vp)) { return '' }
    try {
        $v = Get-Content -LiteralPath $vp -Raw | ConvertFrom-Json
        return "$($v.Version)".Trim()
    } catch { return '' }
}

function Get-BuildFolderName {
    # adminconsole3.1_2026-08-13_0930
    #
    # yyyy-MM-dd_HHmm, matching every other dated thing the toolkit writes --
    # reports, log files, the undo folders. It also sorts chronologically in
    # Explorer, which ddmm does not, and is not ambiguous in November.
    $v = Get-SourceVersion
    $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
    return ('{0}{1}_{2}' -f $SourceLeaf, $v, $stamp)
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Deploy'
$form.ClientSize = New-Object System.Drawing.Size(760, 470)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblTop = New-Object System.Windows.Forms.Label
$lblTop.Text = 'Builds a handout copy of this toolkit and zips it. One build -- everyone gets the same one. Your logs, reports, retry queue and settings are not included.'
$lblTop.Location = New-Object System.Drawing.Point(12, 12)
$lblTop.Size = New-Object System.Drawing.Size(736, 32)
$lblTop.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblTop)

$lblFrom = New-Object System.Windows.Forms.Label
$lblFrom.Text = "From:  $SourceRoot"
$lblFrom.Location = New-Object System.Drawing.Point(12, 50)
$lblFrom.Size = New-Object System.Drawing.Size(736, 18)
$form.Controls.Add($lblFrom)

$lblTo = New-Object System.Windows.Forms.Label
$lblTo.Location = New-Object System.Drawing.Point(12, 70)
$lblTo.Size = New-Object System.Drawing.Size(736, 18)
$form.Controls.Add($lblTo)

$btnDeploy = New-Object System.Windows.Forms.Button
$btnDeploy.Text = 'Build and Zip'
$btnDeploy.Location = New-Object System.Drawing.Point(12, 96)
$btnDeploy.Size = New-Object System.Drawing.Size(130, 28)
$btnDeploy.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnDeploy)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(152, 102)
$lblStatus.Size = New-Object System.Drawing.Size(596, 18)
$form.Controls.Add($lblStatus)

$pbar = New-Object System.Windows.Forms.ProgressBar
$pbar.Location = New-Object System.Drawing.Point(12, 130)
$pbar.Size = New-Object System.Drawing.Size(736, 12)
$form.Controls.Add($pbar)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(12, 152)
$txtOut.Size = New-Object System.Drawing.Size(736, 268)
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap = $false
$txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor = 'Top,Left,Right,Bottom'
$form.Controls.Add($txtOut)

$btnFolder = New-Object System.Windows.Forms.Button
$btnFolder.Text = 'Open Folder'
$btnFolder.Location = New-Object System.Drawing.Point(12, 432)
$btnFolder.Size = New-Object System.Drawing.Size(110, 26)
$btnFolder.Enabled = $false
$btnFolder.Anchor = 'Bottom,Left'
$form.Controls.Add($btnFolder)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(648, 432)
$btnClose.Size = New-Object System.Drawing.Size(100, 26)
$btnClose.Anchor = 'Bottom,Right'
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$script:LastTarget = ''

function Update-TargetLabel {
    # Recomputed on show AND after each build, because the name carries a
    # timestamp -- a label written once would be stale by the second build and
    # would name a folder that is not the one about to be created.
    $lblTo.Text = "To:  $(Join-Path $ParentDir (Get-BuildFolderName))  (+ .zip)"
    $lblTo.ForeColor = [System.Drawing.Color]::FromArgb(0, 110, 0)
}

$btnFolder.Add_Click({ if ($script:LastTarget) { Open-ReportFolder -Path $script:LastTarget } })

$btnDeploy.Add_Click({
    $buildName = Get-BuildFolderName
    $target    = Join-Path $ParentDir $buildName
    $version   = Get-SourceVersion
    if (-not $version) { $version = '(unstamped)' }

    $btnDeploy.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        # A timestamped name to the minute makes this all but impossible, but
        # two builds inside one minute would otherwise merge into each other
        # silently and produce a zip of mixed content.
        if (Test-Path -LiteralPath $target) {
            throw "'$target' already exists. Wait a minute and build again, or delete it."
        }
        New-Item -Path $target -ItemType Directory -Force -ErrorAction Stop | Out-Null

        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar
        $r = Invoke-Deploy -Source $SourceRoot -Target $target -Progress $progress

        & $progress 'Rebuilding the manifest...' 90
        $hashed = Write-CopyManifest -Target $target -SourceVersionPath (Join-Path $SourceRoot 'Version.json')

        & $progress 'Writing README.txt...' 94
        $readme = Write-RecipientReadme -Target $target -Version $version -BuildName $buildName

        & $progress 'Zipping...' 96
        $zip = Compress-Build -Folder $target
        $zipMb = [math]::Round((Get-Item -LiteralPath $zip).Length / 1MB, 1)

        & $progress 'Done.' 100
        $script:LastTarget = $ParentDir
        $btnFolder.Enabled = $true
        $lblStatus.Text = "Built $($r.Copied.Count) file(s) -- $zipMb MB zip"

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("BUILD  $buildName")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("  Version        : $version")
        [void]$sb.AppendLine("  Staged folder  : $target")
        [void]$sb.AppendLine("  HAND THIS OVER : $zip   ($zipMb MB)")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("  Files copied   : $($r.Copied.Count)")
        [void]$sb.AppendLine("  Files skipped  : $($r.Skipped.Count)")
        [void]$sb.AppendLine("  Manifest       : $hashed script(s) hashed")
        [void]$sb.AppendLine("  Readme         : $(Split-Path -Leaf $readme)")
        [void]$sb.AppendLine('')

        $tops = @($r.Copied | ForEach-Object {
            $p = ($_ -split '\\')
            if ($p.Count -gt 1) { $p[0] } else { '(root)' }
        } | Sort-Object -Unique)
        [void]$sb.AppendLine("  Folders included: $($tops -join ', ')")
        [void]$sb.AppendLine('')

        [void]$sb.AppendLine('LEFT OUT:')
        foreach ($grp in ($r.Skipped | Group-Object Why | Sort-Object Name)) {
            [void]$sb.AppendLine("  $($grp.Name)  --  $($grp.Count) item(s)")
            foreach ($s in ($grp.Group | Select-Object -First 6)) {
                [void]$sb.AppendLine("      $($s.Path)")
            }
            if ($grp.Count -gt 6) { [void]$sb.AppendLine("      ... and $($grp.Count - 6) more") }
        }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('The staged folder is only there so you can look inside it before you')
        [void]$sb.AppendLine('send the zip. Delete it whenever you like -- the zip is the artefact.')
        $txtOut.Text = $sb.ToString()

        Write-CompletedLog -Action 'Deploy' -Target $buildName -Result 'Success' `
            -Detail "$($r.Copied.Count) file(s), $zipMb MB -> $zip"
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Deploy'
        Write-FailedLog -Action 'Deploy' -Target $buildName -ErrorRecord $_ -Detail "target: $target"
    } finally {
        $btnDeploy.Enabled = $true
        $form.Cursor = $prev
        Complete-ProgressReporter -Activity 'Deploy'
        Update-TargetLabel
    }
})

Update-TargetLabel

Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Deploy' | Out-Null
Set-ConsoleTheme -Form $form
Set-FormResizable -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
$form.Dispose()
#endregion
