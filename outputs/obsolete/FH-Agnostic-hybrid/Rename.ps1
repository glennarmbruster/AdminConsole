#Requires -Version 5.1
<#
    Rename.ps1
    Packages this toolkit for email and unpacks it at the far end. 1 Package -- append .txt to every blocked extension, write a manifest 2 Restore -- stri

    Reasoning: ScriptNotes/Rename.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-RUNS: nothing. It renames files in a folder on this machine.
# TOOLKIT-HELP-BEGIN
# Packages this toolkit so it survives being emailed, and unpacks it again at
# the other end.
#
#   PACKAGE   appends .txt to every extension mail filters block, and writes a
#             manifest of what was renamed
#   RESTORE   reads the manifest and puts the original extensions back
#
# Mail gateways strip .ps1 attachments silently. Renaming is a transport
# workaround and not a security measure -- the recipient still has to trust
# where the files came from.
#
# RESTORE FIRST, RUN SECOND. A .ps1.txt file looks like a script and is not
# one; PowerShell will not execute it and the error does not explain why.
#
# Console tool, no window -- so no ? button of its own. Touches only this
# folder and connects nowhere.
# TOOLKIT-HELP-END

# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
$ExcludedFolders = @('reports', 'Logs', 'captures', '__pycache__', '.git')

$BlockedExtensions = @(
    '.ps1', '.psm1', '.psd1', '.ps1xml',
    '.json', '.xml', '.config',
    '.cmd', '.bat', '.vbs', '.js'
)

$SelfNames = @('Rename.ps1', 'Rename.txt', 'ShipManifest.json')

$ManifestName = 'ShipManifest.json'
$ManifestPath = Join-Path $PSScriptRoot $ManifestName

# $IsWindows only exists in PowerShell 6+. In 5.1 it is undefined, and 5.1 only
# runs on Windows, so undefined means Windows.
if ($null -eq $IsWindows) { $script:OnWindows = $true }
else                      { $script:OnWindows = [bool]$IsWindows }

# ---------------------------------------------------------------------------

function Write-Head {
    param([string]$Text)
    Write-Host ''
    Write-Host $Text -ForegroundColor Cyan
    Write-Host ('-' * $Text.Length) -ForegroundColor DarkCyan
}

function Get-RelativePath {
    param([string]$FullPath)
    $root = $PSScriptRoot.TrimEnd([IO.Path]::DirectorySeparatorChar)
    $p = $FullPath
    if ($p.StartsWith($root)) { $p = $p.Substring($root.Length).TrimStart('\', '/') }
    return $p.Replace('/', '\')
}

function Resolve-RelativePath {
    param([string]$Relative)
    $parts = $Relative.Split('\')
    $out = $PSScriptRoot
    foreach ($part in $parts) { $out = Join-Path $out $part }
    return $out
}

function Get-ShippableFile {
    $all = Get-ChildItem -LiteralPath $PSScriptRoot -Recurse -File -Force -ErrorAction SilentlyContinue
    $keep = @()
    foreach ($f in $all) {
        $rel = Get-RelativePath $f.FullName
        $skip = $false

        foreach ($ex in $ExcludedFolders) {
            if ($rel -like "$ex\*" -or $rel -eq $ex) { $skip = $true; break }
        }
        if ($skip) { continue }
        if ($SelfNames -contains $f.Name) { continue }
        if ($f.Name -eq '.DS_Store') { continue }

        $keep += $f
    }
    return $keep
}

function Remove-MacJunk {
    $junk = Get-ChildItem -LiteralPath $PSScriptRoot -Recurse -Force -File -Filter '.DS_Store' -ErrorAction SilentlyContinue
    $n = 0
    foreach ($j in $junk) {
        try { Remove-Item -LiteralPath $j.FullName -Force -ErrorAction Stop; $n++ }
        catch { Write-Host "  could not remove $(Get-RelativePath $j.FullName)" -ForegroundColor DarkYellow }
    }
    if ($n -gt 0) { Write-Host "  removed $n .DS_Store file(s)" -ForegroundColor DarkGray }
}

function Get-Sha {
    param([string]$Path)
    try { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash }
    catch { return $null }
}

function Test-Ps1Encoding {
    param([string]$Path)
    try { $bytes = [IO.File]::ReadAllBytes($Path) } catch { return 'unreadable' }
    if ($bytes.Length -lt 4) { return 'empty or truncated' }

    $bom = ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    if (-not $bom) { return 'missing UTF-8 BOM' }
    if ($bytes.Length -ge 6 -and $bytes[3] -eq 0xEF -and $bytes[4] -eq 0xBB -and $bytes[5] -eq 0xBF) {
        return 'doubled BOM'
    }

    $hasCrLf = $false
    for ($i = 0; $i -lt $bytes.Length - 1; $i++) {
        if ($bytes[$i] -eq 0x0D -and $bytes[$i + 1] -eq 0x0A) { $hasCrLf = $true; break }
    }
    if (-not $hasCrLf) { return 'LF line endings (expected CRLF)' }

    return $null
}

# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
function Invoke-Package {
    Write-Head 'Package for email'
    Remove-MacJunk

    $files = Get-ShippableFile
    if ($files.Count -eq 0) { Write-Host '  nothing to package.' -ForegroundColor Yellow; return }

    $entries = @()
    $renamed = 0
    $already = 0

    foreach ($f in $files) {
        $rel = Get-RelativePath $f.FullName
        $ext = [IO.Path]::GetExtension($f.Name)

        # Idempotent: a file already carrying .txt over a blocked extension is
        # recorded under its ORIGINAL name and left alone, so running this twice
        # never produces Common.ps1.txt.txt.
        if ($ext -eq '.txt') {
            $inner = [IO.Path]::GetExtension([IO.Path]::GetFileNameWithoutExtension($f.Name))
            if ($BlockedExtensions -contains $inner) {
                $entries += [PSCustomObject]@{
                    Path    = $rel.Substring(0, $rel.Length - 4)
                    Sha256  = Get-Sha $f.FullName
                    Bytes   = $f.Length
                    Renamed = $true
                }
                $already++
                continue
            }
        }

        $needsRename = ($BlockedExtensions -contains $ext)
        $entries += [PSCustomObject]@{
            Path    = $rel
            Sha256  = Get-Sha $f.FullName
            Bytes   = $f.Length
            Renamed = $needsRename
        }

        if ($needsRename) {
            try {
                Rename-Item -LiteralPath $f.FullName -NewName ($f.Name + '.txt') -ErrorAction Stop
                $renamed++
            } catch {
                Write-Host "  FAILED $rel -- $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    if ($script:OnWindows) { $builtOn = 'Windows' } else { $builtOn = 'macOS/Linux' }

    $manifest = [PSCustomObject]@{
        Created     = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        CreatedOn   = $builtOn
        PSVersion   = $PSVersionTable.PSVersion.ToString()
        FileCount   = $entries.Count
        RenameCount = ($entries | Where-Object { $_.Renamed }).Count
        Files       = $entries
    }
    $manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8

    Write-Host ''
    Write-Host "  files catalogued : $($entries.Count)"
    Write-Host "  renamed to .txt  : $renamed"
    if ($already -gt 0) { Write-Host "  already renamed  : $already" -ForegroundColor DarkGray }
    Write-Host "  manifest written : $ManifestName"
    Write-Host ''
    Write-Host '  Now rename this script by hand: Rename.ps1 -> Rename.txt' -ForegroundColor Yellow
    Write-Host '  Send the folder WITHOUT reports\, Logs\ or captures\.' -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
function Invoke-Restore {
    Write-Head 'Restore to production names'

    $haveManifest = Test-Path -LiteralPath $ManifestPath
    if (-not $haveManifest) {
        # Not fatal: stripping .txt still recovers a working toolkit, without
        # the integrity check -- and the operator is told so.
        Write-Host "  $ManifestName not found -- restoring by name only." -ForegroundColor Yellow
        Write-Host '  No integrity verification will be possible.' -ForegroundColor Yellow
    }

    $restored = 0
    $stragglers = Get-ChildItem -LiteralPath $PSScriptRoot -Recurse -File -Force -Filter '*.txt' -ErrorAction SilentlyContinue
    foreach ($f in $stragglers) {
        if ($SelfNames -contains $f.Name) { continue }
        $inner = [IO.Path]::GetExtension([IO.Path]::GetFileNameWithoutExtension($f.Name))
        if ($BlockedExtensions -notcontains $inner) { continue }

        $target = $f.FullName.Substring(0, $f.FullName.Length - 4)
        if (Test-Path -LiteralPath $target) {
            Write-Host "  SKIP $(Get-RelativePath $f.FullName) -- target already exists" -ForegroundColor Yellow
            continue
        }
        try {
            Rename-Item -LiteralPath $f.FullName -NewName ([IO.Path]::GetFileName($target)) -ErrorAction Stop
            $restored++
        } catch {
            Write-Host "  FAILED $(Get-RelativePath $f.FullName) -- $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Write-Host "  restored to original names : $restored"

    if (-not $haveManifest) {
        Write-Host ''
        Write-Host '  Run 3 to unblock, then spot-check a script opens correctly.' -ForegroundColor Yellow
        return
    }

    # --- verification ------------------------------------------------------
    Write-Host ''
    Write-Host '  Verifying against manifest...'
    $m = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json

    $missing = @(); $changed = @(); $encoding = @(); $ok = 0
    foreach ($e in @($m.Files)) {
        $full = Resolve-RelativePath $e.Path
        if (-not (Test-Path -LiteralPath $full)) { $missing += $e.Path; continue }

        $sha = Get-Sha $full
        if ($sha -ne $e.Sha256) { $changed += $e.Path; continue }

        if ([IO.Path]::GetExtension($e.Path) -in @('.ps1', '.psm1', '.psd1')) {
            $fault = Test-Ps1Encoding $full
            if ($fault) { $encoding += "$($e.Path) -- $fault"; continue }
        }
        $ok++
    }

    Write-Host ''
    Write-Host "  intact           : $ok of $($m.FileCount)" -ForegroundColor Green
    if ($missing.Count)  { Write-Host "  MISSING          : $($missing.Count)"  -ForegroundColor Red;    $missing  | ForEach-Object { Write-Host "      $_" -ForegroundColor Red } }
    if ($changed.Count)  { Write-Host "  ALTERED IN TRANSIT: $($changed.Count)" -ForegroundColor Red;    $changed  | ForEach-Object { Write-Host "      $_" -ForegroundColor Red } }
    if ($encoding.Count) { Write-Host "  ENCODING DAMAGED : $($encoding.Count)" -ForegroundColor Red;    $encoding | ForEach-Object { Write-Host "      $_" -ForegroundColor Red } }

    Write-Host ''
    if ($missing.Count -or $changed.Count -or $encoding.Count) {
        Write-Host '  DO NOT DEPLOY. Something rewrote files in transit.' -ForegroundColor Red
        Write-Host '  Re-send inside a password-protected zip -- that survives' -ForegroundColor Red
        Write-Host '  gateways that rewrite text attachments.' -ForegroundColor Red
    } else {
        Write-Host '  All files match the manifest. Run 3 to unblock.' -ForegroundColor Green
    }
}

# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
function Invoke-UnblockAll {
    Write-Head 'Unblock (clear mark-of-the-web)'

    if (-not $script:OnWindows) {
        Write-Host '  Not Windows -- nothing to do.' -ForegroundColor Yellow
        Write-Host '  The zone marker is an NTFS stream. Run this after copying' -ForegroundColor DarkGray
        Write-Host '  the folder to the target machine.' -ForegroundColor DarkGray
        return
    }

    $files = Get-ChildItem -LiteralPath $PSScriptRoot -Recurse -File -Force -ErrorAction SilentlyContinue
    $n = 0
    foreach ($f in $files) {
        try { Unblock-File -LiteralPath $f.FullName -ErrorAction Stop; $n++ }
        catch { Write-Host "  could not unblock $(Get-RelativePath $f.FullName)" -ForegroundColor DarkYellow }
    }
    Write-Host "  processed $n file(s)."
    Write-Host ''
    Write-Host '  Unblock-File is not the same as execution policy. If scripts' -ForegroundColor DarkGray
    Write-Host '  still refuse to run, that is Set-ExecutionPolicy, separately.' -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
function Show-Menu {
    Write-Host ''
    Write-Host '=====================================================' -ForegroundColor Cyan
    Write-Host '  Toolkit packaging' -ForegroundColor Cyan
    Write-Host '=====================================================' -ForegroundColor Cyan
    Write-Host "  Folder : $PSScriptRoot"
    if ($script:OnWindows) { Write-Host '  Host   : Windows' }
    else                   { Write-Host '  Host   : macOS / Linux  (option 3 unavailable)' }
    Write-Host ''
    Write-Host '  1  Package for email   -- rename to .txt, write manifest'
    Write-Host '  2  Restore             -- rename back, verify integrity'
    Write-Host '  3  Unblock all files   -- Windows only'
    Write-Host '  Q  Quit'
    Write-Host ''
}

do {
    Show-Menu
    $choice = (Read-Host '  Choose').Trim().ToUpper()
    switch ($choice) {
        '1' { Invoke-Package }
        '2' { Invoke-Restore }
        '3' { Invoke-UnblockAll }
        'Q' { }
        default { Write-Host '  Not an option.' -ForegroundColor Yellow }
    }
    if ($choice -ne 'Q') {
        Write-Host ''
        [void](Read-Host '  Press Enter')
    }
} while ($choice -ne 'Q')

Write-Host ''
