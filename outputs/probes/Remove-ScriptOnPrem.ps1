#Requires -Version 5.1
# ---------------------------------------------------------------------------
<#
    Remove-ScriptOnPrem.ps1

    Strips the retired ScriptOnPrem field from your Tools.json, and reports the
    scrambled twins it finds on the way out.

    OPTIONAL. Nothing needs this -- the console ignores ScriptOnPrem now, and
    ToolManager drops it the next time you Save. This just does it in one step,
    and more usefully SHOWS you what the drift had done, which a silent Save
    would not.

    Writes a .bak beside the file before changing anything.

    Author   : Glenn Armbruster
    Revised  : 2026-08-13
#>

param(
    [string]$Path,
    # See what would change without changing it.
    [switch]$WhatIfOnly
)

. (Join-Path (Split-Path -Parent $PSScriptRoot) 'AdminConsole\Common.ps1') -ErrorAction SilentlyContinue

if (-not $Path) {
    try { $Path = Get-ConfigFilePath -Name 'Tools.json' } catch { }
}
if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
    Write-Host "Tools.json not found. Pass -Path explicitly." -ForegroundColor Red
    return
}

Write-Host ''
Write-Host "Tools.json: $Path" -ForegroundColor Cyan
Write-Host ('-' * 72) -ForegroundColor DarkGray

$raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
try {
    $parsed = $raw | ConvertFrom-Json
} catch {
    Write-Host "  Not valid JSON: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Accepts a bare array or a { Tools: [...] } wrapper -- both shapes exist in
# the wild, and guessing wrong here would rewrite the file into the other one.
$wrapped = $false
$tools = $null
if ($parsed -is [array]) {
    $tools = @($parsed)
} elseif ($parsed.PSObject.Properties['Tools']) {
    $tools = @($parsed.Tools); $wrapped = $true
} else {
    $tools = @($parsed)
}

$hits = @($tools | Where-Object { "$($_.ScriptOnPrem)".Trim() })

if ($hits.Count -eq 0) {
    Write-Host '  No ScriptOnPrem values found. Nothing to do.' -ForegroundColor Green
    return
}

Write-Host "  $($hits.Count) tool(s) still carry a twin:" -ForegroundColor Yellow
Write-Host ''
foreach ($t in $hits) {
    # The point of printing this: a twin whose name has nothing to do with the
    # tool is one the reorder bug moved. That is the evidence, and it is worth
    # seeing once before it is deleted.
    $script = [System.IO.Path]::GetFileNameWithoutExtension(("$($t.Script)" -replace '/', '\'))
    $twin   = [System.IO.Path]::GetFileNameWithoutExtension(("$($t.ScriptOnPrem)" -replace '/', '\'))
    $looksWrong = ($twin -notmatch [regex]::Escape($script))
    $flag = if ($looksWrong) { '   <-- unrelated; almost certainly drift' } else { '' }
    $colour = if ($looksWrong) { 'Red' } else { 'Gray' }
    Write-Host ("    {0,-28} {1,-26} -> {2}{3}" -f "$($t.Name)", $script, $twin, $flag) -ForegroundColor $colour
}

Write-Host ''
if ($WhatIfOnly) {
    Write-Host '  -WhatIfOnly: nothing written.' -ForegroundColor Cyan
    return
}

$bak = "$Path.bak-$(Get-Date -Format 'yyyy-MM-dd_HHmm')"
Copy-Item -LiteralPath $Path -Destination $bak -Force
Write-Host "  Backup: $bak" -ForegroundColor DarkGray

$clean = foreach ($t in $tools) {
    $o = [ordered]@{}
    foreach ($p in $t.PSObject.Properties) {
        if ($p.Name -eq 'ScriptOnPrem') { continue }
        $o[$p.Name] = $p.Value
    }
    [PSCustomObject]$o
}

$out = if ($wrapped) { [PSCustomObject]@{ Tools = @($clean) } } else { @($clean) }

# Depth 6: ObjectTypes is an array inside an object inside the array. The
# default of 2 would render it as a type name string and quietly destroy it.
$json = $out | ConvertTo-Json -Depth 6
Set-Content -LiteralPath $Path -Value $json -Encoding UTF8

Write-Host "  Removed ScriptOnPrem from $($hits.Count) tool(s)." -ForegroundColor Green
Write-Host ''
Write-Host '  Open the console and check the buttons still point where you expect.' -ForegroundColor Cyan
Write-Host '  Any AD- tool you actually want a button for needs its own row in' -ForegroundColor Cyan
Write-Host '  Tool Manager, or a shared Group value with its cloud counterpart.' -ForegroundColor Cyan
Write-Host ''
