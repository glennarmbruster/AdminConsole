#Requires -Version 5.1
<#
    FixCommon.ps1
    Check -- and where it is provably safe, repair -- a Common.ps1 that has
    been damaged by mail link-rewriting in transit.

    Author   : Glenn Armbruster
    Revised  : 2026-09-21

    THE PROBLEM THIS EXISTS FOR (see ARCHITECTURE 71).

    Common.ps1 is mailed to the other site as Common.ps1.txt. The .txt rename
    gets it past attachment filters; it does NOTHING about link rewriting,
    which happens to the CONTENT. Proofpoint URL Defense does not know the
    difference between a link in a message and a string literal in a script,
    so a line like

        return "http://$configured/PowerShell/"

    arrives as a rewritten wrapper. Get-ExchangeServerUri then returns an
    unreachable address and Connect-ExSE fails on a line that looks perfectly
    correct in every editor unless you scroll right. It is not cosmetic.

    WHY THIS SCRIPT DOES NOT SIMPLY DECODE AND HOPE.

    The v3 wrapper replaces certain characters in the URL with '*' and stores
    them base64-encoded in the tail. Decoding the single-character case is
    exact. There is also a RUN-LENGTH form for consecutive replacements that
    this script does not implement, because a decoder that is subtly wrong
    produces a URL that looks right and is not -- which is the same failure
    the rewrite itself causes, and no improvement at all.

    So the trust here comes from arithmetic, not from cleverness:

        1. The SHA-256 of the known-good file is embedded below.
        2. If the file on disk matches it, there is nothing to do. Certain.
        3. If it does not, the damage is identified and the SAFE repairs are
           applied to a copy.
        4. The repaired copy is hashed again. If it now matches the reference,
           the repair is PROVEN correct -- not assumed -- and is written out.
        5. If it does not match, the original is left untouched and the script
           says so. A file it cannot prove it fixed is a file it will not
           silently change.

    THE HASH MUST BE UPDATED WHEN Common.ps1 CHANGES. Ship this file and
    Common.ps1 together, always. A stale hash makes this report a false alarm
    on a perfectly good file, which is the failure mode most likely to get the
    whole check ignored.

    NOTE ON THIS SCRIPT'S OWN SAFETY IN TRANSIT: the wrapper host is never
    written here as a complete URL. It is assembled at run time from pieces,
    so that mailing THIS file does not get its own detection patterns
    rewritten by the very product it detects.
#>

[CmdletBinding()]
param(
    # The file to check. Defaults to Common.ps1 beside this script.
    [string]$Path,

    # Without this, the script only reports. Nothing is written.
    [switch]$Apply
)

# ---------------------------------------------------------------------------
# THE REFERENCE. Regenerate with:
#     (Get-FileHash .\Common.ps1 -Algorithm SHA256).Hash
# ---------------------------------------------------------------------------
$script:GoodHash  = '6ACAAB33E41ABA207EEA29CB1D16E9B3AA12CA4DB622336CB9E6B6859D7A82E8'
$script:GoodBytes = 924441
$script:GoodStamp = '2026-09-25'

if (-not $Path) { $Path = Join-Path $PSScriptRoot 'Common.ps1' }

function Write-Head { param([string]$T) Write-Host ''; Write-Host $T -ForegroundColor Cyan }
function Write-Good { param([string]$T) Write-Host "  $T" -ForegroundColor Green }
function Write-Bad  { param([string]$T) Write-Host "  $T" -ForegroundColor Red }
function Write-Warn { param([string]$T) Write-Host "  $T" -ForegroundColor Yellow }
function Write-Info { param([string]$T) Write-Host "  $T" }

# Assembled, never written whole -- see the note in the header.
$script:WrapHost = 'urldefense' + '.' + 'com'

Write-Head "FixCommon -- checking $Path"

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Bad "Not found: $Path"
    Write-Info 'Run this from the folder Common.ps1 lives in, or pass -Path.'
    exit 2
}

$item = Get-Item -LiteralPath $Path
$hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash

Write-Info ("size    : {0:N0} bytes   (reference {1:N0})" -f $item.Length, $script:GoodBytes)
Write-Info ("sha256  : $hash")
Write-Info ("expected: $($script:GoodHash)   (build $($script:GoodStamp))")

if ($hash -eq $script:GoodHash) {
    Write-Good 'MATCH -- this file is byte-for-byte the shipped original. Nothing to do.'
    exit 0
}

Write-Bad 'DOES NOT MATCH the shipped original. Looking for known damage...'

# ---------------------------------------------------------------------------
# READ AS BYTES, NOT AS TEXT.
#
# Get-Content -Raw decodes, and re-encoding on write can change the BOM or the
# line endings on its own -- which would be a second corruption layered on the
# one being repaired. The file is UTF-8 with a BOM and CRLF line endings, and
# it must still be exactly that afterwards.
# ---------------------------------------------------------------------------
$bytes = [System.IO.File]::ReadAllBytes($Path)
$enc   = New-Object System.Text.UTF8Encoding($true)
$text  = $enc.GetString($bytes)
if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }

$findings = New-Object System.Collections.Generic.List[string]

# --- damage 1: the rewrite wrapper -----------------------------------------
$rxWrap = [regex]("(?i)https?://[^\s""']*" + [regex]::Escape($script:WrapHost) + "/v3/__(.+?)__;(.*?)!!.*?\$")
$wraps  = $rxWrap.Matches($text)
if ($wraps.Count) { $findings.Add("$($wraps.Count) rewritten URL wrapper(s)") }

# v2 took a different shape; detected but never auto-repaired.
$rxV2 = [regex]("(?i)https?://[^\s""']*" + [regex]::Escape($script:WrapHost) + "[^\s""']*?/v2/url\?")
$v2   = $rxV2.Matches($text)
if ($v2.Count) { $findings.Add("$($v2.Count) OLD-STYLE (v2) wrapper(s) -- not auto-repaired") }

# --- damage 2: a URL string that gained a trailing space --------------------
# ARCHITECTURE 71: one string picked up a trailing space, which no wrapper
# search would ever have found.
$rxSpace = [regex]'"(?<u>[^"]*https?://[^"]*?) +"'
$spaces  = $rxSpace.Matches($text)
if ($spaces.Count) { $findings.Add("$($spaces.Count) URL string(s) ending in a space") }

if ($findings.Count -eq 0) {
    Write-Head 'No damage this script recognises.'
    Write-Warn 'The file differs from the original but not in a way this script can identify.'
    Write-Info 'It may be a NEWER build than this script knows about -- check the date above.'
    Write-Info 'If it should be the same build, ask for a fresh copy rather than editing it.'
    exit 1
}

Write-Head 'Found:'
foreach ($f in $findings) { Write-Bad $f }

# ---------------------------------------------------------------------------
# REPAIR, ON A COPY
# ---------------------------------------------------------------------------
$fixed   = $text
$skipped = 0

foreach ($m in $wraps) {
    $inner   = $m.Groups[1].Value
    $payload = $m.Groups[2].Value

    if ($inner -match '\*\*') {
        # Run-length form -- see the header. Not implemented, not guessed at.
        $skipped++
        continue
    }

    $url = $inner
    if ($payload) {
        try {
            # URL-safe base64: '-' and '_' stand in for '+' and '/'.
            $b64 = $payload.Replace('-', '+').Replace('_', '/')
            switch ($b64.Length % 4) { 2 { $b64 += '==' } 3 { $b64 += '=' } }
            $chars = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($b64)).ToCharArray()
            $i = 0
            $sb = New-Object System.Text.StringBuilder
            foreach ($c in $url.ToCharArray()) {
                if ($c -eq '*' -and $i -lt $chars.Length) { [void]$sb.Append($chars[$i]); $i++ }
                else { [void]$sb.Append($c) }
            }
            $url = $sb.ToString()
        } catch {
            $skipped++
            continue
        }
    }
    $fixed = $fixed.Replace($m.Value, $url)
}

foreach ($m in $spaces) {
    $fixed = $fixed.Replace($m.Value, ('"' + $m.Groups['u'].Value + '"'))
}

if ($skipped) { Write-Warn "$skipped wrapper(s) use an encoding this script will not guess at." }

# ---------------------------------------------------------------------------
# THE PROOF. Hash what the repair produced, in the exact bytes it would write.
# ---------------------------------------------------------------------------
$out = New-Object System.Collections.Generic.List[byte]
$out.AddRange($enc.GetPreamble())
$out.AddRange($enc.GetBytes($fixed))
$outBytes = $out.ToArray()

$sha = [System.Security.Cryptography.SHA256]::Create()
try   { $newHash = ([BitConverter]::ToString($sha.ComputeHash($outBytes)) -replace '-', '') }
finally { $sha.Dispose() }

Write-Head 'After repair:'
Write-Info ("sha256  : $newHash")

if ($newHash -ne $script:GoodHash) {
    Write-Bad 'The repair did NOT reproduce the original. Nothing has been written.'
    Write-Info 'Something else is different too -- a second rewrite form, an encoding'
    Write-Info 'change in transit, or a genuinely newer build. Ask for a fresh copy.'
    exit 1
}

Write-Good 'PROVEN -- the repaired file is byte-for-byte the shipped original.'

if (-not $Apply) {
    Write-Head 'Report only. Re-run with -Apply to write the repair:'
    Write-Info "    .\FixCommon.ps1 -Apply"
    exit 0
}

$backup = "$Path.damaged-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
Copy-Item -LiteralPath $Path -Destination $backup -Force
[System.IO.File]::WriteAllBytes($Path, $outBytes)

$check = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
Write-Head 'Written.'
Write-Info "backup  : $backup"
if ($check -eq $script:GoodHash) { Write-Good "verified: $check" }
else { Write-Bad "POST-WRITE CHECK FAILED: $check -- restore from $backup" ; exit 1 }
exit 0
