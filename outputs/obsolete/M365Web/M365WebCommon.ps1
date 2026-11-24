#Requires -Version 5.1
<#
    M365WebCommon.ps1
    Host infrastructure for M365_Host_Web.ps1 -- ONE long-running process,
    one port, one Exchange Online connection, all tools served as views.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

#region Logic module import -------------------------------------------------------
$script:M365LogicManifest = Join-Path $PSScriptRoot 'M365Logic.psd1'
if (Test-Path -LiteralPath $script:M365LogicManifest) {
    Import-Module $script:M365LogicManifest -Force -ErrorAction Stop
} else {
    $script:M365LogicModule = Join-Path $PSScriptRoot 'M365Logic.psm1'
    if (-not (Test-Path -LiteralPath $script:M365LogicModule)) {
        throw "M365Logic.psd1/.psm1 not found in '$PSScriptRoot'. The web host needs the logic module alongside this file -- copy M365Logic.psm1 and M365Logic.psd1 into the same folder."
    }
    Import-Module $script:M365LogicModule -Force -ErrorAction Stop
}

$requiredLogic = @(
    'Test-M365ExoConnection'
    'Get-TargetSmtp'
    'Format-ErrorDetail'
    'Get-MailboxHealthSnapshot'
    'Invoke-MessageTrace'
    'Get-MailboxState'
    'Disable-M365User'
)
$missingLogic = @($requiredLogic | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
if ($missingLogic.Count -gt 0) {
    throw ("M365Logic imported but these expected functions are not available: {0}. Check M365Logic.psd1's FunctionsToExport matches the functions defined in M365Logic.psm1." -f ($missingLogic -join ', '))
}
#endregion

#region View registry ------------------------------------------------------------
$script:M365Views = @()

function Register-M365View {
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][string]$Name,
        [string]$Description = '',
        [switch]$Destructive,
        [Parameter(Mandatory)][scriptblock]$GetHtml,
        [hashtable]$Routes = @{}
    )
    if ($script:M365Views | Where-Object { $_.Id -eq $Id }) {
        throw "A view with Id '$Id' is already registered -- view Ids must be unique (they're used as HTML element ids)."
    }
    $script:M365Views += [PSCustomObject]@{
        Id          = $Id
        Name        = $Name
        Description = $Description
        Destructive = [bool]$Destructive
        GetHtml     = $GetHtml
        Routes      = $Routes
    }
}

function Get-M365Views {
    return $script:M365Views
}
#endregion

#region Shared CSS ----------------------------------------------------------------
function Get-M365SharedCss {
    return @'
  body { font-family: "Segoe UI", -apple-system, sans-serif; margin:0; background:#f5f5f5; color:#222; }
  .wrap { max-width:1200px; margin:0 auto; padding:20px 24px 40px 24px; }
  h1 { font-size:19px; margin:0 0 4px 0; }
  h2 { font-size:17px; margin:0 0 10px 0; }
  h3 { font-size:15px; margin:0 0 4px 0; }

  .topbar { background:#1f2937; color:#f9fafb; padding:10px 24px; display:flex; align-items:center; gap:16px; flex-wrap:wrap; }
  .topbar .brand { font-weight:700; font-size:15px; }
  .topbar .spacer { flex:1; }
  .connpill { font-size:12px; font-weight:700; padding:4px 10px; border-radius:999px; background:#374151; color:#e5e7eb; }
  .connpill.ok { background:#065f46; color:#d1fae5; }
  .connpill.bad { background:#7f1d1d; color:#fee2e2; }

  .warn { background:#fdecea; color:#b3261e; border:1px solid #f5c2c0; padding:10px 14px; border-radius:6px; font-weight:600; margin-bottom:16px; }
  .note { background:#eff6ff; color:#1e40af; border:1px solid #bfdbfe; padding:10px 14px; border-radius:6px; margin-bottom:16px; font-size:13px; }

  .row { display:flex; gap:8px; margin-bottom:12px; align-items:flex-end; flex-wrap:wrap; }
  label { font-size:13px; color:#444; }
  input[type=text], input[type=date], select, textarea { padding:8px; font-family: Consolas, monospace; font-size:13px; border:1px solid #ccc; border-radius:4px; background:#fff; }
  textarea { width:100%; box-sizing:border-box; }
  button { padding:8px 16px; border:0; border-radius:4px; background:#2563eb; color:#fff; font-weight:600; cursor:pointer; font-size:13px; }
  button:hover:not(:disabled) { background:#1d4ed8; }
  button:disabled { background:#9ca3af; cursor:default; }
  button.secondary { background:#6b7280; }
  button.secondary:hover:not(:disabled) { background:#4b5563; }
  button.danger { background:#b3261e; }
  button.danger:hover:not(:disabled) { background:#991b1b; }

  pre, .out { background:#fff; border:1px solid #ddd; border-radius:6px; padding:14px; font-family: Consolas, monospace; font-size:13px; white-space:pre-wrap; overflow-x:auto; }
  .card { background:#fff; border:1px solid #ddd; border-radius:6px; padding:14px 18px; margin-bottom:10px; }
  .card.destructive { border-left:4px solid #b3261e; }
  .card .desc { color:#555; font-size:13px; margin-bottom:10px; }
  .status { font-size:13px; color:#444; margin:8px 0; }
  .status.err { color:#b3261e; }
  fieldset { border:1px solid #ccc; border-radius:6px; margin:12px 0; padding:12px 16px; }
  legend { font-size:13px; font-weight:700; color:#374151; }

  table { border-collapse:collapse; width:100%; background:#fff; }
  th, td { border:1px solid #ddd; padding:6px 8px; font-size:13px; text-align:left; }
  th { background:#eef1f5; cursor:pointer; user-select:none; }
  tbody tr:hover { background:#f3f4f6; }
  tbody tr.selected { background:#dbeafe; }

  .view { display:none; }
  .view.active { display:block; }
'@
}
#endregion

#region Request authorization (CSRF defence) ---------------------------------------
$script:M365RequestToken = $null

function New-M365RequestToken {
    $bytes = New-Object 'byte[]' 32
    $rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::Create()
    try {
        $rng.GetBytes($bytes)
    } finally {
        if ($rng -is [System.IDisposable]) { $rng.Dispose() }
    }
    $script:M365RequestToken = ([System.BitConverter]::ToString($bytes) -replace '-', '').ToLowerInvariant()
    return $script:M365RequestToken
}

function Get-M365RequestToken {
    return $script:M365RequestToken
}

# GUARD: request authorisation check. Do not remove.
function Test-M365RequestAuthorized {
    param(
        [Parameter(Mandatory)]$Request,
        [Parameter(Mandatory)][int]$Port
    )

    $expected = @("http://localhost:$Port", "http://127.0.0.1:$Port")
    $origin = $Request.Headers['Origin']
    if ($origin) {
        if ($expected -notcontains $origin.TrimEnd('/')) { return $false }
    } else {
        $referer = $Request.Headers['Referer']
        if ($referer) {
            $ok = $false
            foreach ($e in $expected) { if ($referer.StartsWith($e)) { $ok = $true; break } }
            if (-not $ok) { return $false }
        }
    }

    $token = $Request.Headers['X-M365-Token']
    if (-not $token) { return $false }
    if ($token -ne $script:M365RequestToken) { return $false }

    return $true
}
#endregion

#region Connection gate -----------------------------------------------------------
function Wait-M365ExoConnection {
    param([string]$AlsoNote = '')

    $conn = Test-M365ExoConnection
    if ($conn.Connected) {
        Write-Host "Exchange Online: $($conn.Detail)" -ForegroundColor Green
        return
    }

    Write-Host ""
    Write-Host "Exchange Online is not usable in this session yet." -ForegroundColor Yellow
    Write-Host "  $($conn.Detail)" -ForegroundColor Yellow
    Write-Host "Run Connect-ExchangeOnline below (sign in when prompted)$AlsoNote, then type 'exit' to continue starting the host." -ForegroundColor Yellow
    Write-Host ""
    try {
        $host.EnterNestedPrompt()
    } catch {
        Write-Host "This host doesn't support pausing here ($($_.Exception.Message)) -- connect first, then re-run this script." -ForegroundColor Yellow
        return
    }

    Clear-M365DefaultDomainCache

    $conn2 = Test-M365ExoConnection
    if ($conn2.Connected) {
        Write-Host "Connected. Starting host..." -ForegroundColor Green
    } else {
        Write-Host "Still not connected -- starting anyway. The page will show a red connection banner, and lookups will fail with a clear error until you connect." -ForegroundColor Yellow
    }
}
#endregion

#region The web host --------------------------------------------------------------
function Start-M365WebHost {
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][int]$Port,
        [Parameter(Mandatory)][scriptblock]$GetHtml,
        [hashtable]$Routes = @{},
        [switch]$NoBrowser,
        [switch]$ShowDirectoryPill
    )

    [void](New-M365RequestToken)

    $prefix = "http://localhost:$Port/"
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add($prefix)
    try {
        $listener.Start()
    } catch {
        Write-Host "Could not start the web host on $prefix -- $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Two common causes:" -ForegroundColor Yellow
        Write-Host "  1. A previous run is still holding the port. Stop it, or pass a different -Port." -ForegroundColor Yellow
        Write-Host "  2. Access denied: Windows needs a one-time URL reservation for this prefix." -ForegroundColor Yellow
        Write-Host "     Run ONCE in an elevated prompt (then re-run this normally):" -ForegroundColor Yellow
        Write-Host "       netsh http add urlacl url=$prefix user=$env:USERDOMAIN\$env:USERNAME" -ForegroundColor Cyan
        return
    }

    Write-Host ""
    Write-Host "$HostName running at $prefix" -ForegroundColor Green
    Write-Host "All tools are served from this one process, sharing this one Exchange connection." -ForegroundColor Gray
    Write-Host "Press Ctrl+C here, or use Shut Down on the page, to stop." -ForegroundColor Yellow

    if (-not $NoBrowser) {
        try {
            $isPs5   = ($PSVersionTable.PSVersion.Major -lt 6)
            $onWin   = $isPs5 -or ($IsWindows -eq $true)
            $onMac   = (-not $isPs5) -and ($IsMacOS -eq $true)
            $onLinux = (-not $isPs5) -and ($IsLinux -eq $true)

            if     ($onWin)   { Start-Process $prefix }
            elseif ($onMac)   { Start-Process 'open' -ArgumentList $prefix }
            elseif ($onLinux) { Start-Process 'xdg-open' -ArgumentList $prefix }
            else {
                Write-Host "Could not determine the platform to open a browser -- open $prefix manually." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Could not auto-open a browser -- open $prefix manually. ($($_.Exception.Message))" -ForegroundColor Yellow
        }
    }

    $script:M365HostStopRequested = $false
    $cancelHandler = {
        param($sender, $e)
        $e.Cancel = $true
        Write-Host "`nCtrl+C received -- stopping host..." -ForegroundColor Yellow
        $script:M365HostStopRequested = $true
        try { $listener.Stop() } catch { }
    }
    [Console]::add_CancelKeyPress($cancelHandler)

    try {
        while ($listener.IsListening -and -not $script:M365HostStopRequested) {
            try {
                $context = $listener.GetContext()
            } catch [System.Net.HttpListenerException] {
                break
            } catch [System.ObjectDisposedException] {
                break
            }
            $request  = $context.Request
            $response = $context.Response
            try {
                $routeKey = "$($request.HttpMethod) $($request.Url.AbsolutePath)"

                if ($request.HttpMethod -eq 'GET' -and $request.Url.AbsolutePath -eq '/') {
                    $html  = & $GetHtml
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($html)
                    $response.ContentType = 'text/html; charset=utf-8'
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                # GUARD: request authorisation check. Do not remove.
                elseif ($request.Url.AbsolutePath -like '/api/*' -and -not (Test-M365RequestAuthorized -Request $request -Port $Port)) {
                    Write-Host "Rejected unauthorized $routeKey (bad/missing token or foreign origin)." -ForegroundColor Red
                    $response.StatusCode = 403
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes('{"ok":false,"error":"Request rejected: missing or invalid API token, or the request did not originate from this host''s own page. Reload the page in your browser."}')
                    $response.ContentType = 'application/json; charset=utf-8'
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                elseif ($routeKey -eq 'GET /api/connection') {
                    $full = ($request.QueryString['full'] -eq '1')
                    $sum  = Get-M365ConnectionSummary -IncludeExpensive:$full -IncludeDirectory:$ShowDirectoryPill
                    $payload = [ordered]@{
                        checkedExpensive = $full
                        directoryMonitored = [bool]$ShowDirectoryPill
                        tenantDomain = "$(Get-M365DefaultDomain)"
                        tenant = $(
                            $tg = Test-M365ExpectedTenant
                            [ordered]@{
                                armed    = [bool]$tg.Armed
                                match    = [bool]$tg.Match
                                expected = "$($tg.Expected)"
                                actual   = "$($tg.Actual)"
                                detail   = "$($tg.Detail)"
                            }
                        )
                        exo = [ordered]@{
                            connected    = [bool]$sum.Exo.Connected
                            detail       = "$($sum.Exo.Detail)"
                            organization = "$($sum.Exo.Organization)"
                            skipped      = $false
                        }
                        spo = [ordered]@{
                            connected = [bool]$sum.Spo.Connected
                            detail    = "$($sum.Spo.Detail)"
                            skipped   = [bool]($sum.Spo.PSObject.Properties['Skipped'] -and $sum.Spo.Skipped)
                        }
                        aad = [ordered]@{
                            connected = [bool]$sum.Aad.Connected
                            detail    = "$($sum.Aad.Detail)"
                            provider  = "$($sum.Aad.Provider)"
                            account   = "$($sum.Aad.Account)"
                            skipped   = [bool]($sum.Aad.PSObject.Properties['Skipped'] -and $sum.Aad.Skipped)
                        }
                    }
                    $json  = $payload | ConvertTo-Json -Depth 5 -Compress
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $response.ContentType = 'application/json; charset=utf-8'
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                elseif ($routeKey -eq 'POST /api/shutdown' -or $routeKey -eq 'GET /api/shutdown') {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes('{"ok":true}')
                    $response.ContentType = 'application/json; charset=utf-8'
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                    $script:M365HostStopRequested = $true
                }
                elseif ($Routes.ContainsKey($routeKey)) {
                    $result = & $Routes[$routeKey] $request
                    $bytes  = [System.Text.Encoding]::UTF8.GetBytes($result.Body)
                    $response.ContentType = $result.ContentType
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                else {
                    $response.StatusCode = 404
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes("No route for $routeKey")
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
            } catch {
                Write-Host "Handler error on $($request.HttpMethod) $($request.Url.AbsolutePath): $($_.Exception.Message)" -ForegroundColor Red
                try {
                    $response.StatusCode = 500
                    $errBytes = [System.Text.Encoding]::UTF8.GetBytes("Server error: $($_.Exception.Message)")
                    $response.OutputStream.Write($errBytes, 0, $errBytes.Length)
                } catch { }
            } finally {
                $response.OutputStream.Close()
            }
        }
    } finally {
        [Console]::remove_CancelKeyPress($cancelHandler)
        try { $listener.Stop() } catch { }
        try { $listener.Close() } catch { }
        Write-Host "$HostName stopped." -ForegroundColor Yellow
    }
}
#endregion
