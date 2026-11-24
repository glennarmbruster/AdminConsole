#Requires -Version 5.1
<#
    Convert-DLSourceOfAuthority.ps1

    Transfers Source of Authority for directory-synchronized distribution
    groups to the cloud, so they become editable in Exchange Online.

    PREVIEW BY DEFAULT. Pass -Apply to make changes.

    Input CSV needs one column, Identity: the group's primary SMTP address.
    Nested groups must appear BEFORE their parents -- the conversion does
    not recurse.

    Graph permission required:
        Group-OnPremisesSyncBehavior.ReadWrite.All   (read and write)
        Group.Read.All                               (pre-flight check)

    Role required: Hybrid Identity Administrator.
        Template ID 8ac3fc64-6eca-42ea-9e69-59f4c7b60eb2. This is a directory
        role, not a Graph scope -- the consent alone is not enough.

    Version prerequisite: Entra Connect Sync 2.5.76.0 or later on EVERY sync
        server, including staging. Older clients ignore isCloudManaged, so the
        conversion appears to succeed and then silently reverts on the next
        sync cycle.

    Author   : Glenn Armbruster
    Revised  : 2026-08-18
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$InputCsv,
    [string]$OutputCsv = '',
    [switch]$Apply,
    [int]$DelayMs = 250        # gentle pacing; raise if throttled
)

$ErrorActionPreference = 'Stop'

if (-not $OutputCsv) {
    $OutputCsv = ".\soa-results-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
}

#region Connect ---------------------------------------------------------------
# Group.Read.All is used for the pre-flight check: it confirms each group is
# currently directory-synchronized before the script touches it, so a bad row
# in the CSV cannot send the conversion at the wrong object.
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Connect-MgGraph -Scopes @(
    'Group-OnPremisesSyncBehavior.ReadWrite.All'
    'Group.Read.All'
) -NoWelcome

$ctx = Get-MgContext
if (-not $ctx) { throw 'Not connected to Microsoft Graph.' }

Write-Host "Connected as $($ctx.Account)" -ForegroundColor Cyan
if (-not $Apply) {
    Write-Host 'PREVIEW ONLY -- nothing will be changed. Add -Apply to run.' -ForegroundColor Yellow
}
#endregion

#region Input -----------------------------------------------------------------
$rows = @(Import-Csv -LiteralPath $InputCsv)
if ($rows.Count -eq 0) { throw "No rows in $InputCsv" }
if (-not $rows[0].PSObject.Properties['Identity']) {
    throw "$InputCsv has no 'Identity' column."
}
Write-Host "$($rows.Count) group(s) to process." -ForegroundColor Cyan
#endregion

#region Helpers ---------------------------------------------------------------
function Invoke-GraphWithRetry {
    <#
        One retry wrapper for the whole script.

        429 and 503 are transient and expected at this volume. Anything else
        is a real error and is allowed to surface rather than being retried
        into confusion -- a bad group name retried five times is still a bad
        group name, and the retries only hide when it happened.
    #>
    param(
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [string]$Body,
        [int]$MaxTries = 5
    )
    for ($try = 1; $try -le $MaxTries; $try++) {
        try {
            if ($Body) {
                return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Body -ContentType 'application/json'
            }
            return Invoke-MgGraphRequest -Method $Method -Uri $Uri
        } catch {
            $msg = "$($_.Exception.Message)"
            if ($msg -notmatch '429|503|throttl|too many' -or $try -eq $MaxTries) { throw }
            # Exponential backoff. Graph does return Retry-After; this is the
            # simple form that does not depend on reading the header.
            $wait = [Math]::Pow(2, $try)
            Write-Host "  throttled -- waiting $wait s" -ForegroundColor DarkGray
            Start-Sleep -Seconds $wait
        }
    }
}
#endregion

#region Main ------------------------------------------------------------------
$results = New-Object System.Collections.Generic.List[object]
$n = 0

foreach ($row in $rows) {
    $n++
    $identity = "$($row.Identity)".Trim()
    if (-not $identity) { continue }

    $rec = [ordered]@{
        Identity = $identity
        GroupId  = ''
        Before   = ''
        Action   = ''
        Result   = ''
        Detail   = ''
    }

    Write-Progress -Activity 'Source of Authority' -Status $identity -PercentComplete ([int](100 * $n / $rows.Count))

    try {
        # Find the group by mail address, which is unique in the tenant.
        $enc  = [uri]::EscapeDataString($identity)
        $uri  = 'https://graph.microsoft.com/v1.0/groups' +
                '?$filter=mail eq ' + "'$enc'" +
                '&$select=id,displayName,mail,onPremisesSyncEnabled'
        $find = Invoke-GraphWithRetry -Method GET -Uri $uri

        $hits = @($find.value)
        if ($hits.Count -eq 0) { throw 'No group in Entra has that mail address.' }
        if ($hits.Count -gt 1) { throw "Ambiguous -- $($hits.Count) groups share that address." }

        $g = $hits[0]
        $rec.GroupId = "$($g.id)"
        $rec.Before  = "onPremisesSyncEnabled=$($g.onPremisesSyncEnabled)"

        # ALREADY CONVERTED GROUPS ARE SKIPPED, NOT RE-CONVERTED.
        #
        # A converted group reports onPremisesSyncEnabled as null. Checking
        # this makes the script safe to re-run over a batch that partly
        # succeeded -- which is exactly what you want at 02:00 after a
        # failure, rather than having to work out where it stopped.
        if ($null -eq $g.onPremisesSyncEnabled) {
            $rec.Action = 'Skip'
            $rec.Result = 'AlreadyCloudManaged'
            $rec.Detail = 'Not directory-synchronized; nothing to convert.'
            $results.Add([pscustomobject]$rec)
            continue
        }

        if (-not $Apply) {
            $rec.Action = 'WouldConvert'
            $rec.Result = 'Preview'
            $results.Add([pscustomobject]$rec)
            continue
        }

        # The body is sent as a JSON STRING with an explicit content type,
        # rather than relying on how a hashtable happens to serialize.
        $patch = 'https://graph.microsoft.com/v1.0/groups/' + $g.id + '/onPremisesSyncBehavior'
        $null  = Invoke-GraphWithRetry -Method PATCH -Uri $patch -Body '{"isCloudManaged": true}'

        $rec.Action = 'Convert'
        $rec.Result = 'Success'
        $rec.Detail = 'Converted. Entra Connect stops synchronizing it on the next cycle; verify then.'
    }
    catch {
        $rec.Action = 'Convert'
        $rec.Result = 'Failed'
        $rec.Detail = "$($_.Exception.Message)"
        Write-Host "  FAILED $identity -- $($_.Exception.Message)" -ForegroundColor Red
    }

    $results.Add([pscustomobject]$rec)
    if ($Apply -and $DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
}
Write-Progress -Activity 'Source of Authority' -Completed
#endregion

#region Report ----------------------------------------------------------------
$results | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8

$ok   = @($results | Where-Object { $_.Result -eq 'Success' }).Count
$skip = @($results | Where-Object { $_.Result -eq 'AlreadyCloudManaged' }).Count
$bad  = @($results | Where-Object { $_.Result -eq 'Failed' }).Count
$prev = @($results | Where-Object { $_.Result -eq 'Preview' }).Count

Write-Host ''
if ($prev) { Write-Host "Preview: $prev group(s) would be converted." -ForegroundColor Yellow }
if ($ok)   { Write-Host "Converted: $ok" -ForegroundColor Green }
if ($skip) { Write-Host "Already cloud-managed: $skip" -ForegroundColor DarkGray }
if ($bad)  { Write-Host "Failed: $bad -- see the results file" -ForegroundColor Red }
Write-Host "Results: $OutputCsv"

# VERIFY AFTER THE NEXT SYNC CYCLE, NOT NOW.
#
# The PATCH returns immediately, but Entra Connect only stops synchronizing
# the object on its next run. Re-run this script in preview mode afterwards:
# everything that converted should come back as AlreadyCloudManaged. That is
# the confirmation, not the success count above.
#endregion
