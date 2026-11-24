#Requires -Version 5.1
<#
    2-Export-Usage.ps1
    DL cloud migration -- delivery evidence for retire-vs-migrate.

    Reasoning: ScriptNotes/DLCloudMigration.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-08
#>
# TOOLKIT-HELP-BEGIN
# Counts messages actually delivered to each group over the last few days.
# READ ONLY -- this script changes nothing.
#
# Run 1-Export-Groups.ps1 first. This reads the group list from that export, so
# both files describe the same set of groups.
#
# WHY IT EXISTS: the object's last-changed date says when somebody edited the
# group, not whether anyone emails it. A group untouched for three years can be
# in daily use, and a group edited last week can have been dead for a decade.
# The only evidence that distinguishes them is delivery.
#
# WHAT IT CANNOT TELL YOU: message trace only reaches back about ten days.
# Traffic here proves a group is LIVE. Silence here does NOT prove a group is
# dead -- monthly reports, quarterly notices and seasonal lists all look dead
# in a ten day window. Never retire on silence alone.
# TOOLKIT-HELP-END

# Directory names are NOT in a ValidateSet.
#
# This said [ValidateSet('HQ','Franchise')] -- one organisation's internal
# structure written into a type constraint, in a script meant to be portable.
# It is the exact pattern removed from Get-ADServerString for the same reason,
# reintroduced here by the person who removed it.
#
# Two ways it failed. A tenant whose directories are named anything else could
# not run these at all. And an EMPTY value -- which the console used to pass
# before a lookup -- was rejected by the validator before the default could
# apply, so the script died on its own default value.
#
# Validated against Tenant.json at runtime instead, where the answer lives.
param(
    [string]$Environment = '',
    [ValidateRange(1, 10)]
    [int]$DaysBack = 10,
    [switch]$NoPrompt
)

$toolkitRoot = $PSScriptRoot
while ($toolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $toolkitRoot 'Common.ps1'))) {
    $toolkitRoot = Split-Path -Parent $toolkitRoot
}
if (-not $toolkitRoot) { throw "Common.ps1 was not found above '$PSScriptRoot'." }
. (Join-Path $toolkitRoot 'Common.ps1')

# Resolve the directory before anything uses it.
$script:DirIds = @(@(Get-TenantDirectory) | ForEach-Object { "$($_.Id)" } | Where-Object { $_ })
if (-not "$Environment".Trim()) {
    $def = @(@(Get-TenantDirectory) | Where-Object { $_.IsDefault })
    if ($def.Count -gt 0) { $Environment = "$($def[0].Id)" }
    elseif ($script:DirIds.Count -gt 0) { $Environment = $script:DirIds[0] }
}
if ($script:DirIds.Count -gt 0 -and ($script:DirIds -notcontains $Environment)) {
    throw ("Directory '$Environment' is not defined in Tenant.json. " +
           "Configured directories: $($script:DirIds -join ', ')")
}
Add-Type -AssemblyName System.Windows.Forms

$envFolder = if ($Environment -eq 'HQ') { 'HQ' } else { 'Fran' }
$outDir    = Join-Path (Join-Path $PSScriptRoot 'Exports') $envFolder

#region Locate the source export ------------------------------------------------
# The newest Groups file, and its timestamp is reused for this file so the usage
# data is visibly tied to the snapshot it describes. Usage counted against a
# different day's group list is a subtly wrong answer that looks right.
$groupsCsv = @(Get-ChildItem -LiteralPath $outDir -Filter "Groups_$Environment`_*.csv" -File -ErrorAction SilentlyContinue |
               Sort-Object LastWriteTime -Descending | Select-Object -First 1)
if (@($groupsCsv).Count -eq 0) {
    Write-Host ''
    Write-Host "No Groups_$Environment`_*.csv in $outDir." -ForegroundColor Red
    Write-Host 'Run 1-Export-Groups.ps1 first.' -ForegroundColor Yellow
    return
}
$srcName = $groupsCsv[0].Name
$stamp   = ''
if ($srcName -match "Groups_$Environment`_(.+)\.csv$") { $stamp = $Matches[1] }
if (-not $stamp) { $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm' }

$groups = @(Import-Csv -LiteralPath $groupsCsv[0].FullName)
if ($groups.Count -eq 0) {
    Write-Host "'$srcName' has no rows. Nothing to measure." -ForegroundColor Yellow
    return
}
#endregion

if (-not $NoPrompt) {
    $msg = @"
Measure delivery to $($groups.Count) $Environment group(s) over the last $DaysBack day(s).

READ ONLY. This script changes nothing.

Source list : $srcName
Queries     : Exchange Online message trace, one call per group
Written to  : Exports\$envFolder

This takes roughly one second per group, so about $([int]([math]::Ceiling($groups.Count / 60.0))) minute(s).

REMEMBER: traffic proves a group is live. Silence does NOT prove it is dead --
a ten day window cannot see a monthly or quarterly list.

Continue?
"@
    $ans = [System.Windows.Forms.MessageBox]::Show($msg, 'DL Cloud Migration -- Usage',
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Information)
    if ($ans -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host 'Cancelled.' -ForegroundColor Yellow
        return
    }
}

#region Measure -----------------------------------------------------------------
$start = (Get-Date).AddDays(-$DaysBack)
$end   = Get-Date

Write-Host ''
Write-Host "=== DL Cloud Migration -- usage ($Environment, $DaysBack days) ===" -ForegroundColor Cyan
Write-Host ''

# Pinned to the cloud module. Both Exchanges are typically connected in this
# process and Get-MessageTrace does not exist on-premises in this form -- but a
# bare call would still be resolved by import order rather than by intent.
$traceCmd = 'Get-MessageTraceV2'
try {
    if (-not (Get-EXOCommand -Name $traceCmd)) { throw 'not found' }
} catch {
    $traceCmd = 'Get-MessageTrace'
    Write-Host 'Get-MessageTraceV2 unavailable; falling back to Get-MessageTrace.' -ForegroundColor DarkGray
}

$rows = @(); $i = 0; $errors = 0
foreach ($g in $groups) {
    $i++
    $smtp = "$($g.PrimarySmtpAddress)".Trim()
    if (-not $smtp) { continue }

    if ($i % 25 -eq 0 -or $i -eq 1) {
        Write-Host ("  [{0}/{1}] {2}" -f $i, $groups.Count, $smtp) -ForegroundColor DarkGray
    }

    $msgs = @()
    try {
        $msgs = @(Invoke-EXOCommand -Name $traceCmd -Parameters @{
            RecipientAddress = $smtp
            StartDate        = $start
            EndDate          = $end
            ResultSize       = 1000
            ErrorAction      = 'Stop'
            # A deliberately limited ResultSize makes Exchange warn on every
            # call. Silenced, not raised: a group with more than 1000 messages
            # in ten days is emphatically live and the exact figure changes no
            # decision.
            WarningAction    = 'SilentlyContinue'
        })
    } catch {
        $errors++
        $rows += [PSCustomObject]@{
            GroupPrimarySmtp = $smtp
            GroupName        = "$($g.Name)"
            DaysMeasured     = $DaysBack
            MessageCount     = ''
            DistinctSenders  = ''
            LastDelivery     = ''
            DaysWithTraffic  = ''
            Verdict          = 'Unknown'
            Note             = "Trace failed: $($_.Exception.Message)"
        }
        continue
    }

    $senders = @($msgs | ForEach-Object { "$($_.SenderAddress)".ToLower() } | Sort-Object -Unique)
    $dates   = @($msgs | ForEach-Object { try { ([datetime]$_.Received).Date } catch { $null } } |
                 Where-Object { $_ } | Sort-Object -Unique)
    $last    = if ($dates.Count -gt 0) { ($dates | Select-Object -Last 1).ToString('yyyy-MM-dd') } else { '' }

    # Truncated at the cap means "at least this many", and the row says so
    # rather than reporting 1000 as if it were the true figure.
    $note = if ($msgs.Count -ge 1000) { 'Truncated at 1000 -- actual volume is higher' } else { '' }

    $rows += [PSCustomObject]@{
        GroupPrimarySmtp = $smtp
        GroupName        = "$($g.Name)"
        DaysMeasured     = $DaysBack
        MessageCount     = $msgs.Count
        DistinctSenders  = $senders.Count
        LastDelivery     = $last
        DaysWithTraffic  = $dates.Count
        # 'NoTrafficInWindow', not 'Dead'. The window cannot see a monthly list,
        # and a column called Dead would get acted on as though it could.
        Verdict          = if ($msgs.Count -eq 0) { 'NoTrafficInWindow' }
                           elseif ($msgs.Count -ge 50) { 'Active' }
                           else { 'Light' }
        Note             = $note
    }
}
#endregion

#region Write -------------------------------------------------------------------
$path = Join-Path $outDir "Usage_$Environment`_$stamp.csv"
if ($rows.Count -eq 0) {
    Set-Content -LiteralPath $path -Value '' -Encoding UTF8
} else {
    $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
}

$active = @($rows | Where-Object { $_.Verdict -eq 'Active' }).Count
$light  = @($rows | Where-Object { $_.Verdict -eq 'Light' }).Count
$silent = @($rows | Where-Object { $_.Verdict -eq 'NoTrafficInWindow' }).Count

Write-Host ''
Write-Host "--- $Environment usage over $DaysBack day(s) ---" -ForegroundColor Cyan
Write-Host ("  Active (50+ messages) : {0}" -f $active) -ForegroundColor Green
Write-Host ("  Light  (1-49)         : {0}" -f $light)
Write-Host ("  No traffic in window  : {0}" -f $silent) -ForegroundColor Yellow
if ($errors -gt 0) { Write-Host ("  Trace failed          : {0}" -f $errors) -ForegroundColor Red }
Write-Host ''
Write-Host 'No traffic in a ten day window is NOT evidence a group is dead.' -ForegroundColor Yellow
Write-Host 'Confirm with the owner before retiring anything.' -ForegroundColor Yellow
Write-Host ''
Write-Host "Saved to: $path" -ForegroundColor Green
Write-Host ''

try {
    Write-CompletedLog -Action 'DLCloudMigration-Usage' -Target $Environment -Result 'Success' `
        -Detail "$($rows.Count) group(s) over $DaysBack day(s): $active active, $light light, $silent silent"
} catch { }
#endregion
