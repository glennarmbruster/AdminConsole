#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- Get-DistributionGroup exists in BOTH
#                                 Exchanges and Get-RemoteMailbox is
#                                 on-premises only; Move-ADObject writes to
#                                 AD. This script must not share a process
#                                 with a cloud-connected console.
# TOOLKIT-HELP-BEGIN
# STEP 3 of the Migration project. THE ONE THAT STARTS THE CLOCK.
#
# SHARED MAILBOXES ONLY (2026-10-04). Moves every mailbox in SMB-Input.csv
# into the OU that directory sync does not read. On the next sync the cloud
# copy is deleted: the mailbox's user goes to the Entra recycle bin and the
# mailbox becomes soft-deleted (30 days). DL-Input.csv is NOT read here --
# distribution lists are a separate project.
#
# WHERE THE OU COMES FROM. NonSyncOu on the directory in Tenant.json (Tenant
# Manager), in the canonical form the other OUs use:
#     contoso.com/Accounts/NoSync
# Until it is filled in this script STOPS before touching anything and says
# so. That is deliberate: the OU is the one fact in this project that only
# the person running it knows.
#
# WHAT IT REFUSES. Any object not found, or already in the target OU, is
# reported and skipped. If Steps DL-1, DL-2, 1 and 2 have no export newer than a day it
# warns, because a stale export is what gets restamped in Step 5.
#
# AFTER THIS: WAIT. Sync runs every 30 minutes; allow 75 minutes (two
# cycles plus slack) before Step 4, then confirm in the cloud that the
# mailboxes show as soft-deleted (Step 4 shows you). If the batch is larger
# than the Entra Connect export deletion threshold (500 by default), the
# sync will refuse the deletes and Step 4 will find nothing -- see
# Disable-ADSyncExportDeletionThreshold on the sync server.
#
# UNDO: 3R-ONPREM-MoveBack.ps1 moves everything in the output file back to
# its original OU. Run it before a sync and nothing happened; after a sync,
# the objects re-match on the next cycle (ImmutableId is kept) and the cloud
# copies are restored -- but any list is recreated as new.
#
# Output lands next to this script:
#     Migration-Moved_<date>_<time>.csv   what moved, from where -- Steps 4
#                                         and 7 read the newest one
# TOOLKIT-HELP-END

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$toolkitRoot = $PSScriptRoot
while ($toolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $toolkitRoot 'Common.ps1'))) {
    $toolkitRoot = Split-Path -Parent $toolkitRoot
}
if (-not $toolkitRoot) { throw "Common.ps1 was not found in any folder above '$PSScriptRoot'." }
. (Join-Path $toolkitRoot 'Common.ps1')
. (Join-Path $PSScriptRoot '_Shared.ps1')

$basePath = $PSScriptRoot

function ConvertTo-OuDistinguishedName {
    # 'contoso.com/Accounts/NoSync' -> 'OU=NoSync,OU=Accounts,DC=contoso,DC=com'.
    # A value that already looks like a DN is returned as-is.
    param([string]$Path)
    $p = "$Path".Trim()
    if (-not $p) { return '' }
    if ($p -match '^(OU|CN|DC)=') { return $p }
    $parts = @($p -split '/' | Where-Object { $_ })
    if ($parts.Count -lt 2) { return '' }
    $dom = ConvertTo-DomainDn -Domain $parts[0]
    $ous = @($parts[1..($parts.Count - 1)])
    [array]::Reverse($ous)
    return ((@($ous | ForEach-Object { "OU=$_" }) + @($dom)) -join ',')
}

# ---- THE PENDING PIECE: the OU comes from Tenant.json, never from here. ----
$dirs = @(Get-TenantDirectory)
$configured = @($dirs | Where-Object { $_.PSObject.Properties['NonSyncOu'] -and "$($_.NonSyncOu)".Trim() })
if ($configured.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        ("PENDING -- the non-syncing OU is not configured yet." + "`r`n`r`n" +
         "Open Tenant Manager and fill in 'NonSync OU' for the directory these objects live in, in the same canonical form as the other OUs " +
         "(domain.com/Container/OU). Nothing has been moved."),
        '3 - ONPREM Move to NonSync OU', 'OK', 'Warning') | Out-Null
    Write-Host 'STOPPED: NonSyncOu is not set on any directory in Tenant.json. Nothing was moved.' -ForegroundColor Yellow
    return
}

$go = Show-StepInstructions -Title '3 - ONPREM Move to NonSync OU' -ScriptPath $PSCommandPath -Text @"
STEP 3 of 7 -- move the on-premises objects OUT of directory sync scope.

* Reads SMB-Input.csv (shared mailboxes only; lists are a separate project).
* Every mailbox is moved to the NonSync OU from Tenant.json:
$(($configured | ForEach-Object { "      $($_.Id): $($_.NonSyncOu)" }) -join "`r`n")
* On the next sync the cloud copies are DELETED. Mailboxes become
  soft-deleted (30 days, restorable in Step 4).
* Then WAIT 75 MINUTES before Step 4.

Undo: 3R-ONPREM-MoveBack.ps1. This window must be connected to ON-PREMISES
Exchange. Click OK to review the input files and continue.
"@
if (-not $go) { return }

Import-Module ActiveDirectory -ErrorAction Stop
Enter-MigrationOnPremScope

# Stale-export guard: the point of Steps 1 and 2 is that they describe THESE objects
# as they are NOW.
foreach ($pat in 'Migration-SMB-ONPREM_*.csv', 'Migration-SMB-CLOUD_*.csv') {
    $f = Get-NewestExport -Folder $basePath -Pattern $pat
    if (-not $f) { Write-Host "WARNING: no $pat export found. Steps 1 and 2 should run before this." -ForegroundColor Yellow }
    elseif ($f.LastWriteTime -lt (Get-Date).AddDays(-1)) { Write-Host "WARNING: newest $pat is $([int]((Get-Date) - $f.LastWriteTime).TotalHours)h old." -ForegroundColor Yellow }
}

# Mailboxes only. DL-Input.csv is deliberately not read: lists are a
# separate project, and opening a second file here was one more thing to
# review for nothing.
$dlNames  = @(); $smbNames = @()
try { $smbNames = Read-InputList -Path (Join-Path $basePath 'SMB-Input.csv') -What 'mailbox' } catch { Write-Host $_.Exception.Message -ForegroundColor Yellow }
if ($smbNames.Count -eq 0) { Write-Host 'Nothing to move.' -ForegroundColor Yellow; return }

# Resolve everything FIRST, move NOTHING until the whole list resolves. A
# typo in row 12 found after rows 1-11 moved is a half-done batch with a
# sync clock already running.
$plan = @()
foreach ($n in $dlNames) {
    try {
        $dg = Get-DistributionGroup -Identity $n -ErrorAction Stop
        $plan += [PSCustomObject]@{ Type = 'DL'; Name = "$($dg.Name)"; PrimarySmtpAddress = "$($dg.PrimarySmtpAddress)"; Guid = "$($dg.Guid)"; ExchangeGuid = ''; Dn = "$($dg.DistinguishedName)"; Error = '' }
    } catch { $plan += [PSCustomObject]@{ Type = 'DL'; Name = $n; PrimarySmtpAddress = ''; Guid = ''; ExchangeGuid = ''; Dn = ''; Error = "$($_.Exception.Message)" } }
}
foreach ($n in $smbNames) {
    try {
        $rm = Get-RemoteMailbox -Identity $n -ErrorAction Stop
        $plan += [PSCustomObject]@{ Type = 'SMB'; Name = "$($rm.Name)"; PrimarySmtpAddress = "$($rm.PrimarySmtpAddress)"; Guid = "$($rm.Guid)"; ExchangeGuid = "$($rm.ExchangeGuid)"; Dn = "$($rm.DistinguishedName)"; Error = '' }
    } catch { $plan += [PSCustomObject]@{ Type = 'SMB'; Name = $n; PrimarySmtpAddress = ''; Guid = ''; ExchangeGuid = ''; Dn = ''; Error = "$($_.Exception.Message)" } }
}

$unresolved = @($plan | Where-Object { $_.Error })
if ($unresolved.Count) {
    Write-Host ''
    Write-Host "$($unresolved.Count) object(s) did not resolve -- NOTHING was moved:" -ForegroundColor Red
    $unresolved | ForEach-Object { Write-Host "  $($_.Type)  $($_.Name)  --  $($_.Error)" -ForegroundColor Red }
    Write-Host 'Fix the input files and re-run.' -ForegroundColor Yellow
    return
}

# Target OU per object, from the directory its DN belongs to.
foreach ($p in $plan) {
    $d = Resolve-DistinguishedNameDirectory -DistinguishedName $p.Dn
    $ou = ''
    if ($d -and $d.PSObject.Properties['NonSyncOu']) { $ou = ConvertTo-OuDistinguishedName -Path "$($d.NonSyncOu)" }
    $p | Add-Member -NotePropertyName TargetOu -NotePropertyValue $ou -Force
    $p | Add-Member -NotePropertyName OriginalOu -NotePropertyValue (($p.Dn -split '(?<!\\),', 2)[1]) -Force
    # A controller in the object's OWN domain. Without it the AD module asks
    # the jump box's domain, which does not hold a child-domain object.
    $p | Add-Member -NotePropertyName Server -NotePropertyValue (Get-MigrationAdServer -DistinguishedName $p.Dn) -Force
}
$noOu = @($plan | Where-Object { -not $_.TargetOu })
if ($noOu.Count) {
    Write-Host ''
    Write-Host "$($noOu.Count) object(s) live in a directory with no NonSyncOu configured -- NOTHING was moved:" -ForegroundColor Red
    $noOu | ForEach-Object { Write-Host "  $($_.Type)  $($_.Name)  ($($_.Dn))" -ForegroundColor Red }
    return
}
foreach ($ou in ($plan.TargetOu | Sort-Object -Unique)) {
    $srv = "$(($plan | Where-Object { $_.TargetOu -eq $ou } | Select-Object -First 1).Server)"
    $sp = @{}; if ($srv) { $sp['Server'] = $srv }
    $exists = $null
    try { $exists = Get-ADObject -Identity $ou @sp -ErrorAction Stop } catch { }
    if (-not $exists) {
        Write-Host "Target OU does not exist (asked $(if ($srv) { $srv } else { 'the default domain' })): $ou -- NOTHING was moved." -ForegroundColor Red
        return
    }
}

Write-Host ''
Write-Host "About to move $($plan.Count) object(s):" -ForegroundColor Cyan
$plan | ForEach-Object { Write-Host "  $($_.Type.PadRight(4)) $($_.PrimarySmtpAddress.PadRight(45)) $($_.OriginalOu)  ->  $($_.TargetOu)   [via $($_.Server)]" }
$a = [System.Windows.Forms.MessageBox]::Show(
    ("Move $($plan.Count) shared mailbox(es) out of sync scope?" +
     "`r`n`r`nThe next directory sync deletes their cloud copies. This is the point of no easy return."),
    'Confirm', 'OKCancel', 'Warning')
if ($a -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host 'Cancelled. Nothing was moved.' -ForegroundColor Yellow; return }

$outFile = Join-Path $basePath "Migration-Moved_$(Get-MigrationTimestamp).csv"
$rows = @(); $ok = 0; $failed = 0
foreach ($p in $plan) {
    $res = ''; $newDn = ''
    if ($p.OriginalOu -ieq $p.TargetOu) {
        $res = 'AlreadyThere'; $newDn = $p.Dn
        Write-Host "  skip  $($p.PrimarySmtpAddress)  already in target OU" -ForegroundColor Yellow
    } else {
        try {
            # Learn (Move-ADObject): -Identity <DN or GUID>, -TargetPath <DN>.
            $sp = @{}; if ("$($p.Server)") { $sp['Server'] = $p.Server }
            Move-ADObject -Identity $p.Guid -TargetPath $p.TargetOu @sp -ErrorAction Stop
            $newDn = (($p.Dn -split '(?<!\\),', 2)[0]) + ',' + $p.TargetOu
            $res = 'Moved'; $ok++
            Write-Host "  moved $($p.PrimarySmtpAddress)" -ForegroundColor Green
        } catch {
            $res = "FAILED: $($_.Exception.Message)"; $failed++
            Write-Host "  FAIL  $($p.PrimarySmtpAddress)  $($_.Exception.Message)" -ForegroundColor Red
            Write-FailedLog -Action 'MigrationMoveToNonSyncOu' -Target (Get-LogTarget $p.PrimarySmtpAddress $p.Name) -ErrorRecord $_
        }
    }
    $rows += [PSCustomObject]@{
        Type = $p.Type; Name = $p.Name; PrimarySmtpAddress = $p.PrimarySmtpAddress; Guid = $p.Guid; ExchangeGuid = $p.ExchangeGuid
        OriginalDn = $p.Dn; OriginalOu = $p.OriginalOu; NewDn = $newDn; TargetOu = $p.TargetOu; Server = "$($p.Server)"
        MovedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'); Result = $res
    }
}
$rows | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-ExportDone -Path $outFile -Rows $rows.Count
Write-Host ''
Write-Host "$ok moved, $failed failed. Sync runs every 30 minutes -- WAIT 75 MINUTES, then run Step 4." -ForegroundColor Cyan
Write-CompletedLog -Action 'MigrationMoveToNonSyncOu' -Target "$($plan.Count) objects" -Result $(if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }) `
    -Detail "$ok moved, $failed failed -> $([System.IO.Path]::GetFileName($outFile))"
