#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# STEP 0 of the Migration project. Run BEFORE Step 5, after the input files
# are final. Reads nothing but the cloud and tells you, per object, what a
# restore or recreate CANNOT bring back -- so the decision is made while it
# is still a decision.
#
# READ ONLY. Nothing is changed.
#
# STOP means do not migrate this object as-is:
#   * Security group. A recreated group has a NEW object ID. Every permission
#     granted TO the group -- Full Access / Send As on other mailboxes,
#     SharePoint, Teams, Conditional Access, app assignments -- dies with the
#     old ID. Only trustees on mailboxes IN THIS BATCH are re-granted (Step 7
#     works from the Step 4 Perms file). Everything else is silent loss.
#   * Transport rule names the list. The rule points at the old object and
#     stops matching after recreate. Edit the rule to the new list afterwards.
#
# WARN means it will work, but something changes:
#   * Nested in / owner of / moderator of a list that stays synced. That
#     reference is managed on-premises and the object will not be there.
#   * Archive, litigation hold, in-place hold, or more than 50 GB on a
#     mailbox. Each needs a licence, and the licence comes from a synced AD
#     group the account is about to leave. Nothing breaks on day one.
#   * Not a SharedMailbox. Step 6 converts on restore; licence rules above
#     apply.
#
# INFO is context: already cloud-only, batch above the sync deletion
# threshold, and so on.
#
# Cannot be checked here and is not claimed: permissions granted to a
# security group on mailboxes OUTSIDE this batch, and Entra-side use of a
# group (roles, Conditional Access, SharePoint). If a group is in doubt, it
# is in doubt.
#
# Output lands next to this script:
#     Migration-Preflight_<date>_<time>.csv   Type, Object, Severity, Check, Detail
#
# Needs Exchange Online connected in this window.
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

$go = Show-StepInstructions -Title '0 - CLOUD Preflight' -ScriptPath $PSCommandPath -Text @"
STEP 0 -- what a restore cannot bring back. Run before Step 5.

* Reads DL-Input.csv and SMB-Input.csv.
* Nothing is changed. This only reads the cloud.
* STOP = do not migrate as-is.  WARN = works, something changes.
* Output: Migration-Preflight_<date>_<time>.csv next to this script.

Exchange Online must be connected in this window. Click OK to review the
input files and run.
"@
if (-not $go) { return }

function exo { param([string]$n, [hashtable]$p) Invoke-EXOCommand -Name $n -Parameters $p }

$dlNames = @(); $smbNames = @()
try { $dlNames  = Read-InputList -Path (Join-Path $basePath 'DL-Input.csv')  -What 'list' }    catch { Write-Host $_.Exception.Message -ForegroundColor Yellow }
try { $smbNames = Read-InputList -Path (Join-Path $basePath 'SMB-Input.csv') -What 'mailbox' } catch { Write-Host $_.Exception.Message -ForegroundColor Yellow }
if ($dlNames.Count + $smbNames.Count -eq 0) { Write-Host 'Nothing to check.' -ForegroundColor Yellow; return }

$findings = New-Object System.Collections.ArrayList
function Flag {
    param([string]$Type, [string]$Object, [ValidateSet('STOP', 'WARN', 'INFO')][string]$Severity, [string]$Check, [string]$Detail)
    [void]$findings.Add([PSCustomObject]@{ Type = $Type; Object = $Object; Severity = $Severity; Check = $Check; Detail = $Detail })
    $c = switch ($Severity) { 'STOP' { 'Red' } 'WARN' { 'Yellow' } default { 'DarkGray' } }
    Write-Host ("    {0,-4} {1,-22} {2}" -f $Severity, $Check, $Detail) -ForegroundColor $c
}

# ---- resolve the batch first: addresses and DNs for cross-referencing ----
Write-Host ''
Write-Host 'Resolving the batch...' -ForegroundColor Cyan
$dls = @(); $smbs = @()
foreach ($n in $dlNames) {
    try { $dls += exo 'Get-DistributionGroup' @{ Identity = $n; ErrorAction = 'Stop' } }
    catch { Flag 'DL' $n 'STOP' 'not found' "no cloud list answers to '$n' -- fix DL-Input.csv" }
}
foreach ($n in $smbNames) {
    try { $smbs += exo 'Get-Mailbox' @{ Identity = $n; ErrorAction = 'Stop' } }
    catch { Flag 'SMB' $n 'STOP' 'not found' "no cloud mailbox answers to '$n' -- fix SMB-Input.csv" }
}
$batchSmtp = @(($dls + $smbs) | ForEach-Object { "$($_.PrimarySmtpAddress)".ToLower() })
$batchDn   = @(($dls + $smbs) | ForEach-Object { "$($_.DistinguishedName)".ToLower() })
$batchName = @(($dls + $smbs) | ForEach-Object { "$($_.Name)".ToLower() })
function Test-InBatch { param([string]$Ref) $r = "$Ref".ToLower(); return ($batchSmtp -contains $r -or $batchDn -contains $r -or $batchName -contains $r) }

# ---- transport rules, read once ----
$rules = @()
try { $rules = @(exo 'Get-TransportRule' @{ ResultSize = 'Unlimited'; ErrorAction = 'Stop' }) } catch { Write-Host "  (transport rules not read: $($_.Exception.Message))" -ForegroundColor Yellow }
$ruleProps = @('SentTo', 'From', 'SentToMemberOf', 'FromMemberOf', 'ExceptIfSentTo', 'ExceptIfFrom', 'ExceptIfSentToMemberOf', 'ExceptIfFromMemberOf',
               'RedirectMessageTo', 'BlindCopyTo', 'CopyTo', 'AddToRecipients', 'ModerateMessageByUser', 'AddManagerAsRecipientType')
function Get-RulesNaming {
    param([object]$Recipient)
    $hits = @()
    $keys = @("$($Recipient.PrimarySmtpAddress)", "$($Recipient.Name)", "$($Recipient.DisplayName)", "$($Recipient.DistinguishedName)") | Where-Object { $_ } | ForEach-Object { $_.ToLower() }
    foreach ($r in $rules) {
        foreach ($p in $ruleProps) {
            if (-not $r.PSObject.Properties[$p]) { continue }
            foreach ($v in @($r.$p)) { if ($keys -contains "$v".ToLower()) { $hits += "$($r.Name) ($p)"; break } }
        }
    }
    return @($hits | Select-Object -Unique)
}

# ---- lists ----
foreach ($dg in $dls) {
    $smtp = "$($dg.PrimarySmtpAddress)"
    Write-Host "DL   $smtp" -ForegroundColor Cyan
    $n = $findings.Count
    if ("$($dg.IsDirSynced)" -ne 'True') { Flag 'DL' $smtp 'INFO' 'already cloud-only' 'not directory-synced; Step 5 has nothing to move for it' }
    if ("$($dg.GroupType)" -match 'SecurityEnabled') {
        Flag 'DL' $smtp 'STOP' 'security group' 'recreated with a new object ID -- every permission granted TO this group outside this batch is lost (mailboxes, SharePoint, Teams, Conditional Access)'
    }
    foreach ($h in (Get-RulesNaming -Recipient $dg)) { Flag 'DL' $smtp 'STOP' 'transport rule' "rule '$h' references this list and will stop matching after recreate" }

    $dnQ = "$($dg.DistinguishedName)" -replace "'", "''"
    try {
        foreach ($p in @(exo 'Get-DistributionGroup' @{ Filter = "Members -eq '$dnQ'"; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })) {
            if (Test-InBatch $p.PrimarySmtpAddress) { Flag 'DL' $smtp 'INFO' 'nested in' "$($p.PrimarySmtpAddress) -- also migrating; Step 7 re-adds" }
            elseif ("$($p.IsDirSynced)" -eq 'True') { Flag 'DL' $smtp 'WARN' 'nested in' "$($p.PrimarySmtpAddress) stays synced -- membership is lost; add it to DL-Input.csv or re-add on-premises with a contact" }
            else { Flag 'DL' $smtp 'INFO' 'nested in' "$($p.PrimarySmtpAddress) is cloud-only; Step 7 re-adds" }
        }
    } catch { Flag 'DL' $smtp 'INFO' 'nested in' "could not read: $($_.Exception.Message)" }
    try {
        foreach ($p in @(exo 'Get-DistributionGroup' @{ Filter = "ManagedBy -eq '$dnQ'"; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })) {
            Flag 'DL' $smtp 'WARN' 'owner of' "$($p.PrimarySmtpAddress) -- ManagedBy points at the old object; re-add the new list as owner afterwards$(if ("$($p.IsDirSynced)" -eq 'True') { ' (on-premises)' })"
        }
    } catch { }
    try {
        foreach ($p in @(exo 'Get-DistributionGroup' @{ Filter = "ModeratedBy -eq '$dnQ'"; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })) {
            Flag 'DL' $smtp 'WARN' 'moderator of' "$($p.PrimarySmtpAddress) -- ModeratedBy points at the old object"
        }
    } catch { }
    # A security group used as a trustee on mailboxes IN this batch is fine
    # (Step 7 re-grants from the Perms file). Outside the batch is unknown.
    if ($findings.Count -eq $n) { Write-Host '    clean' -ForegroundColor DarkGray }
}

# ---- mailboxes ----
foreach ($mb in $smbs) {
    $smtp = "$($mb.PrimarySmtpAddress)"
    Write-Host "SMB  $smtp" -ForegroundColor Cyan
    $n = $findings.Count
    if ("$($mb.IsDirSynced)" -ne 'True') { Flag 'SMB' $smtp 'INFO' 'already cloud-only' 'not directory-synced; Step 5 has nothing to move for it' }
    if ("$($mb.RecipientTypeDetails)" -ne 'SharedMailbox') { Flag 'SMB' $smtp 'WARN' 'not shared' "$($mb.RecipientTypeDetails) -- Step 6 converts to Shared on restore; check it is not a person" }
    $empty = '00000000-0000-0000-0000-000000000000'
    if ("$($mb.ArchiveGuid)" -and "$($mb.ArchiveGuid)" -ne $empty) { Flag 'SMB' $smtp 'WARN' 'archive' 'has an online archive -- needs a licence once the synced licence group no longer applies' }
    if ("$($mb.LitigationHoldEnabled)" -eq 'True') { Flag 'SMB' $smtp 'WARN' 'litigation hold' 'needs a licence once the synced licence group no longer applies' }
    if (@($mb.InPlaceHolds).Count -gt 0) { Flag 'SMB' $smtp 'WARN' 'in-place hold' "$(@($mb.InPlaceHolds).Count) hold(s) -- needs a licence once the synced licence group no longer applies" }
    try {
        $st = exo 'Get-MailboxStatistics' @{ Identity = $smtp; ErrorAction = 'Stop' }
        $bytes = 0
        if ("$($st.TotalItemSize)" -match '\(([\d,]+) bytes\)') { $bytes = [int64]($Matches[1] -replace ',', '') }
        $gb = [math]::Round($bytes / 1GB, 1)
        if ($bytes -gt 50GB) { Flag 'SMB' $smtp 'WARN' 'size' "$gb GB -- over the 50 GB unlicensed shared-mailbox limit; needs a licence" }
        else { Flag 'SMB' $smtp 'INFO' 'size' "$gb GB" }
    } catch { Flag 'SMB' $smtp 'INFO' 'size' "could not read: $($_.Exception.Message)" }
    foreach ($h in (Get-RulesNaming -Recipient $mb)) { Flag 'SMB' $smtp 'INFO' 'transport rule' "rule '$h' references this mailbox -- the object is restored intact, so this keeps working" }

    $dnQ = "$($mb.DistinguishedName)" -replace "'", "''"
    try {
        foreach ($p in @(exo 'Get-DistributionGroup' @{ Filter = "Members -eq '$dnQ'"; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })) {
            if (Test-InBatch $p.PrimarySmtpAddress) { Flag 'SMB' $smtp 'INFO' 'member of' "$($p.PrimarySmtpAddress) -- also migrating; Step 7 re-adds" }
            elseif ("$($p.IsDirSynced)" -eq 'True') { Flag 'SMB' $smtp 'WARN' 'member of' "$($p.PrimarySmtpAddress) stays synced -- membership is lost; add it to DL-Input.csv or re-add on-premises with a contact" }
            else { Flag 'SMB' $smtp 'INFO' 'member of' "$($p.PrimarySmtpAddress) is cloud-only; Step 7 re-adds" }
        }
    } catch { Flag 'SMB' $smtp 'INFO' 'member of' "could not read: $($_.Exception.Message)" }
    if ($findings.Count -eq $n) { Write-Host '    clean' -ForegroundColor DarkGray }
}

# ---- batch-level ----
$total = $dls.Count + $smbs.Count
if ($total -gt 500) { Flag 'BATCH' "$total objects" 'WARN' 'deletion threshold' 'over the default Entra Connect export deletion threshold (500) -- the sync will refuse the deletes; raise or disable it first, or split the batch' }

# ---- write ----
$outFile = Join-Path $basePath "Migration-Preflight_$(Get-MigrationTimestamp).csv"
if ($findings.Count -eq 0) { '"Type","Object","Severity","Check","Detail"' | Set-Content -LiteralPath $outFile -Encoding UTF8 }
else { $findings | Sort-Object @{ E = { switch ($_.Severity) { 'STOP' { 0 } 'WARN' { 1 } default { 2 } } } }, Type, Object | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8 }
$nStop = @($findings | Where-Object { $_.Severity -eq 'STOP' }).Count
$nWarn = @($findings | Where-Object { $_.Severity -eq 'WARN' }).Count
Write-ExportDone -Path $outFile -Rows $findings.Count
Write-Host "$nStop STOP, $nWarn WARN across $total object(s)." -ForegroundColor $(if ($nStop) { 'Red' } elseif ($nWarn) { 'Yellow' } else { 'Green' })
if ($nStop) { Write-Host 'Resolve every STOP -- remove the object from the input file or accept the loss in writing -- before Step 5.' -ForegroundColor Red }
Write-CompletedLog -Action 'MigrationPreflight' -Target "$total objects" -Result 'Success' -Detail "$nStop STOP, $nWarn WARN -> $([System.IO.Path]::GetFileName($outFile))"
