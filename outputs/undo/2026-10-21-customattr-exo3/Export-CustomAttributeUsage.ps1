#Requires -Version 5.1
<#
.SYNOPSIS
    One CSV of CustomAttribute1-15 on every Exchange object, on-premises and
    in Exchange Online, plus a summary of which attributes are in use.

.DESCRIPTION
    STANDALONE. Not a console tool -- run it from a PowerShell window on the
    jump box:

        .\Export-CustomAttributeUsage.ps1
        .\Export-CustomAttributeUsage.ps1 -ExchangeServer mail01.contoso.local
        .\Export-CustomAttributeUsage.ps1 -SkipCloud        # on-premises only

    Three sources, one file:

      OnPrem   every mail-enabled recipient in on-premises Exchange (all
               domains in the forest -- Set-ADServerSettings -ViewEntireForest),
               read through the window's existing on-premises session if it
               has one, otherwise a new Kerberos session to an Exchange server
               (CAS array names are skipped -- they have no Kerberos SPN)
      Cloud    every recipient in Exchange Online (Get-EXORecipient with the
               Custom property set), including Microsoft 365 groups
               -- through THIS window's existing Exchange Online session.
               The script never signs in; connect first (Connect-ExchangeOnline).
      AD       directory objects that are NOT mail-enabled but still carry an
               extensionAttribute1-15 value. CustomAttributeN IS
               extensionAttributeN on the AD object, so a value on a plain user
               (HR feed, sync rule, Intune) is in use even though Exchange
               never shows it. Skipped if the ActiveDirectory module is missing.

    A synced object appears twice -- once OnPrem, once Cloud -- because it
    exists in both places; the Source column says which copy each row is.

    Output (Desktop by default):
      CustomAttributes-<timestamp>.csv           every object, one row each
      CustomAttributes-<timestamp>-Summary.csv   per attribute: how many
                                                 objects use it in each source,
                                                 how many distinct values, and
                                                 the most common values

    WHY ONE SCRIPT WORKS. The on-premises cmdlets are run INSIDE a remote
    session (Invoke-Command), never imported, so they cannot collide with the
    Exchange Online cmdlets of the same name in this window.

    Read-only. Nothing is changed anywhere.

    Cmdlets checked against Microsoft Learn: Get-EXORecipient -PropertySets
    Custom (CustomAttribute1-15), Get-Recipient (cloud needs
    -RecipientTypeDetails GroupMailbox for Microsoft 365 groups),
    Set-ADServerSettings -ViewEntireForest.

    Author   : Glenn Armbruster
    Created  : 2026-10-21
#>
[CmdletBinding()]
param(
    # On-premises Exchange server FQDN. Found from Active Directory if omitted;
    # asked for if that fails.
    [string]$ExchangeServer,
    # Credentials for the on-premises session. Default: the signed-in account.
    [System.Management.Automation.PSCredential]$Credential,
    [string]$OutputFolder = [Environment]::GetFolderPath('Desktop'),
    [switch]$SkipOnPrem,
    [switch]$SkipCloud,
    [switch]$SkipAD,
    # Kept so earlier command lines still run. The script never signs in to
    # Exchange Online now, with or without it.
    [switch]$NoCloudSignIn
)

$ErrorActionPreference = 'Stop'
$stamp   = Get-Date -Format 'yyyyMMdd-HHmm'
$outFile = Join-Path $OutputFolder "CustomAttributes-$stamp.csv"
$sumFile = Join-Path $OutputFolder "CustomAttributes-$stamp-Summary.csv"
$attrs   = 1..15 | ForEach-Object { "CustomAttribute$_" }
$rows    = New-Object System.Collections.Generic.List[object]
$notes   = New-Object System.Collections.Generic.List[string]

function Write-Step { param([string]$Text) Write-Host "`n=== $Text ===" -ForegroundColor Cyan }

function Add-Row {
    # One output row. $Values holds attributes 1-15 in order.
    param([string]$Source, [string]$Type, [string]$Display, [string]$Smtp, $Synced, [string]$Id, [object[]]$Values)
    $row = [ordered]@{
        Source               = $Source
        RecipientTypeDetails = $Type
        DisplayName          = $Display
        PrimarySmtpAddress   = $Smtp
        IsDirSynced          = $(if ($null -eq $Synced) { '' } else { "$Synced" })
    }
    for ($i = 1; $i -le 15; $i++) { $row["CustomAttribute$i"] = "$($Values[$i - 1])".Trim() }
    $row['Identity'] = $Id
    $rows.Add([PSCustomObject]$row)
}

#region On-premises Exchange ----------------------------------------------------
if (-not $SkipOnPrem) {
    Write-Step 'On-premises Exchange'
    $session = $null; $ownSession = $false
    # 1. REUSE THE WINDOW'S OWN ON-PREMISES SESSION if it has one (Exchange
    #    Management Shell, or a toolkit connect). It already works, so no
    #    server name, SPN or Kerberos question arises. It is NOT closed at
    #    the end -- it belongs to the window.
    if (-not $ExchangeServer) {
        $session = @(Get-PSSession -ErrorAction SilentlyContinue | Where-Object {
                $_.ConfigurationName -eq 'Microsoft.Exchange' -and "$($_.State)" -eq 'Opened' -and
                "$($_.ComputerName)" -notmatch '(?i)outlook\.office365\.com|protection\.outlook\.com' }) | Select-Object -First 1
        if ($session) { Write-Host "Using this window's on-premises session to $($session.ComputerName)." -ForegroundColor DarkGray }
    }
    # 2. Otherwise open one. Candidate servers: -ExchangeServer, else every
    #    Exchange SERVER object in AD -- CAS ARRAY objects excluded, because
    #    an array name is a load-balancer alias with no Kerberos SPN
    #    ('unknown to Kerberos', found 2026-10-21). Each is tried in turn.
    if (-not $session) {
        $candidates = @()
        if ($ExchangeServer) { $candidates = @($ExchangeServer) }
        else {
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
                $cfgNc = (Get-ADRootDSE).configurationNamingContext
                $candidates = @(Get-ADObject -SearchBase $cfgNc -LDAPFilter '(&(objectClass=msExchExchangeServer)(!(objectClass=msExchClientAccessArray)))' -Properties networkAddress |
                    ForEach-Object { @($_.networkAddress) | Where-Object { "$_" -like 'ncacn_ip_tcp:*' } | ForEach-Object { "$_" -replace '^ncacn_ip_tcp:', '' } } |
                    Where-Object { $_ } | Sort-Object -Unique)
                if ($candidates.Count) { Write-Host "Exchange servers found in AD: $($candidates -join ', ')" -ForegroundColor DarkGray }
            } catch { }
            if (-not $candidates.Count) { $candidates = @((Read-Host 'On-premises Exchange server FQDN').Trim()) }
        }
        $lastErr = ''
        foreach ($cand in $candidates) {
            try {
                Write-Host "Connecting to http://$cand/PowerShell/ (Kerberos)..." -ForegroundColor DarkGray
                $sp = @{ ConfigurationName = 'Microsoft.Exchange'; ConnectionUri = "http://$cand/PowerShell/"; Authentication = 'Kerberos'; ErrorAction = 'Stop' }
                if ($Credential) { $sp['Credential'] = $Credential }
                $session = New-PSSession @sp
                $ownSession = $true
                break
            } catch {
                $lastErr = "$($_.Exception.Message)"
                Write-Host "  $cand failed: $lastErr" -ForegroundColor Yellow
            }
        }
    }
    try {
        if (-not $session) { throw "No on-premises Exchange session could be opened. Last error: $lastErr" }
        # Every domain in the forest, not just the server's own.
        Invoke-Command -Session $session -ScriptBlock { Set-ADServerSettings -ViewEntireForest $true } | Out-Null
        Write-Host 'Reading every on-premises recipient (this can take several minutes)...' -ForegroundColor DarkGray
        $onprem = @(Invoke-Command -Session $session -ScriptBlock { Get-Recipient -ResultSize Unlimited })
        foreach ($r in $onprem) {
            Add-Row -Source 'OnPrem' -Type "$($r.RecipientTypeDetails)" -Display "$($r.DisplayName)" -Smtp "$($r.PrimarySmtpAddress)" `
                -Synced $null -Id "$($r.DistinguishedName)" -Values @(1..15 | ForEach-Object { $r."CustomAttribute$_" })
        }
        Write-Host "On-premises: $($onprem.Count) recipients." -ForegroundColor Green
    } catch {
        $notes.Add("On-premises Exchange NOT included: $($_.Exception.Message)")
        Write-Host "On-premises Exchange failed: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        if ($session -and $ownSession) { Remove-PSSession $session -ErrorAction SilentlyContinue }
    }
}
#endregion

#region Active Directory: not mail-enabled, but carrying an extension attribute --
if (-not $SkipAD) {
    Write-Step 'Active Directory (objects that are not mail-enabled)'
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $any = (1..15 | ForEach-Object { "(extensionAttribute$_=*)" }) -join ''
        # mailNickname absent = not an Exchange recipient, so not already in
        # the on-premises rows above.
        $ldap = "(&(|$any)(!(mailNickname=*)))"
        $props = @('displayName', 'userPrincipalName', 'mail', 'objectClass') + (1..15 | ForEach-Object { "extensionAttribute$_" })
        $adCount = 0
        foreach ($dom in @((Get-ADForest).Domains)) {
            try {
                $objs = @(Get-ADObject -Server $dom -LDAPFilter $ldap -Properties $props -ResultPageSize 1000)
                foreach ($adObj in $objs) {
                    Add-Row -Source 'AD (not mail-enabled)' -Type "$(@($adObj.objectClass)[-1])" `
                        -Display $(if ("$($adObj.displayName)") { "$($adObj.displayName)" } else { "$($adObj.Name)" }) `
                        -Smtp "$($adObj.mail)" -Synced $null -Id "$($adObj.DistinguishedName)" -Values @(1..15 | ForEach-Object { $adObj."extensionAttribute$_" })
                }
                $adCount += $objs.Count
                Write-Host "  $dom : $($objs.Count)" -ForegroundColor DarkGray
            } catch {
                $notes.Add("AD domain $dom NOT included: $($_.Exception.Message)")
                Write-Host "  $dom failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        Write-Host "Active Directory: $adCount non-mail objects with an extension attribute." -ForegroundColor Green
    } catch {
        $notes.Add("Active Directory sweep NOT included: $($_.Exception.Message)")
        Write-Host "Active Directory sweep skipped: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}
#endregion

#region Exchange Online ---------------------------------------------------------
if (-not $SkipCloud) {
    Write-Step 'Exchange Online'
    try {
        # THIS SCRIPT NEVER SIGNS IN. It uses the Exchange Online session that
        # THIS PowerShell window already has, or skips the cloud with a
        # message. No Connect-ExchangeOnline, no module import, and no probe
        # query that could trigger a sign-in prompt (2026-10-21).
        #
        # PER WINDOW: Connect-ExchangeOnline belongs to the PowerShell process
        # that ran it. A session in the Admin Console or in another window is
        # not visible here.
        $gci = $null
        if (Get-Module -Name ExchangeOnlineManagement) { $gci = Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue }
        $conns = @()
        if ($gci) { try { $conns = @(& $gci -ErrorAction Stop) } catch { $conns = @() } }
        if (-not $conns.Count) {
            throw ("This PowerShell window has no Exchange Online session. A connection made in the Admin Console or another window " +
                   "does not carry over. Run Connect-ExchangeOnline in THIS window and re-run the script, or use -SkipCloud.")
        }
        $c0 = $conns | Select-Object -First 1
        Write-Host "Using this window's Exchange Online session ($($c0.UserPrincipalName), state $($c0.State), token $($c0.TokenStatus))." -ForegroundColor DarkGray
        Write-Host 'Reading every Exchange Online recipient (this can take several minutes)...' -ForegroundColor DarkGray
        $q = @{ ResultSize = 'Unlimited'; PropertySets = @('Minimum', 'Custom'); Properties = @('DisplayName', 'PrimarySmtpAddress', 'IsDirSynced') }
        try { $cloud = @(Get-EXORecipient @q) }
        catch {
            # IsDirSynced is a nice-to-have column; never lose the whole cloud
            # half because one property name is not accepted.
            Write-Host "Retrying without IsDirSynced ($($_.Exception.Message))" -ForegroundColor Yellow
            $q['Properties'] = @('DisplayName', 'PrimarySmtpAddress')
            $cloud = @(Get-EXORecipient @q)
        }
        # Microsoft 365 groups are only returned when asked for by type; the
        # Id check stops a double row if this module version already included them.
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($r in $cloud) { [void]$seen.Add("$($r.ExternalDirectoryObjectId)") }
        try {
            foreach ($g in @(Get-EXORecipient @q -RecipientTypeDetails GroupMailbox)) {
                if ($seen.Add("$($g.ExternalDirectoryObjectId)")) { $cloud += $g }
            }
        } catch { $notes.Add("Microsoft 365 groups may be missing: $($_.Exception.Message)") }
        foreach ($r in $cloud) {
            Add-Row -Source 'Cloud' -Type "$($r.RecipientTypeDetails)" -Display "$($r.DisplayName)" -Smtp "$($r.PrimarySmtpAddress)" `
                -Synced $r.IsDirSynced -Id "$($r.ExternalDirectoryObjectId)" -Values @(1..15 | ForEach-Object { $r."CustomAttribute$_" })
        }
        Write-Host "Exchange Online: $($cloud.Count) recipients." -ForegroundColor Green
    } catch {
        $notes.Add("Exchange Online NOT included: $($_.Exception.Message)")
        Write-Host "Exchange Online failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}
#endregion

#region Output ------------------------------------------------------------------
Write-Step 'Writing the report'
if ($rows.Count -eq 0) { Write-Host 'Nothing was collected -- see the errors above. No file written.' -ForegroundColor Red; return }
$rows | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-Host "All objects : $outFile  ($($rows.Count) rows)" -ForegroundColor Green

$sources = @($rows | ForEach-Object { $_.Source } | Sort-Object -Unique)
$summary = foreach ($a in $attrs) {
    $used = @($rows | Where-Object { $_.$a })
    $o = [ordered]@{ Attribute = $a; InUseTotal = $used.Count }
    foreach ($s in $sources) { $o["InUse_$s"] = @($used | Where-Object { $_.Source -eq $s }).Count }
    $vals = @($used | Group-Object -Property $a | Sort-Object Count -Descending)
    $o['DistinctValues'] = $vals.Count
    $o['TopValues'] = (@($vals | Select-Object -First 10 | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ' | ')
    $o['Verdict'] = $(if ($used.Count -eq 0) { 'FREE - not used anywhere scanned' } else { 'in use' })
    [PSCustomObject]$o
}
$summary | Export-Csv -LiteralPath $sumFile -NoTypeInformation -Encoding UTF8
Write-Host "Summary     : $sumFile" -ForegroundColor Green
$summary | Select-Object Attribute, InUseTotal, DistinctValues, Verdict | Format-Table -AutoSize | Out-Host

if ($notes.Count) {
    Write-Host 'NOT COMPLETE -- these sources were missing or partial, so a FREE verdict may be wrong:' -ForegroundColor Yellow
    $notes | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
}
#endregion
