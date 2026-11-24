#Requires -Version 5.1
<#
    Export-CustomAttributeUsage-Cloud.ps1

    Exchange Online only. Every cloud recipient with DisplayName,
    PrimarySmtpAddress and CustomAttribute1-15, to one CSV on the Desktop.
    Prints how many objects use each attribute. Read-only.

    Run in a new PowerShell window:
        .\Export-CustomAttributeUsage-Cloud.ps1

    Author   : Glenn Armbruster
    Created  : 2026-10-21
#>

# The sign-in failed with "not valid in the 'SslProtocolType' enumeration":
# the window's TLS setting had been changed. TLS 1.2 fixes it.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop

Write-Host 'Reading every Exchange Online recipient (this can take several minutes)...' -ForegroundColor Cyan
$recipients = Get-EXORecipient -ResultSize Unlimited -PropertySets Minimum, Custom -Properties DisplayName, PrimarySmtpAddress

$file = Join-Path ([Environment]::GetFolderPath('Desktop')) ("CustomAttributes-Cloud-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmm'))
$recipients |
    Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails,
        CustomAttribute1, CustomAttribute2, CustomAttribute3, CustomAttribute4, CustomAttribute5,
        CustomAttribute6, CustomAttribute7, CustomAttribute8, CustomAttribute9, CustomAttribute10,
        CustomAttribute11, CustomAttribute12, CustomAttribute13, CustomAttribute14, CustomAttribute15 |
    Export-Csv -LiteralPath $file -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $(@($recipients).Count) rows to $file" -ForegroundColor Green

foreach ($n in 1..15) {
    $used = @($recipients | Where-Object { "$($_."CustomAttribute$n")".Trim() }).Count
    Write-Host ("CustomAttribute{0,-3} {1,7} in use" -f $n, $used)
}
