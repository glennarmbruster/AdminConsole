# Set-AccessGroupOwners.ps1
# Updates ManagedBy for all SMB-Access-* groups based on managers.csv

$Managers = Import-Csv "E:\scripts\m365\dlconversion\managers.csv"

# Resolve all manager emails to valid recipients
$ResolvedManagers = @()

foreach ($m in $Managers.email) {
    $m = $m.Trim()
    if ($m -eq "") { continue }

    $Resolved = Get-Recipient -Identity $m -ErrorAction SilentlyContinue
    if ($Resolved) {
        $ResolvedManagers += $Resolved.PrimarySmtpAddress
    }
}

# Update all SMB-Access-* groups
$Groups = Get-DistributionGroup -Filter "Name -like 'SMB-Access-*'"

foreach ($g in $Groups) {
    Write-Host "Setting ManagedBy for $($g.Name)"
    Set-DistributionGroup -Identity $g.Identity `
        -ManagedBy $ResolvedManagers `
        -BypassSecurityGroupManagerCheck
}
