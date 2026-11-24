# Set-AccessGroupOwners.ps1
# Assigns owners from manager-user.csv to DLs listed in manager-dl.csv

# Load manager ? user mapping
$Managers = Import-Csv "E:\scripts\m365\dlconversion\set-manager\manager-user.csv"

# Build list of manager email addresses exactly as provided
$ResolvedManagers = $Managers.email

# Load DLs to update
$DLs = Import-Csv "E:\scripts\m365\dlconversion\set-manager\manager-dl.csv"

# Update each DL
foreach ($dl in $DLs.Name) {
    Write-Host "Setting ManagedBy for $dl"
    Set-DistributionGroup -Identity $dl `
        -ManagedBy $ResolvedManagers `
        -BypassSecurityGroupManagerCheck
}
