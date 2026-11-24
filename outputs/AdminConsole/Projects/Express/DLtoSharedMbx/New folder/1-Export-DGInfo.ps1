# Export-DGInfo.ps1
# Exports DG name, owners, members, and email addresses for DLs listed in DL-Input.csv

$InputFile  = ".\DL-Input.csv"
$OutputFile = ".\DGExport.csv"
$Results    = @()

$List = Import-Csv $InputFile

foreach ($entry in $List) {
    $DG = Get-DistributionGroup -Identity $entry.Name -ErrorAction SilentlyContinue
    if ($DG) {
        $Owners  = ($DG.ManagedBy) -join ";"
        $Members = (Get-DistributionGroupMember -Identity $DG.Identity -ResultSize Unlimited | Select-Object -ExpandProperty PrimarySmtpAddress) -join ";"
        $Emails  = $DG.EmailAddresses -join ";"

        $Results += [pscustomobject]@{
            DGName     = $DG.DisplayName
            DGIdentity = $DG.Identity
            Owners     = $Owners
            Members    = $Members
            Emails     = $Emails
        }
    }
}

$Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
