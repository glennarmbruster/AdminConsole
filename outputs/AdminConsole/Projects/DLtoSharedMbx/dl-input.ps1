#Requires -Version 5.1
# TOOLKIT-HELP-RUNS: nothing. It opens a CSV in whatever your machine uses for .csv files.
# TOOLKIT-HELP-BEGIN
# Opens DL-Input.csv for editing -- the list of distribution lists that Step 1
# will export.
#
# Two columns, and every row needs the comma:
#
#     Name,NewName
#     Jobs.ColoradoSpringsCO,
#     Payroll_F26,Payroll.ColoradoSpringsCO
#
# NewName blank means the list gets DELETED on-premises at Step 2. A NewName
# present means it is mail-disabled and KEPT -- which is how the Payroll groups
# survive. Getting that column wrong is the one mistake in this pipeline that
# cannot be undone.
#
# Save the file before running Step 1.
# TOOLKIT-HELP-END
<#
    dl-input.ps1
    Opens DL-Input.csv.

    Replaces a one-line version that read

        ii: e:\scripts\AdminConsole\Projects\DLConversion\dl-input.ps1

    (that folder has since been renamed DLtoSharedMbx; the line is quoted as
    it was, because the point is the two faults in it, not the path)

    which had two faults: 'ii:' with a colon is not valid PowerShell, and the
    path pointed at this script rather than at the CSV. The button could never
    have worked.

    Author   : Glenn Armbruster
    Revised  : 2026-08-07
#>

$csv = Join-Path $PSScriptRoot 'DL-Input.csv'

if (-not (Test-Path -LiteralPath $csv)) {
    # Created rather than reported missing: an empty input file with the right
    # headers is more useful than an error, and Step 1 needs those two columns
    # to exist regardless.
    'Name,NewName' | Set-Content -LiteralPath $csv -Encoding UTF8
    Write-Host "Created $csv"
}

Write-Host "Opening $csv"
Invoke-Item -LiteralPath $csv
