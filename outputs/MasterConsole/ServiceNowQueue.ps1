#Requires -Version 5.1
<#
    ServiceNowQueue.ps1
    Opens the team's ServiceNow ticket queue in the everyday Edge profile.

    Author   : Glenn Armbruster
    Revised  : 2026-07-29
#>
# TOOLKIT-HELP-BEGIN
# Opens the team's ServiceNow ticket queue in the browser.
#
# It uses the everyday Edge profile rather than whichever one Edge happened to
# have open, so it signs in as the right account without prompting.
#
# That is all it does -- no connection to anything, nothing read or changed.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

. (Join-Path $PSScriptRoot 'Common.ps1')

$Url = 'https://expresspros.service-now.com/task_list.do?sysparm_first_row=1&sysparm_query=assignment_group%3D743e336d1b2fbb00ce6ddd77cc4bcb17%5Estate%21%3D6%5EORstate%3D%5Estate%21%3D3%5EORstate%3D%5Estate%21%3D7%5EORstate%3D%5Estate%21%3D4%5EORstate%3D%5Estate%21%3D8%5EORstate%3DNULL&sysparm_view='

# Purpose rather than a profile name: the mapping lives in Common.ps1
# ($script:BrowserProfiles), so renaming the profile is one edit in one place
# rather than a hunt through every script that opens a browser.
$status = Start-UrlInBrowserProfile -Url $Url -Purpose 'ServiceNow'
Write-Host $status
