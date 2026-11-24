<#
Script Authored by: Glenn A
Date: 05/13/2026
Script Purpose: Process HQ disable requests from CSV:
                - Convert mailboxes to Shared
                - Hide from GAL
                - Delegate mailbox access
                - Grant OneDrive access and notify new owner
                - Disable inbox forwarding rules
                - Output final state and log all actions
#>

#===========================
# Global Settings / Config
#===========================
$ErrorActionPreference = "Stop"
$WarningPreference     = "SilentlyContinue"

# Paths
$CsvFilePath       = "E:\Scripts\M365\DisabledRequests\disablelist.csv"
$LogDir            = "E:\scripts\m365\logs\disabledrequests"

# OneDrive / SPO / SMTP
$OneDriveBaseUrl   = "https://expresspros-my.sharepoint.com/personal"
$SmtpServer        = "smtp.expresspros.com"
$DoNotReplyAddress = "DO_NOT_REPLY@ExpressPros.com"

#===========================
# Build Transcript Log
#===========================
$csvForNames = Import-Csv -Path $CsvFilePath
$names       = $csvForNames | ForEach-Object { $_.mb }
$timestamp   = Get-Date -Format "MM-dd-yyyy.HH.mm.ss"
$logFileName = "Disabled Requests - " + ($names -join "-") + "." + $timestamp + ".log"
$logFilePath = Join-Path -Path $LogDir -ChildPath $logFileName

Start-Transcript -Path $logFilePath

#===========================
# Import CSV (Single Source)
#===========================
$csv = Import-Csv -Path $CsvFilePath

#===========================
# EXECUTION PHASE
#===========================
Write-Host ""
Write-Host "===== EXECUTION PHASE: CHANGES WILL NOW BE APPLIED =====" -ForegroundColor Green
Write-Host ""

try {

    #-----------------------
    # Mailbox Conversion / Hide from GAL / Delegate
    #-----------------------
    foreach ($row in $csv) {
        $mbAlias = $row.mb
        $mailbox = Get-Mailbox -Identity $mbAlias
        $email   = $mailbox.PrimarySmtpAddress

        # Build delegate email (if provided)
        $delegateEmail = $null
        if ($row.delegate -and $row.delegate.Trim() -ne "") {
            $delegateEmail = $row.delegate.Trim() + "@expresspros.com"
        }

        Write-Host ""
        Write-Host "Processing mailbox conversion for $mbAlias ($email)..." -ForegroundColor Magenta

        # Convert to Shared
        Set-Mailbox       -Identity $mbAlias -Type Shared
        Set-RemoteMailbox -Identity $mbAlias -Type Shared

        # Hide from GAL
        Set-RemoteMailbox -Identity $mbAlias -HiddenFromAddressListsEnabled $true

        # Delegate
        if ($delegateEmail) {
            Write-Host "Granting FullAccess to $delegateEmail on $mbAlias..." -ForegroundColor Yellow
            Add-MailboxPermission -Identity $mbAlias -User $delegateEmail -AccessRights FullAccess
        } else {
            Write-Host "No delegate specified for $mbAlias." -ForegroundColor DarkGray
        }
    }

    #-----------------------
    # OneDrive Ownership / Email Notifications
    #-----------------------
    Write-Host ""
    Write-Host "Setting OneDrive permissions and sending notifications..." -ForegroundColor Cyan

    foreach ($row in $csv) {
        if ($row.OD -and $row.OD.Trim() -ne "") {

            $mbAlias = $row.mb
            $mailbox = Get-Mailbox -Identity $mbAlias
            $email   = $mailbox.PrimarySmtpAddress

            $oneDriveOwnerEmail = $row.OD.Trim() + "@expresspros.com"
            $oneDriveUrl        = "$OneDriveBaseUrl/$($email.Replace('@','_').Replace('.','_'))"

            Write-Host "Granting OneDrive Site Collection Admin to $oneDriveOwnerEmail for $email..." -ForegroundColor Yellow
            Set-SPOUser -Site $oneDriveUrl -LoginName $oneDriveOwnerEmail -IsSiteCollectionAdmin $true

            $subject  = "OneDrive Ownership Change Notification"
            $subject2 = "OneDrive Ownership Change Notification - Copy"

            $body = @"
<html>
<body>
<p>To access the user's OneDrive for Business, you can use the following link: <a href='$oneDriveUrl'>OneDrive for Business</a></p>
</body>
</html>
"@

            # Primary notification
            Send-MailMessage -To $oneDriveOwnerEmail -Subject $subject -BodyAsHtml $body -From $DoNotReplyAddress -SmtpServer $SmtpServer

            # Copy to Glenn
            Send-MailMessage -To "glenn.armbruster@expresspros.com" -Subject $subject2 -BodyAsHtml $body -From $DoNotReplyAddress -SmtpServer $SmtpServer
        }
    }

    #-----------------------
    # Final Exchange State Output
    #-----------------------
    Write-Host ""
    Write-Host "Final state from Exchange 2019 (RemoteMailbox):" -ForegroundColor Cyan
    $csv | ForEach-Object { Get-RemoteMailbox $_.mb } |
        Format-Table PrimarySmtpAddress, RecipientTypeDetails, HiddenFromAddressListsEnabled -AutoSize

    Write-Host ""
    Write-Host "Final state from Exchange Online (Mailbox):" -ForegroundColor Cyan
    $csv | ForEach-Object { Get-Mailbox $_.mb } |
        Format-Table PrimarySmtpAddress, RecipientTypeDetails, ForwardingAddress -AutoSize

    #-----------------------
    # Build Output File Path (unchanged behavior)
    #-----------------------
    $date     = Get-Date -Format FileDate
    $outfile1 = "C:\Users\"
    $id1      = whoami
    $id       = $id1 -split '\\' | Select-Object -Last 1
    $outfile2 = "\Express Employment Professionals\Network Services - Messaging Team - Daily Processes\DisabledUserRequests\logs.txt"
    $outfile  = $outfile1 + $id + $outfile2

    Write-Host ""
    Write-Host "All selected actions completed successfully." -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Script execution halted due to error (fail-fast behavior)." -ForegroundColor Red
    Stop-Transcript
    exit
}

#===========================
# End Transcript
#===========================
Stop-Transcript

# =========================
# LOGGING - DISABLE USER
# =========================
$CompletedTaskFolder = "E:\Scripts\M365\Logs\Completed"
$CompletedTaskFile   = Join-Path $CompletedTaskFolder "CompletedTasks.csv"

if (-not (Test-Path -LiteralPath $CompletedTaskFolder)) {
    New-Item -Path $CompletedTaskFolder -ItemType Directory -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $CompletedTaskFile)) {
    "Task,DateTime" | Out-File -FilePath $CompletedTaskFile -Encoding utf8
}

$LogDateTime = Get-Date -Format "yyyy-MM-dd HHmm"

foreach ($row in $csv) {
    if ($row.mb -and $row.mb.Trim() -ne "") {
        Add-Content -Path $CompletedTaskFile -Value ('"Disable User {0}","{1}"' -f $row.mb.Trim(), $LogDateTime)
    }
}