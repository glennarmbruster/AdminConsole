#Requires -Version 5.1
<#
    Set-FSC2027-CalendarPerms.ps1
    One-off: strip fsc2027-informationsupportservices back to zero, then grant
    the agreed calendar access.

    Run it in an Exchange Online session. Not a console tool -- no form, no
    Common.ps1, no logging.

    WHAT IT DOES
      1. removes ALL Full Access
      2. removes ALL Send As
      3. clears Send on Behalf
      4. removes all explicit Inbox and Calendar folder permissions
      5. sets Default and Anonymous on the calendar to None
      6. grants Editor   (add + delete) to 7 people
      7. grants Reviewer (view only)     to 17 people
      8. prints the result

    FOUR THINGS TO KNOW BEFORE RUNNING IT

      NT AUTHORITY\SELF IS NEVER REMOVED. That entry is the mailbox's access to
      itself, not a delegation. Stripping it breaks the mailbox.

      STEP 5 IS THE ONE TO THINK ABOUT. Default is what EVERYONE gets, and on
      most calendars it is AvailabilityOnly -- free/busy in the scheduling
      assistant. Setting it to None is what "zero permissions" literally means
      and it is what this does, but it also means nobody outside the list below
      can see even busy blocks. If you want free/busy preserved for the company,
      delete the two Set-MailboxFolderPermission lines in step 5.

      Default and Anonymous are SET, not removed. Every folder has both and
      Remove-MailboxFolderPermission refuses them.

      ONLY INBOX AND CALENDAR FOLDERS ARE CLEARED. A full walk of every folder
      in the mailbox is a different and much slower job. If this mailbox has
      permissions on some other folder, they survive.

      THE NAMES ARE DISPLAY NAMES. Exchange resolves them, and it will fail on
      any that are ambiguous or spelled differently in the directory. Each grant
      is independent, so one bad name costs that grant and not the run -- check
      the table at the end for anyone missing.
#>

$mbx = 'fsc2027-InformationSupportServices@ExpressPros.com'

# Folder NAMES, not the words 'Inbox' and 'Calendar'. Both are localised and
# renameable; FolderType is the invariant.
$inboxName = (Get-MailboxFolderStatistics -Identity $mbx -FolderScope Inbox |
              Where-Object { $_.FolderType -eq 'Inbox' } | Select-Object -First 1).Name
$calName   = (Get-MailboxFolderStatistics -Identity $mbx -FolderScope Calendar |
              Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1).Name
$inbox = "${mbx}:\${inboxName}"
$cal   = "${mbx}:\${calName}"

$editor = @(
    'Jodi Strum', 'Karen Glenn', 'Ben Meyer', 'Jennifer DeSpain',
    'Debi Nimz', 'Cleasha Baeriswyl', 'Tracey Lonteen'
)

$reviewer = @(
    'Kim Baig', 'Daniel Priest', 'Steven Jones', 'Brent Wendorff', 'Joshua Bounds',
    'Seth Copple', 'Bradley Murray', 'Lori Nissen', 'Katherine Lopez', 'Michelle Lane',
    'Brenna Gay', 'Andrew Williams', 'Roger Dolese', 'Barbara Hardwick',
    'Richard Edwards', 'Christopher Cobb', 'Mireya Reyes'
)

# --- 1. Full Access -----------------------------------------------------------
# Inherited entries come from the database, not from a decision about this
# mailbox, and Remove-MailboxPermission throws on them.
Get-MailboxPermission -Identity $mbx |
    Where-Object { $_.User -notlike 'NT AUTHORITY\*' -and -not $_.IsInherited -and -not $_.Deny } |
    ForEach-Object { Remove-MailboxPermission -Identity $mbx -User $_.User -AccessRights $_.AccessRights -Confirm:$false -ErrorAction Continue }

# --- 2. Send As ---------------------------------------------------------------
Get-RecipientPermission -Identity $mbx |
    Where-Object { $_.Trustee -notlike 'NT AUTHORITY\*' -and $_.AccessControlType -eq 'Allow' } |
    ForEach-Object { Remove-RecipientPermission -Identity $mbx -Trustee $_.Trustee -AccessRights SendAs -Confirm:$false -ErrorAction Continue }

# --- 3. Send on Behalf --------------------------------------------------------
Set-Mailbox -Identity $mbx -GrantSendOnBehalfTo $null

# --- 4. Inbox and Calendar folder permissions ---------------------------------
foreach ($f in @($inbox, $cal)) {
    Get-MailboxFolderPermission -Identity $f |
        Where-Object { $_.User.DisplayName -notin @('Default', 'Anonymous') } |
        ForEach-Object { Remove-MailboxFolderPermission -Identity $f -User $_.User.DisplayName -Confirm:$false -ErrorAction Continue }
}

# --- 5. the two that cannot be removed ----------------------------------------
# Delete these two lines if the company should keep free/busy. See the header.
Set-MailboxFolderPermission -Identity $cal -User Default   -AccessRights None
Set-MailboxFolderPermission -Identity $cal -User Anonymous -AccessRights None

# --- 6 + 7. grant -------------------------------------------------------------
$editor   | ForEach-Object { Add-MailboxFolderPermission -Identity $cal -User $_ -AccessRights Editor   -ErrorAction Continue }
$reviewer | ForEach-Object { Add-MailboxFolderPermission -Identity $cal -User $_ -AccessRights Reviewer -ErrorAction Continue }

# --- 8. check -----------------------------------------------------------------
'--- Calendar ---'
Get-MailboxFolderPermission -Identity $cal |
    Select-Object @{ n = 'User'; e = { $_.User.DisplayName } }, @{ n = 'Rights'; e = { $_.AccessRights -join ', ' } } |
    Sort-Object Rights, User | Format-Table -AutoSize

'--- Should all be empty ---'
"Full Access   : $(@(Get-MailboxPermission -Identity $mbx | Where-Object { $_.User -notlike 'NT AUTHORITY\*' -and -not $_.IsInherited }).Count)"
"Send As       : $(@(Get-RecipientPermission -Identity $mbx | Where-Object { $_.Trustee -notlike 'NT AUTHORITY\*' }).Count)"
"Send on Behalf: $(@((Get-Mailbox -Identity $mbx).GrantSendOnBehalfTo).Count)"
"Inbox perms   : $(@(Get-MailboxFolderPermission -Identity $inbox | Where-Object { $_.User.DisplayName -notin @('Default','Anonymous') }).Count)"
