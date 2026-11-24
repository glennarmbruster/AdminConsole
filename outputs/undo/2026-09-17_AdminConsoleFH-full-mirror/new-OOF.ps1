#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-RUNS: Exchange Online
# TOOLKIT-HELP-BEGIN
# Reads and sets automatic replies on a mailbox.
# TOOLKIT-HELP-END
<#
.SYNOPSIS
    Console utility to read and set the Out of Office / automatic reply configuration of
    any Exchange Online mailbox, with separate INTERNAL and EXTERNAL messages.

.DESCRIPTION
    Console refactor of Set-OutOfOffice-GUI.ps1. Same pipeline, no WinForms:

        validate mailbox -> compose HTML -> Set-MailboxAutoReplyConfiguration ->
        read the configuration back and display it -> (optionally) send an HTML test email

    The rich-text editors are replaced by plain text plus a small inline markup, so the
    same formatting can be expressed on a command line or in a text file:

        **bold**            <b>
        *italic*            <i>
        __underline__       <u>
        ~~strike~~          <s>
        [text](url)         <a href="url">
        - item              bullet list item (also * item)
        blank line          paragraph break

    Anything not listed is passed through as literal text and HTML-encoded, so an
    ampersand or an angle bracket in the message cannot break the markup.

    Messages can come from -InternalMessage / -ExternalMessage, from a file with
    -InternalMessagePath / -ExternalMessagePath, or be typed at the prompt. Text that
    already looks like HTML (starts with a tag) is stored as-is; use -NoHtml to store
    plain text untouched.

.PARAMETER Mailbox
    Mailbox to configure. Aliases and proxy addresses are accepted - the script resolves
    them to an identity form the auto-reply cmdlets can actually bind to.

.PARAMETER State
    Enabled, Scheduled or Disabled. Scheduled requires -StartTime and -EndTime.

.PARAMETER ShowOnly
    Read and display the current configuration, then exit without writing anything.

.EXAMPLE
    .\new-OOF.ps1 -Mailbox jane.doe@contoso.com -ShowOnly

.EXAMPLE
    .\new-OOF.ps1 -Mailbox jane.doe@contoso.com -State Scheduled `
        -StartTime '2026-09-08 08:00' -EndTime '2026-09-15 17:00' `
        -InternalMessagePath .\oof.txt -SameAsInternal -ExternalAudience All

.EXAMPLE
    .\new-OOF.ps1 -Mailbox jane.doe@contoso.com -State Disabled -Force

.NOTES
    Requires: Install-Module ExchangeOnlineManagement -Scope CurrentUser
    Test email uses Microsoft Graph when Microsoft.Graph.Users.Actions is present and
    falls back to authenticated SMTP.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Position = 0)]
    [Alias('Email', 'Identity', 'SmtpAddress')]
    [string] $Mailbox,

    [ValidateSet('Enabled', 'Scheduled', 'Disabled')]
    [string] $State = 'Enabled',

    [datetime] $StartTime,
    [datetime] $EndTime,

    [ValidateSet('All', 'Known', 'None')]
    [string] $ExternalAudience = 'All',

    [string] $InternalMessage,
    [string] $ExternalMessage,
    [string] $InternalMessagePath,
    [string] $ExternalMessagePath,

    [switch] $SameAsInternal,
    [switch] $NoHtml,
    [switch] $ShowOnly,
    [switch] $LoadCurrent,
    [string] $SaveHtmlPath,

    [switch] $SendTestEmail,
    # BLANK ON PURPOSE, both of them. These shipped holding one named
    # person's address and a mail server, which is a tenant fact in a .ps1
    # and an individual's address in a file that goes to other people.
    # -SendTestEmail now requires them to be supplied, which is the only
    # honest default: nothing here can know which relay this site uses.
    [string] $From       = '',
    [string] $SmtpServer = '',
    [int]    $SmtpPort   = 587,

    [switch] $Force,
    [switch] $NonInteractive
)

$ErrorActionPreference = 'Stop'
# --------------------------------------------------------------------------- #
# region  Toolkit integration                                                  #
# --------------------------------------------------------------------------- #

$root = $PSScriptRoot
while ($root -and -not (Test-Path -LiteralPath (Join-Path $root 'Common.ps1'))) {
    $parent = Split-Path -Parent $root
    if (-not $parent -or $parent -eq $root) { $root = ''; break }
    $root = $parent
}
if (-not $root) { throw "Common.ps1 not found above '$PSScriptRoot'." }
. (Join-Path $root 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

# EVERY Write-Host IN THIS FILE IS CAPTURED, without editing any of them.
#
# The engine functions below were written for a console and say useful things
# with Write-Host -- Show-OofConfiguration in particular explains a schedule
# that has already closed, or an audience of None that means nobody outside
# gets a reply. Throwing that away and re-deriving it in the window would be
# two versions of the same warning, and the second one would rot.
#
# So Write-Host is shadowed for the duration of this script: it appends to a
# sink the window reads, and still calls the real one for anybody running this
# file from a shell. The real cmdlet is named in full so this cannot recurse.
$script:OutSink = New-Object System.Collections.ArrayList
function Write-Host {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]$Object,
        [switch]$NoNewline,
        $Separator,
        $ForegroundColor,
        $BackgroundColor
    )
    [void]$script:OutSink.Add("$Object")
    Microsoft.PowerShell.Utility\Write-Host @PSBoundParameters
}

# endregion ----------------------------------------------------------------- #


# --------------------------------------------------------------------------- #
# region  Console output                                                       #
# --------------------------------------------------------------------------- #

function Write-Log {
    param([string] $Message, [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string] $Level = 'INFO')

    $color = switch ($Level) {
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Gray' }
    }
    Write-Host ('[{0:HH:mm:ss}] {1,-5} {2}' -f (Get-Date), $Level, $Message) -ForegroundColor $color
}

function Write-Heading {
    param([string] $Text)
    Write-Host ''
    Write-Host "== $Text " -ForegroundColor Cyan -NoNewline
    Write-Host ('=' * [math]::Max(0, 74 - $Text.Length)) -ForegroundColor DarkCyan
}

function Read-Line {
    param([string] $Prompt, [string] $Default)

    if ($NonInteractive) { return $Default }
    $suffix = if ($Default) { " [$Default]" } else { '' }
    $answer = Read-Host -Prompt "$Prompt$suffix"
    if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
    return $answer.Trim()
}

function Read-Choice {
    <# Asks for one of $Options, repeating until a valid answer or the default. #>
    param([string] $Prompt, [string[]] $Options, [string] $Default)

    if ($NonInteractive) { return $Default }
    while ($true) {
        $answer = Read-Line -Prompt "$Prompt ($($Options -join '/'))" -Default $Default
        $match  = $Options | Where-Object { $_ -like "$answer*" } | Select-Object -First 1
        if ($match) { return $match }
        Write-Log "'$answer' is not one of: $($Options -join ', ')" 'WARN'
    }
}

function Confirm-Action {
    param([string] $Question)
    if ($Force -or $NonInteractive) { return $true }
    return (Read-Choice -Prompt $Question -Options @('Yes', 'No') -Default 'No') -eq 'Yes'
}

function Read-MultiLineText {
    <#
        Replaces the GUI's rich-text editor. Reads lines until a single '.' is entered.
        Returns $null if the operator cancels, so callers can tell empty from abandoned.
    #>
    param([string] $Title, [string] $Seed)

    Write-Heading $Title
    Write-Host '  Markup:  **bold**  *italic*  __underline__  ~~strike~~  [text](url)' -ForegroundColor DarkGray
    Write-Host '           lines starting "- " become a bullet list; a blank line is a paragraph break' -ForegroundColor DarkGray
    Write-Host '  Finish with a single "." on its own line.  ":cancel" abandons the message.' -ForegroundColor DarkGray
    if ($Seed) {
        Write-Host '  Current text (copy or retype what you want to keep):' -ForegroundColor DarkGray
        foreach ($line in ($Seed -split "\r\n|\n|\r")) { Write-Host "    | $line" -ForegroundColor DarkGray }
    }
    Write-Host ''

    $lines = New-Object System.Collections.Generic.List[string]
    while ($true) {
        $line = Read-Host
        if ($line -eq '.')       { break }
        if ($line -eq ':cancel') { return $null }
        $lines.Add($line)
    }
    return ($lines -join [Environment]::NewLine)
}

# endregion ----------------------------------------------------------------- #

# --------------------------------------------------------------------------- #
# region  HTML generation (pure functions - unit testable)                     #
# --------------------------------------------------------------------------- #

$script:HtmlHead = @'
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
</head>
<body style="margin:0; padding:0;">
'@

function Format-InlineMarkup {
    <# HTML-encodes a line, then turns the inline markup into tags. #>
    param([string] $Text)

    if ([string]::IsNullOrEmpty($Text)) { return '' }

    # Encode first: after this the text holds no literal < > & that markup could produce,
    # so every tag emitted below is one this function put there deliberately.
    $html = [System.Net.WebUtility]::HtmlEncode($Text)
    $html = $html -replace '  ', '&nbsp;&nbsp;'

    $html = [regex]::Replace($html, '\[([^\]]+)\]\(\s*([^)\s]+)\s*\)', {
        param($m) '<a href="{0}">{1}</a>' -f $m.Groups[2].Value, $m.Groups[1].Value })
    $html = [regex]::Replace($html, '\*\*(?=\S)(.+?)(?<=\S)\*\*',        '<b>$1</b>')
    $html = [regex]::Replace($html, '(?<!\*)\*(?=\S)([^*]+?)(?<=\S)\*(?!\*)', '<i>$1</i>')
    $html = [regex]::Replace($html, '__(?=\S)(.+?)(?<=\S)__',            '<u>$1</u>')
    $html = [regex]::Replace($html, '~~(?=\S)(.+?)(?<=\S)~~',            '<s>$1</s>')
    return $html
}

function ConvertTo-HtmlFromMarkup {
    <#
        Builds the HTML body Exchange will store, from plain text plus inline markup.
        Paragraphs come from lines, bullet lists from lines starting "- " or "* ".
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Text,
        [string] $DefaultFamily = 'Segoe UI',
        [double] $DefaultSize   = 11
    )

    $out      = New-Object System.Collections.Generic.List[string]
    $inList   = $false
    $emptyRun = 0

    foreach ($raw in ($Text -split "\r\n|\n|\r")) {
        $line   = $raw.TrimEnd()
        $bullet = $line -match '^\s*[-*+]\s+(.*)$'
        $body   = if ($bullet) { $Matches[1] } else { $line.Trim() }

        if ([string]::IsNullOrWhiteSpace($body)) {
            if ($inList) { $out.Add('</ul>'); $inList = $false }
            $emptyRun++
            continue
        }

        # A blank line between text blocks becomes paragraph separation.
        if ($emptyRun -gt 0 -and $out.Count -gt 0) { $out.Add('<p style="margin:0;">&nbsp;</p>') }
        $emptyRun = 0

        $content = Format-InlineMarkup -Text $body

        if ($bullet) {
            if (-not $inList) { $out.Add('<ul style="margin:0 0 0 18px; padding:0;">'); $inList = $true }
            $out.Add("  <li style=""margin:0;"">$content</li>")
        }
        else {
            if ($inList) { $out.Add('</ul>'); $inList = $false }
            $out.Add("<p style=""margin:0 0 6px 0;"">$content</p>")
        }
    }
    if ($inList) { $out.Add('</ul>') }

    $wrapperStyle = "margin:0; padding:0; font-family:'$DefaultFamily', Arial, Helvetica, sans-serif; font-size:$([math]::Round($DefaultSize,1))pt; color:#201f1e; line-height:1.5;"
    return @"
$script:HtmlHead<div style="$wrapperStyle">
$($out -join "`n")
</div>
</body>
</html>
"@
}

function ConvertTo-OofBody {
    <# Decides whether text needs converting, is already HTML, or must be left alone. #>
    param([AllowEmptyString()][string] $Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    if ($NoHtml)                             { return $Text }
    if ($Text.TrimStart() -match '^<(!DOCTYPE|html|div|p|body|span|table|font)\b') { return $Text }
    return ConvertTo-HtmlFromMarkup -Text $Text
}

function Test-EmailAddress {
    param([string] $Address)
    if ([string]::IsNullOrWhiteSpace($Address)) { return $false }
    try   { $null = [System.Net.Mail.MailAddress]::new($Address.Trim()); return $true }
    catch { return $false }
}

function Get-HtmlTextLength {
    <# Rough visible-text length, used to detect an empty message. #>
    param([string] $Html)
    if (-not $Html) { return 0 }
    $text = [regex]::Replace($Html, '(?is)<(script|style).*?</\1>', '')
    $text = [regex]::Replace($text, '(?s)<[^>]+>', '')
    $text = [System.Net.WebUtility]::HtmlDecode($text) -replace '\s|\u00a0', ''
    return $text.Length
}

function ConvertFrom-HtmlToText {
    <# Turns a stored HTML message back into something readable in a console. #>
    param([string] $Html)

    if ([string]::IsNullOrWhiteSpace($Html)) { return '' }
    $text = [regex]::Replace($Html, '(?is)<(script|style|head).*?</\1>', '')
    $text = [regex]::Replace($text, '(?i)<br\s*/?>', "`n")
    $text = [regex]::Replace($text, '(?i)</(p|div|li|tr|h\d)\s*>', "`n")
    $text = [regex]::Replace($text, '(?i)<li\b[^>]*>', '- ')
    $text = [regex]::Replace($text, '(?s)<[^>]+>', '')
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    $text = $text -replace "[ \t]+\r?\n", "`n"
    $text = [regex]::Replace($text, "(\r?\n){3,}", "`n`n")
    return $text.Trim()
}

# endregion ----------------------------------------------------------------- #

# --------------------------------------------------------------------------- #
# region  Exchange helpers                                                     #
# --------------------------------------------------------------------------- #

function Connect-ExchangeIfNeeded {
    param([string] $UserPrincipalName)

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw 'The ExchangeOnlineManagement module is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $connected = $false
    try {
        $connected = @(Get-ConnectionInformation -ErrorAction Stop |
                       Where-Object { $_.State -eq 'Connected' }).Count -gt 0
    }
    catch { $connected = $false }

    if ($connected) { return 'Already connected to Exchange Online.' }

    if ($UserPrincipalName) { Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false }
    else                    { Connect-ExchangeOnline -ShowBanner:$false }
    return 'Connected to Exchange Online.'
}

function Resolve-Mailbox {
    <#
        Resolves an address to a mailbox and returns the identity form that the
        auto-reply cmdlets can actually bind to (aliases/proxy addresses often can't).
    #>
    param([Parameter(Mandatory)][string] $Address)

    $mailbox = $null
    try { $mailbox = Get-Mailbox -Identity $Address -ErrorAction Stop }
    catch {
        $recipient = $null
        try { $recipient = Get-Recipient -Identity $Address -ErrorAction Stop } catch { $recipient = $null }
        if ($recipient) {
            throw "'$Address' exists as $($recipient.RecipientTypeDetails) ($($recipient.RecipientType)), which has no Out of Office configuration. Only mailboxes support automatic replies."
        }
        throw "No mailbox found for '$Address'. Check the spelling, try the primary SMTP address or UPN instead of an alias, confirm the mailbox is in Exchange Online (not on-premises), and confirm your account has rights to see it."
    }

    $candidates = @(
        [string]$mailbox.ExchangeGuid
        [string]$mailbox.PrimarySmtpAddress
        [string]$mailbox.UserPrincipalName
        [string]$mailbox.DistinguishedName
        [string]$mailbox.Alias
        $Address
    ) | Where-Object { $_ -and $_ -ne '00000000-0000-0000-0000-000000000000' } | Select-Object -Unique

    foreach ($candidate in $candidates) {
        try {
            $config = Get-MailboxAutoReplyConfiguration -Identity $candidate -ErrorAction Stop
            if ($config) {
                return [pscustomobject]@{ Mailbox = $mailbox; Identity = $candidate; Config = $config }
            }
        }
        catch { continue }
    }

    throw "Found mailbox '$($mailbox.DisplayName)' <$($mailbox.PrimarySmtpAddress)> but could not read its automatic reply configuration with any identity form ($($candidates -join ', ')). Verify your Exchange role assignment (Mail Recipients / Organization Management)."
}

function Show-OofConfiguration {
    <#
        Prints a stored configuration and, more usefully, the ways it can be switched on
        and still send nothing: a schedule that has closed, or an audience of None.
    #>
    param([Parameter(Mandatory)] $Config, [string] $Title = 'Stored in Exchange')

    $state    = [string]$Config.AutoReplyState
    $audience = [string]$Config.ExternalAudience

    Write-Heading $Title
    Write-Host ('  State            : {0}' -f $state)
    Write-Host ('  ExternalAudience : {0}' -f $audience)
    if ($state -eq 'Scheduled') {
        Write-Host ('  Window           : {0}  ->  {1}' -f $Config.StartTime, $Config.EndTime)
    }

    $notes = New-Object System.Collections.Generic.List[string]
    if ($state -eq 'Disabled') {
        $notes.Add('AutoReplyState is Disabled - no automatic reply will be sent.')
    }
    if ($state -eq 'Scheduled') {
        # Scheduled with a past end date still reads as "on" in most views and sends
        # nothing at all, which is why the window is printed rather than just the state.
        try { if ([datetime]$Config.EndTime   -lt (Get-Date)) { $notes.Add('THIS SCHEDULE HAS ALREADY ENDED. The state still says Scheduled but nothing is being sent.') } } catch { }
        try { if ([datetime]$Config.StartTime -gt (Get-Date)) { $notes.Add('This schedule has not started yet, so nothing is being sent right now.') } } catch { }
    }
    if ($state -ne 'Disabled') {
        if ($audience -eq 'None')  { $notes.Add('External audience is None, so NOBODY outside the organisation receives a reply, whatever the external message says.') }
        if ($audience -eq 'Known') { $notes.Add('External audience is Known, so only senders in the mailbox Contacts get the external reply.') }
        if ((Get-HtmlTextLength ([string]$Config.InternalMessage)) -eq 0) { $notes.Add('The internal message is EMPTY while auto-reply is on.') }
        if ((Get-HtmlTextLength ([string]$Config.ExternalMessage)) -eq 0) { $notes.Add('The external message is EMPTY while auto-reply is on.') }
    }

    foreach ($note in $notes) { Write-Log $note 'WARN' }

    Write-Host ''
    Write-Host '  INTERNAL  -- people inside the organisation' -ForegroundColor White
    $internal = ConvertFrom-HtmlToText ([string]$Config.InternalMessage)
    foreach ($line in (($internal, '(no message stored)')[[int][string]::IsNullOrWhiteSpace($internal)] -split "\r\n|\n|\r")) {
        Write-Host "    | $line" -ForegroundColor Gray
    }
    Write-Host ''
    Write-Host '  EXTERNAL  -- everyone else' -ForegroundColor White
    $external = ConvertFrom-HtmlToText ([string]$Config.ExternalMessage)
    foreach ($line in (($external, '(no message stored)')[[int][string]::IsNullOrWhiteSpace($external)] -split "\r\n|\n|\r")) {
        Write-Host "    | $line" -ForegroundColor Gray
    }
}

function Send-OofTestEmail {
    <# Sends an HTML message to trigger the reply. Graph if available, otherwise SMTP. #>
    param(
        [Parameter(Mandatory)][string] $Recipient,
        [Parameter(Mandatory)][AllowEmptyString()][string] $StoredHtml,
        [string] $ConfiguredState = 'unknown'
    )

    $subject = "Test Out of Office Trigger - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    $body = @"
$script:HtmlHead<div style="font-family:'Segoe UI', Arial, Helvetica, sans-serif; font-size:11pt; color:#201f1e;">
  <p>This is an <b>automated HTML test message</b> sent to trigger the Out of Office reply for
     <a href="mailto:$Recipient">$Recipient</a>.</p>
  <p style="color:#605e5c; font-size:9pt;">
     Sent $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz') by $From.<br />
     Configured state: $ConfiguredState.
  </p>
  <hr style="border:0; border-top:1px solid #edebe9;" />
  <p style="font-size:9pt; color:#605e5c;">The reply you receive should match the message below:</p>
  <blockquote style="margin:0; padding:8px 12px; border-left:3px solid #0078d4; background:#f3f2f1;">
$StoredHtml
  </blockquote>
</div>
</body>
</html>
"@

    # -From IS NOW REQUIRED, because its default was removed.
    #
    # It used to carry one named person's address. Sending as somebody else by
    # accident is not a failure anyone notices in testing -- the mail goes, it
    # just goes from the wrong person -- so this refuses rather than guesses.
    # The Graph path below sends AS -From; the SMTP path authenticates as it.
    if (-not "$From".Trim()) {
        Write-Log 'Test email NOT sent: -From is required. Pass the address the test should come from.' 'ERROR'
        return
    }

    if (Get-Module -ListAvailable -Name Microsoft.Graph.Users.Actions) {
        try {
            Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
            if (-not (Get-MgContext)) { Connect-MgGraph -Scopes 'Mail.Send' -NoWelcome }
            Send-MgUserMail -UserId $From -BodyParameter @{
                Message = @{
                    Subject      = $subject
                    Body         = @{ ContentType = 'HTML'; Content = $body }
                    ToRecipients = @(@{ EmailAddress = @{ Address = $Recipient } })
                }
                SaveToSentItems = $true
            }
            Write-Log "Test email sent via Microsoft Graph from $From to $Recipient." 'OK'
            return
        }
        catch { Write-Log "Graph send failed ($($_.Exception.Message)); falling back to SMTP." 'WARN' }
    }

    $credential = Get-Credential -UserName $From -Message "Credentials for SMTP host $SmtpServer"
    if (-not $credential) { Write-Log 'Test email cancelled (no credentials).' 'WARN'; return }

    Send-MailMessage -From $From -To $Recipient -Subject $subject -Body $body -BodyAsHtml `
        -Encoding ([System.Text.Encoding]::UTF8) -SmtpServer $SmtpServer -Port $SmtpPort `
        -UseSsl -Credential $credential -WarningAction SilentlyContinue
    Write-Log "Test email sent via ${SmtpServer}:${SmtpPort} from $From to $Recipient." 'OK'
}

function Get-MessageText {
    <# Resolves one message from -Message, -MessagePath, or the prompt. #>
    param(
        [string] $Title,
        [string] $Literal,
        [string] $Path,
        [string] $Seed
    )

    if ($Path) {
        if (-not (Test-Path -LiteralPath $Path)) { throw "Message file not found: $Path" }
        return (Get-Content -LiteralPath $Path -Raw)
    }
    if ($PSBoundParameters.ContainsKey('Literal') -and $Literal) { return $Literal }
    if ($NonInteractive) { return '' }
    return (Read-MultiLineText -Title $Title -Seed $Seed)
}

# endregion ----------------------------------------------------------------- #

# --------------------------------------------------------------------------- #
# region  Window                                                               #
# --------------------------------------------------------------------------- #
#
# THIS WAS A COMMAND-LINE SCRIPT AND IS NOW A WINDOW. Everything above this
# line is unchanged: the HTML conversion, the mailbox resolution, the read-back
# and the test mail are the same functions they were. Only the front end moved,
# because a console prompt inside a toolkit of windows is a tool nobody can
# hand to anybody else.
#
# THE ADDRESS BOXES COMPLETE A BARE NAME. Get-TargetSmtp is the toolkit's own
# rule for that -- anything without an @ gets the tenant's mail domain appended
# -- so "jsmith" and "jsmith@contoso.com" are the same instruction. Used
# for the mailbox and for the test recipient, so the two behave alike.
#
# THE TEST EMAIL IS A BUTTON, NOT A QUESTION. It used to be the last prompt in
# the run, which meant answering it before knowing whether the apply had
# worked. It is optional, separate, and switched on only once a configuration
# has actually been applied and read back.
#
# THE MESSAGES ARE EDITED AS RICH TEXT: font, size, bold, italic, underline,
# strikethrough, colour, highlight, alignment, bullets and links. ONE toolbar
# serves both messages and acts on whichever tab is in front -- two identical
# toolbars would double the controls and still leave the operator guessing
# which was live.
#
# WHAT EXCHANGE STORES IS HTML, so the formatting is read off the control and
# written out as HTML on apply. See ConvertTo-HtmlFromRichText.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:Resolved     = $null   # what Resolve-Mailbox returned for the current box
$script:AppliedCfg   = $null   # the read-back after a successful apply
$script:ActiveEditor = $null   # which message box the toolbar acts on

# The wording that goes in when Insert template is pressed. Angle brackets
# rather than prose, so an unfilled placeholder is obvious in the sent reply
# instead of reading as a sentence somebody meant.
$script:OofTemplate = @'
I am currently out of the office and will not have access to email until <return date>.
For immediate assistance, please contact <name> at <email> or call <phone>.

I will respond to your message as soon as possible after my return.

Kind regards,
<your name>
'@

$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Out of Office (Automatic Replies) -- Exchange Online'
$form.ClientSize      = New-Object System.Drawing.Size(994, 782)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$lblHead          = New-Object System.Windows.Forms.Label
$lblHead.Text     = 'Out of Office Manager'
$lblHead.Location = New-Object System.Drawing.Point(14, 10)
$lblHead.Size     = New-Object System.Drawing.Size(600, 22)
$lblHead.Font     = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($lblHead)

# --------------------------------------------------------------------------- #
# Mailbox and state
# --------------------------------------------------------------------------- #
$grpTop          = New-Object System.Windows.Forms.GroupBox
$grpTop.Text     = 'Mailbox and state'
$grpTop.Location = New-Object System.Drawing.Point(14, 38)
$grpTop.Size     = New-Object System.Drawing.Size(966, 152)
[void]$form.Controls.Add($grpTop)

$lblBox          = New-Object System.Windows.Forms.Label
$lblBox.Text     = 'Mailbox (email)'
$lblBox.Location = New-Object System.Drawing.Point(16, 26)
$lblBox.Size     = New-Object System.Drawing.Size(112, 20)
[void]$grpTop.Controls.Add($lblBox)

$txtMailbox          = New-Object System.Windows.Forms.TextBox
$txtMailbox.Location = New-Object System.Drawing.Point(132, 23)
$txtMailbox.Size     = New-Object System.Drawing.Size(300, 22)
$txtMailbox.Text     = "$Mailbox"
[void]$grpTop.Controls.Add($txtMailbox)

# CAPTIONED AS THE OPERATOR ASKED, BUT IT VALIDATES RATHER THAN CONNECTS.
# B2: a sign-in prompt raised from a modal form draws and then ignores every
# click, so the console owns connecting. If there is no session this says so
# and names the button that makes one.
$btnValidate           = New-Object System.Windows.Forms.Button
$btnValidate.Text      = 'Connect and validate'
$btnValidate.Location  = New-Object System.Drawing.Point(444, 22)
$btnValidate.Size      = New-Object System.Drawing.Size(156, 26)
$btnValidate.FlatStyle = 'Flat'
[void]$grpTop.Controls.Add($btnValidate)

$lblValid          = New-Object System.Windows.Forms.Label
$lblValid.Text     = 'Not validated yet.'
$lblValid.Location = New-Object System.Drawing.Point(610, 26)
$lblValid.Size     = New-Object System.Drawing.Size(340, 20)
[void]$grpTop.Controls.Add($lblValid)

$lblState          = New-Object System.Windows.Forms.Label
$lblState.Text     = 'Auto-reply state'
$lblState.Location = New-Object System.Drawing.Point(16, 58)
$lblState.Size     = New-Object System.Drawing.Size(112, 20)
[void]$grpTop.Controls.Add($lblState)

# RADIO BUTTONS, NOT A DROPDOWN. All three states are visible at once, so
# "Disabled" cannot be sitting one scroll away from an operator who thinks
# automatic replies are off.
$rbEnabled          = New-Object System.Windows.Forms.RadioButton
$rbEnabled.Text     = 'Enabled'
$rbEnabled.Location = New-Object System.Drawing.Point(132, 56)
$rbEnabled.Size     = New-Object System.Drawing.Size(84, 22)
$rbEnabled.Checked  = $true
[void]$grpTop.Controls.Add($rbEnabled)

$rbScheduled          = New-Object System.Windows.Forms.RadioButton
$rbScheduled.Text     = 'Scheduled'
$rbScheduled.Location = New-Object System.Drawing.Point(220, 56)
$rbScheduled.Size     = New-Object System.Drawing.Size(94, 22)
[void]$grpTop.Controls.Add($rbScheduled)

$rbDisabled          = New-Object System.Windows.Forms.RadioButton
$rbDisabled.Text     = 'Disabled (turn OOF off)'
$rbDisabled.Location = New-Object System.Drawing.Point(318, 56)
$rbDisabled.Size     = New-Object System.Drawing.Size(174, 22)
[void]$grpTop.Controls.Add($rbDisabled)

$lblStart          = New-Object System.Windows.Forms.Label
$lblStart.Text     = 'Start'
$lblStart.Location = New-Object System.Drawing.Point(500, 58)
$lblStart.Size     = New-Object System.Drawing.Size(40, 20)
[void]$grpTop.Controls.Add($lblStart)

# A PICKER RATHER THAN A TEXT BOX. The console version parsed typed dates and
# had to refuse the ones it could not read; a picker cannot produce an
# unparseable value at all, which deletes a whole class of error message.
$dtpStart              = New-Object System.Windows.Forms.DateTimePicker
$dtpStart.Location     = New-Object System.Drawing.Point(542, 55)
$dtpStart.Size         = New-Object System.Drawing.Size(140, 22)
$dtpStart.Format       = 'Custom'
$dtpStart.CustomFormat = 'yyyy-MM-dd HH:mm'
$dtpStart.ShowUpDown   = $true
[void]$grpTop.Controls.Add($dtpStart)

$lblEnd          = New-Object System.Windows.Forms.Label
$lblEnd.Text     = 'End'
$lblEnd.Location = New-Object System.Drawing.Point(692, 58)
$lblEnd.Size     = New-Object System.Drawing.Size(32, 20)
[void]$grpTop.Controls.Add($lblEnd)

$dtpEnd              = New-Object System.Windows.Forms.DateTimePicker
$dtpEnd.Location     = New-Object System.Drawing.Point(726, 55)
$dtpEnd.Size         = New-Object System.Drawing.Size(140, 22)
$dtpEnd.Format       = 'Custom'
$dtpEnd.CustomFormat = 'yyyy-MM-dd HH:mm'
$dtpEnd.ShowUpDown   = $true
$dtpEnd.Value        = (Get-Date).AddDays(7)
[void]$grpTop.Controls.Add($dtpEnd)

$lblAud          = New-Object System.Windows.Forms.Label
$lblAud.Text     = 'External audience'
$lblAud.Location = New-Object System.Drawing.Point(16, 90)
$lblAud.Size     = New-Object System.Drawing.Size(112, 20)
[void]$grpTop.Controls.Add($lblAud)

$cboAud               = New-Object System.Windows.Forms.ComboBox
$cboAud.Location      = New-Object System.Drawing.Point(132, 87)
$cboAud.Size          = New-Object System.Drawing.Size(120, 24)
$cboAud.DropDownStyle = 'DropDownList'
[void]$cboAud.Items.AddRange(@('All', 'Known', 'None'))
$cboAud.SelectedIndex = 0
[void]$grpTop.Controls.Add($cboAud)

$chkSame          = New-Object System.Windows.Forms.CheckBox
$chkSame.Text     = 'Use the internal message for external senders too'
$chkSame.Location = New-Object System.Drawing.Point(266, 88)
$chkSame.Size     = New-Object System.Drawing.Size(340, 22)
$chkSame.Checked  = $true
[void]$grpTop.Controls.Add($chkSame)

$lblAudHint          = New-Object System.Windows.Forms.Label
$lblAudHint.Text     = 'None means nobody outside gets a reply, whatever the external message says.'
$lblAudHint.Location = New-Object System.Drawing.Point(614, 90)
$lblAudHint.Size     = New-Object System.Drawing.Size(340, 20)
[void]$grpTop.Controls.Add($lblAudHint)

$lblTestTo          = New-Object System.Windows.Forms.Label
$lblTestTo.Text     = 'Test email to'
$lblTestTo.Location = New-Object System.Drawing.Point(16, 122)
$lblTestTo.Size     = New-Object System.Drawing.Size(112, 20)
[void]$grpTop.Controls.Add($lblTestTo)

$txtTestTo          = New-Object System.Windows.Forms.TextBox
$txtTestTo.Location = New-Object System.Drawing.Point(132, 119)
$txtTestTo.Size     = New-Object System.Drawing.Size(300, 22)
[void]$grpTop.Controls.Add($txtTestTo)

$lblTestHint          = New-Object System.Windows.Forms.Label
$lblTestHint.Text     = 'Optional, and only after applying. NOT pre-filled -- type the address. A bare name gets the tenant domain.'
$lblTestHint.Location = New-Object System.Drawing.Point(444, 122)
$lblTestHint.Size     = New-Object System.Drawing.Size(510, 20)
[void]$grpTop.Controls.Add($lblTestHint)

# --------------------------------------------------------------------------- #
# The formatting toolbar -- above the tabs, so one copy serves both messages
# --------------------------------------------------------------------------- #
$cboFont               = New-Object System.Windows.Forms.ComboBox
$cboFont.Location      = New-Object System.Drawing.Point(14, 198)
$cboFont.Size          = New-Object System.Drawing.Size(140, 24)
$cboFont.DropDownStyle = 'DropDownList'
# A SHORT LIST OF FONTS THAT SURVIVE THE JOURNEY. Enumerating every font on
# this machine would offer faces the recipient does not have, and an automatic
# reply is read on somebody else's computer.
[void]$cboFont.Items.AddRange(@('Segoe UI', 'Calibri', 'Arial', 'Verdana', 'Tahoma', 'Times New Roman', 'Georgia', 'Courier New'))
$cboFont.SelectedIndex = 0
[void]$form.Controls.Add($cboFont)

$cboSize               = New-Object System.Windows.Forms.ComboBox
$cboSize.Location      = New-Object System.Drawing.Point(160, 198)
$cboSize.Size          = New-Object System.Drawing.Size(60, 24)
$cboSize.DropDownStyle = 'DropDownList'
[void]$cboSize.Items.AddRange(@('8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28'))
$cboSize.SelectedItem  = '11'
[void]$form.Controls.Add($cboSize)

$btnBold           = New-Object System.Windows.Forms.Button
$btnBold.Text      = 'B'
$btnBold.Location  = New-Object System.Drawing.Point(226, 197)
$btnBold.Size      = New-Object System.Drawing.Size(32, 26)
$btnBold.FlatStyle = 'Flat'
$btnBold.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($btnBold)

$btnItalic           = New-Object System.Windows.Forms.Button
$btnItalic.Text      = 'I'
$btnItalic.Location  = New-Object System.Drawing.Point(262, 197)
$btnItalic.Size      = New-Object System.Drawing.Size(32, 26)
$btnItalic.FlatStyle = 'Flat'
$btnItalic.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Italic)
[void]$form.Controls.Add($btnItalic)

$btnUnder           = New-Object System.Windows.Forms.Button
$btnUnder.Text      = 'U'
$btnUnder.Location  = New-Object System.Drawing.Point(298, 197)
$btnUnder.Size      = New-Object System.Drawing.Size(32, 26)
$btnUnder.FlatStyle = 'Flat'
$btnUnder.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Underline)
[void]$form.Controls.Add($btnUnder)

$btnStrike           = New-Object System.Windows.Forms.Button
$btnStrike.Text      = 'S'
$btnStrike.Location  = New-Object System.Drawing.Point(334, 197)
$btnStrike.Size      = New-Object System.Drawing.Size(32, 26)
$btnStrike.FlatStyle = 'Flat'
$btnStrike.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Strikeout)
[void]$form.Controls.Add($btnStrike)

$btnColour           = New-Object System.Windows.Forms.Button
$btnColour.Text      = 'Text colour'
$btnColour.Location  = New-Object System.Drawing.Point(374, 197)
$btnColour.Size      = New-Object System.Drawing.Size(100, 26)
$btnColour.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnColour)

$btnHighlight           = New-Object System.Windows.Forms.Button
$btnHighlight.Text      = 'Highlight'
$btnHighlight.Location  = New-Object System.Drawing.Point(478, 197)
$btnHighlight.Size      = New-Object System.Drawing.Size(90, 26)
$btnHighlight.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnHighlight)

$btnPlain           = New-Object System.Windows.Forms.Button
$btnPlain.Text      = 'Clear format'
$btnPlain.Location  = New-Object System.Drawing.Point(572, 197)
$btnPlain.Size      = New-Object System.Drawing.Size(104, 26)
$btnPlain.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnPlain)

$btnBullet           = New-Object System.Windows.Forms.Button
$btnBullet.Text      = 'Bullets'
$btnBullet.Location  = New-Object System.Drawing.Point(14, 229)
$btnBullet.Size      = New-Object System.Drawing.Size(80, 26)
$btnBullet.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnBullet)

$btnLeft           = New-Object System.Windows.Forms.Button
$btnLeft.Text      = 'Left'
$btnLeft.Location  = New-Object System.Drawing.Point(98, 229)
$btnLeft.Size      = New-Object System.Drawing.Size(66, 26)
$btnLeft.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnLeft)

$btnCentre           = New-Object System.Windows.Forms.Button
$btnCentre.Text      = 'Center'
$btnCentre.Location  = New-Object System.Drawing.Point(168, 229)
$btnCentre.Size      = New-Object System.Drawing.Size(66, 26)
$btnCentre.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnCentre)

$btnRight           = New-Object System.Windows.Forms.Button
$btnRight.Text      = 'Right'
$btnRight.Location  = New-Object System.Drawing.Point(238, 229)
$btnRight.Size      = New-Object System.Drawing.Size(66, 26)
$btnRight.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnRight)

$btnLink           = New-Object System.Windows.Forms.Button
$btnLink.Text      = 'Hyperlink...'
$btnLink.Location  = New-Object System.Drawing.Point(308, 229)
$btnLink.Size      = New-Object System.Drawing.Size(100, 26)
$btnLink.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnLink)

$btnTemplate           = New-Object System.Windows.Forms.Button
$btnTemplate.Text      = 'Insert template'
$btnTemplate.Location  = New-Object System.Drawing.Point(412, 229)
$btnTemplate.Size      = New-Object System.Drawing.Size(124, 26)
$btnTemplate.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnTemplate)

$lblToolHint          = New-Object System.Windows.Forms.Label
$lblToolHint.Text     = 'Formatting is converted to HTML automatically. Acts on the message tab in front.'
$lblToolHint.Location = New-Object System.Drawing.Point(544, 233)
$lblToolHint.Size     = New-Object System.Drawing.Size(436, 20)
[void]$form.Controls.Add($lblToolHint)

# --------------------------------------------------------------------------- #
# The two messages and the preview
# --------------------------------------------------------------------------- #
# PLACED, NOT DOCKED. A docked fill would leave Add-ToolHelpButton no gap in
# the top row -- the '?' would be widened into a gutter and then covered by the
# tab strip. That reason is unchanged by the removal of the scaler on
# 2026-09-11; it was never the scaler that required this.
$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(14, 262)
$tabs.Size     = New-Object System.Drawing.Size(966, 296)
[void]$form.Controls.Add($tabs)

$tabInt          = New-Object System.Windows.Forms.TabPage
$tabInt.Text     = 'Internal message'
$tabExt          = New-Object System.Windows.Forms.TabPage
$tabExt.Text     = 'External message'
$tabPrev         = New-Object System.Windows.Forms.TabPage
$tabPrev.Text    = 'HTML preview'
$tabSrc          = New-Object System.Windows.Forms.TabPage
$tabSrc.Text     = 'HTML source'
foreach ($t in @($tabInt, $tabExt, $tabPrev, $tabSrc)) { [void]$tabs.TabPages.Add($t) }

$rtbInternal               = New-Object System.Windows.Forms.RichTextBox
$rtbInternal.Location      = New-Object System.Drawing.Point(8, 8)
$rtbInternal.Size          = New-Object System.Drawing.Size(942, 250)
$rtbInternal.ScrollBars    = 'Vertical'
$rtbInternal.Font          = New-Object System.Drawing.Font('Segoe UI', 11)
$rtbInternal.HideSelection = $false
[void]$tabInt.Controls.Add($rtbInternal)

$rtbExternal               = New-Object System.Windows.Forms.RichTextBox
$rtbExternal.Location      = New-Object System.Drawing.Point(8, 8)
$rtbExternal.Size          = New-Object System.Drawing.Size(942, 250)
$rtbExternal.ScrollBars    = 'Vertical'
$rtbExternal.Font          = New-Object System.Drawing.Font('Segoe UI', 11)
$rtbExternal.HideSelection = $false
[void]$tabExt.Controls.Add($rtbExternal)

# TWO PREVIEWS, BECAUSE THEY ANSWER DIFFERENT QUESTIONS.
#
# The first version of this showed only the markup, on the argument that a
# WebBrowser renders with whatever Internet Explorer engine this machine has
# while the reply is read in Outlook, OWA and a hundred phones -- so a preview
# that looks right could still send wrong.
#
# That reasoning was sound and the conclusion was wrong. "Is my message going
# to look right" is the question an operator actually has, and answering it
# with a wall of tags does not answer it at all. The rendering is approximate;
# approximate is enormously more use than nothing.
#
# So: HTML preview renders, HTML source shows the markup. The markup still
# matters -- it is where you see that a font-size span has crept onto every run
# -- it is just not the first thing anybody wants.
$wbPreview                                = New-Object System.Windows.Forms.WebBrowser
$wbPreview.Location                       = New-Object System.Drawing.Point(8, 8)
$wbPreview.Size                           = New-Object System.Drawing.Size(942, 250)
$wbPreview.IsWebBrowserContextMenuEnabled = $false
$wbPreview.WebBrowserShortcutsEnabled     = $false
$wbPreview.AllowWebBrowserDrop            = $false
# An unhandled script error in a preview must not raise a dialog over a tool
# that is about to write to a mailbox.
$wbPreview.ScriptErrorsSuppressed         = $true
[void]$tabPrev.Controls.Add($wbPreview)

$txtPreview            = New-Object System.Windows.Forms.TextBox
$txtPreview.Location   = New-Object System.Drawing.Point(8, 8)
$txtPreview.Size       = New-Object System.Drawing.Size(942, 250)
$txtPreview.Multiline  = $true
$txtPreview.ReadOnly   = $true
$txtPreview.ScrollBars = 'Both'
$txtPreview.WordWrap   = $false
$txtPreview.Font       = New-Object System.Drawing.Font('Consolas', 9)
[void]$tabSrc.Controls.Add($txtPreview)

# --------------------------------------------------------------------------- #
# The log
# --------------------------------------------------------------------------- #
$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(14, 566)
$txtOut.Size       = New-Object System.Drawing.Size(966, 158)
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.WordWrap   = $true
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
[void]$form.Controls.Add($txtOut)

# --------------------------------------------------------------------------- #
# The buttons
# --------------------------------------------------------------------------- #
$btnPreview           = New-Object System.Windows.Forms.Button
$btnPreview.Text      = 'Refresh HTML preview'
$btnPreview.Location  = New-Object System.Drawing.Point(14, 734)
$btnPreview.Size      = New-Object System.Drawing.Size(170, 30)
$btnPreview.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnPreview)

$btnApply           = New-Object System.Windows.Forms.Button
$btnApply.Text      = 'Apply Out of Office'
$btnApply.Location  = New-Object System.Drawing.Point(190, 734)
$btnApply.Size      = New-Object System.Drawing.Size(170, 30)
$btnApply.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnApply)

$btnShowCfg           = New-Object System.Windows.Forms.Button
$btnShowCfg.Text      = 'Show current config'
$btnShowCfg.Location  = New-Object System.Drawing.Point(366, 734)
$btnShowCfg.Size      = New-Object System.Drawing.Size(160, 30)
$btnShowCfg.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnShowCfg)

$btnTest           = New-Object System.Windows.Forms.Button
$btnTest.Text      = 'Send test email'
$btnTest.Location  = New-Object System.Drawing.Point(532, 734)
$btnTest.Size      = New-Object System.Drawing.Size(150, 30)
$btnTest.FlatStyle = 'Flat'
$btnTest.Enabled   = $false
[void]$form.Controls.Add($btnTest)

$btnSaveHtml           = New-Object System.Windows.Forms.Button
$btnSaveHtml.Text      = 'Save HTML...'
$btnSaveHtml.Location  = New-Object System.Drawing.Point(688, 734)
$btnSaveHtml.Size      = New-Object System.Drawing.Size(130, 30)
$btnSaveHtml.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnSaveHtml)

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(824, 734)
$btnClose.Size      = New-Object System.Drawing.Size(100, 30)
$btnClose.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnClose)

# --------------------------------------------------------------------------- #
# region  Rich text to HTML                                                    #
# --------------------------------------------------------------------------- #

function ConvertTo-HtmlFromRichText {
    <#
        What the operator sees in the box, as the HTML Exchange will store.

        READ OFF THE CONTROL, NOT OUT OF THE RTF. A RichTextBox holds RTF, and
        parsing RTF properly is a project in itself -- but the control will
        answer a much easier question: for any position, what font, colour and
        highlight is the character there? So this walks the text once, asks
        that, and groups neighbouring characters that agree into a run.
        Whatever it emits therefore matches what is on the screen by
        construction, which is the only definition of correct that matters.

        A DELIBERATELY SMALL SUBSET: family, size, bold, italic, underline,
        strikethrough, colour, highlight, alignment, paragraphs, bullets and
        links. An automatic reply is read in a hundred different clients and the
        plainer the HTML, the more of them render it the same way.

        The wrapper is the same one ConvertTo-HtmlFromMarkup produces, so the
        stored HTML looks the same whether it came from here or from the console
        path -- and ConvertTo-OofBody passes it straight through because it
        already starts with a tag.
    #>
    param(
        [Parameter(Mandatory)]$Rtb,
        [string] $DefaultFamily = 'Segoe UI',
        [double] $DefaultSize   = 11
    )

    $text = "$($Rtb.Text)"
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }

    # Selecting moves the caret, so put it back afterwards. Nothing in the
    # window should shift because a value was read.
    $keepStart = $Rtb.SelectionStart
    $keepLen   = $Rtb.SelectionLength

    $out      = New-Object System.Collections.Generic.List[string]
    $inList   = $false
    $emptyRun = 0

    $pos = 0
    foreach ($rawLine in ($text -split "`n")) {
        $line = $rawLine.TrimEnd([char]13)
        $len  = $line.Length

        if ([string]::IsNullOrWhiteSpace($line)) {
            if ($inList) { $out.Add('</ul>'); $inList = $false }
            $emptyRun++
            $pos += $rawLine.Length + 1
            continue
        }

        if ($emptyRun -gt 0 -and $out.Count -gt 0) { $out.Add('<p style="margin:0;">&nbsp;</p>') }
        $emptyRun = 0

        # Bullets and alignment are per paragraph, so they are read once for
        # the line rather than per character.
        $Rtb.Select($pos, [Math]::Max(1, $len))
        $bullet = [bool]$Rtb.SelectionBullet
        $align  = "$($Rtb.SelectionAlignment)"

        $frag    = New-Object System.Collections.Generic.List[string]
        $runText = New-Object System.Text.StringBuilder
        $curKey  = ''
        $curFont = $null
        $curCol  = [System.Drawing.Color]::Empty
        $curBack = [System.Drawing.Color]::Empty

        # LINKS ARE FOUND ON THE WHOLE LINE, BEFORE IT IS CUT INTO RUNS.
        #
         # Measured 2026-09-03: converting them per run produced
        #   Call <b>[the desk](<a href="https://x.contoso.">...</a></b>com/d)
        # because bolding part of the markup split [text](url) across two runs,
        # the pattern stopped matching, and the bare-URL fallback then linkified
        # the fragment "https://x.contoso." -- a broken link, produced
        # silently, in a message that goes to everyone who writes to the mailbox.
        #
        # So each link is located here and emitted as ONE run whatever
        # formatting sits inside it, taking its appearance from its first
        # character. A link cannot be half bold, and pretending otherwise is
        # what broke it.
        $linkAt = @{}
        foreach ($lm in [regex]::Matches($line, '\[(?<t>[^\]]+)\]\((?<u>https?://[^\s)]+)\)|(?<u2>https?://[^\s<>"]+)')) {
            $linkAt[$lm.Index] = $lm
        }

        $i = 0
        while ($i -lt $len) {
            $lm = $linkAt[$i]
            if ($lm) {
                # Flush whatever plain text was accumulating, then the anchor.
                if ($runText.Length -gt 0) {
                    $frag.Add((Format-HtmlRun -Text $runText.ToString() -Font $curFont -Colour $curCol -Highlight $curBack `
                        -DefaultFamily $DefaultFamily -DefaultSize $DefaultSize))
                    [void]$runText.Clear()
                }
                $Rtb.Select($pos + $i, 1)
                $lf2 = $Rtb.SelectionFont
                if (-not $lf2) { $lf2 = $Rtb.Font }
                $shown = if ($lm.Groups['t'].Success) { $lm.Groups['t'].Value } else { $lm.Groups['u2'].Value }
                $href  = if ($lm.Groups['u'].Success)  { $lm.Groups['u'].Value }  else { $lm.Groups['u2'].Value }
                $frag.Add((Format-HtmlRun -Text $shown -Font $lf2 -Colour $Rtb.SelectionColor -Highlight $Rtb.SelectionBackColor `
                    -LinkHref $href -DefaultFamily $DefaultFamily -DefaultSize $DefaultSize))
                $curKey = ''   # force a fresh run after the anchor
                $i += $lm.Length
                continue
            }

            $Rtb.Select($pos + $i, 1)
            $f = $Rtb.SelectionFont
            if (-not $f) { $f = $Rtb.Font }
            $c = $Rtb.SelectionColor
            $b = $Rtb.SelectionBackColor
            $key = '{0}|{1}|{2}|{3}|{4}' -f $f.Name, [math]::Round($f.Size, 1), [int]$f.Style, $c.ToArgb(), $b.ToArgb()

            if ($key -ne $curKey) {
                if ($runText.Length -gt 0) {
                    $frag.Add((Format-HtmlRun -Text $runText.ToString() -Font $curFont -Colour $curCol -Highlight $curBack `
                        -DefaultFamily $DefaultFamily -DefaultSize $DefaultSize))
                    [void]$runText.Clear()
                }
                $curKey = $key; $curFont = $f; $curCol = $c; $curBack = $b
            }
            [void]$runText.Append($line[$i])
            $i++
        }
        if ($runText.Length -gt 0) {
            $frag.Add((Format-HtmlRun -Text $runText.ToString() -Font $curFont -Colour $curCol -Highlight $curBack `
                -DefaultFamily $DefaultFamily -DefaultSize $DefaultSize))
        }

        $content = ($frag -join '')

        $alignCss = ''
        if ($align -eq 'Center') { $alignCss = ' text-align:center;' }
        elseif ($align -eq 'Right') { $alignCss = ' text-align:right;' }

        if ($bullet) {
            if (-not $inList) { $out.Add('<ul style="margin:0 0 0 18px; padding:0;">'); $inList = $true }
            $out.Add("  <li style=""margin:0;$alignCss"">$content</li>")
        }
        else {
            if ($inList) { $out.Add('</ul>'); $inList = $false }
            $out.Add("<p style=""margin:0 0 6px 0;$alignCss"">$content</p>")
        }

        $pos += $rawLine.Length + 1
    }
    if ($inList) { $out.Add('</ul>') }

    $Rtb.Select($keepStart, $keepLen)

    $wrapperStyle = "margin:0; padding:0; font-family:'$DefaultFamily', Arial, Helvetica, sans-serif; font-size:$([math]::Round($DefaultSize,1))pt; color:#201f1e; line-height:1.5;"
    return @"
$script:HtmlHead<div style="$wrapperStyle">
$($out -join "`n")
</div>
</body>
</html>
"@
}

function Format-HtmlRun {
    <#
        One stretch of identically formatted characters, as HTML.

        ONLY WHAT DIFFERS FROM THE DEFAULT IS WRITTEN. A span repeating the
        wrapper's own font on every run triples the size of the message and
        gives Outlook more to disagree with.

        b/i/u/s tags rather than CSS for the four styles: they are the most
        universally honoured things in mail HTML, and an automatic reply is
        rendered by clients that were old when this was written.

        LINKS ARE MARKUP, NOT A CONTROL PROPERTY. A RichTextBox has nowhere to
        keep a link target per run -- DetectUrls only colours what already
        looks like a URL -- so the Hyperlink button writes [text](url) into the
        text and this turns it into an anchor. That also means a link survives
        being retyped, copied or pasted, which a hidden property would not.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Text,
        $Font,
        $Colour,
        $Highlight,
        [string] $LinkHref,
        [string] $DefaultFamily,
        [double] $DefaultSize
    )

    if ($Text -eq '') { return '' }

    $html = [System.Net.WebUtility]::HtmlEncode($Text)
    # Runs of spaces collapse in HTML; the operator meant them.
    $html = $html -replace '  ', '&nbsp;&nbsp;'

    # THE HREF IS GIVEN, NOT GUESSED. This used to run the link patterns over
    # the run''s own text, which fails the moment formatting splits a link --
    # see the note in ConvertTo-HtmlFromRichText. The caller finds the link and
    # passes the address; this only has to render it.
    if ($LinkHref) {
        $safeHref = [System.Net.WebUtility]::HtmlEncode($LinkHref)
        $html = "<a href=""$safeHref"">$html</a>"
    }

    if ($Font) {
        if ($Font.Bold)      { $html = "<b>$html</b>" }
        if ($Font.Italic)    { $html = "<i>$html</i>" }
        if ($Font.Underline) { $html = "<u>$html</u>" }
        if ($Font.Strikeout) { $html = "<s>$html</s>" }
    }

    $bits = New-Object System.Collections.Generic.List[string]
    if ($Font -and $Font.Name -and $Font.Name -ne $DefaultFamily) {
        $bits.Add("font-family:'$($Font.Name)', Arial, Helvetica, sans-serif")
    }
    # COMPARED WITH A TOLERANCE, NOT FOR EQUALITY. A RichTextBox stores its text
    # as RTF and reports an 11pt font back as 11.2 -- so an exact test said
    # "different" for every character and stamped a redundant font-size span on
    # every run. Measured 2026-09-03; 0.35pt is below the smallest step the size
    # list offers and well above the drift.
    if ($Font -and [Math]::Abs($Font.Size - $DefaultSize) -gt 0.35) {
        $bits.Add("font-size:$([math]::Round($Font.Size,0))pt")
    }
    if ($Colour -and $Colour.ToArgb() -ne ([System.Drawing.Color]::Black).ToArgb() -and $Colour.ToArgb() -ne 0) {
        $bits.Add(('color:#{0:X2}{1:X2}{2:X2}' -f $Colour.R, $Colour.G, $Colour.B))
    }
    # WHITE IS NOT A HIGHLIGHT. An unhighlighted character reports the control's
    # own background, so writing that out would put a white box behind every
    # word -- visible the moment somebody reads the reply in dark mode.
    if ($Highlight -and $Highlight.ToArgb() -ne 0 -and
        $Highlight.ToArgb() -ne ([System.Drawing.Color]::White).ToArgb() -and
        $Highlight.ToArgb() -ne ([System.Drawing.SystemColors]::Window).ToArgb()) {
        $bits.Add(('background-color:#{0:X2}{1:X2}{2:X2}' -f $Highlight.R, $Highlight.G, $Highlight.B))
    }
    if ($bits.Count -gt 0) { $html = "<span style=""$($bits -join '; ')"">$html</span>" }
    return $html
}

# --------------------------------------------------------------------------- #
# region  Window behaviour                                                     #
# --------------------------------------------------------------------------- #

function Add-Log {
    <#
        One timestamped line in the log pane, newest at the bottom.

        APPENDED, NOT REPLACED. A tool that overwrites its own output loses the
        sentence that explained the state the next one is complaining about.
    #>
    param([string] $Text, [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string] $Level = 'INFO')

    $stamp = '[{0:HH:mm:ss}] {1,-5} ' -f (Get-Date), $Level
    $lines = @("$Text" -split "\r\n|\n|\r")
    $block = New-Object System.Collections.Generic.List[string]
    foreach ($l in $lines) {
        if ($block.Count -eq 0) { $block.Add("$stamp$l") }
        else { $block.Add((' ' * $stamp.Length) + $l) }
    }
    $joined = ($block -join [Environment]::NewLine)
    if ("$($txtOut.Text)".Trim()) { $txtOut.AppendText([Environment]::NewLine + $joined) }
    else { $txtOut.Text = $joined }
    $txtOut.SelectionStart = $txtOut.TextLength
    $txtOut.ScrollToCaret()
}

function Get-SinkText {
    <# Whatever the engine functions printed since the sink was last cleared. #>
    if (-not $script:OutSink) { return '' }
    return (@($script:OutSink.ToArray()) -join [Environment]::NewLine)
}

function Clear-Sink {
    if ($script:OutSink) { $script:OutSink.Clear() }
}

function Get-SelectedState {
    if ($rbScheduled.Checked) { return 'Scheduled' }
    if ($rbDisabled.Checked)  { return 'Disabled' }
    return 'Enabled'
}

function Get-ActiveEditor {
    <#
        Which message the toolbar acts on: whichever tab is in front.

        NOT ActiveControl -- clicking a toolbar button moves the focus to the
        button, so by then the answer would always be "the button". The tab is
        what the operator can see, which makes it the honest answer.
    #>
    if ($tabs.SelectedTab -eq $tabExt) { return $rtbExternal }
    return $rtbInternal
}

function Set-RunStyle {
    <#
        Toggle bold, italic, underline or strikethrough over the selection.

        SelectionFont IS NULL WHEN THE SELECTION IS NOT UNIFORM -- select a word
        that is half bold and the control declines to name one font. Applying
        the toggle to the whole run would flatten the differences it could not
        describe, so a mixed selection is handled a character at a time and each
        keeps everything except the bit being changed.
    #>
    param([Parameter(Mandatory)][System.Drawing.FontStyle] $Style)

    $rtb = Get-ActiveEditor
    if (-not $rtb -or -not $rtb.Enabled) { return }
    $start = $rtb.SelectionStart
    $len   = $rtb.SelectionLength

    if ($len -eq 0) {
        # No selection: set it for what gets typed next, as a word processor does.
        $f = $rtb.SelectionFont
        if (-not $f) { $f = $rtb.Font }
        $rtb.SelectionFont = New-Object System.Drawing.Font($f, ($f.Style -bxor $Style))
        return
    }

    $uniform = $rtb.SelectionFont
    if ($uniform) {
        $rtb.SelectionFont = New-Object System.Drawing.Font($uniform, ($uniform.Style -bxor $Style))
        $rtb.Select($start, $len)
        return
    }

    # Mixed. Decide once from the first character so the whole selection ends up
    # the same way, which is what pressing B is understood to mean.
    $rtb.Select($start, 1)
    $first = $rtb.SelectionFont
    if (-not $first) { $first = $rtb.Font }
    $turnOn = -not ($first.Style -band $Style)

    for ($i = 0; $i -lt $len; $i++) {
        $rtb.Select($start + $i, 1)
        $f = $rtb.SelectionFont
        if (-not $f) { $f = $rtb.Font }
        $s = if ($turnOn) { $f.Style -bor $Style } else { $f.Style -band (-bnot $Style) }
        $rtb.SelectionFont = New-Object System.Drawing.Font($f, $s)
    }
    $rtb.Select($start, $len)
}

function Set-RunFamilyOrSize {
    param([string] $Family, [double] $Size)

    $rtb = Get-ActiveEditor
    if (-not $rtb -or -not $rtb.Enabled) { return }
    $start = $rtb.SelectionStart
    $len   = $rtb.SelectionLength

    if ($len -eq 0) {
        $f = $rtb.SelectionFont
        if (-not $f) { $f = $rtb.Font }
        $fam = if ($Family) { $Family } else { $f.Name }
        $sz  = if ($Size -gt 0) { $Size } else { $f.Size }
        $rtb.SelectionFont = New-Object System.Drawing.Font($fam, $sz, $f.Style)
        return
    }

    # Per character, so a selection that mixes bold and plain keeps doing so.
    for ($i = 0; $i -lt $len; $i++) {
        $rtb.Select($start + $i, 1)
        $f = $rtb.SelectionFont
        if (-not $f) { $f = $rtb.Font }
        $fam = if ($Family) { $Family } else { $f.Name }
        $sz  = if ($Size -gt 0) { $Size } else { $f.Size }
        $rtb.SelectionFont = New-Object System.Drawing.Font($fam, $sz, $f.Style)
    }
    $rtb.Select($start, $len)
}

function Set-RunAlignment {
    param([Parameter(Mandatory)][System.Windows.Forms.HorizontalAlignment] $Alignment)
    $rtb = Get-ActiveEditor
    if (-not $rtb -or -not $rtb.Enabled) { return }
    $rtb.SelectionAlignment = $Alignment
    $rtb.Focus()
}

function Update-StateEnabled {
    $state = Get-SelectedState
    $sched = ($state -eq 'Scheduled')
    foreach ($c in @($dtpStart, $dtpEnd, $lblStart, $lblEnd)) { $c.Enabled = $sched }

    $off = ($state -eq 'Disabled')
    foreach ($c in @($rtbInternal, $rtbExternal, $chkSame, $cboAud, $cboFont, $cboSize,
                     $btnBold, $btnItalic, $btnUnder, $btnStrike, $btnColour, $btnHighlight,
                     $btnPlain, $btnBullet, $btnLeft, $btnCentre, $btnRight, $btnLink, $btnTemplate)) {
        $c.Enabled = -not $off
    }
    if (-not $off) { $rtbExternal.Enabled = -not $chkSame.Checked }
}

function Test-ExoReady {
    <#
        Connected? Asked here rather than connected here.

        B2: a sign-in prompt raised from a modal form draws and then ignores
        every click, which is why the console owns connecting and this window
        only checks. Connect-ExchangeIfNeeded is still in this file and still
        works when the script is run on its own from a shell.
    #>
    try {
        $s = Test-ExoConnectionState
        if ($s.Connected) { return $true }
        $lblValid.Text = 'Not connected to Exchange Online.'
        Add-Log ("Exchange Online is not connected in this window. $($s.Detail)" + [Environment]::NewLine +
            'Press Connect-ExchangeOnline on the console, then try again. This window cannot sign in itself -- a prompt raised from a modal form ignores every click.') 'ERROR'
        return $false
    } catch {
        Add-Log (Format-ErrorDetail -ErrorRecord $_ -Context 'Checking Exchange Online') 'ERROR'
        return $false
    }
}

function Invoke-MailboxLookup {
    $typed = "$($txtMailbox.Text)".Trim()
    if (-not $typed) { Add-Log 'Type a mailbox address, or just the first part of one.' 'WARN'; return }

    # A bare name becomes an address, by the toolkit's own rule.
    $addr = Get-TargetSmtp -Text $typed
    if ("$addr" -ne $typed) {
        $txtMailbox.Text = "$addr"
        Add-Log "Completed '$typed' to '$addr'."
    }

    if (-not (Test-ExoReady)) { return }

    $script:Resolved   = $null
    $script:AppliedCfg = $null
    $btnTest.Enabled   = $false
    $lblValid.Text     = 'Validating...'

    $old = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Clear-Sink
        $script:Resolved = Resolve-Mailbox -Address "$addr"
        $box     = $script:Resolved.Mailbox
        $cfg     = $script:Resolved.Config
        $primary = [string]$box.PrimarySmtpAddress

        $lblValid.Text  = "$($box.DisplayName) <$primary>"
        # THE TEST ADDRESS IS NOT PRE-FILLED, DELIBERATELY. It used to be set to
        # the mailbox just resolved, which put a live recipient in the box of a
        # button that sends mail -- one stray click away from mailing somebody
        # who was only being looked at. It is typed, and a bare name still gets
        # the tenant domain.
        Add-Log "Validated: $($box.DisplayName) <$primary> [$($box.RecipientTypeDetails)]" 'OK'

        switch ("$($cfg.AutoReplyState)") {
            'Scheduled' { $rbScheduled.Checked = $true }
            'Disabled'  { $rbDisabled.Checked  = $true }
            default     { $rbEnabled.Checked   = $true }
        }
        $aud = "$($cfg.ExternalAudience)"
        if ($cboAud.Items.Contains($aud)) { $cboAud.SelectedItem = $aud }

        try { if ($cfg.StartTime) { $dtpStart.Value = [datetime]$cfg.StartTime } } catch { }
        try { if ($cfg.EndTime -and ([datetime]$cfg.EndTime) -gt (Get-Date)) { $dtpEnd.Value = [datetime]$cfg.EndTime } } catch { }

        Update-StateEnabled
        Show-CurrentConfig
    } catch {
        $lblValid.Text = 'Not validated.'
        Add-Log (Format-ErrorDetail -ErrorRecord $_ -Context "Looking up '$addr'") 'ERROR'
    } finally {
        $form.Cursor = $old
    }
}

function Show-CurrentConfig {
    <# The engine's own account of what is stored, warnings and all. #>
    if (-not $script:Resolved) { Add-Log 'Look a mailbox up first.' 'WARN'; return }
    Clear-Sink
    Show-OofConfiguration -Config $script:Resolved.Config -Title 'Stored in Exchange'
    Add-Log (Get-SinkText)
}

function Import-StoredText {
    <#
        Puts what is already stored into the edit boxes.

        PLAIN TEXT, AND THE LOG SAYS SO. Reading HTML back into formatted rich
        text means an HTML parser, and a half-working one would show formatting
        that is not what Exchange holds -- worse than showing none, because it
        would be believed. The wording comes back; the bold does not.
    #>
    if (-not $script:Resolved) { Add-Log 'Look a mailbox up first.' 'WARN'; return }
    $cfg = $script:Resolved.Config
    $rtbInternal.Text = ConvertFrom-HtmlToText ([string]$cfg.InternalMessage)
    $rtbExternal.Text = ConvertFrom-HtmlToText ([string]$cfg.ExternalMessage)
    $chkSame.Checked  = $false
    Update-StateEnabled
    Add-Log 'Loaded the stored wording as PLAIN TEXT. Any bold, colour or sizing already in Exchange is not shown here and will be replaced by whatever you apply next. Nothing has been changed yet.' 'WARN'
}

function Get-HtmlBodyInner {
    <#
        Just the body of a generated document.

        The two messages are complete html documents in their own right, and
        nesting <html> inside <html> is not something any renderer promises to
        handle. So the wrapper is stripped and only the content is put into the
        preview document.
    #>
    param([AllowEmptyString()][string] $Html)
    if (-not $Html) { return '<p style="color:#999; font:italic 10pt Segoe UI;">(empty)</p>' }
    $m = [regex]::Match($Html, '(?is)<body[^>]*>(?<b>.*?)</body>')
    if ($m.Success) { return $m.Groups['b'].Value }
    return $Html
}

function Set-PreviewDocument {
    <#
        Put a document in the browser control.

        VIA A FILE, BECAUSE DocumentText DOES NOT WORK HERE.

        Measured 2026-09-03: assigning $wbPreview.DocumentText left the control
        holding a 14-character empty document and an empty body. The property
        only takes effect once the control already has a loaded document, and
        it loads asynchronously -- so an assignment made before that is thrown
        away with no error raised anywhere. THAT was "the HTML preview is not
        working": the markup was correct, the control was wired up correctly,
        and the one line that mattered did nothing at all.

        Navigating to a file has none of that behaviour, needs no COM interop
        with IHTMLDocument, and renders the same bytes a browser would --
        which is the point of a preview. One file per process, overwritten on
        each refresh, removed when the window closes.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Html)

    try {
        if (-not $script:PreviewFile) {
            $script:PreviewFile = Join-Path ([System.IO.Path]::GetTempPath()) "oof-preview-$PID.html"
        }
        [System.IO.File]::WriteAllText($script:PreviewFile, $Html, (New-Object System.Text.UTF8Encoding($false)))
        if (-not $wbPreview.IsHandleCreated) { [void]$wbPreview.CreateControl() }
        $wbPreview.Navigate((New-Object System.Uri $script:PreviewFile))
    } catch {
        Add-Log ("The rendered preview could not be drawn: $($_.Exception.Message)  The HTML source tab still shows exactly what would be stored.") 'WARN'
    }
}
function Update-HtmlPreview {
    <# The markup that would be stored, without storing it. #>
    $state = Get-SelectedState
    if ($state -eq 'Disabled') {
        $txtPreview.Text = '(Disabled -- the stored message text is left untouched, so there is no new HTML.)'
        Set-PreviewDocument -Html '<html><body style="font-family:Segoe UI; color:#666;"><p>Disabled &mdash; the stored message text is left untouched, so there is no new HTML.</p></body></html>'
        Add-Log 'State is Disabled, so no message HTML would be written.'
        return
    }
    $old = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $int = ConvertTo-OofBody -Text (ConvertTo-HtmlFromRichText -Rtb $rtbInternal)
        $ext = if ($chkSame.Checked) { $int } else { ConvertTo-OofBody -Text (ConvertTo-HtmlFromRichText -Rtb $rtbExternal) }

        # --- the markup, for the source tab ---
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('=== INTERNAL ===')
        [void]$sb.AppendLine($int)
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('=== EXTERNAL ===')
        if ($chkSame.Checked) { [void]$sb.AppendLine('(the same as the internal message)') }
        [void]$sb.AppendLine($ext)
        $txtPreview.Text = $sb.ToString()

        # --- the rendering, for the preview tab ---
        # BOTH MESSAGES IN ONE DOCUMENT, labelled. The external one is the
        # easier to get wrong and the harder to notice, so it is shown rather
        # than left to be imagined.
        $doc = New-Object System.Text.StringBuilder
        [void]$doc.Append('<html><head><meta http-equiv="X-UA-Compatible" content="IE=edge"><meta charset="utf-8"></head><body style="margin:0; padding:0; background:#ffffff;">')
        [void]$doc.Append('<div style="font:bold 9pt Segoe UI; background:#dde6f0; color:#17293c; padding:4px 8px;">INTERNAL &mdash; people inside the organisation</div>')
        [void]$doc.Append('<div style="padding:8px;">' + (Get-HtmlBodyInner -Html $int) + '</div>')
        [void]$doc.Append('<div style="font:bold 9pt Segoe UI; background:#dde6f0; color:#17293c; padding:4px 8px;">EXTERNAL &mdash; everyone else')
        if ($chkSame.Checked) { [void]$doc.Append(' (the same as the internal message)') }
        [void]$doc.Append('</div>')
        [void]$doc.Append('<div style="padding:8px;">' + (Get-HtmlBodyInner -Html $ext) + '</div>')
        [void]$doc.Append('</body></html>')
        Set-PreviewDocument -Html $doc.ToString()

        Add-Log "HTML preview refreshed -- internal $($int.Length) characters, external $($ext.Length)."
    } finally {
        $form.Cursor = $old
    }
}

function Invoke-ApplyOof {
    if (-not $script:Resolved) { Add-Log 'Look a mailbox up first.' 'WARN'; return }
    if (-not (Test-ExoReady)) { return }

    $identity = $script:Resolved.Identity
    $primary  = [string]$script:Resolved.Mailbox.PrimarySmtpAddress
    $state    = Get-SelectedState

    $params = @{ Identity = $identity; AutoReplyState = $state }

    if ($state -ne 'Disabled') {
        if ($state -eq 'Scheduled') {
            # The picker cannot produce an unparseable value, but it can still
            # produce a nonsensical window.
            if ($dtpEnd.Value -le $dtpStart.Value) { Add-Log 'The end must be after the start. Nothing was changed.' 'WARN'; return }
            if ($dtpEnd.Value -le (Get-Date))      { Add-Log 'The end is in the past, so the reply would never fire. Nothing was changed.' 'WARN'; return }
            $params.StartTime = $dtpStart.Value
            $params.EndTime   = $dtpEnd.Value
        }

        # The formatting comes off the control here. ConvertTo-OofBody passes
        # the result through untouched because it already begins with a tag.
        $internalHtml = ConvertTo-OofBody -Text (ConvertTo-HtmlFromRichText -Rtb $rtbInternal)
        if ((Get-HtmlTextLength $internalHtml) -eq 0) { Add-Log 'The internal message is empty. Nothing was changed.' 'WARN'; return }

        $externalHtml = if ($chkSame.Checked) { $internalHtml }
                        else { ConvertTo-OofBody -Text (ConvertTo-HtmlFromRichText -Rtb $rtbExternal) }
        if ((Get-HtmlTextLength $externalHtml) -eq 0) { Add-Log 'The external message is empty. Tick "use the internal message" or type one. Nothing was changed.' 'WARN'; return }

        $params.InternalMessage  = $internalHtml
        $params.ExternalMessage  = $externalHtml
        $params.ExternalAudience = "$($cboAud.SelectedItem)"
    }

    $sentence = "Set automatic replies on $primary to $state."
    if ($state -eq 'Scheduled') { $sentence += "$([Environment]::NewLine)Window: $($params.StartTime.ToString('yyyy-MM-dd HH:mm')) -> $($params.EndTime.ToString('yyyy-MM-dd HH:mm'))" }
    if ($state -ne 'Disabled')  { $sentence += "$([Environment]::NewLine)External audience: $($params.ExternalAudience)" }
    if ($state -eq 'Disabled')  { $sentence += "$([Environment]::NewLine)The stored message text is left untouched." }

    $answer = Show-ToolDialog($sentence, 'Apply Out of Office',
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button1)
    if ($answer -ne [System.Windows.Forms.DialogResult]::OK) { Add-Log 'Apply cancelled. Nothing was changed.' 'WARN'; return }

    $old = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnApply.Enabled = $false
    try {
        Set-MailboxAutoReplyConfiguration @params

        # READ BACK. An absent exception is not evidence the change landed.
        $applied = Get-MailboxAutoReplyConfiguration -Identity $identity -ErrorAction Stop
        if (-not $applied) { throw "Could not read the configuration back for '$identity'." }
        $script:AppliedCfg      = $applied
        $script:Resolved.Config = $applied

        Add-Log "APPLIED to $primary -- state is now $($applied.AutoReplyState), read back from Exchange." 'OK'
        Show-CurrentConfig

        # Only now is there something worth testing.
        $btnTest.Enabled = ("$($applied.AutoReplyState)" -ne 'Disabled')
        if ($btnTest.Enabled) { Add-Log 'The optional test email is now available. Type the address to send it to -- it is not filled in for you.' }
        else { Add-Log 'Automatic replies are off, so there is nothing to test.' }

        Write-CompletedLog -Action 'Modify' -Target (Get-LogTarget $primary 'Out of Office') -Result 'Success' `
            -Detail "AutoReplyState=$($applied.AutoReplyState); ExternalAudience=$($applied.ExternalAudience)." -Environment $Environment
    } catch {
        Add-Log (Format-ErrorDetail -ErrorRecord $_ -Context "Applying to $primary") 'ERROR'
        Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $primary 'Out of Office') `
            -Detail "Setting AutoReplyState=$state failed." -ErrorRecord $_ -Environment $Environment
    } finally {
        $btnApply.Enabled = $true
        $form.Cursor = $old
    }
}

function Invoke-TestEmail {
    if (-not $script:AppliedCfg) { Add-Log 'Apply a configuration first -- there is nothing to test yet.' 'WARN'; return }

    $typed = "$($txtTestTo.Text)".Trim()
    if (-not $typed) { Add-Log 'Type an address to send the test to.' 'WARN'; return }

    # Same completion rule as the mailbox box, so the two behave alike.
    $to = Get-TargetSmtp -Text $typed
    if ("$to" -ne $typed) { $txtTestTo.Text = "$to" }

    $old = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnTest.Enabled = $false
    try {
        Clear-Sink
        Send-OofTestEmail -Recipient "$to" -StoredHtml ([string]$script:AppliedCfg.InternalMessage) `
                          -ConfiguredState ([string]$script:AppliedCfg.AutoReplyState)
        Add-Log (Get-SinkText)
        Add-Log "Test message sent to $to. The automatic reply comes back to $From, so look there rather than here." 'OK'
        Write-CompletedLog -Action 'Report' -Target (Get-LogTarget "$to" 'Out of Office') -Result 'Success' `
            -Detail "Sent an Out of Office test message to $to from $From." -Environment $Environment
    } catch {
        Add-Log (Format-ErrorDetail -ErrorRecord $_ -Context "Sending the test message to $to") 'ERROR'
        Add-Log (Get-SinkText)
    } finally {
        $btnTest.Enabled = $true
        $form.Cursor = $old
    }
}

function Save-HtmlToFile {
    <# The generated markup, as a file a browser can open. #>
    if ((Get-SelectedState) -eq 'Disabled') { Add-Log 'State is Disabled, so there is no message HTML to save.' 'WARN'; return }
    $html = ConvertTo-OofBody -Text (ConvertTo-HtmlFromRichText -Rtb (Get-ActiveEditor))
    if (-not $html) { Add-Log 'That message is empty, so there is nothing to save.' 'WARN'; return }

    $which = if ((Get-ActiveEditor) -eq $rtbExternal) { 'external' } else { 'internal' }
    $stem  = 'oof'
    if ($script:Resolved) {
        $stem = (("$($script:Resolved.Mailbox.PrimarySmtpAddress)" -split '@')[0] -replace '[^A-Za-z0-9._-]', '').ToLowerInvariant()
        if (-not $stem) { $stem = 'oof' }
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'HTML file (*.html)|*.html|All files (*.*)|*.*'
    $dlg.FileName = "$stem-oof-$which.html"
    try { $dlg.InitialDirectory = Get-ReportPath } catch { }
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        Set-Content -LiteralPath $dlg.FileName -Value $html -Encoding UTF8 -ErrorAction Stop
        Add-Log "Saved the $which message HTML to $($dlg.FileName)" 'OK'
    } catch {
        Add-Log (Format-ErrorDetail -ErrorRecord $_ -Context "Writing $($dlg.FileName)") 'ERROR'
    }
}

function Add-Hyperlink {
    <#
        Writes [text](url) at the cursor, which Format-HtmlRun turns into an
        anchor. See its docstring for why the markup rather than a property.
    #>
    $rtb = Get-ActiveEditor
    if (-not $rtb -or -not $rtb.Enabled) { return }

    $dlg          = New-Object System.Windows.Forms.Form
    $dlg.Text     = 'Insert hyperlink'
    $dlg.ClientSize = New-Object System.Drawing.Size(520, 150)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.StartPosition   = 'CenterParent'
    $dlg.MinimizeBox = $false
    $dlg.MaximizeBox = $false

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = 'Text to show'
    $l1.Location = New-Object System.Drawing.Point(14, 18)
    $l1.Size = New-Object System.Drawing.Size(100, 20)
    [void]$dlg.Controls.Add($l1)

    $t1 = New-Object System.Windows.Forms.TextBox
    $t1.Location = New-Object System.Drawing.Point(120, 15)
    $t1.Size = New-Object System.Drawing.Size(380, 22)
    $t1.Text = "$($rtb.SelectedText)"
    [void]$dlg.Controls.Add($t1)

    $l2 = New-Object System.Windows.Forms.Label
    $l2.Text = 'Address'
    $l2.Location = New-Object System.Drawing.Point(14, 52)
    $l2.Size = New-Object System.Drawing.Size(100, 20)
    [void]$dlg.Controls.Add($l2)

    $t2 = New-Object System.Windows.Forms.TextBox
    $t2.Location = New-Object System.Drawing.Point(120, 49)
    $t2.Size = New-Object System.Drawing.Size(380, 22)
    $t2.Text = 'https://'
    [void]$dlg.Controls.Add($t2)

    $l3 = New-Object System.Windows.Forms.Label
    $l3.Text = 'Only http and https addresses become links.'
    $l3.Location = New-Object System.Drawing.Point(120, 76)
    $l3.Size = New-Object System.Drawing.Size(380, 20)
    [void]$dlg.Controls.Add($l3)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Insert'
    $ok.Location = New-Object System.Drawing.Point(320, 108)
    $ok.Size = New-Object System.Drawing.Size(84, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    [void]$dlg.Controls.Add($ok)

    $no = New-Object System.Windows.Forms.Button
    $no.Text = 'Cancel'
    $no.Location = New-Object System.Drawing.Point(412, 108)
    $no.Size = New-Object System.Drawing.Size(88, 28)
    $no.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    [void]$dlg.Controls.Add($no)

    $dlg.AcceptButton = $ok
    $dlg.CancelButton = $no
    Set-ConsoleTheme -Form $dlg

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { $dlg.Dispose(); return }
    $text = "$($t1.Text)".Trim()
    $url  = "$($t2.Text)".Trim()
    $dlg.Dispose()

    if ($url -notmatch '^https?://.+') { Add-Log "'$url' is not an http or https address, so it would not become a link. Nothing was inserted." 'WARN'; return }
    if (-not $text) { $text = $url }

    $rtb.SelectedText = "[$text]($url)"
    Add-Log "Inserted a link to $url. It becomes an anchor when the HTML is generated."
}

# --------------------------------------------------------------------------- #
# region  Handlers                                                             #
# --------------------------------------------------------------------------- #
# B1: every handler wrapped, or an escaping exception is caught by .NET and the
# operator gets a CLR dialog with no script name on it.
$btnValidate.Add_Click({ try { Invoke-MailboxLookup } catch { Show-HandlerError -ErrorRecord $_ -Context 'Connect and validate' } })
$btnApply.Add_Click({ try { Invoke-ApplyOof } catch { Show-HandlerError -ErrorRecord $_ -Context 'Apply' } })
$btnShowCfg.Add_Click({ try { Show-CurrentConfig } catch { Show-HandlerError -ErrorRecord $_ -Context 'Show current config' } })
$btnTest.Add_Click({ try { Invoke-TestEmail } catch { Show-HandlerError -ErrorRecord $_ -Context 'Test email' } })
$btnPreview.Add_Click({ try { Update-HtmlPreview; $tabs.SelectedTab = $tabPrev } catch { Show-HandlerError -ErrorRecord $_ -Context 'HTML preview' } })
$btnSaveHtml.Add_Click({ try { Save-HtmlToFile } catch { Show-HandlerError -ErrorRecord $_ -Context 'Save HTML' } })
$btnClose.Add_Click({ try { $form.Close() } catch { Show-HandlerError -ErrorRecord $_ -Context 'Close' } })

foreach ($rb in @($rbEnabled, $rbScheduled, $rbDisabled)) {
    $rb.Add_CheckedChanged({ try { Update-StateEnabled } catch { Show-HandlerError -ErrorRecord $_ -Context 'State' } })
}
$chkSame.Add_CheckedChanged({ try { Update-StateEnabled } catch { Show-HandlerError -ErrorRecord $_ -Context 'Same as internal' } })

$btnBold.Add_Click({ try { Set-RunStyle -Style ([System.Drawing.FontStyle]::Bold); (Get-ActiveEditor).Focus() } catch { Show-HandlerError -ErrorRecord $_ -Context 'Bold' } })
$btnItalic.Add_Click({ try { Set-RunStyle -Style ([System.Drawing.FontStyle]::Italic); (Get-ActiveEditor).Focus() } catch { Show-HandlerError -ErrorRecord $_ -Context 'Italic' } })
$btnUnder.Add_Click({ try { Set-RunStyle -Style ([System.Drawing.FontStyle]::Underline); (Get-ActiveEditor).Focus() } catch { Show-HandlerError -ErrorRecord $_ -Context 'Underline' } })
$btnStrike.Add_Click({ try { Set-RunStyle -Style ([System.Drawing.FontStyle]::Strikeout); (Get-ActiveEditor).Focus() } catch { Show-HandlerError -ErrorRecord $_ -Context 'Strikethrough' } })

$btnLeft.Add_Click({ try { Set-RunAlignment -Alignment ([System.Windows.Forms.HorizontalAlignment]::Left) } catch { Show-HandlerError -ErrorRecord $_ -Context 'Align left' } })
$btnCentre.Add_Click({ try { Set-RunAlignment -Alignment ([System.Windows.Forms.HorizontalAlignment]::Center) } catch { Show-HandlerError -ErrorRecord $_ -Context 'Align centre' } })
$btnRight.Add_Click({ try { Set-RunAlignment -Alignment ([System.Windows.Forms.HorizontalAlignment]::Right) } catch { Show-HandlerError -ErrorRecord $_ -Context 'Align right' } })

$cboFont.Add_SelectedIndexChanged({
    try { Set-RunFamilyOrSize -Family "$($cboFont.SelectedItem)" -Size 0; (Get-ActiveEditor).Focus() }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Font' }
})
$cboSize.Add_SelectedIndexChanged({
    try {
        $sz = 0.0
        if ([double]::TryParse("$($cboSize.SelectedItem)", [ref]$sz) -and $sz -gt 0) { Set-RunFamilyOrSize -Family '' -Size $sz }
        (Get-ActiveEditor).Focus()
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Size' }
})

$btnColour.Add_Click({
    try {
        $rtb = Get-ActiveEditor
        $dlg = New-Object System.Windows.Forms.ColorDialog
        $dlg.AllowFullOpen = $true
        if ($rtb.SelectionColor) { $dlg.Color = $rtb.SelectionColor }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $rtb.SelectionColor = $dlg.Color }
        $rtb.Focus()
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Text colour' }
})

$btnHighlight.Add_Click({
    try {
        $rtb = Get-ActiveEditor
        $dlg = New-Object System.Windows.Forms.ColorDialog
        $dlg.AllowFullOpen = $true
        $dlg.Color = [System.Drawing.Color]::Yellow
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $rtb.SelectionBackColor = $dlg.Color }
        $rtb.Focus()
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Highlight' }
})

$btnBullet.Add_Click({
    try {
        $rtb = Get-ActiveEditor
        $rtb.SelectionBullet = -not $rtb.SelectionBullet
        $rtb.Focus()
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Bullets' }
})

$btnPlain.Add_Click({
    try {
        $rtb = Get-ActiveEditor
        $start = $rtb.SelectionStart
        $len   = $rtb.SelectionLength
        if ($len -eq 0) { Add-Log 'Select the text to clear the formatting from.' 'WARN'; return }
        $rtb.SelectionFont      = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Regular)
        $rtb.SelectionColor     = [System.Drawing.Color]::Black
        $rtb.SelectionBackColor = [System.Drawing.Color]::White
        $rtb.SelectionBullet    = $false
        $rtb.SelectionAlignment = [System.Windows.Forms.HorizontalAlignment]::Left
        $rtb.Select($start, $len)
        $rtb.Focus()
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Clear format' }
})

$btnLink.Add_Click({ try { Add-Hyperlink } catch { Show-HandlerError -ErrorRecord $_ -Context 'Hyperlink' } })

$btnTemplate.Add_Click({
    try {
        $rtb = Get-ActiveEditor
        if (-not $rtb.Enabled) { return }
        if ("$($rtb.Text)".Trim()) {
            $r = Show-ToolDialog(
                'Replace what is in this message with the template?', 'Insert template',
                [System.Windows.Forms.MessageBoxButtons]::OKCancel,
                [System.Windows.Forms.MessageBoxIcon]::Question,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
            if ($r -ne [System.Windows.Forms.DialogResult]::OK) { return }
        }
        $rtb.Text = $script:OofTemplate
        $rtb.Select(0, 0)
        $rtb.Focus()
        Add-Log 'Template inserted. Replace every <placeholder> before applying -- they are sent exactly as they appear.' 'WARN'
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Insert template' }
})

# Enter in the address box validates, which is what pressing Enter in an
# address box should always do.
$txtMailbox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $_.SuppressKeyPress = $true
        try { Invoke-MailboxLookup } catch { Show-HandlerError -ErrorRecord $_ -Context 'Connect and validate' }
    }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Out of Office' | Out-Null

# THE LOG KEEPS ITS CONSOLE COLOURS, and this has to happen AFTER
# Set-ConsoleTheme or the theme would repaint it like every other text box.
# Deliberate: it is a transcript, and a transcript that looks like a transcript
# is read as one.
$txtOut.BackColor = [System.Drawing.Color]::FromArgb(24, 24, 24)
$txtOut.ForeColor = [System.Drawing.Color]::FromArgb(210, 220, 210)

$form.Add_Shown({
    try {
        Update-StateEnabled
        Add-Log 'Ready. Enter a mailbox, press Connect and validate, compose the message, then Apply.'
        Add-Log 'Formatting (font, size, bold, italic, underline, strikethrough, colour, highlight, alignment, bullets, links) is converted to HTML automatically.'
        if ("$($txtMailbox.Text)".Trim()) { Invoke-MailboxLookup }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Out of Office' }
})

# The preview file is this process's, so it goes when the window does.
$form.Add_FormClosed({
    try {
        if ($script:PreviewFile -and (Test-Path -LiteralPath $script:PreviewFile)) {
            Remove-Item -LiteralPath $script:PreviewFile -Force -ErrorAction SilentlyContinue
        }
    } catch { }
})

[void]$form.ShowDialog()

# endregion ----------------------------------------------------------------- #
# End of Script
