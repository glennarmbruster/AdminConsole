#Requires -Version 5.1
<#
    HeaderAnalyzer.ps1
    Paste raw internet message headers, get the Received chain with per-hop
    delays, the authentication verdicts, and the Microsoft antispam headers
    decoded into English.

    Offline. Nothing leaves this machine.

    Author   : Glenn Armbruster
    Revised  : 2026-08-06
#>
# TOOLKIT-HELP-BEGIN
# Paste the raw internet headers from an email and get a readable summary of
# what happened to it.
#
#   * the delivery path, oldest hop first, with how long each hop took
#   * SPF, DKIM and DMARC results, and Microsoft's composite verdict
#   * the Exchange Online filtering result -- spam confidence, bulk level,
#     and WHY it was filtered, in words rather than codes
#   * anything worth a second look, like a Reply-To that points somewhere
#     other than the From address
#
# Useful for "why did this go to junk", "why was this blocked", and "is this
# actually from who it says it is".
#
# Everything happens on this machine. Nothing is uploaded, which is the main
# reason it exists -- headers contain internal server names, addresses and
# subject lines that should not be pasted into a public website.
#
# To get the headers: in Outlook, open the message, File > Properties, and
# copy the Internet headers box.
# TOOLKIT-HELP-END

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

#region Reference data ----------------------------------------------------------
<#
    Every value below is transcribed from
    https://learn.microsoft.com/en-us/defender-office-365/message-headers-eop-mdo
    (page last updated 2025-10-08), read on 2026-08-06.

    Checked rather than recalled, and it mattered twice:

      * SFV:SFE (safe-senders list) would have been missed entirely.
      * PCL -- phishing confidence level -- is NOT in the current header set.
        Decoding it would have produced a field that never appears, which is
        the kind of wrong that looks like working code forever.

    Microsoft revises these. When a value shows as "unknown" below, re-read the
    page before assuming the parser is broken.
#>

$script:CatMap = @{
    'AMP'    = 'Anti-malware'
    'BIMP'   = 'Brand impersonation (Defender for Office 365)'
    'BULK'   = 'Bulk'
    'DIMP'   = 'Domain impersonation (Defender for Office 365)'
    'FTBP'   = 'Anti-malware common attachments filter'
    'GIMP'   = 'Mailbox intelligence impersonation (Defender for Office 365)'
    'HPHSH'  = 'High confidence phishing'
    'HPHISH' = 'High confidence phishing'
    'HSPM'   = 'High confidence spam'
    'INTOS'  = 'Intra-organization phishing'
    'MALW'   = 'Malware'
    'OSPM'   = 'Outbound spam'
    'PHSH'   = 'Phishing'
    'SAP'    = 'Safe Attachments (Defender for Office 365)'
    'SPM'    = 'Spam'
    'SPOOF'  = 'Spoofing'
    'UIMP'   = 'User impersonation (Defender for Office 365)'
}

$script:SfvMap = @{
    'BLK'  = "Blocked -- sender is on the recipient's Blocked Senders list (filtering skipped)"
    'NSPM' = 'Not spam -- delivered to the intended recipients'
    'SFE'  = "Allowed -- sender is on the recipient's Safe Senders list (filtering skipped)"
    'SKA'  = 'Allowed -- sender/domain is on an anti-spam policy allow list (filtering skipped)'
    'SKB'  = 'Marked as spam -- sender/domain is on an anti-spam policy block list'
    'SKN'  = 'Marked NOT spam BEFORE filtering -- e.g. SCL -1 set by a mail flow rule'
    'SKQ'  = 'Released from quarantine and delivered'
    'SKS'  = 'Marked as spam BEFORE filtering -- e.g. SCL 5-9 set by a mail flow rule'
    'SPM'  = 'Marked as spam by spam filtering'
}

$script:DirMap = @{ 'INB' = 'Inbound'; 'OUT' = 'Outbound'; 'INT' = 'Internal' }
$script:IpvMap = @{
    'CAL' = 'Source IP was on the IP Allow List -- spam filtering skipped'
    'NLI' = 'Source IP was not found on any IP reputation list'
}
$script:SftyMap = @{
    '9.19' = 'Domain impersonation -- sending domain is impersonating a protected domain'
    '9.20' = 'User impersonation -- sender is impersonating a protected user'
    '9.25' = 'First contact safety tip -- may indicate a suspicious or phishing message'
}

# compauth reason codes. Exact codes first, then the documented ranges.
$script:CompAuthReason = @{
    '000' = 'Failed explicit authentication. DMARC fail with p=quarantine or p=reject.'
    '001' = 'Failed implicit authentication. Sender published no auth records, or weak policy (SPF ~all/?all, DMARC p=none).'
    '002' = 'Sender/domain pair is explicitly prohibited from spoofing by an admin-configured policy.'
    '010' = 'Failed DMARC with p=reject/quarantine AND the sending domain is one of your accepted domains (self-to-self spoofing).'
    '100' = 'SPF or DKIM passed and MAIL FROM / From domains are aligned.'
    '101' = 'DKIM signed by the domain in the From address.'
    '102' = 'MAIL FROM and From domains aligned, and SPF passed.'
    '103' = 'From domain aligns with the DNS PTR record of the source IP.'
    '104' = 'DNS PTR record of the source IP aligns with the From domain.'
    '108' = 'DKIM failed due to body modification by a previous legitimate hop -- e.g. your own on-premises environment.'
    '109' = 'No DMARC record for the sender, but the message would pass anyway.'
    '111' = 'DMARC temp/perm error, but SPF or DKIM domain aligns with the From domain.'
    '112' = 'A DNS timeout prevented the DMARC record from being retrieved.'
    '115' = 'Sent from a Microsoft 365 organization where the From domain is an accepted domain.'
    '116' = 'MX record of the From domain aligns with the PTR record of the connecting IP.'
    '130' = 'A trusted ARC sealer overrode the DMARC failure.'
    '201' = 'PTR of the From domain aligns with the subnet of the connecting IP PTR.'
    '202' = 'From domain aligns with the domain of the connecting IP PTR.'
    '501' = 'DMARC not enforced -- valid NDR, prior contact established between sender and recipient.'
    '502' = 'DMARC not enforced -- valid NDR for a message sent from this organization.'
    '601' = 'Sending domain is an accepted domain in your organization (self-to-self / intra-org spoofing).'
    '905' = 'DMARC not enforced due to complex routing -- e.g. routed via on-premises Exchange before reaching Microsoft 365.'
}
$script:CompAuthRange = @(
    @{ Match = '^1'; Text = 'Passed explicit or implicit authentication.' }
    @{ Match = '^2'; Text = 'Soft-passed implicit authentication.' }
    @{ Match = '^3'; Text = 'Not checked for composite authentication.' }
    @{ Match = '^4'; Text = 'Bypassed composite authentication.' }
    @{ Match = '^6'; Text = 'Failed implicit email authentication.' }
    @{ Match = '^70'; Text = 'DMARC not enforced -- this org has a history of legitimate mail from this infrastructure.' }
    @{ Match = '^7'; Text = 'Passed implicit authentication.' }
    @{ Match = '^9'; Text = 'Bypassed composite authentication.' }
)

function Get-CompAuthReasonText {
    param([string]$Code)
    if (-not $Code) { return $null }
    if ($script:CompAuthReason.ContainsKey($Code)) { return $script:CompAuthReason[$Code] }
    foreach ($r in $script:CompAuthRange) {
        if ($Code -match $r.Match) { return $r.Text }
    }
    return "Unknown reason code -- re-read the Microsoft Learn anti-spam headers page."
}
#endregion

#region Parsing -----------------------------------------------------------------
function ConvertFrom-EncodedWord {
    <#
        Decodes RFC 2047 encoded-words: =?charset?B?base64?= and =?charset?Q?quoted?=

        Subjects and display names arrive encoded whenever they contain
        anything outside ASCII, and an analyzer that shows the raw form makes
        the operator do the decoding -- which is the job.

        Matches are replaced by walking them explicitly rather than by handing
        a scriptblock to Regex.Replace as a MatchEvaluator. The delegate
        conversion does work in PowerShell, but this runs on untrusted input
        and a plain loop has no conversion semantics to be wrong about.
    #>
    param([string]$Text)
    if (-not $Text) { return $Text }
    if ($Text -notmatch '=\?') { return $Text }

    $result = $Text
    foreach ($m in [regex]::Matches($Text, '=\?([^?]+)\?([BbQq])\?([^?]*)\?=')) {
        $decoded = $m.Value
        try {
            $enc  = [System.Text.Encoding]::GetEncoding($m.Groups[1].Value)
            $kind = $m.Groups[2].Value.ToUpperInvariant()
            $data = $m.Groups[3].Value
            if ($kind -eq 'B') {
                $decoded = $enc.GetString([Convert]::FromBase64String($data))
            } else {
                # Q-encoding: underscore means space, =XX is a hex byte.
                $s     = $data.Replace('_', ' ')
                $bytes = New-Object System.Collections.ArrayList
                $i = 0
                while ($i -lt $s.Length) {
                    if ($s[$i] -eq '=' -and ($i + 2) -lt $s.Length) {
                        [void]$bytes.Add([Convert]::ToByte($s.Substring($i + 1, 2), 16))
                        $i += 3
                    } else {
                        [void]$bytes.Add([byte][char]$s[$i])
                        $i++
                    }
                }
                $decoded = $enc.GetString([byte[]]$bytes.ToArray())
            }
        } catch {
            # An undecodable word is shown as-is. Better a visible =?utf-8?B?...
            # than a silently dropped subject line.
            $decoded = $m.Value
        }
        $result = $result.Replace($m.Value, $decoded)
    }
    return $result
}

function ConvertFrom-RawHeaders {
    <#
        Raw text to an ordered list of header objects, duplicates preserved.

        UNFOLDING IS THE WHOLE JOB. A header continues onto the next line
        whenever that line starts with a space or tab (RFC 5322 section 2.2.3),
        and Received headers are almost always folded across three or four
        lines. Parsing line-by-line without unfolding produces a Received chain
        with no dates in it, because the date is on the last fold.

        Duplicates are kept in a list rather than a hashtable: Received appears
        once per hop and Authentication-Results can appear several times, and a
        hashtable would silently keep exactly one of each.
    #>
    param([string]$Raw)

    $headers = New-Object System.Collections.ArrayList
    if (-not "$Raw".Trim()) { return $headers }

    # Normalise line endings first. Pasted headers arrive with whatever the
    # source used, and a lone CR would otherwise end up inside a value.
    $lines = ($Raw -replace "`r`n", "`n" -replace "`r", "`n") -split "`n"

    $name = $null
    $val  = New-Object System.Text.StringBuilder

    $flush = {
        if ($name) {
            [void]$headers.Add([PSCustomObject]@{
                Name  = $name
                Value = $val.ToString().Trim()
            })
        }
    }

    foreach ($line in $lines) {
        if ($line -match '^[ \t]' -and $name) {
            # Folded continuation. The fold is replaced by a single space,
            # which is what unfolding means -- not removed, or words either
            # side of the fold would run together.
            [void]$val.Append(' ').Append($line.Trim())
            continue
        }
        if ($line -match '^([!-9;-~]+):[ \t]*(.*)$') {
            & $flush
            $name = $Matches[1]
            $val  = New-Object System.Text.StringBuilder
            [void]$val.Append($Matches[2])
            continue
        }
        if ("$line".Trim() -eq '') {
            # Blank line ends the header block and starts the body. Anything
            # after it is not a header, whatever it looks like.
            & $flush
            $name = $null
            break
        }
        # A line that is neither a fold nor a header. Pasting usually picks up
        # a stray line or two; ignoring them beats guessing.
    }
    & $flush
    return $headers
}

function Get-HeaderValue {
    param($Headers, [string]$Name)
    $h = @($Headers | Where-Object { $_.Name -ieq $Name })
    if ($h.Count -eq 0) { return $null }
    return $h[0].Value
}

function Get-HeaderValues {
    param($Headers, [string]$Name)
    return @($Headers | Where-Object { $_.Name -ieq $Name } | ForEach-Object { $_.Value })
}

function ConvertFrom-SemicolonFields {
    <#
        'CIP:1.2.3.4;CTRY:US;SCL:1;SFV:NSPM' -> hashtable.

        Split on ';' then on the FIRST ':' only. Values contain colons --
        IPv6 connecting addresses most obviously -- and splitting on every
        colon truncates them to their first group.
    #>
    param([string]$Value)
    $map = @{}
    if (-not $Value) { return $map }
    foreach ($part in ($Value -split ';')) {
        $p = $part.Trim()
        if (-not $p) { continue }
        $i = $p.IndexOf(':')
        if ($i -lt 1) { continue }
        $k = $p.Substring(0, $i).Trim()
        $v = $p.Substring($i + 1).Trim()
        if ($k) { $map[$k] = $v }
    }
    return $map
}

function ConvertFrom-ReceivedHeader {
    <#
        One Received: header to its parts.

        Every clause is optional and matched independently. There is no single
        canonical Received format -- RFC 5321 defines the tokens but not the
        order or which appear, and every MTA differs. A monolithic regex that
        expects from/by/with/id/for in sequence fails on the first appliance
        that omits one, and fails silently by matching nothing.

        The DATE is the reliable part: RFC 5321 puts it after the final ';'.
        That is what the per-hop timing depends on, so it is taken from the
        end rather than searched for.
    #>
    param([string]$Value)

    $o = [ordered]@{
        From = $null; FromIp = $null; By = $null; With = $null
        Id = $null; For = $null; Date = $null; DateText = $null
        Raw = $Value
    }

    $semi = $Value.LastIndexOf(';')
    if ($semi -ge 0) {
        $dateText = $Value.Substring($semi + 1).Trim()
        # Trailing comments like '(UTC)' are legal and break DateTimeOffset.
        $dateText = ($dateText -replace '\([^)]*\)\s*$', '').Trim()
        $o.DateText = $dateText
        $dt = [DateTimeOffset]::MinValue
        if ([DateTimeOffset]::TryParse($dateText, [ref]$dt)) { $o.Date = $dt }
    }

    if ($Value -match '(?i)\bfrom\s+([^\s;]+)')        { $o.From = $Matches[1] }
    if ($Value -match '(?i)\bby\s+([^\s;]+)')          { $o.By   = $Matches[1] }
    if ($Value -match '(?i)\bwith\s+([^\s;(]+)')       { $o.With = $Matches[1] }
    if ($Value -match '(?i)\bid\s+([^\s;]+)')          { $o.Id   = $Matches[1] }
    if ($Value -match '(?i)\bfor\s+<([^>]+)>')         { $o.For  = $Matches[1] }

    # The connecting IP, searched ONLY in the 'from' clause.
    #
    # Two things learned by testing this against real Microsoft 365 headers:
    #
    #  1. Restricting to the text before 'by' is not tidiness. The 'id' clause
    #     carries an Exchange version like 'id 15.20.771.19', which a loose
    #     IPv4 pattern happily reports as the connecting IP -- a plausible,
    #     completely wrong answer on a header that has no IP at all.
    #
    #  2. Microsoft writes IPv6 in PARENTHESES -- (2603:10b6:5:1f5::8) -- not
    #     brackets. A bracket-only pattern silently returns nothing for every
    #     hop inside Exchange Online, which is most of them.
    #
    # IPv6 is tried first: an IPv6 hop can still contain a dotted version
    # string, but an IPv4 hop never contains a colon-group token.
    $fromPart = $Value
    $byAt = [regex]::Match($Value, '(?i)\bby\b')
    if ($byAt.Success) { $fromPart = $Value.Substring(0, $byAt.Index) }

    if ($fromPart -match '[\[(]((?:[0-9A-Fa-f]{0,4}:){2,}[0-9A-Fa-f]{0,4})[\])]') {
        $o.FromIp = $Matches[1]
    } elseif ($fromPart -match '[\[(]?((?:\d{1,3}\.){3}\d{1,3})[\])]?') {
        $candidate = $Matches[1]
        # Octet range check. '15.20.771.19' matches the shape and is not an
        # address, and the shape is all a regex can see.
        $octetsOk = $true
        foreach ($o8 in ($candidate -split '\.')) { if ([int]$o8 -gt 255) { $octetsOk = $false } }
        if ($octetsOk) { $o.FromIp = $candidate }
    }

    return [PSCustomObject]$o
}

function ConvertFrom-AuthResults {
    <#
        Pulls spf / dkim / dmarc / compauth out of Authentication-Results.

        There can be more than one such header -- an on-premises hop and
        Microsoft 365 both stamp one -- so all of them are concatenated before
        matching. Taking only the first would report the on-premises verdict
        for a message whose cloud verdict is the one in question.
    #>
    param($Headers)

    $all = @(Get-HeaderValues -Headers $Headers -Name 'Authentication-Results')
    $all += @(Get-HeaderValues -Headers $Headers -Name 'Authentication-Results-Original')
    $joined = ($all -join ' ; ')

    $o = [ordered]@{
        Spf = $null; SpfDetail = $null; SmtpMailFrom = $null
        Dkim = $null; DkimDetail = $null; DkimDomain = $null
        Dmarc = $null; DmarcAction = $null; HeaderFrom = $null
        CompAuth = $null; Reason = $null
        Present = [bool]$joined
        Raw = $all
    }
    if (-not $joined) { return [PSCustomObject]$o }

    if ($joined -match '(?i)\bspf=([a-z]+)\s*(\([^)]*\))?') {
        $o.Spf = $Matches[1].ToLowerInvariant()
        if ($Matches[2]) { $o.SpfDetail = $Matches[2].Trim('(', ')') }
    }
    if ($joined -match '(?i)\bsmtp\.mailfrom=([^\s;]+)') { $o.SmtpMailFrom = $Matches[1] }

    if ($joined -match '(?i)\bdkim=([a-z]+)\s*(\([^)]*\))?') {
        $o.Dkim = $Matches[1].ToLowerInvariant()
        if ($Matches[2]) { $o.DkimDetail = $Matches[2].Trim('(', ')') }
    }
    if ($joined -match '(?i)\bheader\.d=([^\s;]+)') { $o.DkimDomain = $Matches[1] }

    if ($joined -match '(?i)\bdmarc=([a-z]+)')      { $o.Dmarc       = $Matches[1].ToLowerInvariant() }
    if ($joined -match '(?i)\baction=([^\s;]+)')    { $o.DmarcAction = $Matches[1] }
    if ($joined -match '(?i)\bheader\.from=([^\s;]+)') { $o.HeaderFrom = $Matches[1] }

    if ($joined -match '(?i)\bcompauth=([a-z]+)')   { $o.CompAuth = $Matches[1].ToLowerInvariant() }
    if ($joined -match '(?i)\breason=(\d{3})')      { $o.Reason   = $Matches[1] }

    return [PSCustomObject]$o
}

function Get-DomainFromAddress {
    <#
        NOT named Get-AddressDomain. That form starts with the literal string
        'Get-AD', which the toolkit's capability scanner matches as an Active
        Directory cmdlet -- so this offline tool reported itself as talking to
        AD. A cosmetic-looking name choice with a factual consequence.

        Domain out of 'Display Name <user@domain>' or a bare address.
        Used for the From / Return-Path alignment check, so it has to cope
        with both forms and give back $null rather than nonsense when it
        cannot.
    #>
    param([string]$Value)
    if (-not $Value) { return $null }
    $v = $Value
    if ($v -match '<([^>]+)>') { $v = $Matches[1] }
    $v = $v.Trim().Trim('<', '>')
    if ($v -match '@([^\s>;,]+)$') { return $Matches[1].ToLowerInvariant().TrimEnd('.') }
    return $null
}
#endregion

#region Report ------------------------------------------------------------------
function Format-Delay {
    param([TimeSpan]$Span)
    $s = [Math]::Round($Span.TotalSeconds)
    if ($s -lt 0)    { return "$s s  (clock skew)" }
    if ($s -lt 60)   { return "$s s" }
    # [Math]::Floor, not [int] -- the cast rounds rather than truncating, so
    # 119 seconds rendered as "2 m 59 s". Same defect as Format-PimSpan.
    if ($s -lt 3600) { return "{0} m {1} s" -f [int][Math]::Floor($s / 60), ($s % 60) }
    return "{0} h {1} m" -f [int][Math]::Floor($s / 3600), [int][Math]::Floor(($s % 3600) / 60)
}

function New-HeaderReport {
    param([string]$Raw)

    $sb = New-Object System.Text.StringBuilder
    $add = { param($t) [void]$sb.AppendLine($t) }

    $headers = ConvertFrom-RawHeaders -Raw $Raw
    if ($headers.Count -eq 0) {
        & $add 'Nothing parsed. Paste the full internet headers, starting at the first "Received:" or "From:" line.'
        return $sb.ToString()
    }

    # ---------------------------------------------------------------- summary
    $from       = ConvertFrom-EncodedWord (Get-HeaderValue $headers 'From')
    $to         = ConvertFrom-EncodedWord (Get-HeaderValue $headers 'To')
    $subject    = ConvertFrom-EncodedWord (Get-HeaderValue $headers 'Subject')
    $replyTo    = ConvertFrom-EncodedWord (Get-HeaderValue $headers 'Reply-To')
    $returnPath = Get-HeaderValue $headers 'Return-Path'
    $sender     = ConvertFrom-EncodedWord (Get-HeaderValue $headers 'Sender')
    $date       = Get-HeaderValue $headers 'Date'
    $msgId      = Get-HeaderValue $headers 'Message-ID'

    & $add '=============================================================='
    & $add ' MESSAGE'
    & $add '=============================================================='
    & $add ("Subject      : {0}" -f $subject)
    & $add ("From         : {0}" -f $from)
    & $add ("To           : {0}" -f $to)
    if ($replyTo)    { & $add ("Reply-To     : {0}" -f $replyTo) }
    if ($sender)     { & $add ("Sender       : {0}" -f $sender) }
    & $add ("Return-Path  : {0}" -f $returnPath)
    & $add ("Date         : {0}" -f $date)
    & $add ("Message-ID   : {0}" -f $msgId)
    & $add ''
    & $add 'Paste the Message-ID into Message Trace to follow this message through the service.'
    & $add ''

    # -------------------------------------------------------- authentication
    $auth = ConvertFrom-AuthResults -Headers $headers
    & $add '=============================================================='
    & $add ' AUTHENTICATION'
    & $add '=============================================================='
    if (-not $auth.Present) {
        & $add 'No Authentication-Results header.'
        & $add 'Expected on inbound internet mail. Its absence usually means this is'
        & $add 'an internal message, or the headers were captured before the service'
        & $add 'stamped them.'
    } else {
        $spfLine = "SPF          : $($auth.Spf)"
        if ($auth.SmtpMailFrom) { $spfLine += "   (smtp.mailfrom=$($auth.SmtpMailFrom))" }
        & $add $spfLine
        if ($auth.SpfDetail) { & $add "               $($auth.SpfDetail)" }

        $dkimLine = "DKIM         : $($auth.Dkim)"
        if ($auth.DkimDomain) { $dkimLine += "   (header.d=$($auth.DkimDomain))" }
        & $add $dkimLine
        if ($auth.DkimDetail) { & $add "               $($auth.DkimDetail)" }

        $dmarcLine = "DMARC        : $($auth.Dmarc)"
        if ($auth.DmarcAction) { $dmarcLine += "   action=$($auth.DmarcAction)" }
        if ($auth.HeaderFrom)  { $dmarcLine += "   (header.from=$($auth.HeaderFrom))" }
        & $add $dmarcLine

        if ($auth.CompAuth) {
            & $add "compauth     : $($auth.CompAuth)"
            if ($auth.Reason) {
                & $add "  reason $($auth.Reason) -- $(Get-CompAuthReasonText -Code $auth.Reason)"
            }
        }

        if ($auth.Dmarc -eq 'bestguesspass') {
            & $add ''
            & $add 'NOTE: bestguesspass means the sender has NO DMARC record. It did not'
            & $add '      pass DMARC -- it would have, if a record existed.'
        }
    }
    & $add ''

    # --------------------------------------------------------- MS antispam
    $ffRaw = Get-HeaderValue $headers 'X-Forefront-Antispam-Report'
    $msRaw = Get-HeaderValue $headers 'X-Microsoft-Antispam'
    & $add '=============================================================='
    & $add ' MICROSOFT FILTERING VERDICT'
    & $add '=============================================================='
    if (-not $ffRaw -and -not $msRaw) {
        & $add 'No X-Forefront-Antispam-Report or X-Microsoft-Antispam header.'
        & $add 'These are stamped by Exchange Online Protection, so a message that never'
        & $add 'traversed the service will not have them.'
    } else {
        $ff = ConvertFrom-SemicolonFields -Value $ffRaw
        $ms = ConvertFrom-SemicolonFields -Value $msRaw

        if ($ff.ContainsKey('SCL')) {
            $scl = $ff['SCL']
            $sclNote = switch -Regex ("$scl") {
                '^-1$'  { 'filtering bypassed -- trusted sender, or a mail flow rule' ; break }
                '^[01]$' { 'not spam' ; break }
                '^[45]$' { 'spam' ; break }
                '^[5-9]$' { 'spam / high confidence spam' ; break }
                default { 'see the SCL reference' }
            }
            & $add ("SCL          : {0}   ({1})" -f $scl, $sclNote)
        }
        if ($ms.ContainsKey('BCL')) {
            $bcl = $ms['BCL']
            $bclNote = if ("$bcl" -eq '0') { 'not bulk' } else { 'bulk -- higher means more complaint-prone' }
            & $add ("BCL          : {0}   ({1})" -f $bcl, $bclNote)
        }
        if ($ff.ContainsKey('SFV')) {
            $sfv = $ff['SFV']
            $txt = if ($script:SfvMap.ContainsKey($sfv)) { $script:SfvMap[$sfv] }
                   else { 'unknown SFV value -- re-read the Microsoft Learn anti-spam headers page' }
            & $add ("SFV          : {0}   {1}" -f $sfv, $txt)
        }
        if ($ff.ContainsKey('CAT') -and $ff['CAT']) {
            $cat = $ff['CAT']
            $txt = if ($script:CatMap.ContainsKey($cat)) { $script:CatMap[$cat] }
                   else { 'unknown CAT value -- re-read the Microsoft Learn anti-spam headers page' }
            & $add ("CAT          : {0}   {1}" -f $cat, $txt)
        }
        if ($ff.ContainsKey('SFTY') -and $ff['SFTY']) {
            $sfty = $ff['SFTY']
            $txt = if ($script:SftyMap.ContainsKey($sfty)) { $script:SftyMap[$sfty] } else { 'phishing safety tip applied' }
            & $add ("SFTY         : {0}   {1}" -f $sfty, $txt)
        }
        if ($ff.ContainsKey('DIR') -and $ff['DIR']) {
            $dir = $ff['DIR']
            & $add ("Direction    : {0}   {1}" -f $dir, $(if ($script:DirMap.ContainsKey($dir)) { $script:DirMap[$dir] } else { '' }))
        }
        if ($ff.ContainsKey('IPV') -and $ff['IPV']) {
            $ipv = $ff['IPV']
            & $add ("IPV          : {0}   {1}" -f $ipv, $(if ($script:IpvMap.ContainsKey($ipv)) { $script:IpvMap[$ipv] } else { '' }))
        }
        if ($ff.ContainsKey('SRV') -and $ff['SRV']) {
            & $add ("SRV          : {0}   {1}" -f $ff['SRV'], $(if ($ff['SRV'] -eq 'BULK') { 'identified as bulk email' } else { '' }))
        }
        & $add ''
        if ($ff.ContainsKey('CIP') -and $ff['CIP']) { & $add ("Connecting IP: {0}" -f $ff['CIP']) }
        if ($ff.ContainsKey('PTR') -and $ff['PTR']) { & $add ("Reverse DNS  : {0}" -f $ff['PTR']) }
        if ($ff.ContainsKey('H')   -and $ff['H'])   { & $add ("HELO/EHLO    : {0}" -f $ff['H']) }
        if ($ff.ContainsKey('CTRY') -and $ff['CTRY']) { & $add ("Source ctry  : {0}" -f $ff['CTRY']) }
        if ($ff.ContainsKey('LANG') -and $ff['LANG']) { & $add ("Language     : {0}" -f $ff['LANG']) }
    }
    & $add ''

    # ------------------------------------------------------------- hop chain
    & $add '=============================================================='
    & $add ' DELIVERY PATH  (oldest first)'
    & $add '=============================================================='
    $recv = @(Get-HeaderValues -Headers $headers -Name 'Received')
    if ($recv.Count -eq 0) {
        & $add 'No Received headers found.'
    } else {
        # Received headers are PREPENDED by each hop, so the first one in the
        # message is the LAST hop. Reversed here so the chain reads in the
        # direction the message actually travelled.
        [array]::Reverse($recv)
        $hops = @($recv | ForEach-Object { ConvertFrom-ReceivedHeader -Value $_ })

        $prev = $null
        $i = 0
        $totalKnown = $true
        foreach ($h in $hops) {
            $i++
            $delay = ''
            if ($h.Date -and $prev -and $prev.Date) {
                $delay = "  +$(Format-Delay -Span ($h.Date - $prev.Date))"
            } elseif (-not $h.Date) {
                $totalKnown = $false
            }
            & $add ("{0,2}. {1}{2}" -f $i, $(if ($h.DateText) { $h.DateText } else { '(no parseable date)' }), $delay)
            if ($h.From) { & $add ("    from : {0}{1}" -f $h.From, $(if ($h.FromIp) { "  [$($h.FromIp)]" } else { '' })) }
            if ($h.By)   { & $add ("    by   : {0}" -f $h.By) }
            if ($h.With) { & $add ("    with : {0}" -f $h.With) }
            if ($h.For)  { & $add ("    for  : {0}" -f $h.For) }
            & $add ''
            $prev = $h
        }

        $first = @($hops | Where-Object { $_.Date })
        if ($first.Count -ge 2) {
            $total = $first[-1].Date - $first[0].Date
            & $add ("Total transit: {0}{1}" -f (Format-Delay -Span $total),
                    $(if (-not $totalKnown) { '   (some hops had no parseable date -- this is a floor, not a total)' } else { '' }))
        }
        & $add ''
        & $add 'Hop delays are the difference between each server''s own clock. A server'
        & $add 'with a wrong clock produces a negative or absurd delay -- that is the'
        & $add 'clock being wrong, not the message being slow.'
    }
    & $add ''

    # ----------------------------------------------------------- observations
    & $add '=============================================================='
    & $add ' WORTH A LOOK'
    & $add '=============================================================='
    $notes = New-Object System.Collections.ArrayList

    $fromDom = Get-DomainFromAddress $from
    $rpDom   = Get-DomainFromAddress $returnPath
    $replyDom = Get-DomainFromAddress $replyTo

    if ($fromDom -and $rpDom -and $fromDom -ne $rpDom) {
        [void]$notes.Add("From domain ($fromDom) differs from Return-Path domain ($rpDom). Normal for mailing lists and bulk senders; also what a spoof looks like. Check SPF and DMARC above.")
    }
    if ($fromDom -and $replyDom -and $fromDom -ne $replyDom) {
        [void]$notes.Add("Reply-To domain ($replyDom) differs from the From domain ($fromDom). Common in phishing -- the reply goes somewhere other than where it appears to come from.")
    }
    if ($auth.Present) {
        if ($auth.Spf -in @('fail', 'softfail'))  { [void]$notes.Add("SPF $($auth.Spf) -- the sending IP is not authorised by the From domain's SPF record.") }
        if ($auth.Dkim -eq 'fail')                { [void]$notes.Add('DKIM failed. If the message passed through a list server or an on-premises hop that modified the body, this can be legitimate (see compauth reason 108).') }
        if ($auth.Dmarc -eq 'fail')               { [void]$notes.Add("DMARC failed with action=$($auth.DmarcAction). This is the verdict that decides spoofing.") }
        if ($auth.CompAuth -eq 'fail')            { [void]$notes.Add('Composite authentication FAILED -- Microsoft did not believe the sender is who the From address claims.') }
    }
    if ($recv.Count -gt 12) {
        [void]$notes.Add("$($recv.Count) hops. A long chain usually means forwarding or a mailing list, and each extra hop is another place DKIM can break.")
    }
    if (Get-HeaderValue $headers 'X-MS-Exchange-CrossTenant-Id') {
        [void]$notes.Add('Cross-tenant headers present -- this message crossed a Microsoft 365 tenant boundary.')
    }

    if ($notes.Count -eq 0) {
        & $add 'Nothing obviously off. That is not the same as safe -- it means nothing in'
        & $add 'the headers contradicts itself.'
    } else {
        foreach ($n in $notes) { & $add "  * $n" }
    }
    & $add ''
    & $add ("Parsed {0} header(s), {1} hop(s). Nothing left this machine." -f $headers.Count, $recv.Count)

    return $sb.ToString()
}
#endregion

#region UI ----------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Header Analyzer'
$form.Size          = New-Object System.Drawing.Size(1000, 820)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblIn          = New-Object System.Windows.Forms.Label
$lblIn.Text     = 'Paste the raw internet message headers:'
$lblIn.Location = New-Object System.Drawing.Point(12, 10)
$lblIn.Size     = New-Object System.Drawing.Size(400, 18)
$form.Controls.Add($lblIn)

$lblPrivacy           = New-Object System.Windows.Forms.Label
$lblPrivacy.Tag       = 'NoTheme'
$lblPrivacy.Text      = 'Offline -- nothing is sent anywhere.'
$lblPrivacy.Location  = New-Object System.Drawing.Point(700, 10)
$lblPrivacy.Size      = New-Object System.Drawing.Size(270, 18)
$lblPrivacy.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
$lblPrivacy.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblPrivacy)

$txtIn            = New-Object System.Windows.Forms.TextBox
$txtIn.Multiline  = $true
$txtIn.ScrollBars = 'Vertical'
$txtIn.WordWrap   = $false
$txtIn.Location   = New-Object System.Drawing.Point(12, 32)
$txtIn.Size       = New-Object System.Drawing.Size(958, 170)
$txtIn.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtIn)

$btnGo          = New-Object System.Windows.Forms.Button
$btnGo.Text     = 'Analyze'
$btnGo.Location = New-Object System.Drawing.Point(12, 210)
$btnGo.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnGo)

$btnPaste          = New-Object System.Windows.Forms.Button
$btnPaste.Text     = 'Paste + Analyze'
$btnPaste.Location = New-Object System.Drawing.Point(130, 210)
$btnPaste.Size     = New-Object System.Drawing.Size(130, 30)
$form.Controls.Add($btnPaste)

$btnCopy          = New-Object System.Windows.Forms.Button
$btnCopy.Text     = 'Copy Report'
$btnCopy.Location = New-Object System.Drawing.Point(268, 210)
$btnCopy.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnCopy)

$btnClear          = New-Object System.Windows.Forms.Button
$btnClear.Text     = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(386, 210)
$btnClear.Size     = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnClear)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(880, 210)
$btnClose.Size     = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnClose)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
# WordWrap on. A truncated diagnostic line is how the write-scope error hid its
# own explanation for three rounds -- not repeating that here.
$txtOut.WordWrap   = $true
$txtOut.Location   = New-Object System.Drawing.Point(12, 250)
$txtOut.Size       = New-Object System.Drawing.Size(958, 500)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.BackColor  = [System.Drawing.Color]::WhiteSmoke
$txtOut.Tag        = 'NoTheme'
$form.Controls.Add($txtOut)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(12, 758)
$lblStatus.Size     = New-Object System.Drawing.Size(958, 18)
$form.Controls.Add($lblStatus)

function Invoke-Analyze {
    if (-not "$($txtIn.Text)".Trim()) {
        $lblStatus.Text = 'Nothing to analyze -- paste headers first.'
        return
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $txtOut.Text = New-HeaderReport -Raw $txtIn.Text
        $lblStatus.Text = "Analyzed $(Get-Date -Format 'HH:mm:ss')."
    } catch {
        $txtOut.Text = "Could not analyze these headers.`r`n`r`n$(Format-ErrorDetail -ErrorRecord $_ -Context 'HeaderAnalyzer')"
        $lblStatus.Text = 'Failed -- see the output.'
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

$btnGo.Add_Click({ Invoke-Analyze })

$btnPaste.Add_Click({
    try {
        $c = [System.Windows.Forms.Clipboard]::GetText()
        if ("$c".Trim()) { $txtIn.Text = $c; Invoke-Analyze }
        else { $lblStatus.Text = 'Clipboard has no text.' }
    } catch {
        $lblStatus.Text = "Could not read the clipboard: $($_.Exception.Message)"
    }
})

$btnCopy.Add_Click({
    if (-not "$($txtOut.Text)".Trim()) { $lblStatus.Text = 'Nothing to copy.'; return }
    try {
        [System.Windows.Forms.Clipboard]::SetText($txtOut.Text)
        $lblStatus.Text = 'Report copied to the clipboard.'
    } catch {
        $lblStatus.Text = "Could not write to the clipboard: $($_.Exception.Message)"
    }
})

$btnClear.Add_Click({ $txtIn.Clear(); $txtOut.Clear(); $lblStatus.Text = '' })
$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Consistent '?' in the top-right of every tool. Added after the theme so
# it is not restyled into a normal button, and BringToFront'd inside the
# helper because WinForms draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Header Analyzer' | Out-Null

[void]$form.ShowDialog()
$form.Dispose()
#endregion
