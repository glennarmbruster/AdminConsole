#Requires -Version 5.1
<#
    HybridMessageTrace.ps1
    One search box that queries Exchange Online AND on-premises Exchange and
    shows the result as one correlated view.

    A SEPARATE TOOL FROM MessageTrace.ps1, ON PURPOSE. That one is the
    Exchange Online tool and it works; this does not replace it, extend it, or
    depend on it. Two buttons, two tools, and the cloud-only path stays exactly
    as it was for anyone who does not want the on-premises half.

    THIS TOOL READS ONLY, and must never gain a write. Quarantine release lives
    in MessageTrace.ps1 and stays there.
#>
# TOOLKIT-HELP-RUNS: READS Exchange Online and on-premises message tracking. Writes nothing.
# TOOLKIT-VARIANT: Hybrid
# TOOLKIT-HELP-BEGIN
# Traces a message across BOTH halves of a hybrid tenant in one search.
#
# In hybrid a message crosses two systems that answer different questions, and
# answering "where did it go" normally means running two tools with two
# vocabularies and stitching the result together by hand. This does that part.
#
# THE TWO SYSTEMS DISAGREE ABOUT WHAT A ROW IS. Exchange Online returns one row
# per recipient with a final status. On-premises returns one row per EVENT --
# received, then sent, then delivered -- so one message is three or four rows.
# This groups them, so a message appears ONCE per recipient with a single
# outcome. Double-click it to see every hop in order across both systems.
#
# ON-PREMISES IS THE SLOW HALF because there is no organisation-wide view: each
# transport server has to be asked separately. That is why the source starts on
# Exchange Online only.
#
# IF A SERVER DOES NOT ANSWER IT SAYS SO -- "3 of 4 servers answered", naming
# the one that failed. Rows will be missing when that happens. A trace that
# quietly returns less is how somebody concludes a message never existed.
#
# THE TWO 90-DAY BUTTONS ARE NOT THE SAME REPORT, and the difference decides
# which question you can answer:
#
#   90 day Enhanced   up to 100,000 rows. Summary fields plus direction and
#                     original client IP. Use it for "did this arrive, and
#                     where did it go" across a lot of mail.
#
#   90 day Extended   full routing and per-event detail, but capped at 1,000
#                     ROWS -- a hundredth of Enhanced. Use it on ONE message or
#                     a very narrow search. Pointed at a busy mailbox for 90
#                     days it stops at 1,000 and does not say so, and the CSV
#                     looks complete when it is not.
#
# Both are emailed as a CSV and neither fills the grid. Both need a sender,
# recipient or message ID -- the service refuses without one. Both are Exchange
# Online only.
#
# ARCHIVED DATA LAGS LIVE DATA BY 1-4 HOURS, so a message sent this morning may
# be missing from a report whose range covers today. Use a normal search for
# anything recent.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

#region Load Common -------------------------------------------------------------
$root = $PSScriptRoot
while ($root -and -not (Test-Path -LiteralPath (Join-Path $root 'Common.ps1'))) {
    $parent = Split-Path -Parent $root
    if (-not $parent -or $parent -eq $root) { $root = ''; break }
    $root = $parent
}
if (-not $root) { throw "Common.ps1 not found above '$PSScriptRoot'." }
. (Join-Path $root 'Common.ps1')
#endregion

if (-not (Test-ConsoleLaunch)) { return }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Hybrid trace core -------------------------------------------------------
<#
    THE NORMALISED ROW, AND THE TWO SOURCES THAT PRODUCE IT.

    Exchange Online and on-premises Exchange answer different questions. EXO
    returns ONE row per recipient carrying a final STATUS. On-premises returns
    one row per EVENT -- RECEIVE, then SEND, then DELIVER -- so a single message
    is three or four rows before anything is done to it.

    Merged naively, every on-premises message appears several times and reads as
    a duplication bug. The grouping in Get-OnPremTraceRows is the fix, and it is
    the reason this region exists rather than a second Get- call bolted onto the
    existing one.

    READS ONLY. Neither of these cmdlets has a counterpart in the other edition,
    so loading both environments cannot produce a name collision here (C1) --
    but both calls are routed explicitly anyway, so the intent survives the day
    somebody adds a third cmdlet.
#>

function ConvertTo-TraceMessageId {
    <#
        The join key, made comparable.

        MessageId is what survives the hybrid hop, and it is the entire reason
        one message can be shown crossing both systems. The two sources do not
        agree on its punctuation:

        Learn 2026-08-23 (message tracking log fields, message-id): "For
        messages created in Exchange, the value is in the format
        <GUID@ServerFQDN>, INCLUDING THE ANGLE BRACKETS. Other messaging systems
        may use different syntax or values."

        Exchange Online is not consistent about them. So: strip brackets,
        trim, lowercase. Stripping rather than adding, because a value from
        "other messaging systems" may have no brackets to add back and would
        then differ from itself depending on which source saw it first.
    #>
    param([string]$Value)

    $v = "$Value".Trim()
    if (-not $v) { return '' }
    $v = $v.TrimStart('<').TrimEnd('>')
    return $v.ToLowerInvariant()
}

function Get-OnPremOutcome {
    <#
        EventId -> a single word for the flat grid.

        VERIFIED AGAINST THE FULL LIST, NOT THE COMMON ONES. Learn 2026-08-23
        (Event types in the message tracking log) lists 30-odd EventId values.
        The obvious eight are not the whole surface, and an unmapped event that
        silently becomes "Unknown" is how a failure gets read as a mystery.

        TWO CORRECTIONS TO THE OBVIOUS MAPPING, both from that list:

          - THE POISON EVENT IS 'POISONMESSAGE', NOT 'POISON'. A mapping keyed
            on POISON matches nothing, and poisoned messages report as Unknown
            forever.
          - DELIVERFAIL, BADMAIL, MODERATORREJECT, MODERATIONEXPIRE,
            RESUBMITFAIL and HAREDIRECTFAIL are all failures and none of them
            contain the word FAIL in a way a prefix match would catch
            reliably -- MODERATORREJECT does not contain it at all.

        ANYTHING NOT LISTED RETURNS '' AND IS TREATED AS HOP DETAIL, not as an
        outcome. That is deliberate: a new event type in a future Exchange build
        should make the timeline longer, never invent a wrong verdict.
    #>
    param([string]$EventId)

    switch ("$EventId".Trim().ToUpperInvariant()) {
        # --- delivered ---
        'DELIVER'            { return 'Delivered' }
        'DUPLICATEDELIVER'   { return 'Delivered (duplicate)' }
        # --- handed on ---
        'SEND'               { return 'Relayed' }
        'TRANSFER'           { return 'Relayed' }
        'REDIRECT'           { return 'Redirected' }
        'RESOLVE'            { return 'Redirected' }
        # --- not yet ---
        'DEFER'              { return 'Pending' }
        'RESUBMITDEFER'      { return 'Pending' }
        'THROTTLE'           { return 'Pending' }
        # --- failed ---
        'FAIL'               { return 'Failed' }
        'DELIVERFAIL'        { return 'Failed' }
        'BADMAIL'            { return 'Failed' }
        'RESUBMITFAIL'       { return 'Failed' }
        'HAREDIRECTFAIL'     { return 'Failed' }
        'MODERATORREJECT'    { return 'Rejected by moderator' }
        'MODERATIONEXPIRE'   { return 'Moderation expired' }
        'MODERATORSALLNDR'   { return 'Failed' }
        # --- gone, quietly ---
        'DROP'               { return 'Dropped' }
        'POISONMESSAGE'      { return 'Dropped' }
        # --- bounced ---
        'DSN'                { return 'Bounced' }
        # --- structural ---
        'EXPAND'             { return 'Expanded' }
        'DUPLICATEEXPAND'    { return 'Expanded' }
    }
    # RECEIVE, SUBMIT, CLIENTSUBMISSION, NOTIFYMAPI, NOTIFYSHADOW, HARECEIVE,
    # HAREDIRECT, HADISCARD, AGENTINFO, LOAD, PROCESS, PROCESSMEETINGMESSAGE,
    # DUPLICATEREDIRECT, INITMESSAGECREATED, MODERATORAPPROVE, RESUBMIT --
    # every one belongs in the timeline and none of them is an outcome.
    return ''
}

function Get-TraceOutcomeRank {
    <#
        Which outcome wins when one recipient has several.

        NOT "the last event", which is the tempting answer and is wrong here.
        Events arrive from several servers whose clocks are only approximately
        aligned, so chronological order across servers is not reliable enough to
        decide a verdict from.

        Ranked by how final and how bad the state is: a FAIL anywhere in a
        recipient's history matters more than a SEND that preceded it, and a
        DELIVER outranks the DEFER it followed.

        THE EXCHANGE ONLINE STATUSES WERE MISSING AND THAT WAS A REAL BUG.
        Get-MessageTraceV2 returns Status values this function had never heard
        of -- Quarantined, FilteredAsSpam, GettingStatus, None -- and an unknown
        outcome falls through to 0. Zero loses to EVERYTHING, so a recipient
        with a Quarantined cloud row and a Relayed on-premises row reported as
        RELAYED: the tool said the message went out fine when it was sitting in
        quarantine. That is the single worst answer this tool can give, and it
        was reachable with entirely ordinary data.

        THE FULL EXCHANGE ONLINE SET, from the Get-MessageTraceV2 reference:
        None, GettingStatus, Failed, Pending, Delivered, Expanded, Quarantined,
        FilteredAsSpam. All eight are ranked below.
    #>
    param([string]$Outcome)

    switch ("$Outcome") {
        # --- terminal and bad -------------------------------------------------
        'Failed'                { return 100 }
        'Bounced'               { return 95 }
        'Rejected by moderator' { return 90 }
        'Moderation expired'    { return 85 }
        'Dropped'               { return 80 }

        # --- held, and the message has NOT arrived ----------------------------
        #
        # ABOVE Delivered ON PURPOSE. A message quarantined for one recipient and
        # delivered for another is a message somebody did not get, and the row
        # has to say so -- ranking it below Delivered would hide exactly the
        # case the operator is searching for.
        'Quarantined'           { return 78 }

        # --- arrived, but not where they looked -------------------------------
        # Also above Delivered: "it never arrived" is usually this, and reporting
        # a plain Delivered sends the operator hunting for a mail flow fault that
        # does not exist.
        'FilteredAsSpam'        { return 75 }

        # --- arrived ----------------------------------------------------------
        'Delivered'             { return 70 }
        'Delivered (duplicate)' { return 69 }
        'Expanded'              { return 60 }
        'Redirected'            { return 50 }
        'Relayed'               { return 40 }

        # --- still moving -----------------------------------------------------
        'Pending'               { return 30 }
        'Deferred'              { return 28 }

        # --- Exchange Online has not finished writing the record --------------
        # Lowest of the real values, because it is not a verdict at all: it means
        # ask again in a minute.
        'GettingStatus'         { return 10 }
        'None'                  { return 5 }
    }

    # UNKNOWN RANKS ABOVE NOTHING, NOT AT ZERO.
    #
    # 1 rather than 0 so a status this build has never seen still beats the
    # absence of a status. A future Exchange value should degrade to "shown, and
    # ranked last among real answers", never to "silently discarded in favour of
    # a lesser one".
    if ("$Outcome".Trim()) { return 1 }
    return 0
}

function New-TraceRow {
    <#
        THE ONE CONSTRUCTOR. Both sources call this and neither builds a row by
        hand.

        That is the whole guarantee behind "both sources must produce the same
        shape". Two PSCustomObject literals in two functions are the same shape
        on the day they are written and diverge the first time one gains a
        field -- and the symptom is a blank column for half the rows, which
        reads as missing data rather than as a bug.
    #>
    param(
        [datetime]$TimestampUtc,
        [string]$Sender,
        [string]$Recipient,
        [string]$Subject,
        [string]$Outcome,
        [string]$MessageId,
        [long]$SizeBytes = 0,
        [string]$Source,
        [string]$Server = '',
        [string]$RawEvent = '',
        [string]$Detail = '',
        $Events = @(),

        # THE HANDLE FOR THE DETAIL LOOKUP. Get-MessageTraceDetailV2 takes a
        # MessageTraceId GUID and nothing else will do -- not the Message-ID,
        # not the subject. It has to be captured on the way past, because
        # re-finding it later means running the whole trace again.
        [string]$TraceGuid = ''
    )

    return [PSCustomObject]@{
        TimestampUtc = $TimestampUtc
        Sender       = "$Sender"
        Recipient    = "$Recipient"
        Subject      = "$Subject"
        Outcome      = "$Outcome"
        MessageId    = (ConvertTo-TraceMessageId -Value $MessageId)
        MessageIdRaw = "$MessageId"
        SizeBytes    = $SizeBytes
        Source       = "$Source"
        Server       = "$Server"
        RawEvent     = "$RawEvent"
        Detail       = "$Detail"
        TraceGuid    = "$TraceGuid"
        # The individual on-prem events behind this row, kept for the timeline.
        # Empty for Exchange Online, which has no equivalent.
        Events       = @($Events)
    }
}

function ConvertTo-TraceUtc {
    <#
        Both sources onto one clock, WITHOUT assuming which clock they started
        on.

        THE BRIEF FOR THIS TOOL SAID ON-PREMISES TIMESTAMPS ARE SERVER-LOCAL.
        Learn 2026-08-23 (Fields in the message tracking log files, date-time)
        says the opposite: "The UTC date-time of the message tracking event...
        Z signifies Zulu, which is another way to denote UTC."

        Both can be true in practice. The LOG FILE stores UTC; what the cmdlet
        hands back is a .NET DateTime whose Kind depends on how the remote
        session deserialised it, and Unspecified is common over PSRemoting.

        So this does not assume either way -- it reads Kind:

            Utc          taken as-is
            Local        converted
            Unspecified  TREATED AS UTC, per Learn, and recorded in UNVERIFIED

        Guessing wrong here shifts every on-premises row by the server's offset
        and silently destroys the correlation this whole tool exists to produce.
        A four-hour skew between the two halves of one message is not obviously
        a bug -- it looks like a slow hop.
    #>
    param($Value)

    if (-not $Value) { return $null }
    try {
        $dt = [datetime]$Value
        switch ($dt.Kind) {
            'Utc'         { return $dt }
            'Local'       { return $dt.ToUniversalTime() }
            default       { return [DateTime]::SpecifyKind($dt, [DateTimeKind]::Utc) }
        }
    } catch {
        return $null
    }
}

function Get-ExoTraceRows {
    <#
        Exchange Online -> normalised rows.

        One row per recipient already, carrying a final status, so there is
        nothing to group. This is the easy half and it is unchanged in
        behaviour from the tool as it stands -- deliberately, because the
        on-premises half must not be able to alter what the cloud half returns.
    #>
    param(
        [string]$SenderAddress,
        [string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [string]$MessageId
    )

    $rows = New-Object System.Collections.ArrayList
    $params = @{
        StartDate     = $StartDate
        EndDate       = $EndDate
        # C3: explicit on every enumerating call.
        ResultSize    = 5000
        WarningAction = 'SilentlyContinue'
        ErrorAction   = 'Stop'
    }
    if ($SenderAddress)    { $params.SenderAddress    = $SenderAddress }
    if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }
    if ($MessageId)        { $params.MessageId        = $MessageId }

    # Get-MessageTraceV2, not Get-MessageTrace: the V1 cmdlets are what the
    # existing tool already calls, and this half is not being changed.
    foreach ($m in @(Invoke-EXOCommand -Name 'Get-MessageTraceV2' -Parameters $params)) {
        [void]$rows.Add((New-TraceRow `
            -TimestampUtc (ConvertTo-TraceUtc -Value $m.Received) `
            -Sender    "$($m.SenderAddress)" `
            -Recipient "$($m.RecipientAddress)" `
            -Subject   "$($m.Subject)" `
            -Outcome   "$($m.Status)" `
            -MessageId "$($m.MessageId)" `
            -SizeBytes ([long]0 + $m.Size) `
            -Source    'Exchange Online' `
            -TraceGuid "$($m.MessageTraceId)"))
    }
    return @($rows.ToArray())
}

function Get-OnPremTransportServer {
    <#
        Which servers to ask.

        EDGE TRANSPORT IS EXCLUDED AND SAID SO, never silently dropped. An Edge
        server is not in the Active Directory site and typically will not answer
        this call from here at all -- but a hop that happened and is not shown
        is exactly the failure this tool is meant to end. The caller surfaces
        the exclusion in the results header.
    #>
    param()

    $r = [PSCustomObject]@{ Ok = $false; Servers = @(); Edge = @(); Excluded = @(); Reason = '' }

    # NEVER-QUERY LIST, from Tenant.json rather than from this file.
    #
    # Non-production and decommissioned transport servers fail every single
    # search, and four red rows on every run is how an operator learns to stop
    # reading the results header -- which is the one line in this tool that must
    # never be ignored.
    #
    # EXCLUDED, NOT SUPPRESSED. Silencing the error while still querying would
    # leave rows genuinely missing with nothing on screen to say so. Not asking
    # at all means "4 of 4 servers answered" is the truth rather than a
    # comforting version of "4 of 8 failed".
    #
    # They are still NAMED in the header, quietly. A message that genuinely
    # transited one of these would not appear, and that has to be visible even
    # when the exclusion is entirely deliberate.
    $never = @()
    try { $never = @(Get-Prop -Object (Get-Prop -Object (Get-TenantConfig) -Name 'OnPremExchange' -Default $null) -Name 'ExcludedServers' -Default @()) } catch { }
    try {
        # C13: ask before filtering on a property that may not exist.
        $all = @(Invoke-EmsCommand -Name 'Get-TransportService' -Parameters @{ ErrorAction = 'Stop' })
        if (@($all).Count -eq 0) {
            $r.Reason = 'Get-TransportService returned nothing, so there are no servers to query.'
            return $r
        }
        foreach ($s in $all) {
            $role = "$(Get-Prop -Object $s -Name 'ServerRole' -Default '')"
            $name = "$(Get-Prop -Object $s -Name 'Name' -Default $s)"
            # C12: exact, not a substring on 'Edge'.
            if ($role -eq 'Edge') { $r.Edge += $name; continue }
            # Case-insensitive: the list is typed by a person and the directory
            # does not agree with them about capitalisation.
            if (@($never | ForEach-Object { "$_".ToLowerInvariant() }) -contains $name.ToLowerInvariant()) {
                $r.Excluded += $name
                continue
            }
            $r.Servers += $name
        }
        $r.Ok = $true
    } catch {
        $r.Reason = "$($_.Exception.Message)"
    }
    return $r
}

function Get-OnPremTraceRows {
    <#
        On-premises -> normalised rows, one per message per recipient.

        THREE THINGS HAPPEN HERE AND THE ORDER MATTERS.

        1. PER-SERVER QUERY. Get-MessageTrackingLog has no organisation-wide
           view -- Learn 2026-08-23 documents -Server as the scope and there is
           no -Organization equivalent. Every transport server is asked
           separately and the answers are merged.

        2. FAN OUT THE RECIPIENTS. Learn 2026-08-23 (recipient-address): "The
           email addresses of the message's recipients. Multiple email addresses
           are separated by the semicolon character." The cmdlet surfaces this
           as an ARRAY, while Exchange Online returns one scalar recipient per
           row. Fan out first or every count in the grid is wrong -- a message
           to four people reads as one.

        3. GROUP TO ONE ROW. Only after fanning out, because the grouping key
           includes the recipient.

        PARTIAL RESULTS ARE THE MAIN CORRECTNESS RISK AND ARE RETURNED, NOT
        LOGGED AND FORGOTTEN. A server that does not answer means rows are
        missing, and a tool that returns fewer rows without saying so teaches
        the operator that the message never existed. The caller is given the
        counts and the names and must show them.
    #>
    param(
        [string]$SenderAddress,
        [string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [string]$MessageId,
        [string[]]$Servers,
        [int]$PerServerTimeoutSeconds = 90
    )

    $result = [PSCustomObject]@{
        Rows      = @()
        Asked     = @($Servers).Count
        Answered  = 0
        Failed    = @()   # @{ Server; Reason }
        Truncated = @()
    }

    $flat = New-Object System.Collections.ArrayList

    foreach ($srv in @($Servers)) {
        $params = @{
            Server      = $srv
            Start       = $StartDate
            End         = $EndDate
            # C3, and EXO002: WarningAction SITS DIRECTLY BELOW ResultSize with
            # nothing between them, comment included. A deliberately limited
            # ResultSize makes Exchange warn "There are more results available"
            # on every single run, and the truncation is already reported
            # honestly in the results header -- the warning is noise that
            # teaches the operator to ignore warnings.
            ResultSize    = 5000
            WarningAction = 'SilentlyContinue'
            ErrorAction   = 'Stop'
        }
        # Learn 2026-08-23 (Get-MessageTrackingLog syntax): -Sender takes a
        # single address, -Recipients takes String[], -MessageId takes the
        # header value. -MessageSubject exists but is not used here: subject
        # tracking can be switched off per-server via
        # Set-TransportService -MessageTrackingLogSubjectLoggingEnabled, so a
        # subject filter can silently return nothing on one server and results
        # on another.
        if ($SenderAddress)    { $params.Sender     = $SenderAddress }
        if ($RecipientAddress) { $params.Recipients = @($RecipientAddress) }
        if ($MessageId)        { $params.MessageId  = $MessageId }

        try {
            $raw = @(Invoke-EmsCommand -Name 'Get-MessageTrackingLog' -Parameters $params)
            $result.Answered++

            # C15: shape, never retry. A full ResultSize is a truncated answer
            # and is reported as one rather than silently accepted.
            if (@($raw).Count -ge 5000) { $result.Truncated += $srv }

            foreach ($e in $raw) {
                # 2. FAN OUT. One event, many recipients.
                $recips = @($e.Recipients)
                if (@($recips).Count -eq 0) { $recips = @('') }
                foreach ($rcp in $recips) {
                    [void]$flat.Add([PSCustomObject]@{
                        TimestampUtc = (ConvertTo-TraceUtc -Value $e.Timestamp)
                        Sender       = "$($e.Sender)"
                        Recipient    = "$rcp"
                        Subject      = "$($e.MessageSubject)"
                        EventId      = "$($e.EventId)"
                        MessageId    = (ConvertTo-TraceMessageId -Value $e.MessageId)
                        MessageIdRaw = "$($e.MessageId)"
                        SizeBytes    = ([long]0 + $e.TotalBytes)
                        Server       = "$($e.ServerHostname)"
                        Detail       = (@("$($e.SourceContext)", "$($e.EventData)") | Where-Object { $_ }) -join ' | '
                        # KEPT SEPARATELY AS WELL AS IN Detail.
                        #
                        # Detail is joined for display and is not safe to parse
                        # -- a SourceContext containing a pipe would split a
                        # rule record in half. The transport-rule reader needs
                        # the original.
                        SourceContext = "$($e.SourceContext)"
                        EventData     = @($e.EventData | ForEach-Object { "$_" })
                    })
                }
            }
        } catch {
            $result.Failed += [PSCustomObject]@{ Server = $srv; Reason = "$($_.Exception.Message)" }
        }
    }

    # 3. GROUP. MessageId + recipient -> one row, keeping the events.
    $grouped = New-Object System.Collections.ArrayList
    foreach ($g in @($flat | Group-Object { "$($_.MessageId)|$($_.Recipient.ToLowerInvariant())" })) {
        $events = @($g.Group | Sort-Object TimestampUtc)
        $first  = $events[0]

        $best = ''
        $bestRank = -1
        $bestEvent = $null
        foreach ($e in $events) {
            $o = Get-OnPremOutcome -EventId $e.EventId
            if (-not $o) { continue }
            $rank = Get-TraceOutcomeRank -Outcome $o
            if ($rank -gt $bestRank) { $bestRank = $rank; $best = $o; $bestEvent = $e }
        }
        if (-not $best) {
            # Every event was hop detail -- RECEIVE with nothing after it is a
            # real and interesting state, not an error. Saying "In transit" is
            # honest; saying "Delivered" would not be.
            $best = 'In transit'
            $bestEvent = $events[-1]
        }

        [void]$grouped.Add((New-TraceRow `
            -TimestampUtc $first.TimestampUtc `
            -Sender    $first.Sender `
            -Recipient $first.Recipient `
            -Subject   $first.Subject `
            -Outcome   $best `
            -MessageId $first.MessageIdRaw `
            -SizeBytes $first.SizeBytes `
            -Source    'On-premises' `
            -Server    "$(if ($bestEvent) { $bestEvent.Server } else { $first.Server })" `
            -RawEvent  "$(if ($bestEvent) { $bestEvent.EventId } else { '' })" `
            -Detail    "$(if ($bestEvent) { $bestEvent.Detail } else { '' })" `
            -Events    $events))
    }

    $result.Rows = @($grouped.ToArray())
    return $result
}

# WHICH MESSAGES ARE OPEN, keyed by Get-TraceGroupKey. Collapsed is the default:
# a hybrid message legitimately produces one row per source, and showing both by
# default means every message that crossed the boundary looks like two.
$script:Expanded = @{}
$script:RowGroup = @()

# THE SORT IS OURS, NOT THE GRID'S.
#
# A DataGridView bound to a DataTable sorts by reordering the whole view, which
# tears child rows away from the parent they belong to -- a child carries a
# blank Subject and Sender, so sorting by Subject piles every hop at one end,
# detached from any message. The grid's own sort cannot be made
# hierarchy-aware, so it is switched off and the GROUPS are sorted instead.
$script:SortCol = 'Time'
$script:SortAsc = $false

function Get-TraceGroupKey {
    <#
        What makes two rows the same message.

        MESSAGE-ID AND RECIPIENT TOGETHER, never Message-ID alone. One message
        to four people is four different stories -- delivered to two,
        quarantined for a third, still pending for the fourth -- and collapsing
        them onto one line would pick one verdict and hide three.

        The Message-ID is normalised by ConvertTo-TraceMessageId first, because
        the two sides disagree about angle brackets and a raw comparison puts
        the same message in two groups.
    #>
    param([Parameter(Mandatory)]$Row)
    return ("{0}|{1}" -f (ConvertTo-TraceMessageId -Value "$($Row.MessageIdRaw)"),
                          "$($Row.Recipient)".ToLowerInvariant())
}

function Get-GroupVerdict {
    <#
        The one row that speaks for a group.

        BY OUTCOME RANK, NOT BY TIME. The whole reason Get-TraceOutcomeRank
        exists: events come from several servers whose clocks are only roughly
        aligned, so "the latest row" is not reliably the final one. A message
        relayed on-premises at 10:04 and quarantined in the cloud at 10:03 was
        quarantined, whatever the timestamps say.
    #>
    param($Rows)
    return @($Rows | Sort-Object @{ E = { Get-TraceOutcomeRank -Outcome "$($_.Outcome)" } }, TimestampUtc -Descending)[0]
}
function Merge-TraceRows {
    <#
        Both sources into one ordered view.

        NOT DE-DUPLICATED ACROSS SOURCES, and that is the point rather than an
        omission. A message that crossed the hybrid boundary SHOULD appear
        twice -- once as the on-premises hop and once as the Exchange Online
        outcome -- because those are two different facts about it, and
        collapsing them would throw away the very thing this tool was built to
        show.

        The Source column is what makes that legible, and the timeline on
        double-click is what makes it useful.
    #>
    param($ExoRows = @(), $OnPremRows = @())

    $all = New-Object System.Collections.ArrayList
    foreach ($r in @($ExoRows))    { [void]$all.Add($r) }
    foreach ($r in @($OnPremRows)) { [void]$all.Add($r) }
    return @($all.ToArray() | Sort-Object TimestampUtc)
}

function Get-TraceHopTimeline {
    <#
        One message across both systems, oldest first.

        THIS IS THE PAYOFF -- the view neither system can produce alone. Every
        on-premises event is included, RECEIVE and SUBMIT among them: they are
        not outcomes but they are where the message WAS, and hiding them because
        they are "not interesting" removes the hop that answers the question.
    #>
    param([Parameter(Mandatory)][string]$MessageId, $Rows = @())

    $key = ConvertTo-TraceMessageId -Value $MessageId
    $hops = New-Object System.Collections.ArrayList

    foreach ($r in @($Rows | Where-Object { $_.MessageId -eq $key })) {
        if ($r.Source -eq 'On-premises' -and @($r.Events).Count -gt 0) {
            foreach ($e in @($r.Events)) {
                [void]$hops.Add([PSCustomObject]@{
                    TimestampUtc = $e.TimestampUtc
                    Source       = 'On-premises'
                    Server       = $e.Server
                    Event        = $e.EventId
                    Outcome      = (Get-OnPremOutcome -EventId $e.EventId)
                    Recipient    = $e.Recipient
                    Detail       = $e.Detail
                })
            }
        } else {
            [void]$hops.Add([PSCustomObject]@{
                TimestampUtc = $r.TimestampUtc
                Source       = $r.Source
                Server       = $r.Server
                Event        = $r.RawEvent
                Outcome      = $r.Outcome
                Recipient    = $r.Recipient
                Detail       = $r.Detail
            })
        }
    }
    return @($hops.ToArray() | Sort-Object TimestampUtc)
}
#endregion

#region UI ----------------------------------------------------------------------
$script:TraceTimeZoneIds = [ordered]@{
    'Eastern'  = 'Eastern Standard Time'
    'Central'  = 'Central Standard Time'
    'Mountain' = 'Mountain Standard Time'
    'Pacific'  = 'Pacific Standard Time'
    'UTC'      = 'UTC'
}

$script:AllRows   = @()     # everything the last search returned
$script:LastSearch = ''
$script:EmsState  = $null

function Get-SelectedTimeZone {
    param([string]$Label)
    $id = $script:TraceTimeZoneIds[$Label]
    if (-not $id) { $id = 'UTC' }
    try { return [System.TimeZoneInfo]::FindSystemTimeZoneById($id) }
    catch { return [System.TimeZoneInfo]::Utc }
}

function ConvertTo-DisplayTime {
    <#
        Both sources are already UTC by the time they reach here --
        ConvertTo-TraceUtc saw to that -- so this is one conversion, applied
        identically to both. That is the whole reason the normalised row keeps
        TimestampUtc rather than a display string.
    #>
    param($Utc, [System.TimeZoneInfo]$TimeZone)
    if (-not $Utc) { return $null }
    try { return [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]$Utc, $TimeZone) }
    catch { return $Utc }
}

$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Hybrid Message Trace'
$form.ClientSize      = New-Object System.Drawing.Size(1280, 820)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'
$form.MinimumSize     = New-Object System.Drawing.Size(1000, 640)

# ---- search criteria ----
$grpFind          = New-Object System.Windows.Forms.GroupBox
$grpFind.Text     = 'Search'
$grpFind.Location = New-Object System.Drawing.Point(14, 12)
$grpFind.Size     = New-Object System.Drawing.Size(1200, 180)   # 1200, not 1252: Add-ToolHelpButton owns the top-right corner.
$grpFind.Anchor   = 'Top,Left,Right'
[void]$form.Controls.Add($grpFind)

$lblFrom          = New-Object System.Windows.Forms.Label
$lblFrom.Text     = 'Sender'
$lblFrom.Location = New-Object System.Drawing.Point(12, 28)
$lblFrom.Size     = New-Object System.Drawing.Size(70, 20)
[void]$grpFind.Controls.Add($lblFrom)

$txtFrom          = New-Object System.Windows.Forms.TextBox
$txtFrom.Location = New-Object System.Drawing.Point(88, 25)
$txtFrom.Size     = New-Object System.Drawing.Size(250, 22)
[void]$grpFind.Controls.Add($txtFrom)

$lblTo          = New-Object System.Windows.Forms.Label
$lblTo.Text     = 'Recipient'
$lblTo.Location = New-Object System.Drawing.Point(352, 28)
$lblTo.Size     = New-Object System.Drawing.Size(70, 20)
[void]$grpFind.Controls.Add($lblTo)

$txtTo          = New-Object System.Windows.Forms.TextBox
$txtTo.Location = New-Object System.Drawing.Point(428, 25)
$txtTo.Size     = New-Object System.Drawing.Size(250, 22)
[void]$grpFind.Controls.Add($txtTo)

$lblMid          = New-Object System.Windows.Forms.Label
$lblMid.Text     = 'Message ID'
$lblMid.Location = New-Object System.Drawing.Point(692, 28)
$lblMid.Size     = New-Object System.Drawing.Size(76, 20)
[void]$grpFind.Controls.Add($lblMid)

$txtMid          = New-Object System.Windows.Forms.TextBox
$txtMid.Location = New-Object System.Drawing.Point(774, 25)
$txtMid.Size     = New-Object System.Drawing.Size(300, 22)
[void]$grpFind.Controls.Add($txtMid)

$lblMidHint          = New-Object System.Windows.Forms.Label
$lblMidHint.Text     = 'Angle brackets optional.'
$lblMidHint.Location = New-Object System.Drawing.Point(1082, 28)
$lblMidHint.Size     = New-Object System.Drawing.Size(106, 20)
$lblMidHint.Font     = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grpFind.Controls.Add($lblMidHint)

$lblStart          = New-Object System.Windows.Forms.Label
$lblStart.Text     = 'From'
$lblStart.Location = New-Object System.Drawing.Point(12, 60)
$lblStart.Size     = New-Object System.Drawing.Size(70, 20)
[void]$grpFind.Controls.Add($lblStart)

$dtStart          = New-Object System.Windows.Forms.DateTimePicker
$dtStart.Format   = 'Custom'
$dtStart.CustomFormat = 'yyyy-MM-dd HH:mm'
$dtStart.Location = New-Object System.Drawing.Point(88, 57)
$dtStart.Size     = New-Object System.Drawing.Size(180, 22)
$dtStart.Value    = (Get-Date).AddDays(-1)
[void]$grpFind.Controls.Add($dtStart)

$lblEnd          = New-Object System.Windows.Forms.Label
$lblEnd.Text     = 'To'
$lblEnd.Location = New-Object System.Drawing.Point(282, 60)
$lblEnd.Size     = New-Object System.Drawing.Size(30, 20)
[void]$grpFind.Controls.Add($lblEnd)

$dtEnd          = New-Object System.Windows.Forms.DateTimePicker
$dtEnd.Format   = 'Custom'
$dtEnd.CustomFormat = 'yyyy-MM-dd HH:mm'
$dtEnd.Location = New-Object System.Drawing.Point(318, 57)
$dtEnd.Size     = New-Object System.Drawing.Size(180, 22)
$dtEnd.Value    = (Get-Date)
[void]$grpFind.Controls.Add($dtEnd)

# ---- the source control ----
# ---- range guard controls ----------------------------------------------------
$lblRange           = New-Object System.Windows.Forms.Label
$lblRange.Location  = New-Object System.Drawing.Point(12, 132)
$lblRange.Size      = New-Object System.Drawing.Size(430, 18)
$lblRange.ForeColor = [System.Drawing.Color]::Gray
[void]$grpFind.Controls.Add($lblRange)

# BACK IS FIRST AND ON THE LEFT, because stepping into the past is the direction
# that actually gets used -- forward only exists to undo an overshoot.
$btnBack           = New-Object System.Windows.Forms.Button
$btnBack.Text      = '< Back'
$btnBack.Location  = New-Object System.Drawing.Point(452, 128)
$btnBack.Size      = New-Object System.Drawing.Size(74, 24)
$btnBack.FlatStyle = 'Flat'
[void]$grpFind.Controls.Add($btnBack)

$btnForward           = New-Object System.Windows.Forms.Button
$btnForward.Text      = 'Forward >'
$btnForward.Location  = New-Object System.Drawing.Point(530, 128)
$btnForward.Size      = New-Object System.Drawing.Size(84, 24)
$btnForward.FlatStyle = 'Flat'
[void]$grpFind.Controls.Add($btnForward)

# TWO BUTTONS, NOT A DROPDOWN. The choice between them is a real decision with a
# hundredfold difference in coverage, and a dropdown defaulting to one would make
# it look like a setting rather than a choice.
$btnEnhanced           = New-Object System.Windows.Forms.Button
$btnEnhanced.Text      = '90 day Enhanced'
$btnEnhanced.Location  = New-Object System.Drawing.Point(622, 128)
$btnEnhanced.Size      = New-Object System.Drawing.Size(124, 24)
$btnEnhanced.FlatStyle = 'Flat'
$btnEnhanced.Enabled   = $false
[void]$grpFind.Controls.Add($btnEnhanced)

$btnExtended           = New-Object System.Windows.Forms.Button
$btnExtended.Text      = '90 day Extended'
$btnExtended.Location  = New-Object System.Drawing.Point(750, 128)
$btnExtended.Size      = New-Object System.Drawing.Size(124, 24)
$btnExtended.FlatStyle = 'Flat'
$btnExtended.Enabled   = $false
[void]$grpFind.Controls.Add($btnExtended)

$lblHist           = New-Object System.Windows.Forms.Label
$lblHist.Location  = New-Object System.Drawing.Point(880, 126)
$lblHist.Size      = New-Object System.Drawing.Size(306, 30)
$lblHist.ForeColor = [System.Drawing.Color]::Gray
$lblHist.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
$lblHist.Text      = 'Emailed as CSV, not shown here. Enhanced = 100,000 rows. Extended = full routing detail, only 1,000 rows.'
[void]$grpFind.Controls.Add($lblHist)
$lblSrc          = New-Object System.Windows.Forms.Label
$lblSrc.Text     = 'Search'
$lblSrc.Location = New-Object System.Drawing.Point(512, 60)
$lblSrc.Size     = New-Object System.Drawing.Size(48, 20)
[void]$grpFind.Controls.Add($lblSrc)

# DEFAULTS TO EXCHANGE ONLINE. The on-premises half fans out across every
# transport server and is the slow one; making it opt-in keeps the common case
# as quick as the cloud-only tool.
$rbExo          = New-Object System.Windows.Forms.RadioButton
$rbExo.Text     = 'Exchange Online'
$rbExo.Location = New-Object System.Drawing.Point(566, 58)
$rbExo.Size     = New-Object System.Drawing.Size(130, 22)
$rbExo.Checked  = $true
[void]$grpFind.Controls.Add($rbExo)

$rbOnPrem          = New-Object System.Windows.Forms.RadioButton
$rbOnPrem.Text     = 'On-premises'
$rbOnPrem.Location = New-Object System.Drawing.Point(700, 58)
$rbOnPrem.Size     = New-Object System.Drawing.Size(110, 22)
[void]$grpFind.Controls.Add($rbOnPrem)

$rbBoth          = New-Object System.Windows.Forms.RadioButton
$rbBoth.Text     = 'Both'
$rbBoth.Location = New-Object System.Drawing.Point(814, 58)
$rbBoth.Size     = New-Object System.Drawing.Size(70, 22)
[void]$grpFind.Controls.Add($rbBoth)

$tips = New-Object System.Windows.Forms.ToolTip

$btnSearch           = New-Object System.Windows.Forms.Button
$btnSearch.Text      = 'Search'
$btnSearch.Location  = New-Object System.Drawing.Point(12, 92)
$btnSearch.Size      = New-Object System.Drawing.Size(120, 28)
$btnSearch.FlatStyle = 'Flat'
[void]$grpFind.Controls.Add($btnSearch)

$lblTz          = New-Object System.Windows.Forms.Label
$lblTz.Text     = 'Times in'
$lblTz.Location = New-Object System.Drawing.Point(148, 97)
$lblTz.Size     = New-Object System.Drawing.Size(60, 20)
[void]$grpFind.Controls.Add($lblTz)

$cboTz               = New-Object System.Windows.Forms.ComboBox
$cboTz.DropDownStyle = 'DropDownList'
$cboTz.Location      = New-Object System.Drawing.Point(212, 94)
$cboTz.Size          = New-Object System.Drawing.Size(120, 22)
foreach ($k in $script:TraceTimeZoneIds.Keys) { [void]$cboTz.Items.Add($k) }
$cboTz.SelectedIndex = 0
[void]$grpFind.Controls.Add($cboTz)

$lblFilter          = New-Object System.Windows.Forms.Label
$lblFilter.Text     = 'Outcome'
$lblFilter.Location = New-Object System.Drawing.Point(348, 97)
$lblFilter.Size     = New-Object System.Drawing.Size(58, 20)
[void]$grpFind.Controls.Add($lblFilter)

$cboFilter               = New-Object System.Windows.Forms.ComboBox
$cboFilter.DropDownStyle = 'DropDownList'
$cboFilter.Location      = New-Object System.Drawing.Point(410, 94)
$cboFilter.Size          = New-Object System.Drawing.Size(180, 22)
[void]$cboFilter.Items.Add('All')
$cboFilter.SelectedIndex = 0
[void]$grpFind.Controls.Add($cboFilter)

$lblFilterHint          = New-Object System.Windows.Forms.Label
$lblFilterHint.Text     = 'Filters what is already loaded -- does not re-run the search.'
$lblFilterHint.Location = New-Object System.Drawing.Point(600, 90)
# TWO LINES, AND TALL ENOUGH FOR BOTH. It was 268x20 at 8pt holding a sentence
# that needs about 330px, so the second half was simply never drawn -- and a
# clipped label gives no hint that there is more to read.
$lblFilterHint.Size      = New-Object System.Drawing.Size(268, 32)
$lblFilterHint.Font     = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grpFind.Controls.Add($lblFilterHint)

$btnExpandAll           = New-Object System.Windows.Forms.Button
$btnExpandAll.Text      = 'Expand all'
$btnExpandAll.Location  = New-Object System.Drawing.Point(876, 92)
$btnExpandAll.Size      = New-Object System.Drawing.Size(92, 26)
$btnExpandAll.FlatStyle = 'Flat'
[void]$grpFind.Controls.Add($btnExpandAll)

$btnCollapseAll           = New-Object System.Windows.Forms.Button
$btnCollapseAll.Text      = 'Collapse all'
$btnCollapseAll.Location  = New-Object System.Drawing.Point(972, 92)
$btnCollapseAll.Size      = New-Object System.Drawing.Size(92, 26)
$btnCollapseAll.FlatStyle = 'Flat'
[void]$grpFind.Controls.Add($btnCollapseAll)
$btnExport           = New-Object System.Windows.Forms.Button
$btnExport.Text      = 'Export CSV'
$btnExport.Location  = New-Object System.Drawing.Point(1070, 92)
$btnExport.Size      = New-Object System.Drawing.Size(118, 28)
$btnExport.FlatStyle = 'Flat'
[void]$grpFind.Controls.Add($btnExport)

# ---- THE HONESTY LINE (D6) ----
# Above the grid, not below it and not in a tooltip. If rows are missing this is
# the only thing standing between the operator and "the message never existed".
$lblPartial          = New-Object System.Windows.Forms.Label
$lblPartial.Location = New-Object System.Drawing.Point(14, 198)
$lblPartial.Size     = New-Object System.Drawing.Size(1252, 20)
$lblPartial.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblPartial.Anchor   = 'Top,Left,Right'
[void]$form.Controls.Add($lblPartial)

$pb = Add-ReportProgressBar -Form $form -X 14 -Y 172 -Width 1252 -Height 12

$grid                     = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(14, 238)
$grid.Size                = New-Object System.Drawing.Size(1252, 372)
$grid.ReadOnly            = $true
$grid.AllowUserToAddRows  = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode       = 'FullRowSelect'
$grid.MultiSelect         = $false
$grid.Anchor              = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($grid)

$lblHint          = New-Object System.Windows.Forms.Label
$lblHint.Text     = 'Double-click a row for the hop timeline across both systems.'
$lblHint.Location = New-Object System.Drawing.Point(14, 618)
$lblHint.Size     = New-Object System.Drawing.Size(500, 20)
$lblHint.Anchor   = 'Left,Bottom'
[void]$form.Controls.Add($lblHint)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(14, 642)
$txtOut.Size       = New-Object System.Drawing.Size(1252, 164)
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
# B6.
$txtOut.WordWrap   = $true
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

function Get-GroupSortKey {
    <#
        The value a group sorts on for a given column.

        TAKEN FROM THE VERDICT ROW, which is the line actually on screen. Sorting
        on any other row in the group would order the grid by something the
        operator cannot see.

        SIZE SORTS NUMERICALLY. It is rendered with thousands separators, so a
        text sort puts 9,000 after 10,000 -- and the whole reason to sort by size
        is to find the big one.
    #>
    param($Verdict, [string]$Column, $Rows)

    switch ($Column) {
        'Time'      { return $Verdict.TimestampUtc }
        'Size'      { return [long]$Verdict.SizeBytes }
        'Source'    { return (@($Rows | ForEach-Object { "$($_.Source)" } | Sort-Object -Unique) -join ' + ') }
        'Sender'    { return "$($Verdict.Sender)" }
        'Recipient' { return "$($Verdict.Recipient)" }
        'Subject'   { return "$($Verdict.Subject)" }
        'Outcome'   { return "$($Verdict.Outcome)" }
        'Server'    { return "$($Verdict.Server)" }
    }
    return $Verdict.TimestampUtc
}
function Update-Grid {
    <#
        ONE LINE PER MESSAGE, collapsed, showing the verdict.

        Redraws from $script:AllRows. Filtering and the timezone dropdown both
        come through here, so neither re-runs the search -- which matters on the
        on-premises side where a search is a fan-out across every server.

        WHY COLLAPSING WAS NEEDED. Merge-TraceRows deliberately does NOT
        de-duplicate across sources: a message that crossed the hybrid boundary
        genuinely produces an on-premises row AND an Exchange Online row, and
        both are true. Shown flat, every hybrid message looks like two messages
        -- and the operator has to know the topology to see that it is not.

        The rows are still all there. They are one keystroke away instead of
        being the default view.

        ONLY A GROUP WITH MORE THAN ONE ROW EXPANDS. A message seen by one side
        only is already fully described by its single line, and an expander that
        opens to restate it is a control that does nothing.
    #>
    $tz   = Get-SelectedTimeZone -Label "$($cboTz.SelectedItem)"
    $want = "$($cboFilter.SelectedItem)" -replace '\s+\(\d+\)\s*$', ''

    $groups = @($script:AllRows | Group-Object { Get-TraceGroupKey -Row $_ })

    # THE FILTER TESTS THE GROUP'S VERDICT, not each row.
    #
    # Testing rows would show a group whose winning outcome is Quarantined under
    # a filter of 'Relayed', because one of its hops relayed. The grid claims to
    # be one line per message, so the filter has to mean one line per message
    # too.
    if ($want -and $want -ne 'All') {
        $groups = @($groups | Where-Object { "$((Get-GroupVerdict -Rows $_.Group).Outcome)" -eq $want })
    }

    # SORTED BY GROUP, so a child never leaves its parent.
    $groups = @($groups | Sort-Object -Descending:(-not $script:SortAsc) `
        { Get-GroupSortKey -Verdict (Get-GroupVerdict -Rows $_.Group) -Column $script:SortCol -Rows $_.Group })

    $t = New-Object System.Data.DataTable
    foreach ($c in @(' ', 'Time', 'Source', 'Sender', 'Recipient', 'Subject', 'Outcome', 'Server', 'Size')) {
        [void]$t.Columns.Add($c)
    }

    $flat     = New-Object System.Collections.ArrayList
    $groupOf  = New-Object System.Collections.ArrayList

    foreach ($grp in $groups) {
        $rows      = @($grp.Group)
        $verdict   = Get-GroupVerdict -Rows $rows
        $canExpand = (@($rows).Count -gt 1)
        $open      = ($script:Expanded.ContainsKey("$($grp.Name)") -and $canExpand)

        $marker = if (-not $canExpand) { '   ' } elseif ($open) { '[-]' } else { '[+]' }

        $srcTxt = "$($verdict.Source)"
        if ($canExpand) {
            # NAME BOTH SIDES ON THE COLLAPSED LINE. 'Exchange Online (2)' does
            # not say the other row is on-premises, and "did it reach the cloud"
            # is most of what this tool is asked.
            $srcTxt = (@($rows | ForEach-Object { "$($_.Source)" } | Sort-Object -Unique) -join ' + ')
        }

        $lt = ConvertTo-DisplayTime -Utc $verdict.TimestampUtc -TimeZone $tz
        [void]$t.Rows.Add(
            $marker,
            $(if ($lt) { $lt.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }),
            $srcTxt, $verdict.Sender, $verdict.Recipient, $verdict.Subject, $verdict.Outcome,
            $verdict.Server,
            $(if ($verdict.SizeBytes) { '{0:N0}' -f $verdict.SizeBytes } else { '' }))
        [void]$flat.Add($verdict)
        [void]$groupOf.Add("$($grp.Name)")

        if (-not $open) { continue }

        foreach ($r in @($rows | Sort-Object TimestampUtc)) {
            $rt = ConvertTo-DisplayTime -Utc $r.TimestampUtc -TimeZone $tz
            [void]$t.Rows.Add(
                '',
                $(if ($rt) { '   ' + $rt.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }),
                "   $($r.Source)", '', '', '', $r.Outcome, $r.Server,
                $(if ($r.SizeBytes) { '{0:N0}' -f $r.SizeBytes } else { '' }))
            [void]$flat.Add($r)
            [void]$groupOf.Add("$($grp.Name)")

            # ON-PREMISES EVENTS UNDER THEIR OWN ROW. They are the hop detail
            # the cloud has no equivalent of, and burying them behind a second
            # click when the row is already open is a click for nothing.
            foreach ($ev in @($r.Events | Sort-Object TimestampUtc)) {
                $et = ConvertTo-DisplayTime -Utc $ev.TimestampUtc -TimeZone $tz
                [void]$t.Rows.Add(
                    '',
                    $(if ($et) { '      ' + $et.ToString('HH:mm:ss') } else { '' }),
                    "      $($ev.EventId)", '', '', '',
                    (Get-OnPremOutcome -EventId "$($ev.EventId)"), $ev.Server, '')
                [void]$flat.Add($r)
                [void]$groupOf.Add("$($grp.Name)")
            }
        }
    }

    # SCROLL POSITION IS CAPTURED BEFORE THE SWAP.
    #
    # Assigning DataSource throws it away, so expanding a message at row 80
    # dropped the operator back at row 0 -- on a long trace that loses the
    # message they were reading and they have to find it again.
    $firstVisible = 0
    if ($grid.FirstDisplayedScrollingRowIndex -gt 0) { $firstVisible = $grid.FirstDisplayedScrollingRowIndex }

    $grid.DataSource   = $t
    $grid.Tag          = @($flat.ToArray())
    # WHICH GROUP EACH DISPLAYED LINE BELONGS TO. DataGridView has no second
    # Tag, and re-deriving it from the row text would break on a subject
    # containing the separator. Built in the same loop so it cannot drift.
    $script:RowGroup   = @($groupOf.ToArray())

    if ($grid.Columns.Contains(' ')) { $grid.Columns[' '].Width = 34 }

    # THE GRID'S OWN SORT IS DISABLED ON EVERY COLUMN. See the note by
    # $script:SortCol -- it would scatter children away from their parents.
    # The header click is handled instead and re-sorts the groups.
    foreach ($col in $grid.Columns) {
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
        $col.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None
        if ("$($col.Name)" -eq $script:SortCol) {
            $col.HeaderCell.SortGlyphDirection = $(if ($script:SortAsc) {
                [System.Windows.Forms.SortOrder]::Ascending } else { [System.Windows.Forms.SortOrder]::Descending })
        }
    }

    # Put the view back where it was.
    if ($firstVisible -gt 0 -and $grid.Rows.Count -gt $firstVisible) {
        try { $grid.FirstDisplayedScrollingRowIndex = $firstVisible } catch { }
    }
}

function Toggle-TraceGroup {
    <#
        Open or close the message on the clicked line.

        A BLANK MARKER MEANS DO NOTHING, and it now means two things at once: a
        child line has no marker, and neither does a single-row message. Both
        must ignore the click, and the empty cell says so for both.
    #>
    param([int]$Index)

    if ($Index -lt 0 -or $Index -ge @($script:RowGroup).Count) { return }
    $marker = "$($grid.Rows[$Index].Cells[' '].Value)".Trim()
    if (-not $marker) { return }

    $key = "$(@($script:RowGroup)[$Index])"
    if (-not $key) { return }

    if ($script:Expanded.ContainsKey($key)) { $script:Expanded.Remove($key) }
    else { $script:Expanded[$key] = $true }

    Update-Grid

    # PUT THE SELECTION BACK ON THE SAME MESSAGE. Without this the grid jumps to
    # row 0 on every expand, and the row somebody was about to double-click for
    # detail is no longer the row under the cursor.
    for ($i = 0; $i -lt @($script:RowGroup).Count; $i++) {
        if ("$(@($script:RowGroup)[$i])" -eq $key -and "$($grid.Rows[$i].Cells[' '].Value)".Trim()) {
            $grid.ClearSelection()
            $grid.Rows[$i].Selected = $true
            $grid.CurrentCell = $grid.Rows[$i].Cells['Subject']
            break
        }
    }
}

function Set-AllTraceExpanded {
    param([bool]$Open)
    $script:Expanded = @{}
    if ($Open) {
        # MULTI-ROW GROUPS ONLY, or a single-row message renders children under
        # a line with no marker to close them.
        foreach ($grp in @($script:AllRows | Group-Object { Get-TraceGroupKey -Row $_ })) {
            if (@($grp.Group).Count -le 1) { continue }
            $script:Expanded["$($grp.Name)"] = $true
        }
    }
    Update-Grid
}

#region Transport rules ---------------------------------------------------------
<#
    WHICH MAIL FLOW RULES TOUCHED THIS MESSAGE, AND WHAT THEY DID TO IT.

    This is the question a trace almost never answers and the one that ends the
    investigation. "Delivered to Junk" and "quarantined" and "went to the wrong
    connector" are all usually a rule somebody wrote two years ago, and finding
    it today means reading the whole rule list and guessing which one matched.

    THE TWO SIDES REPORT IT COMPLETELY DIFFERENTLY, which is why this is not one
    parser:

      Exchange Online   Get-MessageTraceDetailV2 returns an event whose Event is
                        'Transport rule' and whose Detail reads
                            Transport rule: 'NAME', ID: ('GUID'), DLP policy: ...
      On-premises       Get-MessageTrackingLog puts it in EventData on the
                        AGENTINFO event, as
                            S:TRA=ETR|ruleId=GUID|st=...|action=X|sev=1|mode=Enforce

    Both give a GUID, which is what Get-TransportRule is then asked for. The
    NAME is deliberately not used as the lookup key: rules get renamed, and a
    trace from last week naming a rule that has since been renamed would fail to
    resolve and report nothing rather than reporting the rename.
#>

$script:RuleCache = @{}

function Get-TransportRuleFact {
    <#
        One rule, looked up once per session.

        CACHED BECAUSE A SINGLE TRACE HITS THE SAME RULE REPEATEDLY -- a
        disclaimer rule touches every message in the grid, and asking Exchange
        per row turns a detail pane into a round trip per hop.

        A MISS IS CACHED TOO. A rule that has been DELETED since it acted on the
        message will never resolve, and retrying it for every row is the same
        wasted call over and over. The cache holds the failure so the pane can
        say "deleted since" instead.
    #>
    param([Parameter(Mandatory)][string]$RuleId, [string]$NameHint = '')

    $key = "$RuleId".Trim().ToLowerInvariant()
    if (-not $key) { return $null }
    if ($script:RuleCache.ContainsKey($key)) { return $script:RuleCache[$key] }

    $fact = [PSCustomObject]@{
        Id          = $RuleId
        Name        = $(if ($NameHint) { $NameHint } else { '(unnamed)' })
        Found       = $false
        State       = ''
        Mode        = ''
        Priority    = ''
        Description = ''
        Comments    = ''
        Reason      = ''
    }

    try {
        $r = Invoke-EXOCommand -Name 'Get-TransportRule' -Parameters @{
            Identity = $RuleId; ErrorAction = 'Stop' }
        if ($r) {
            $fact.Found       = $true
            $fact.Name        = "$($r.Name)"
            $fact.State       = "$($r.State)"
            $fact.Mode        = "$($r.Mode)"
            $fact.Priority    = "$($r.Priority)"
            # EXCHANGE WRITES THIS SENTENCE ITSELF, and it is better than
            # anything this tool could assemble from the condition and action
            # properties: "If the message: ... Take the following actions: ...".
            # It is generated from the live rule, so it cannot drift from what
            # the rule actually does.
            $fact.Description = "$($r.Description)"
            $fact.Comments    = "$($r.Comments)"
        }
    } catch {
        $fact.Reason = "$($_.Exception.Message)"
    }

    $script:RuleCache[$key] = $fact
    return $fact
}

function Get-RuleHitFromExoDetail {
    <#
        Rule hits out of the Exchange Online detail events.

        THE DLP POLICY IS IN THE SAME STRING and is deliberately picked out
        separately. A message stopped by DLP looks identical to one stopped by a
        plain rule until you read the second half of the line, and 'blocked by a
        rule' versus 'blocked by a data loss prevention policy' are different
        conversations with different people.
    #>
    param($DetailEvents)

    $hits = New-Object System.Collections.ArrayList
    foreach ($ev in @($DetailEvents)) {
        $txt = "$($ev.Detail)"
        if ("$($ev.Event)" -notmatch '(?i)transport rule|rule match|DLP') {
            if ($txt -notmatch '(?i)transport rule:') { continue }
        }

        $name = ''
        $id   = ''
        if ($txt -match "(?i)Transport rule:\s*'([^']*)'") { $name = $Matches[1] }
        if ($txt -match "(?i)Transport rule:[^(]*\(\s*'?\{?([0-9A-Fa-f-]{36})\}?'?\s*\)") { $id = $Matches[1] }

        $dlp   = ''
        $dlpId = ''
        if ($txt -match "(?i)DLP policy:\s*'([^']*)'") { $dlp = $Matches[1] }
        if ($txt -match "(?i)DLP policy:[^(]*\(\s*'?\{?([0-9A-Fa-f-]{36})\}?'?\s*\)") { $dlpId = $Matches[1] }
        # The all-zero GUID is what Exchange writes when there is no DLP policy.
        if ($dlpId -match '^0{8}-0{4}-0{4}-0{4}-0{12}$') { $dlpId = ''; $dlp = '' }

        if (-not $id -and -not $name) { continue }

        [void]$hits.Add([PSCustomObject]@{
            Source    = 'Exchange Online'
            RuleId    = $id
            RuleName  = $name
            DlpPolicy = $dlp
            DlpId     = $dlpId
            Action    = ''
            Mode      = ''
            When      = $ev.Date
            Raw       = $txt
        })
    }
    return @($hits.ToArray())
}

function Get-RuleHitFromOnPremEvent {
    <#
        Rule hits out of the on-premises tracking log.

        THE FORMAT IS AN AGENT RECORD, NOT PROSE:
            S:TRA=ETR|ruleId=<guid>|st=<utc>|action=<action>|sev=<n>|mode=<mode>

        ONE EventData ENTRY CAN CARRY SEVERAL RULES, pipe-separated within the
        same S:TRA value when more than one matched at the same stage. Splitting
        on 'ETR|' rather than parsing one match per line is what catches the
        second and third.

        THE MODE MATTERS MORE THAN IT LOOKS. A rule in 'Test without Policy
        Tips' or 'Test with Policy Tips' APPEARS IN THE LOG AS A MATCH AND DOES
        NOTHING. An operator who finds a matching rule and stops reading will
        spend the afternoon disabling a rule that was never acting.
    #>
    param($Events)

    $hits = New-Object System.Collections.ArrayList
    foreach ($e in @($Events)) {
        foreach ($blob in @($e.EventData)) {
            $s = "$blob"
            if ($s -notmatch '(?i)S:TRA=ETR') { continue }

            foreach ($chunk in ($s -split '(?i)S:TRA=ETR\|')) {
                if ($chunk -notmatch '(?i)ruleId=') { continue }

                $id = ''; $act = ''; $mode = ''; $st = ''
                if ($chunk -match '(?i)ruleId=\{?([0-9A-Fa-f-]{36})\}?') { $id   = $Matches[1] }
                if ($chunk -match '(?i)action=([^|]+)')                  { $act  = $Matches[1].Trim() }
                if ($chunk -match '(?i)mode=([^|]+)')                    { $mode = $Matches[1].Trim() }
                if ($chunk -match '(?i)st=([^|]+)')                      { $st   = $Matches[1].Trim() }
                if (-not $id) { continue }

                [void]$hits.Add([PSCustomObject]@{
                    Source    = 'On-premises'
                    RuleId    = $id
                    RuleName  = ''
                    DlpPolicy = ''
                    DlpId     = ''
                    Action    = $act
                    Mode      = $mode
                    When      = $(if ($st) { $st } else { $e.TimestampUtc })
                    Raw       = $chunk.Trim()
                })
            }
        }
    }
    return @($hits.ToArray())
}

function Format-RuleExplanation {
    <#
        One rule hit, written out so somebody can act on it.

        THE ORDER IS DELIBERATE: what it is, whether it was actually enforcing,
        then what it does. State and Mode come BEFORE the description, because
        they decide whether the description is relevant at all -- a disabled or
        test-mode rule that matched is a red herring, and burying that under
        twelve lines of conditions is how an afternoon gets lost.
    #>
    param([Parameter(Mandatory)]$Hit, $Fact)

    $L = New-Object System.Collections.ArrayList

    $title = if ($Fact -and $Fact.Found) { $Fact.Name } elseif ($Hit.RuleName) { $Hit.RuleName } else { '(rule name unavailable)' }
    [void]$L.Add("RULE: $title")
    [void]$L.Add("      seen by $($Hit.Source)$(if ($Hit.When) { "  at $($Hit.When)" })")
    if ($Hit.RuleId) { [void]$L.Add("      id $($Hit.RuleId)") }

    if ($Hit.Action) { [void]$L.Add("      action recorded in the log: $($Hit.Action)") }

    if (-not $Fact -or -not $Fact.Found) {
        [void]$L.Add('')
        if ($Hit.RuleName) {
            [void]$L.Add("      THE RULE COULD NOT BE READ. It matched this message and no rule with")
            [void]$L.Add("      that id exists now -- it has almost certainly been deleted or")
            [void]$L.Add("      recreated since. The name above is the one recorded at the time.")
        } else {
            [void]$L.Add('      THE RULE COULD NOT BE READ and the log did not record its name.')
        }
        if ($Fact -and $Fact.Reason) { [void]$L.Add("      $($Fact.Reason)") }
        return ($L -join [Environment]::NewLine)
    }

    # WAS IT ACTUALLY DOING ANYTHING? Asked before what it does.
    $live = $true
    if ("$($Fact.State)" -match '(?i)disabled') {
        [void]$L.Add('')
        [void]$L.Add('      DISABLED NOW. It matched when this message passed, so it was enabled')
        [void]$L.Add('      then and has been turned off since -- which may well be why the')
        [void]$L.Add('      problem "fixed itself".')
        $live = $false
    }
    if ("$($Fact.Mode)" -match '(?i)audit|test') {
        [void]$L.Add('')
        [void]$L.Add("      MODE IS '$($Fact.Mode)' -- IT MATCHED BUT DID NOT ACT.")
        [void]$L.Add('      A test-mode rule appears in the log exactly like an enforcing one and')
        [void]$L.Add('      changes nothing. This rule is not your cause. Keep looking.')
        $live = $false
    }
    if ($Hit.Mode -and $Hit.Mode -notmatch '(?i)enforce') {
        [void]$L.Add('')
        [void]$L.Add("      THE LOG RECORDED MODE '$($Hit.Mode)' AT THE TIME, so it was not")
        [void]$L.Add('      enforcing when it touched this message even if it is enforcing now.')
        $live = $false
    }
    if ($live) {
        [void]$L.Add("      State: $($Fact.State)   Mode: $($Fact.Mode)   Priority: $($Fact.Priority)")
        [void]$L.Add('      ENFORCING -- this rule acted on the message.')
    }

    if ($Hit.DlpPolicy) {
        [void]$L.Add('')
        [void]$L.Add("      PART OF DLP POLICY '$($Hit.DlpPolicy)'. Changing the rule directly is")
        [void]$L.Add('      usually the wrong move -- it belongs to a policy and the policy will')
        [void]$L.Add('      put it back. Take it up with whoever owns the policy.')
    }

    if ($Fact.Description) {
        [void]$L.Add('')
        [void]$L.Add('      WHAT IT DOES  (Exchange generates this from the live rule):')
        foreach ($line in ("$($Fact.Description)" -split "`r?`n")) {
            $s = "$line".Trim()
            if (-not $s) { continue }
            while ($s.Length -gt 78) {
                $cut = $s.LastIndexOf(' ', 78)
                if ($cut -lt 30) { $cut = 78 }
                [void]$L.Add("          $($s.Substring(0, $cut))")
                $s = $s.Substring($cut).TrimStart()
            }
            [void]$L.Add("          $s")
        }
    }

    if ($Fact.Comments) {
        [void]$L.Add('')
        [void]$L.Add("      NOTE LEFT ON THE RULE: $($Fact.Comments)")
    }

    return ($L -join [Environment]::NewLine)
}

function Get-MessageRuleReport {
    <#
        Every rule that touched one message, from both sides, explained.

        DEDUPLICATED ON RULE ID PER SOURCE. A rule that fires on several hops
        appears once per hop in the raw data, and listing it four times pushes
        the other rules off the screen.
    #>
    param($ExoDetailEvents, $OnPremEvents)

    $hits = New-Object System.Collections.ArrayList
    foreach ($h in @(Get-RuleHitFromExoDetail   -DetailEvents $ExoDetailEvents)) { [void]$hits.Add($h) }
    foreach ($h in @(Get-RuleHitFromOnPremEvent -Events       $OnPremEvents))    { [void]$hits.Add($h) }

    if (@($hits).Count -eq 0) {
        return [PSCustomObject]@{ Count = 0; Text = 'No mail flow rule touched this message on either side.' }
    }

    $seen = @{}
    $out  = New-Object System.Collections.ArrayList
    foreach ($h in $hits) {
        $k = "$($h.Source)|$("$($h.RuleId)".ToLowerInvariant())"
        if ($seen.ContainsKey($k)) { continue }
        $seen[$k] = $true

        $fact = $null
        if ($h.RuleId) { $fact = Get-TransportRuleFact -RuleId $h.RuleId -NameHint $h.RuleName }
        [void]$out.Add((Format-RuleExplanation -Hit $h -Fact $fact))
    }

    return [PSCustomObject]@{
        Count = @($out).Count
        Text  = (@($out) -join ([Environment]::NewLine + [Environment]::NewLine))
    }
}
#endregion

function Get-KnownOutcome {
    <#
        Every outcome this tool can produce, whether or not the current search
        found one.

        THE FILTER USED TO BE BUILT ONLY FROM THE RESULTS, and that reads as the
        tool having no such option. A search returning nothing but Delivered
        offered 'All' and 'Delivered' -- so "are any of this person's messages
        quarantined?" could not be asked, only stumbled upon. Listing the whole
        set with a count against each turns the filter into an answer in itself:
        Quarantined (0) is a fact, and an absent Quarantined entry is not.

        ORDERED BY WHAT AN OPERATOR IS HUNTING FOR, not alphabetically. Nobody
        opens a trace hoping to find Delivered.
    #>
    return @(
        'Failed'
        'Quarantined'
        'FilteredAsSpam'
        'Pending'
        'Deferred'
        'Bounced'
        'Dropped'
        'Rejected by moderator'
        'Moderation expired'
        'Redirected'
        'Expanded'
        'Relayed'
        'Delivered'
        'Delivered (duplicate)'
        'GettingStatus'
        'None'
    )
}

function Update-OutcomeFilter {
    <#
        Rebuild the outcome list: the full set, each with how many rows carry it,
        plus anything the search returned that this build has never heard of.

        THE UNKNOWN BUCKET MATTERS. A future Exchange status would otherwise be
        invisible in the filter while sitting in the grid, and the operator would
        have no way to isolate it.

        THE SELECTION IS PRESERVED across searches. Filtering to Quarantined,
        widening the date range and searching again should keep Quarantined --
        silently resetting to All is how a filtered view gets mistaken for the
        whole answer.
    #>
    $keep = "$($cboFilter.SelectedItem)" -replace '\s+\(\d+\)\s*$', ''

    $counts = @{}
    foreach ($row in @($script:AllRows)) {
        $o = "$($row.Outcome)".Trim()
        if (-not $o) { continue }
        if ($counts.ContainsKey($o)) { $counts[$o]++ } else { $counts[$o] = 1 }
    }

    $cboFilter.Items.Clear()
    [void]$cboFilter.Items.Add("All  ($(@($script:AllRows).Count))")

    foreach ($o in (Get-KnownOutcome)) {
        $seen = 0
        if ($counts.ContainsKey($o)) { $seen = $counts[$o] }
        [void]$cboFilter.Items.Add("$o  ($seen)")
    }

    # Anything returned that is not in the known list, so it can still be picked.
    foreach ($o in @($counts.Keys | Sort-Object)) {
        if ((Get-KnownOutcome) -contains $o) { continue }
        [void]$cboFilter.Items.Add("$o  ($($counts[$o]))")
    }

    $cboFilter.SelectedIndex = 0
    if ($keep -and $keep -ne 'All') {
        for ($i = 0; $i -lt $cboFilter.Items.Count; $i++) {
            if (("$($cboFilter.Items[$i])" -replace '\s+\(\d+\)\s*$', '') -eq $keep) {
                $cboFilter.SelectedIndex = $i
                break
            }
        }
    }
}
function Get-OutcomeMeaning {
    <#
        What a verdict MEANS, and what to do next.

        THE WORD ON ITS OWN IS NOT ENOUGH, and three of these are actively
        misread:

          Delivered      does not mean they can see it. It means Exchange handed
                         it over. Rules and client-side filing happen after.
          Relayed        is an on-premises SEND -- it left this server. It says
                         nothing about whether the next hop accepted it, and it
                         is NOT the same as Delivered.
          GettingStatus  is not a failure. Exchange has not finished writing the
                         record. Asking again in a minute usually answers it.

        Returns lines, not a paragraph, so the caller controls the layout.
    #>
    param([string]$Outcome, $Row)

    $L = New-Object System.Collections.ArrayList
    switch -Regex ("$Outcome") {

        '^Delivered' {
            [void]$L.Add('The message reached the mailbox. Mail flow is not your problem.')
            [void]$L.Add('')
            [void]$L.Add('DELIVERED DOES NOT MEAN THEY CAN SEE IT. Anything that files or hides a')
            [void]$L.Add('message happens after this point:')
            [void]$L.Add('   - an inbox rule -- the usual answer, and the one nobody checks')
            [void]$L.Add('   - the Junk Email folder, if a rule or policy raised the SCL')
            [void]$L.Add('   - a mobile client that already filed or deleted it')
            [void]$L.Add('   - Focused/Other, which hides nothing but is where people stop looking')
            [void]$L.Add('Check the Mail flow rules tab before believing the mailbox is at fault.')
        }

        '^Relayed$' {
            [void]$L.Add('An on-premises SEND: this server passed it on. IT IS NOT A DELIVERY.')
            [void]$L.Add('')
            [void]$L.Add('The next hop may have accepted it, deferred it, or rejected it outright,')
            [void]$L.Add('and none of that is in this row. If the trace stops here, the message')
            [void]$L.Add('left the estate and the answer is on the other side -- the receiving')
            [void]$L.Add('organisation, the smart host, or the cloud connector.')
        }

        '^Quarantined$' {
            [void]$L.Add('HELD. It was not delivered and the recipient has no idea it exists.')
            [void]$L.Add('')
            [void]$L.Add('Review it in the Microsoft Defender portal under Review > Quarantine, or')
            [void]$L.Add('with Get-QuarantineMessage. Release with Release-QuarantineMessage.')
            [void]$L.Add('')
            [void]$L.Add('RELEASING ALONE TEACHES THE FILTER NOTHING -- the next message from the')
            [void]$L.Add('same sender is quarantined identically. Submit it as a false positive at')
            [void]$L.Add('the same time, or you will see this ticket again next week.')
            [void]$L.Add('Check the Mail flow rules tab: a rule setting SCL 9 quarantines mail')
            [void]$L.Add('that the spam filter itself was perfectly happy with.')
        }

        '^FilteredAsSpam$' {
            [void]$L.Add('DELIVERED, to the Junk Email folder. Nothing is lost and nothing is held.')
            [void]$L.Add('')
            [void]$L.Add('Have them mark it Not Junk from inside Outlook, which trains the filter.')
            [void]$L.Add('Dragging it to the Inbox does not.')
            [void]$L.Add('')
            [void]$L.Add('IF THE SENDER IS INTERNAL, THIS IS WORTH CHASING PROPERLY. Internal mail')
            [void]$L.Add('being junked usually means SPF, DKIM or DMARC is wrong on the sending')
            [void]$L.Add('domain, and every recipient is getting the same treatment silently.')
            [void]$L.Add('Check the Mail flow rules tab -- a rule raising the SCL does this too.')
        }

        '^Failed$' {
            [void]$L.Add('PERMANENT FAILURE. It will not arrive and it will not retry.')
            [void]$L.Add('')
            [void]$L.Add('The sender should have received a non-delivery report. The Hops tab')
            [void]$L.Add('carries the status code -- read it before anything else, because it')
            [void]$L.Add('names the cause and a 5.x.x is final while a 4.x.x is not.')
            [void]$L.Add('')
            [void]$L.Add('Common on-premises causes: 5.4.1 relay access denied (the sending device')
            [void]$L.Add('is not in a receive connector scope), 5.1.1 unknown recipient, 5.7.x a')
            [void]$L.Add('rule or connector refusing it.')
        }

        '^(Pending|Deferred)$' {
            [void]$L.Add('STILL MOVING. It has not failed and it is not lost.')
            [void]$L.Add('')
            [void]$L.Add('Something downstream asked for a retry -- greylisting, throttling, or a')
            [void]$L.Add('receiving server that is briefly unavailable. Exchange keeps trying.')
            [void]$L.Add('')
            [void]$L.Add('DO NOT ADVISE A RESEND YET. A second copy gets deferred as well and both')
            [void]$L.Add('eventually arrive. Re-run this trace in an hour; if it is still pending')
            [void]$L.Add('after several, look at the receiving domain rather than at this one.')
        }

        '^GettingStatus$' {
            [void]$L.Add('NOT A VERDICT. Exchange Online has not finished writing the record yet.')
            [void]$L.Add('')
            [void]$L.Add('This is normal for a message from the last few minutes. Run the trace')
            [void]$L.Add('again shortly and it will resolve to a real outcome. Nothing is wrong.')
        }

        '^None$' {
            [void]$L.Add('NO STATUS RECORDED. Exchange Online has the message but no outcome for')
            [void]$L.Add('it, which is usually a very recent message or one still in processing.')
            [void]$L.Add('Re-run the trace before treating it as a fault.')
        }

        '^Expanded$' {
            [void]$L.Add('A DISTRIBUTION GROUP WAS EXPANDED. This row is the group, not a person.')
            [void]$L.Add('')
            [void]$L.Add('The real outcomes are on the member rows. Search again for an individual')
            [void]$L.Add('member to find out whether THEY got it -- an Expanded row says the list')
            [void]$L.Add('was resolved and nothing about delivery.')
        }

        '^Bounced$' {
            [void]$L.Add('The receiving side rejected it and a non-delivery report went back to')
            [void]$L.Add('the sender. The Hops tab has the reason the remote server gave.')
        }

        '^Dropped$' {
            [void]$L.Add('SILENTLY DISCARDED -- no bounce, no NDR, nobody told. Usually a rule or')
            [void]$L.Add('an anti-malware action with a delete rather than a reject.')
            [void]$L.Add('Check the Mail flow rules tab: a rule with a delete action does exactly')
            [void]$L.Add('this and is invisible from every other angle.')
        }

        '^Redirected$' {
            [void]$L.Add('SENT SOMEWHERE ELSE. The recipient on this row is not where it ended up.')
            [void]$L.Add('A forwarding setting or a rule redirected it -- the Mail flow rules tab')
            [void]$L.Add('will name the rule if a rule did it.')
        }

        '^Rejected by moderator$' {
            [void]$L.Add('A moderated group or recipient was rejected by its moderator. Working as')
            [void]$L.Add('configured. The sender was notified.')
        }

        '^Moderation expired$' {
            [void]$L.Add('Nobody moderated it in time and it was discarded. Check the group has a')
            [void]$L.Add('moderator who is still here and still reading -- an unmoderated queue')
            [void]$L.Add('silently discards everything sent to the group.')
        }

        default {
            [void]$L.Add("No specific guidance for '$Outcome' yet. The Hops tab has everything")
            [void]$L.Add('both sides reported, and the Raw tab has it unparsed.')
        }
    }
    return @($L.ToArray())
}
function Get-ExoMessageDetail {
    <#
        The per-event detail Exchange Online holds for one message and one
        recipient.

        BOTH PARAMETERS ARE MANDATORY AND POSITIONAL on the cmdlet, and the
        MessageTraceId is a GUID that only comes from the trace row -- which is
        why New-TraceRow carries TraceGuid. Without it there is no way back to
        this data short of re-running the whole search.

        RETURNS AN OBJECT, NEVER THROWS. The detail is an enrichment: a message
        older than ten days has no detail at all (the cmdlet's own limit,
        against ninety days for the trace itself), and losing the whole pane
        because the enrichment aged out would be worse than showing the trace
        without it.
    #>
    param([string]$TraceGuid, [string]$Recipient)

    $out = [PSCustomObject]@{ Ok = $false; Events = @(); Reason = '' }

    if (-not "$TraceGuid".Trim()) {
        $out.Reason = 'No MessageTraceId on this row -- detail is only available for Exchange Online rows.'
        return $out
    }
    if (-not "$Recipient".Trim()) {
        $out.Reason = 'No recipient on this row, and the detail cmdlet requires one.'
        return $out
    }

    try {
        $out.Events = @(Invoke-EXOCommand -Name 'Get-MessageTraceDetailV2' -Parameters @{
            MessageTraceId   = [guid]"$TraceGuid"
            RecipientAddress = "$Recipient"
            ErrorAction      = 'Stop'
        } | Sort-Object Date)
        $out.Ok = $true
        if (@($out.Events).Count -eq 0) {
            $out.Reason = 'Exchange Online returned no detail events. Detail is kept for 10 days; the trace itself goes back 90, so an older message shows here with no detail.'
        }
    } catch {
        $out.Reason = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MessageTraceDetailV2'
    }
    return $out
}

function Show-MessageDetail {
    <#
        Everything known about one message, in one window.

        THIS REPLACED A HOP-TIMELINE-ONLY DIALOG. The timeline answered "where
        did it go" and left the two questions that actually close a ticket --
        what happened to it, and which rule did that -- to be worked out
        somewhere else.

        FOUR SECTIONS, IN THE ORDER SOMEBODY TROUBLESHOOTS:
            1. what the message is
            2. what the verdict was, and what that verdict MEANS
            3. every hop, both sides
            4. every mail flow rule that touched it, explained
    #>
    param([Parameter(Mandatory)]$Row)

    $tz = Get-SelectedTimeZone -Label "$($cboTz.SelectedItem)"

    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Message detail -- $($Row.Subject)"
    $dlg.ClientSize      = New-Object System.Drawing.Size(1120, 700)
    $dlg.FormBorderStyle = 'Sizable'
    $dlg.StartPosition   = 'CenterParent'

    $tabs          = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(12, 12)
    $tabs.Size     = New-Object System.Drawing.Size(1096, 640)
    $tabs.Anchor   = 'Top,Left,Right,Bottom'
    [void]$dlg.Controls.Add($tabs)

    function New-DetailTab {
        param([string]$Title)
        $tp      = New-Object System.Windows.Forms.TabPage
        $tp.Text = $Title
        $tb            = New-Object System.Windows.Forms.TextBox
        $tb.Multiline  = $true
        $tb.ReadOnly   = $true
        $tb.ScrollBars = 'Both'
        # B6: a wrapped Message-ID or DN is unreadable.
        $tb.WordWrap   = $false
        $tb.Dock       = 'Fill'
        $tb.Font       = New-Object System.Drawing.Font('Consolas', 9)
        [void]$tp.Controls.Add($tb)
        [void]$tabs.TabPages.Add($tp)
        return $tb
    }

    $tbSum   = New-DetailTab -Title 'Summary'
    $tbHops  = New-DetailTab -Title 'Hops'
    $tbRules = New-DetailTab -Title 'Mail flow rules'
    $tbRaw   = New-DetailTab -Title 'Raw'

    $btnClose           = New-Object System.Windows.Forms.Button
    $btnClose.Text      = 'Close'
    $btnClose.Location  = New-Object System.Drawing.Point(998, 660)
    $btnClose.Size      = New-Object System.Drawing.Size(110, 28)
    $btnClose.FlatStyle = 'Flat'
    $btnClose.Anchor    = 'Right,Bottom'
    $btnClose.Add_Click({ $dlg.Close() })
    [void]$dlg.Controls.Add($btnClose)

    $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    # --- the cloud detail, fetched once and shared by three tabs -------------
    $exo = Get-ExoMessageDetail -TraceGuid "$($Row.TraceGuid)" -Recipient "$($Row.Recipient)"

    # --- 1. summary ----------------------------------------------------------
    $sumLines = New-Object System.Collections.ArrayList
    [void]$sumLines.Add("Subject     : $($Row.Subject)")
    [void]$sumLines.Add("From        : $($Row.Sender)")
    [void]$sumLines.Add("To          : $($Row.Recipient)")
    $lt = ConvertTo-DisplayTime -Utc $Row.TimestampUtc -TimeZone $tz
    [void]$sumLines.Add("When        : $(if ($lt) { $lt.ToString('yyyy-MM-dd HH:mm:ss') } else { '' })   ($($cboTz.SelectedItem))")
    [void]$sumLines.Add("Size        : $('{0:N0}' -f $Row.SizeBytes) bytes")
    [void]$sumLines.Add("Seen by     : $($Row.Source)")
    if ($Row.Server)   { [void]$sumLines.Add("Server      : $($Row.Server)") }
    [void]$sumLines.Add("Message-ID  : $($Row.MessageIdRaw)")
    if ($Row.TraceGuid) { [void]$sumLines.Add("TraceId     : $($Row.TraceGuid)") }
    [void]$sumLines.Add('')
    [void]$sumLines.Add("OUTCOME     : $($Row.Outcome)")
    [void]$sumLines.Add('')
    foreach ($line in (Get-OutcomeMeaning -Outcome "$($Row.Outcome)" -Row $Row)) { [void]$sumLines.Add($line) }
    if ($Row.Detail) {
        [void]$sumLines.Add('')
        [void]$sumLines.Add('What the source reported:')
        [void]$sumLines.Add("    $($Row.Detail)")
    }
    $tbSum.Text = ($sumLines -join [Environment]::NewLine)

    # --- 2. hops -------------------------------------------------------------
    $hopLines = New-Object System.Collections.ArrayList
    [void]$hopLines.Add("Times in $($cboTz.SelectedItem).")
    [void]$hopLines.Add('')

    $hops = @(Get-TraceHopTimeline -MessageId $Row.MessageIdRaw -Rows $script:AllRows)
    if (@($hops).Count -gt 0) {
        [void]$hopLines.Add('ACROSS BOTH SIDES, correlated on Message-ID:')
        foreach ($hop in $hops) {
            $ht = ConvertTo-DisplayTime -Utc $hop.TimestampUtc -TimeZone $tz
            [void]$hopLines.Add(("  {0}  {1,-18} {2,-22} {3,-24} {4}" -f `
                $(if ($ht) { $ht.ToString('yyyy-MM-dd HH:mm:ss') } else { '                   ' }),
                "$($hop.Source)", "$($hop.Outcome)", "$($hop.Server)", "$($hop.Detail)"))
        }
        [void]$hopLines.Add('')
    }

    if (@($Row.Events).Count -gt 0) {
        [void]$hopLines.Add('ON-PREMISES TRANSPORT EVENTS for this recipient:')
        foreach ($e in @($Row.Events | Sort-Object TimestampUtc)) {
            $et = ConvertTo-DisplayTime -Utc $e.TimestampUtc -TimeZone $tz
            [void]$hopLines.Add(("  {0}  {1,-18} {2,-22} {3}" -f `
                $(if ($et) { $et.ToString('yyyy-MM-dd HH:mm:ss') } else { '                   ' }),
                "$($e.EventId)", "$($e.Server)", "$($e.Detail)"))
        }
        [void]$hopLines.Add('')
    }

    if ($exo.Ok -and @($exo.Events).Count -gt 0) {
        [void]$hopLines.Add('EXCHANGE ONLINE EVENT DETAIL:')
        foreach ($e in @($exo.Events)) {
            $et = ConvertTo-DisplayTime -Utc (ConvertTo-TraceUtc -Value $e.Date) -TimeZone $tz
            [void]$hopLines.Add(("  {0}  {1,-24} {2}" -f `
                $(if ($et) { $et.ToString('yyyy-MM-dd HH:mm:ss') } else { '                   ' }),
                "$($e.Event)", "$($e.Detail)"))
        }
    } elseif ($exo.Reason) {
        [void]$hopLines.Add("EXCHANGE ONLINE EVENT DETAIL: $($exo.Reason)")
    }
    $tbHops.Text = ($hopLines -join [Environment]::NewLine)

    # --- 3. rules ------------------------------------------------------------
    $rep = Get-MessageRuleReport -ExoDetailEvents @($exo.Events) -OnPremEvents @($Row.Events)
    $ruleLines = New-Object System.Collections.ArrayList
    if ($rep.Count -eq 0) {
        [void]$ruleLines.Add($rep.Text)
        [void]$ruleLines.Add('')
        [void]$ruleLines.Add('WORTH KNOWING WHAT THAT DOES AND DOES NOT RULE OUT.')
        [void]$ruleLines.Add('')
        [void]$ruleLines.Add('  It means no rule reported a match on the hops visible here. It does NOT')
        [void]$ruleLines.Add('  mean no rule was involved:')
        [void]$ruleLines.Add('')
        [void]$ruleLines.Add('    - Exchange Online detail is kept for 10 days. The trace goes back 90.')
        [void]$ruleLines.Add('      An older message shows its verdict with no rule data behind it.')
        [void]$ruleLines.Add('    - A rule that matched on a hop this search did not return is not here.')
        [void]$ruleLines.Add('      Widen the date range or search by Message-ID rather than address.')
        [void]$ruleLines.Add('    - Connectors, anti-spam policies and outbound spam policies are not')
        [void]$ruleLines.Add('      mail flow rules and never appear in this list.')
    } else {
        [void]$ruleLines.Add("$($rep.Count) mail flow rule(s) touched this message.")
        [void]$ruleLines.Add('')
        [void]$ruleLines.Add($rep.Text)
    }
    $tbRules.Text = ($ruleLines -join [Environment]::NewLine)

    # --- 4. raw --------------------------------------------------------------
    $rawLines = New-Object System.Collections.ArrayList
    [void]$rawLines.Add('THE UNPARSED RECORDS. Everything above is derived from these -- if a')
    [void]$rawLines.Add('column looks wrong, this is what it was derived from.')
    [void]$rawLines.Add('')
    [void]$rawLines.Add('--- trace row ---')
    [void]$rawLines.Add(($Row | Select-Object * -ExcludeProperty Events | Format-List | Out-String).Trim())
    if (@($Row.Events).Count -gt 0) {
        [void]$rawLines.Add('')
        [void]$rawLines.Add('--- on-premises events ---')
        foreach ($e in @($Row.Events)) { [void]$rawLines.Add(($e | Format-List | Out-String).Trim()); [void]$rawLines.Add('') }
    }
    if ($exo.Ok -and @($exo.Events).Count -gt 0) {
        [void]$rawLines.Add('')
        [void]$rawLines.Add('--- Exchange Online detail events ---')
        foreach ($e in @($exo.Events)) { [void]$rawLines.Add(($e | Format-List | Out-String).Trim()); [void]$rawLines.Add('') }
    }
    $tbRaw.Text = ($rawLines -join [Environment]::NewLine)

    $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
    Set-ConsoleTheme -Form $dlg
    [void]$dlg.ShowDialog()
    $dlg.Dispose()
}
# Show-HopTimeline REMOVED 2026-08-24, replaced by Show-MessageDetail.
#
# It showed the hop list and nothing else, which answered "where did it go" and
# left the two questions that actually close a ticket -- what happened to it,
# and which rule did that -- to be worked out somewhere else. Its hop list is
# now the Hops tab of the new dialog; Get-TraceHopTimeline, which did the real
# work, is unchanged and still called.

# THE RANGE LIMITS ARE A PROPERTY OF THE SOURCE, NOT A PREFERENCE.
#
#   On-premises   Get-MessageTrackingLog reads flat log files, so one query can
#                 span the whole retention period. 90 days is offered.
#   Cloud         Get-MessageTraceV2 returns 90 days of DATA but accepts at most
#                 10 DAYS PER QUERY. That is a hard cmdlet limit, not throttling
#                 -- a wider range is an error, not a slower answer.
#   Historical    Start-HistoricalSearch covers the full 90 days in one request
#                 and delivers a CSV by email instead of returning rows.
$script:OnPremMaxDays  = 90
$script:CloudMaxDays   = 10
$script:HistoryMaxDays = 90
# Which source the guard last saw, so the window widens on a SOURCE CHANGE and
# not on every touch of the date pickers.
$script:LastSource = ''
# Which source the guard last saw, so widening to a full window happens on a
# SOURCE CHANGE and not on every keystroke in the date pickers.
$script:LastSource = ''

function Update-SourceEmphasis {
    <#
        Make the selected source obvious.

        A THEMED RADIO BUTTON IS NEARLY UNREADABLE AS A STATE. The dot is drawn
        by the OS in the system accent colour, which does not follow the theme --
        on a dark palette it is a dark mark on a dark field, and on a light one
        it is barely there. The operator has to lean in to see which of three
        options is live, on a control that decides what gets searched.

        THE TEXT CARRIES THE STATE INSTEAD. Bold plus the theme's selection
        colour on the chosen one, regular and muted on the others -- weight
        survives any palette, and colour alone would not.
    #>
    $theme = Get-ConsoleTheme
    $on  = $null
    $off = $null
    if ($theme) {
        $on  = ConvertTo-ThemeColor $theme.SelectedText
        if (-not $on) { $on = ConvertTo-ThemeColor $theme.ButtonBorder }
        $off = ConvertTo-ThemeColor $theme.FormText
    }

    foreach ($rb in @($rbExo, $rbOnPrem, $rbBoth)) {
        if (-not $rb) { continue }
        if ($rb.Checked) {
            $rb.Font = New-Object System.Drawing.Font($rb.Font.FontFamily, $rb.Font.Size, [System.Drawing.FontStyle]::Bold)
            if ($on) { $rb.ForeColor = $on }
        } else {
            $rb.Font = New-Object System.Drawing.Font($rb.Font.FontFamily, $rb.Font.Size, [System.Drawing.FontStyle]::Regular)
            if ($off) { $rb.ForeColor = $off }
        }
        # B3/theme: mark these as semantically coloured so a later theme pass
        # does not flatten them back to the form's text colour.
        $rb.Tag = 'SemanticColour'
    }
}
function Get-RangeLimitDays {
    <#
        How many days the CURRENT source selection allows in one search.

        BOTH IS CAPPED AT THE CLOUD LIMIT, not the on-premises one. A 30-day
        "Both" search would return 30 days of on-premises rows against 10 days
        of cloud rows, and the grid would look as though the cloud had lost
        everything older than 10 days. A partial answer that looks complete is
        worse than a refusal.
    #>
    if ($rbOnPrem.Checked) { return $script:OnPremMaxDays }
    return $script:CloudMaxDays
}

function Update-RangeGuard {
    <#
        Keep the pickers inside the limit, and say what the limit is.

        THE START DATE IS MOVED, NOT THE END. When a limit shrinks -- switching
        from On-premises to Cloud with a 60-day range open -- pulling the END
        backwards would silently drop the most recent mail, which is what the
        operator is nearly always looking for. Moving the START forwards keeps
        the newest end of the window intact.
    #>
    param([switch]$Quiet)

    $limit = Get-RangeLimitDays
    $now   = Get-Date

    if ($dtEnd.Value -gt $now) { $dtEnd.Value = $now }
    $floor = $now.AddDays(-1 * $script:HistoryMaxDays)
    if ($dtStart.Value -lt $floor) { $dtStart.Value = $floor }

    $span = ($dtEnd.Value - $dtStart.Value).TotalDays
    $clamped = $false
    if ($span -gt $limit) {
        $dtStart.Value = $dtEnd.Value.AddDays(-1 * $limit)
        $clamped = $true
    }

    # SWITCHING TO A CLOUD SOURCE OPENS THE WINDOW TO A FULL 10 DAYS.
    #
    # Not a snap-to-grid -- any narrower range is still allowed and is left alone
    # once set. This fires only when the SOURCE changes, so the default window
    # tiles the 90 days cleanly with Back and Forward rather than leaving the
    # operator stepping in 1-day hops from the default range.
    if (-not $rbOnPrem.Checked -and $script:LastSource -ne 'cloud' -and $span -lt $limit) {
        $dtStart.Value = $dtEnd.Value.AddDays(-1 * $limit)
        if ($dtStart.Value -lt $floor) { $dtStart.Value = $floor }
    }
    $script:LastSource = $(if ($rbOnPrem.Checked) { 'onprem' } else { 'cloud' })
    if ($dtEnd.Value -le $dtStart.Value) { $dtEnd.Value = $dtStart.Value.AddHours(1) }

    $src = if ($rbOnPrem.Checked) { 'on-premises' } elseif ($rbExo.Checked) { 'Exchange Online' } else { 'both' }
    $lblRange.Text = "$src : up to $limit day(s) per search, $($script:HistoryMaxDays) days of history"
    $btnEnhanced.Enabled = $rbExo.Checked
    $btnExtended.Enabled = $rbExo.Checked

    # STEPPING IS ONLY OFFERED WHERE IT IS NEEDED. On-premises covers the whole
    # 90 days in one query, so a window-stepping control there does nothing.
    $stepping = -not $rbOnPrem.Checked
    $btnBack.Visible    = $stepping
    $btnForward.Visible = $stepping

    if ($clamped -and -not $Quiet) {
        $msg = New-Object System.Collections.ArrayList
        [void]$msg.Add("The range was reduced to $limit day(s) -- the most $src can return in one search.")
        [void]$msg.Add('')
        [void]$msg.Add('The START was moved forward rather than the end, so the newest mail is still')
        [void]$msg.Add('in range. That is almost always what is being looked for.')
        if ($stepping) {
            [void]$msg.Add('')
            [void]$msg.Add('Use Back to step the window a full width into the past. Nothing is lost --')
            [void]$msg.Add('it takes several searches.')
        }
        if ($rbExo.Checked) {
            [void]$msg.Add('')
            [void]$msg.Add('For one search across all 90 days use 90-day search, which emails a CSV')
            [void]$msg.Add('instead of filling the grid.')
        }
        $txtOut.Text = ($msg -join [Environment]::NewLine)
    }
}

function Move-RangeWindow {
    <#
        Step the whole window one full width earlier or later.

        MOVES BY THE WINDOW WIDTH, not by a fixed 10 days. If the operator has
        narrowed to 2 days to keep the result readable, stepping by 10 would
        skip 8 days without saying so -- the gap would simply never be searched
        and nothing would report it.
    #>
    param([int]$Direction)

    $span = $dtEnd.Value - $dtStart.Value
    if ($span.TotalMinutes -lt 1) { $span = New-TimeSpan -Days 1 }

    $newStart = $dtStart.Value.Add([TimeSpan]::FromTicks($span.Ticks * $Direction))
    $newEnd   = $dtEnd.Value.Add([TimeSpan]::FromTicks($span.Ticks * $Direction))
    $now      = Get-Date
    $floor    = $now.AddDays(-1 * $script:HistoryMaxDays)

    if ($newEnd -gt $now) {
        $txtOut.Text = 'Already at the present -- there is nothing later to step into.'
        return
    }
    if ($newStart -lt $floor) {
        $txtOut.Text = 'Already at the 90-day limit. Neither source holds anything older.'
        return
    }

    $dtStart.Value = $newStart
    $dtEnd.Value   = $newEnd
    Update-RangeGuard -Quiet
    $txtOut.Text = "Window moved to $($newStart.ToString('yyyy-MM-dd HH:mm')) - $($newEnd.ToString('yyyy-MM-dd HH:mm')). Press Search."
}

function Get-NotifyAddress {
    <#
        A small modal asking where the report should go.

        A DIALOG RATHER THAN A FIELD ON THE MAIN FORM. This is asked once per
        historical search and never for a normal one, so a permanent box would
        be a permanently confusing control -- and one that looks like it affects
        the ordinary Search button.
    #>
    param([string]$Default, [string]$Kind = 'Enhanced', [string]$RowCap = '', [string]$RangeText = '')

    $askForm                 = New-Object System.Windows.Forms.Form
    $askForm.Text            = "$Kind report -- where should it go?"
    $askForm.ClientSize      = New-Object System.Drawing.Size(520, 190)
    $askForm.FormBorderStyle = 'FixedDialog'
    $askForm.MaximizeBox     = $false
    $askForm.MinimizeBox     = $false
    $askForm.StartPosition   = 'CenterParent'

    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(14, 14)
    $lbl.Size     = New-Object System.Drawing.Size(492, 68)
    $lbl.Text     = "A CSV is emailed when the search finishes." + [Environment]::NewLine +
                    "It can take minutes or hours, and NOTHING will appear in the grid." + [Environment]::NewLine + [Environment]::NewLine +
                    "Send it to:"
    [void]$askForm.Controls.Add($lbl)

    $txt          = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(14, 88)
    $txt.Size     = New-Object System.Drawing.Size(492, 22)
    $txt.Text     = "$Default"
    [void]$askForm.Controls.Add($txt)

    $hint           = New-Object System.Windows.Forms.Label
    $hint.Location  = New-Object System.Drawing.Point(14, 114)
    $hint.Size      = New-Object System.Drawing.Size(492, 18)
    $hint.ForeColor = [System.Drawing.Color]::Gray
    $hint.Text      = 'A shared mailbox is fine. It does not have to be you.'
    [void]$askForm.Controls.Add($hint)

    $ok           = New-Object System.Windows.Forms.Button
    $ok.Text      = 'Start search'
    $ok.Location  = New-Object System.Drawing.Point(268, 144)
    $ok.Size      = New-Object System.Drawing.Size(114, 30)
    $ok.FlatStyle = 'Flat'
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    [void]$askForm.Controls.Add($ok)

    $no           = New-Object System.Windows.Forms.Button
    $no.Text      = 'Cancel'
    $no.Location  = New-Object System.Drawing.Point(392, 144)
    $no.Size      = New-Object System.Drawing.Size(114, 30)
    $no.FlatStyle = 'Flat'
    $no.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    [void]$askForm.Controls.Add($no)

    $askForm.AcceptButton = $ok
    $askForm.CancelButton = $no
    Set-ConsoleTheme -Form $askForm
    $res = $askForm.ShowDialog()
    $val = "$($txt.Text)".Trim()
    $askForm.Dispose()
    if ($res -ne [System.Windows.Forms.DialogResult]::OK) { return '' }
    return $val
}
function Start-NinetyDaySearch {
    <#
        An archived search that arrives by email instead of in the grid.

        THIS IS A DIFFERENT CMDLET, NOT A WIDER RANGE. Get-MessageTraceV2 cannot
        be asked for more than 10 days. Start-HistoricalSearch reads the archive
        and does not return rows at all -- it queues a job and Microsoft emails
        a CSV when it finishes.

        TWO REPORT TYPES, AND THE DIFFERENCE IS NOT COSMETIC:

          ENHANCED  -ReportType MessageTrace
                    Up to 100,000 rows. The summary fields plus direction and
                    original client IP.
                    Use it to answer "did this arrive, and where did it go" over
                    a lot of mail.

          EXTENDED  -ReportType MessageTraceDetail
                    Full routing and per-event detail -- but capped at 1,000
                    ROWS, a hundredth of Enhanced.
                    Use it on ONE message, or a very narrow search. Point it at
                    a busy mailbox for 90 days and it silently stops at 1,000
                    and the answer looks complete when it is not.

        THE ROW CAP IS THE WHOLE TRADE-OFF. Detail costs coverage. Picking
        Extended because "more detail sounds better" is how a 90-day question
        gets answered with the first 1,000 rows.

        CLOUD ONLY. There is no on-premises equivalent, so both buttons are
        disabled unless Exchange Online is the selected source -- a "Both"
        historical search would silently be cloud-only and the on-premises half
        would be missing from a report treated as complete.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('Enhanced', 'Extended')][string]$Kind
    )

    $from = "$($txtFrom.Text)".Trim()
    $to   = "$($txtTo.Text)".Trim()
    $mid  = "$($txtMid.Text)".Trim()

    # REQUIRED BY THE SERVICE, not just by this tool. Start-HistoricalSearch
    # rejects a request with none of MessageID, RecipientAddress or
    # SenderAddress, so catching it here saves a round trip and an opaque error.
    if (-not $from -and -not $to -and -not $mid) {
        $txtOut.Text = 'Give a sender, a recipient or a message ID. The service refuses an archived search without at least one of those.'
        return
    }

    $reportType = $(if ($Kind -eq 'Extended') { 'MessageTraceDetail' } else { 'MessageTrace' })
    $rowCap     = $(if ($Kind -eq 'Extended') { '1,000' } else { '100,000' })

    # THE PICKERS DECIDE THE RANGE, not a hardcoded 90 days. The window is
    # clamped to 10 for the grid, so widen it before pressing this if a longer
    # archived report is wanted -- the clamp does not apply to this path.
    $start = $dtStart.Value
    $end   = $dtEnd.Value
    if ($end -le $start) { $txtOut.Text = 'The end of the range is not after the start.'; return }

    $now   = Get-Date
    $floor = $now.AddDays(-1 * $script:HistoryMaxDays)
    if ($start -lt $floor) { $start = $floor }
    if ($end -gt $now) { $end = $now }

    $default = ''
    try {
        $ci = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object { "$($_.State)" -eq 'Connected' })
        if (@($ci).Count -gt 0) { $default = "$($ci[0].UserPrincipalName)" }
    } catch { }

    $addr = Get-NotifyAddress -Default $default -Kind $Kind -RowCap $rowCap `
        -RangeText "$($start.ToString('yyyy-MM-dd')) to $($end.ToString('yyyy-MM-dd'))"
    if (-not $addr) { $txtOut.Text = 'Cancelled. No archived search was started.'; return }
    if ($addr -notlike '*@*' -or $addr -like '* *') {
        $txtOut.Text = "'$addr' is not an email address, so nothing was started. The report can only be delivered by email."
        return
    }

    $name = "$Kind $((Get-Date).ToString('yyyyMMdd-HHmmss'))"
    $params = @{
        ReportTitle   = $name
        StartDate     = $start
        EndDate       = $end
        ReportType    = $reportType
        NotifyAddress = $addr
        ErrorAction   = 'Stop'
    }
    if ($from) { $params.SenderAddress    = $from }
    if ($to)   { $params.RecipientAddress = $to }
    if ($mid)  { $params.MessageID        = $mid }

    $btnEnhanced.Enabled = $false
    $btnExtended.Enabled = $false
    try {
        $job = Invoke-EXOCommand -Name 'Start-HistoricalSearch' -Parameters $params

        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("$($Kind.ToUpper()) 90-DAY SEARCH QUEUED. Nothing will appear in the grid.")
        [void]$L.Add('')
        [void]$L.Add("Report    : $name")
        if ($job -and "$($job.JobId)") { [void]$L.Add("Job ID    : $($job.JobId)") }
        [void]$L.Add("Type      : $reportType")
        [void]$L.Add("Row limit : $rowCap")
        [void]$L.Add("Range     : $($start.ToString('yyyy-MM-dd HH:mm')) to $($end.ToString('yyyy-MM-dd HH:mm'))  ($([Math]::Round(($end - $start).TotalDays, 1)) days)")
        [void]$L.Add("Emailed   : $addr")
        if ($from) { [void]$L.Add("Sender    : $from") }
        if ($to)   { [void]$L.Add("Recipient : $to") }
        if ($mid)  { [void]$L.Add("MessageID : $mid") }
        [void]$L.Add('')
        if ($Kind -eq 'Extended') {
            [void]$L.Add('EXTENDED IS CAPPED AT 1,000 ROWS. If this search could match more than')
            [void]$L.Add('that, the CSV will stop at 1,000 and will NOT say it was truncated. Narrow')
            [void]$L.Add('the sender, recipient or date range, or run Enhanced instead -- it takes')
            [void]$L.Add('100,000 rows and answers "did it arrive" perfectly well.')
            [void]$L.Add('')
        }
        [void]$L.Add('WHAT HAPPENS NOW')
        [void]$L.Add('')
        [void]$L.Add('  Microsoft runs the search against archived data and emails a CSV when it')
        [void]$L.Add('  is done. That takes minutes to several hours depending on the range and')
        [void]$L.Add('  how many other requests are queued. Closing this window does not cancel it.')
        [void]$L.Add('')
        [void]$L.Add('  THE LAST FEW HOURS MAY BE MISSING. Archived data lags live data by roughly')
        [void]$L.Add('  1-4 hours, so a message sent this morning may not be in the report even')
        [void]$L.Add('  though the range covers today. Use a normal search for anything recent.')
        [void]$L.Add('')
        [void]$L.Add('  Check progress with Get-HistoricalSearch. NotStarted means queued, not')
        [void]$L.Add('  stuck -- there is a per-tenant limit and submitting more will not help.')
        [void]$L.Add('  There is a cap of 250 searches per day, and cancelled ones still count.')
        [void]$L.Add('')
        [void]$L.Add('CLOUD ONLY. Anything that never reached Exchange Online is not in this')
        [void]$L.Add('report. For that, search On-premises directly -- it takes 90 days in one go.')
        $txtOut.Text = ($L -join [Environment]::NewLine)

        Write-CompletedLog -Action 'Report' -Target (Get-LogTarget 'Hybrid Message Trace') -Result 'Success' `
            -Detail "Queued $Kind ($reportType) historical search '$name', $($start.ToString('yyyy-MM-dd')) to $($end.ToString('yyyy-MM-dd')), notify $addr. Sender='$from' Recipient='$to' MessageID='$mid'." `
            -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Start-HistoricalSearch ($reportType)"
        Write-FailedLog -Action 'Report' -Target (Get-LogTarget 'Hybrid Message Trace') `
            -Detail "Start-HistoricalSearch ($reportType) failed. $($_.Exception.Message)" -Environment $Environment
    } finally {
        $btnEnhanced.Enabled = $rbExo.Checked
        $btnExtended.Enabled = $rbExo.Checked
    }
}

function Invoke-HybridSearch {
    $from = "$($txtFrom.Text)".Trim()
    $to   = "$($txtTo.Text)".Trim()
    $mid  = "$($txtMid.Text)".Trim()
    if (-not $from -and -not $to -and -not $mid) {
        $txtOut.Text = 'Give a sender, a recipient or a message ID.'
        return
    }

    $start = $dtStart.Value
    $end   = $dtEnd.Value
    if ($end -le $start) { $txtOut.Text = 'The end of the range is not after the start.'; return }

    # BELT AND BRACES. Update-RangeGuard keeps the pickers legal as they are
    # touched, but a value can still arrive from a restored form state or a
    # future caller, and Get-MessageTraceV2 answers an over-wide range with an
    # error rather than a truncated result. Refusing here is cheaper than
    # decoding that error later.
    $limitDays = Get-RangeLimitDays
    if (($end - $start).TotalDays -gt $limitDays) {
        $txtOut.Text = "That range is $([Math]::Round(($end - $start).TotalDays, 1)) days and the limit for this source is $limitDays." +
            [Environment]::NewLine + 'Narrow it, or use 90-day search for a full-history report by email.'
        return
    }

    $doExo    = $rbExo.Checked -or $rbBoth.Checked
    $doOnPrem = $rbOnPrem.Checked -or $rbBoth.Checked

    $btnSearch.Enabled = $false
    $script:AllRows = @()
    $lblPartial.Text = ''
    $pb.Value = 0
    $L = New-Object System.Collections.ArrayList
    $exoRows = @()
    $onPremRows = @()
    $failedServers = @()
    $askedCount = 0
    $answeredCount = 0
    $edgeExcluded = @()
    $neverQueried = @()
    $truncated = @()

    try {
        if ($doExo) {
            $txtOut.Text = 'Querying Exchange Online ...'
            $form.Refresh()
            try {
                $exoRows = @(Get-ExoTraceRows -SenderAddress $from -RecipientAddress $to `
                             -StartDate $start -EndDate $end -MessageId $mid)
                [void]$L.Add("Exchange Online: $(@($exoRows).Count) row(s).")
            } catch {
                [void]$L.Add("Exchange Online FAILED: $($_.Exception.Message)")
            }
        }

        if ($doOnPrem) {
            $srv = Get-OnPremTransportServer
            if (-not $srv.Ok) {
                [void]$L.Add("On-premises: could not list transport servers -- $($srv.Reason)")
            } else {
                $edgeExcluded = @($srv.Edge)
                $neverQueried = @($srv.Excluded)
                $servers = @($srv.Servers)
                $askedCount = @($servers).Count

                # SEQUENTIAL, WITH THE SERVER NAMED. A visible slow answer beats
                # a silent fast one -- the operator can see which server is
                # taking the time, and that is usually the useful fact.
                $n = 0
                $pb.Maximum = [Math]::Max(1, $askedCount)
                foreach ($s in $servers) {
                    $n++
                    $txtOut.Text = "Querying $s  ($n of $askedCount) ..."
                    $lblPartial.Text = "Querying on-premises server $n of $askedCount -- $s"
                    [System.Windows.Forms.Application]::DoEvents()

                    $one = Get-OnPremTraceRows -SenderAddress $from -RecipientAddress $to `
                            -StartDate $start -EndDate $end -MessageId $mid -Servers @($s)
                    $onPremRows += @($one.Rows)
                    $answeredCount += $one.Answered
                    $failedServers += @($one.Failed)
                    $truncated     += @($one.Truncated)

                    $pb.Value = [Math]::Min($n, $pb.Maximum)
                }
                [void]$L.Add("On-premises: $(@($onPremRows).Count) row(s) from $answeredCount of $askedCount server(s).")
            }
        }

        $script:AllRows = @(Merge-TraceRows -ExoRows $exoRows -OnPremRows $onPremRows)

        # THE FULL OUTCOME SET WITH COUNTS -- see Update-OutcomeFilter for why
        # this is no longer built from the results alone.
        Update-OutcomeFilter
        Update-Grid

        # ---- the honesty line ----
        $partial = @($failedServers).Count -gt 0
        if ($doOnPrem -and $askedCount -gt 0) {
            $txt = "$answeredCount of $askedCount servers answered"
            if ($partial) {
                $txt += " -- FAILED: $((@($failedServers | ForEach-Object { $_.Server }) -join ', '))"
                $txt += '. ROWS ARE MISSING.'
                $lblPartial.ForeColor = [System.Drawing.Color]::FromArgb(255, 105, 97)
            } else {
                $lblPartial.ForeColor = [System.Drawing.Color]::FromArgb(120, 200, 255)
            }
            if (@($edgeExcluded).Count -gt 0) {
                $txt += "   Edge Transport NOT queried: $((@($edgeExcluded) -join ', '))"
            }
            # Stated, not warned about. The exclusion is deliberate, so this is
            # scope rather than a problem -- but a hop on one of these would not
            # appear, and that stays visible.
            if (@($neverQueried).Count -gt 0) {
                $txt += "   Excluded by config: $((@($neverQueried) -join ', '))"
            }
            if (@($truncated).Count -gt 0) {
                $txt += "   TRUNCATED at 5000 rows on: $((@($truncated) -join ', '))"
            }
            $lblPartial.Text = $txt
        } else {
            $lblPartial.Text = ''
        }

        [void]$L.Add('')
        [void]$L.Add("$(@($script:AllRows).Count) row(s) total.")
        foreach ($f in $failedServers) { [void]$L.Add("  FAILED $($f.Server): $($f.Reason)") }
        if (@($edgeExcluded).Count -gt 0) {
            [void]$L.Add('')
            [void]$L.Add("Edge Transport servers exist and were NOT queried: $((@($edgeExcluded) -join ', ')). A hop that happened there is not shown.")
        }
        if ($doOnPrem -and $doExo) {
            [void]$L.Add('')
            [void]$L.Add('A message that crossed the hybrid boundary appears TWICE -- once as the on-premises hop and once as the Exchange Online outcome. That is not duplication: they are two different facts. Double-click either to see the whole path.')
        }
        $txtOut.Text = ($L -join [Environment]::NewLine)

        # ---- audit: one row per search ----
        $target = $from
        if (-not $target) { $target = $to }
        if (-not $target) { $target = 'Message Trace' }
        $detail = "Hybrid trace $($start.ToString('s')) to $($end.ToString('s')). Sources: $(if($doExo){'EXO '})$(if($doOnPrem){'OnPrem'}). Rows: $(@($script:AllRows).Count). Servers: $answeredCount of $askedCount."
        if ($partial) {
            $detail += " FAILED: $((@($failedServers | ForEach-Object { $_.Server }) -join ', '))."
            Write-CompletedLog -Action 'Read' -Target (Get-LogTarget $target 'Hybrid Message Trace') `
                -Result 'Partial' -Detail $detail -Environment $Environment
        } else {
            Write-CompletedLog -Action 'Read' -Target (Get-LogTarget $target 'Hybrid Message Trace') `
                -Result 'Success' -Detail $detail -Environment $Environment
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Hybrid trace'
    } finally {
        $btnSearch.Enabled = $true
        $pb.Value = 0
    }
}

function Export-Rows {
    if (@($script:AllRows).Count -eq 0) { $txtOut.Text = 'Nothing to export.'; return }
    $tz = Get-SelectedTimeZone -Label "$($cboTz.SelectedItem)"
    $out = @($script:AllRows | ForEach-Object {
        $local = ConvertTo-DisplayTime -Utc $_.TimestampUtc -TimeZone $tz
        [PSCustomObject]@{
            Time      = $(if ($local) { $local.ToString('yyyy-MM-dd HH:mm:ss') } else { '' })
            TimeZone  = "$($cboTz.SelectedItem)"
            Source    = $_.Source
            Sender    = $_.Sender
            Recipient = $_.Recipient
            Subject   = $_.Subject
            Outcome   = $_.Outcome
            Server    = $_.Server
            Event     = $_.RawEvent
            MessageId = $_.MessageIdRaw
            SizeBytes = $_.SizeBytes
            Detail    = $_.Detail
        } })
    $p = Write-ReportCsv -Rows $out -Prefix 'HybridMessageTrace'
    $txtOut.Text = "Exported $(@($out).Count) row(s) to:$([Environment]::NewLine)$p"
}

# ---- handlers, all wrapped (B1) ----
$btnBack.Add_Click({ try { Move-RangeWindow -Direction -1 } catch { Show-HandlerError -ErrorRecord $_ -Context 'Back' } })
$btnForward.Add_Click({ try { Move-RangeWindow -Direction 1 } catch { Show-HandlerError -ErrorRecord $_ -Context 'Forward' } })
$btnEnhanced.Add_Click({ try { Start-NinetyDaySearch -Kind 'Enhanced' } catch { Show-HandlerError -ErrorRecord $_ -Context '90 day Enhanced' } })
$btnExtended.Add_Click({ try { Start-NinetyDaySearch -Kind 'Extended' } catch { Show-HandlerError -ErrorRecord $_ -Context '90 day Extended' } })

# THE GUARD RUNS ON EVERY INPUT THAT CAN BREAK THE LIMIT -- both pickers and all
# three source radios. Checking only on Search would let the operator set a
# 60-day cloud range, see nothing stop them, and be told no at the end.
$dtStart.Add_ValueChanged({ try { Update-RangeGuard } catch { } })
$dtEnd.Add_ValueChanged({ try { Update-RangeGuard } catch { } })
$rbExo.Add_CheckedChanged({ try { Update-RangeGuard; Update-SourceEmphasis } catch { } })
$rbOnPrem.Add_CheckedChanged({ try { Update-RangeGuard; Update-SourceEmphasis } catch { } })
$rbBoth.Add_CheckedChanged({ try { Update-RangeGuard; Update-SourceEmphasis } catch { } })

$btnSearch.Add_Click({ try { Invoke-HybridSearch } catch { Show-HandlerError -ErrorRecord $_ -Context 'Search' } })
$btnExport.Add_Click({ try { Export-Rows } catch { Show-HandlerError -ErrorRecord $_ -Context 'Export' } })
$cboFilter.Add_SelectedIndexChanged({ try { Update-Grid } catch { Show-HandlerError -ErrorRecord $_ -Context 'Filter' } })
$cboTz.Add_SelectedIndexChanged({ try { Update-Grid } catch { Show-HandlerError -ErrorRecord $_ -Context 'Timezone' } })
$btnExpandAll.Add_Click({ try { Set-AllTraceExpanded -Open $true } catch { Show-HandlerError -ErrorRecord $_ -Context 'Expand all' } })
$btnCollapseAll.Add_Click({ try { Set-AllTraceExpanded -Open $false } catch { Show-HandlerError -ErrorRecord $_ -Context 'Collapse all' } })
# THE MARKER COLUMN TOGGLES. NOTHING ELSE DOES.
#
# This used to fire on a click anywhere in the row, with a comment arguing that
# a 34px target is fiddly to hit. That was wrong, and wrong in the way that
# matters: SELECTING a message is the common action and TOGGLING is the rare
# one, so every attempt to read a message expanded or collapsed it -- and
# because expanding rebuilds the grid, it also threw away the column sort and
# jumped back to the top of the list.
#
# One misplaced convenience cost the sort, the scroll position and the
# selection, on every single click.
$grid.Add_CellClick({
    param($sender, $e)
    try {
        if ($e.RowIndex -lt 0) { return }
        if ($e.ColumnIndex -ne 0) { return }
        Toggle-TraceGroup -Index $e.RowIndex
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Expand' }
})

# SORTING IS HANDLED HERE because the grid's own sort would separate children
# from parents -- see the note by $script:SortCol.
#
# Clicking the sorted column again reverses it, which is what every grid does
# and what the operator will try first.
$grid.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    try {
        if ($e.ColumnIndex -lt 0) { return }
        $name = "$($grid.Columns[$e.ColumnIndex].Name)"
        if (-not $name -or $name -eq ' ') { return }
        if ($name -eq $script:SortCol) { $script:SortAsc = -not $script:SortAsc }
        else { $script:SortCol = $name; $script:SortAsc = $true }
        Update-Grid
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Sort' }
})
$grid.Add_CellDoubleClick({
    try {
        $rows = @($grid.Tag)
        $i = $grid.CurrentRow.Index
        if ($i -lt 0 -or $i -ge @($rows).Count) { return }
        Show-MessageDetail -Row $rows[$i]
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Hop timeline' }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Hybrid Message Trace' | Out-Null

$form.Add_Shown({
    # Quiet, so opening the tool does not look like a complaint.
    try { Update-RangeGuard -Quiet } catch { }
    # AFTER Set-ConsoleTheme has run, or the theme pass overwrites the emphasis.
    try { Update-SourceEmphasis } catch { }
    try {
        if ("$SmtpAddress".Trim()) { $txtFrom.Text = Get-TargetSmtp -Text $SmtpAddress }

        # CAPABILITY GATING. Without the on-premises shell this tool must behave
        # exactly like the Exchange Online one -- greyed, with the reason, not
        # broken and not silently returning half an answer.
        $script:EmsState = Test-EmsAvailable
        $L = New-Object System.Collections.ArrayList
        if (-not $script:EmsState.Available) {
            $rbOnPrem.Enabled = $false
            $rbBoth.Enabled   = $false
            $rbExo.Checked    = $true
            $why = "On-premises Exchange is not available in this process: $($script:EmsState.Detail)"
            $tips.SetToolTip($rbOnPrem, $why)
            $tips.SetToolTip($rbBoth, $why)
            [void]$L.Add('EXCHANGE ONLINE ONLY.')
            [void]$L.Add("  $why")
            [void]$L.Add('')
            [void]$L.Add('Everything else works exactly as it does in the Exchange Online tool. Connect on-premises and reopen to search both.')
        } else {
            [void]$L.Add('Both Exchange Online and on-premises Exchange are available.')
            [void]$L.Add('')
            [void]$L.Add('Exchange Online is selected because the on-premises half asks every transport server in turn and is the slow one. Pick Both when you need to see a message cross the boundary.')
        }
        [void]$L.Add('')
        [void]$L.Add('Give a sender, a recipient or a message ID, then Search. Double-click any row for the hop timeline.')
        $txtOut.Text = ($L -join [Environment]::NewLine)
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Hybrid Message Trace' }
})

[void]$form.ShowDialog()
#endregion
