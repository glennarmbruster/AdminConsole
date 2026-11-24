#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
# CLOUD ONLY. Invoke-ChangeMeetingOrganizer exists only in Exchange Online.
# ---------------------------------------------------------------------------
<#
    ChangeMeetingOrganizer.ps1
    Transfer a meeting series to a new organizer -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-09-18

    2026-09-18  'Everything they organise...' -- the offboarding sweep. Reads
                the calendar log with NO subject, two years back a quarter at
                a time, and lists every recurring series the mailbox organises
                for a tick-and-transfer. Get-MeetingFacts' row-reading was
                extracted to ConvertTo-MeetingFactList so both paths judge a
                series the same way.
    2026-09-16  The confirmation names the subject or id being transferred. The
                batch loop checks what came back before reporting DONE, the
                same guard the single path already had; a returned candidate
                list is told apart by the absence of NewMeetingId.

    2026-09-15  Ships as ONE FILE: the two Common.ps1 helpers it needs are
                defined locally when Common does not already have them.
    2026-09-15  Root cause of every 'server side error' seen in testing: the
                cmdlet rejects -ErrorAction. Calls now go through
                Invoke-EXOCommandNoEA. Id box takes MeetingId or EventId.
    2026-09-14  Find now reads the meeting off the calendar first
                (Get-CalendarDiagnosticObjects) and says outright when it is
                a single occurrence, already over, or cancelled -- the three
                things the transfer cmdlet reports as one generic error.
#>
# TOOLKIT-HELP-BEGIN
# Moves a meeting or series from one organizer to another -- the usual case
# being someone leaving who owns a recurring meeting nobody else can edit.
#
# Look up the CURRENT organizer, type the subject, name the new organizer,
# press Find, then Transfer.
#
# THE SUBJECT IS A SUBSTRING MATCH. 'Status' matches 'Weekly Status' and
# 'Status Sync' too. Several matches come back as a list with a MeetingId and
# an EventId each. Then either paste one id into the id box for a single
# transfer, or press Transfer with the subject and say yes to 'Transfer ALL'
# -- the tool then runs one transfer per meeting, by id, and reports each.
#
# FIND READS THE CALENDAR FIRST and lists what matches the subject: when it
# is, whether it is a series or a one-off, and whether it is already over or
# cancelled. The transfer cmdlet only moves a SERIES from its next future
# instance; a single meeting, or one that has happened, is refused with a
# generic error. This tells you which before the cmdlet is asked.
#
# OFFBOARDING: 'EVERYTHING THEY ORGANISE...' does the whole handover. Put the
# leaver in the top box and their replacement in the second, press it, and it
# reads the calendar for every recurring series that mailbox organises, shows
# them in a tick-list, and transfers the ones you tick -- one at a time, each
# reported.
#
# RUN IT BEFORE THE MAILBOX IS CONVERTED TO SHARED. Once it is shared the
# transfer cmdlet refuses outright, and the only way back is to convert it to
# a user mailbox and re-license it.
#
# TWO LIMITS OF THAT SWEEP, both from the calendar log rather than this tool:
#   * The two-year window is on when a meeting was LAST EDITED, not when it
#     runs. A series nobody has touched in longer will not appear.
#   * Each quarter caps at 1000 log rows. If one fills up, the tool says the
#     list may be incomplete -- believe it, and check the calendar by hand for
#     anything you expected and cannot see.
#
# AMBIGUOUS in the results means several meetings share that subject and the
# cmdlet will not guess between them. Transfer those from the subject and id
# boxes instead.
#
# READ THIS BEFORE USING IT ON A TEAMS MEETING. Changing the organizer does NOT
# change the organizer of the Teams online meeting attached to it. The new
# organizer will not be able to manage or start that Teams meeting and must
# replace the online meeting details manually afterwards.
#
# WHO NOTICES:
#   Attendees in this tenant  -- nothing. Silently updated, no re-RSVP.
#   External / on-premises    -- a cancellation from the old organizer and a
#                                new invitation from the new one. They must
#                                re-RSVP, and any Reminder / Category / Show As
#                                changes they made are lost.
#
# ONE WAY, ONCE. After a transfer you cannot transfer the part of the series
# left behind on the original calendar again.
#
# Only user mailboxes, only the default calendar. Shared and room mailboxes are
# not supported -- give the new person access to the mailbox instead.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress
)

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

#region Standalone fallbacks -----------------------------------------------------
# THIS TOOL SHIPS AS ONE FILE. The two helpers it needs arrived in Common.ps1
# on 2026-09-15; a console whose Common.ps1 predates that would break on this
# script alone. Defined here ONLY when Common has not already provided them,
# so a toolkit with the current Common uses Common's copy and there is never
# a second, drifting implementation in play.
#
# Both rely on Get-EXOCommand and Get-ExchangeModuleMap, which every version
# of Common.ps1 that has ever shipped provides.
if (-not (Get-Command -Name 'Invoke-EXOCommandNoEA' -ErrorAction SilentlyContinue)) {
    function Invoke-EXOCommandNoEA {
        <#
            Invoke-EXOCommand for the cmdlets that CANNOT be given -ErrorAction.

            Proven 2026-09-15 on Invoke-ChangeMeetingOrganizer and
            Get-CalendarDiagnosticObjects: the identical call works bare and
            fails with 'a server side error has occurred' the moment
            -ErrorAction Stop is on it, typed or splatted. -ErrorVariable is
            accepted, so errors are still turned into a throw and a caller's
            try/catch behaves as it would with -ErrorAction Stop.
        #>
        param(
            [Parameter(Mandatory)][string]$Name,
            [hashtable]$Parameters = @{}
        )
        $p = @{}
        foreach ($k in $Parameters.Keys) { if ($k -ne 'ErrorAction') { $p[$k] = $Parameters[$k] } }
        $cmd = Get-EXOCommand -Name $Name
        $ev  = $null
        $out = & $cmd @p -ErrorVariable ev 2>$null
        if ($ev -and @($ev).Count -gt 0) { throw @($ev)[0] }
        return $out
    }
}

if (-not (Get-Command -Name 'Get-EXOSessionReport' -ErrorAction SilentlyContinue)) {
    function Get-EXOSessionReport {
        <#
            This process's Exchange Online plumbing as text, for pasting into
            an error message: every connection with its token state and
            module, which module the map chose, and what a given cmdlet name
            resolves to. 'A server side error has occurred' says nothing on
            its own; this says which of the usual three it is.
        #>
        param([string]$CommandName)
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('EXCHANGE ONLINE SESSION IN THIS PROCESS')
        try {
            $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
            if (-not $gci) { [void]$sb.AppendLine('  Get-ConnectionInformation is not available (module v2 or older).') }
            else {
                $conns = @(Get-ConnectionInformation -ErrorAction Stop)
                if ($conns.Count -eq 0) { [void]$sb.AppendLine('  No Exchange Online connection at all.') }
                foreach ($c in $conns) {
                    $mn = $(if ($c.PSObject.Properties['ModuleName']) { "$($c.ModuleName)" } else { '?' })
                    [void]$sb.AppendLine("  $($c.UserPrincipalName)   token=$($c.TokenStatus)   state=$($c.State)   module=$mn")
                }
            }
        } catch { [void]$sb.AppendLine("  Could not list connections: $($_.Exception.Message)") }
        try {
            $map = Get-ExchangeModuleMap
            [void]$sb.AppendLine("  Module map: cloud=$(if ($map.Cloud) { $map.Cloud.Name } else { '(none)' })  onprem=$(if ($map.OnPrem) { $map.OnPrem.Name } else { '(none)' })")
        } catch { }
        if ($CommandName) {
            try {
                $cmd = Get-EXOCommand -Name $CommandName
                [void]$sb.AppendLine("  $CommandName resolves to: $($cmd.ModuleName)\$($cmd.Name)  ($($cmd.CommandType))")
            } catch { [void]$sb.AppendLine("  $CommandName does not resolve: $($_.Exception.Message)") }
        }
        return $sb.ToString()
    }
}
#endregion

# NO PROMPT BEFORE THE WINDOW, AND NO ADDRESS BAR EITHER.
#
# The other tools grew an Add-ToolTargetBar row because their target was fixed
# when the form was built. This one never was: the form already opens with
# 'Current organizer' as an editable box, because the whole job is moving a
# meeting FROM one mailbox TO another and both ends have to be typeable.
# Adding a bar would put a second address field above the one that already
# does the work.
#
# So all that changes is the modal in front. An empty address is fine -- the
# box opens blank and the Find/Transfer buttons already refuse an empty From.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Data -------------------------------------------------------------------

function Test-OrganizerMailbox {
    <#
        Both ends must be USER mailboxes in this tenant.

        The cmdlet refuses shared and room mailboxes, and the failure arrives as
        a server-side error after the operator has already filled the form in.
        Checking first turns that into a sentence naming the actual mailbox type
        and what to do instead.
    #>
    param([Parameter(Mandatory)][string]$Smtp)

    $out = [PSCustomObject]@{ Ok = $false; Smtp = ''; Display = ''; Type = ''; Reason = '' }
    try {
        $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop'
        }
    } catch {
        $out.Reason = "not found in this tenant: $($_.Exception.Message)"
        return $out
    }
    $out.Smtp    = "$($mb.PrimarySmtpAddress)"
    $out.Display = "$($mb.DisplayName)"
    $out.Type    = "$($mb.RecipientTypeDetails)"
    if ($out.Type -notmatch '(?i)^UserMailbox$') {
        $out.Reason = ("is a $($out.Type). The cmdlet only transfers meetings on a user's " +
                       "default calendar -- for a shared or room mailbox, give the new person " +
                       "access to the mailbox instead.")
        return $out
    }
    $out.Ok = $true
    return $out
}

function ConvertTo-MeetingFactList {
    <#
        Calendar LOG ROWS IN, one fact object per meeting out.

        EXTRACTED FROM Get-MeetingFacts ON 2026-09-18, unchanged, because the
        sweep that lists everything a person organises needs exactly the same
        reading of exactly the same rows. Two copies of this would be two
        answers to 'is this a series' within one tool -- the second one wrong
        the first time somebody fixed a property name in the other.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string]$Smtp
    )
    $rows = @($Rows)
    $list = @()
    # 'NotFound' IS THE LOG'S NULL. Every absent value comes back as that
    # string, so a plain truthiness test reads it as present. Normalised here.
    $prop = {
        param($o, [string]$n)
        if (-not $o -or -not $o.PSObject.Properties[$n]) { return $null }
        $v = $o.$n
        if ($null -eq $v -or "$v" -eq 'NotFound') { return $null }
        return $v
    }
    # GROUP THE LOG ROWS PER ITEM, then read each fact from whichever row has
    # it. The newest row is not always the fullest. Property names as seen on
    # a real row 2026-09-15: SubjectProperty, StartTime, EndTime,
    # RecurrencePattern (the sentence Outlook shows), IsRecurring,
    # IsSeriesCancelled, AppointmentState (decoded text, 'Meeting, Received'),
    # From, CleanGlobalObjectId. This block is information for the operator;
    # the cmdlet's own answer is the verdict.
    $groups = @{}
    foreach ($r in $rows) {
        $id = "$(& $prop $r 'CleanGlobalObjectId')"
        if (-not $id) { $id = "$(& $prop $r 'EntryId')" }
        if (-not $id) { $id = "$(& $prop $r 'Subject')$(& $prop $r 'NormalizedSubject')" }
        if (-not $groups.ContainsKey($id)) { $groups[$id] = @() }
        $groups[$id] += $r
    }

    $now = Get-Date
    $list = @()
    foreach ($k in $groups.Keys) {
        $g = @($groups[$k])
        $start = $null; $end = $null; $subj = ''; $recurring = $null; $cancelled = $false; $isOrganizer = $null; $pattern = ''
        foreach ($r in $g) {
            # ExDateTime, not DateTime: a cast fails, Parse of the string works
            # (Microsoft's own example on the cmdlet page does the same).
            if (-not $start) { try { $v = & $prop $r 'StartTime'; if ($v) { $start = [DateTime]::Parse("$v") } } catch { } }
            if (-not $end)   { try { $v = & $prop $r 'EndTime';   if ($v) { $end   = [DateTime]::Parse("$v") } } catch { } }
            if (-not $subj) {
                foreach ($n in @('SubjectProperty', 'NormalizedSubject', 'Subject')) {
                    $v = "$(& $prop $r $n)"; if ($v) { $subj = $v; break }
                }
            }
            # Recurrence: the pattern sentence is the best evidence, then the
            # flag, then the blob. Any one of them present means a series.
            $pat = "$(& $prop $r 'RecurrencePattern')"
            if ($pat) { $pattern = $pat; $recurring = $true }
            $ir = & $prop $r 'IsRecurring'
            if ($null -ne $ir -and [bool]$ir) { $recurring = $true }
            if (& $prop $r 'AppointmentRecurrenceBlob') { $recurring = $true }
            $cit = "$(& $prop $r 'CalendarItemType')"
            if ($cit -match '(?i)RecurringMaster|Occurrence|Exception') { $recurring = $true }
            elseif ($cit -match '(?i)^Single$' -and $null -eq $recurring) { $recurring = $false }
            # Cancelled: decoded text on AppointmentState, or the series flag.
            $state = "$(& $prop $r 'AppointmentState')"
            if ($state -match '(?i)cancel') { $cancelled = $true }
            $isc = & $prop $r 'IsSeriesCancelled'
            if ($null -ne $isc -and [bool]$isc) { $cancelled = $true }
            # Organizer: the item's From is this mailbox. ResponseType is
            # usually NotFound on these rows, so it is not relied on.
            $from = "$(& $prop $r 'From')"
            if ($from -and $null -eq $isOrganizer) { $isOrganizer = ($from -match [regex]::Escape($Smtp)) }
        }

        $problems = @()
        if ($cancelled) { $problems += 'Looks CANCELLED -- there may be nothing to transfer.' }
        if ($recurring -eq $false) {
            $problems += 'Looks like a SINGLE OCCURRENCE. The cmdlet moves a series from its next instance.'
        }
        if ($end -and $end -le $now -and $recurring -eq $false) {
            $problems += "Already over (ended $($end.ToString('yyyy-MM-dd HH:mm')))."
        }
        if ($isOrganizer -eq $false) {
            $problems += 'This mailbox looks like an ATTENDEE, not the organizer.'
        }

        $list += [PSCustomObject]@{
            # The GOID. Added 2026-09-18 for the sweep, which dedupes on it
            # across overlapping query slices and shows it so a meeting can be
            # chased with Get-CalendarDiagnosticObjects -MeetingId.
            Id          = $k
            Subject     = $subj
            Pattern     = $pattern
            Start       = $start
            End         = $end
            Recurring   = $recurring
            Cancelled   = $cancelled
            IsOrganizer = $isOrganizer
            Problems    = @($problems)
        }
    }
    return @($list)
}

function Get-MeetingFacts {
    <#
        WHAT IS ACTUALLY ON THE CALENDAR, before the transfer cmdlet is asked.

        Invoke-ChangeMeetingOrganizer reports 'a server side error has occurred'
        for a subject that matches nothing, a single occurrence, a meeting that
        has already happened and a cancelled one alike. The first time this tool
        was tested it was against four one-off meetings on the same afternoon,
        and that message was all it had to say. So Find now reads the calendar
        item first and names the problem in plain words.

        Get-CalendarDiagnosticObjects returns the calendar LOG -- one row per
        change to the item -- so rows are grouped by CleanGlobalObjectId and
        the newest row wins. Restricted to IPM.Appointment: the meeting
        messages (IPM.Schedule.Meeting.*) share the subject and are not the
        thing being transferred.

        ADVISORY, NOT A GATE. Every property is read defensively, and anything
        this cannot tell is reported as 'could not tell' rather than guessed.
        The cmdlet keeps the final say; this exists so its one-line error is
        no longer the first thing the operator learns.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$Subject
    )
    $out = [PSCustomObject]@{ Ok = $false; Meetings = @(); Note = '' }
    try {
        # 90-DAY WINDOW on last-modified time, or -ExactMatch walks every
        # calendar item the mailbox has ever held and a busy calendar takes
        # minutes. A meeting somebody is about to transfer was touched recently.
        # NO -ErrorAction: this cmdlet fails with the generic server-side
        # error when it is given one (see Invoke-EXOCommandNoEA).
        # NO ItemClass FILTER. The rows that state the recurrence plainly
        # (RecurrencePattern, IsRecurring) are the meeting-message rows, and
        # an IPM.Appointment-only filter dropped them. Rows are grouped by
        # CleanGlobalObjectId below, so extra rows cost nothing.
        $rows = @(Invoke-EXOCommandNoEA -Name 'Get-CalendarDiagnosticObjects' -Parameters @{
            Identity    = $Smtp
            Subject     = $Subject
            ExactMatch  = $true
            StartDate   = (Get-Date).AddDays(-90)
            EndDate     = (Get-Date).AddDays(1)
            ResultSize  = 1000
        })
    } catch {
        $out.Note = "Calendar could not be read first ($($_.Exception.Message)). Going straight to the cmdlet."
        return $out
    }
    $out.Ok = $true
    if ($rows.Count -eq 0) { return $out }

    $out.Meetings = @(ConvertTo-MeetingFactList -Rows $rows -Smtp $Smtp | Sort-Object Start)
    return $out
}

function Get-OrganizedSeries {
    <#
        EVERY RECURRING SERIES THIS MAILBOX ORGANISES, for the handover that
        happens when somebody leaves.

        WHY IT SWEEPS IN SLICES, AND WHY THE WINDOW IS TWO YEARS.
        --------------------------------------------------------
        Get-CalendarDiagnosticObjects takes no Subject at all -- only Identity
        is mandatory -- so a date range returns the whole calendar log. Two
        facts about that range decide this whole function:

        1. StartDate/EndDate filter on OriginalLastModifiedTime -- when the
           meeting was last EDITED, not when it runs. That is the trap for an
           offboarding sweep: a weekly meeting set up two years ago and never
           touched since is exactly the one nobody else can manage, and a
           90-day window does not return it. Hence two years by default.

        2. ResultSize caps at 1000, and a row is one CHANGE, not one meeting.
           A single busy series can be dozens of rows, so one wide query
           silently truncates and the list looks complete when it is not.

        So the window is cut into slices and each is queried separately: eight
        quarters rather than one two-year call. Each slice has its own 1000-row
        budget, and a slice that fills it is reported rather than hidden --
        Truncated means 'there may be more than this', which is the only
        honest thing to say.

        Rows from all slices are pooled and grouped by GOID exactly once, by
        ConvertTo-MeetingFactList, so a series touched in three different
        quarters is still one row in the answer.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [int]$MonthsBack = 24,
        [int]$SliceMonths = 3,
        [scriptblock]$Progress
    )
    $out = [PSCustomObject]@{ Ok = $false; Series = @(); Truncated = $false; Slices = 0; Rows = 0; Note = '' }

    $all  = New-Object System.Collections.ArrayList
    $end  = (Get-Date).Date.AddDays(1)
    $stop = (Get-Date).Date.AddMonths(-1 * $MonthsBack)
    $n    = 0
    $total = [Math]::Max(1, [Math]::Ceiling($MonthsBack / [double]$SliceMonths))

    while ($end -gt $stop) {
        $start = $end.AddMonths(-1 * $SliceMonths)
        if ($start -lt $stop) { $start = $stop }
        $n++
        if ($Progress) {
            & $Progress ("Reading the calendar log, slice $n of $total  ($($start.ToString('yyyy-MM-dd')) to $($end.ToString('yyyy-MM-dd')))...")
        }
        $rows = @()
        try {
            # NO -Subject: that is the whole point. NO -ErrorAction: this
            # cmdlet answers a generic server-side error when given one
            # (see Invoke-EXOCommandNoEA).
            $rows = @(Invoke-EXOCommandNoEA -Name 'Get-CalendarDiagnosticObjects' -Parameters @{
                Identity   = $Smtp
                StartDate  = $start
                EndDate    = $end
                ResultSize = 1000
            })
        } catch {
            # ONE BAD SLICE IS NOT A FAILED SWEEP. The window is cut into
            # pieces precisely so a heavy quarter cannot sink the whole run;
            # the note says which piece was lost so the list is not read as
            # complete.
            $out.Note += "Slice $($start.ToString('yyyy-MM-dd')) to $($end.ToString('yyyy-MM-dd')) failed: $($_.Exception.Message). "
            $end = $start
            continue
        }
        if ($rows.Count -ge 1000) { $out.Truncated = $true }
        foreach ($r in $rows) { [void]$all.Add($r) }
        $end = $start
    }

    $out.Ok     = $true
    $out.Slices = $n
    $out.Rows   = $all.Count
    if ($all.Count -eq 0) { return $out }

    $facts = @(ConvertTo-MeetingFactList -Rows $all.ToArray() -Smtp $Smtp)

    # WHAT SURVIVES THE FILTER, AND WHY EACH TEST IS THE WAY ROUND IT IS.
    #
    #   Recurring -eq $true   A single meeting cannot be transferred at all,
    #                         so listing one would be offering something that
    #                         always fails. NOT '-ne $false' -- an item the
    #                         log could not classify ($null) is excluded too,
    #                         because a maybe in a tick-list is a trap.
    #   not Cancelled         Nothing left to move.
    #   IsOrganizer -ne $false  Kept when the log did not say. The From
    #                         property is absent on some rows, and dropping
    #                         everything unproven would hide real series on a
    #                         mailbox whose log is sparse. The transfer
    #                         refuses an attendee's copy anyway.
    #   Subject non-empty     Subject is how each one is transferred below.
    $out.Series = @($facts |
        Where-Object { $_.Recurring -eq $true -and -not $_.Cancelled -and $_.IsOrganizer -ne $false -and "$($_.Subject)".Trim() } |
        Sort-Object Subject)
    return $out
}

function Invoke-SeriesTransfer {
    <#
        ONE SERIES, BY ITS EXACT SUBJECT, WITH THE AMBIGUITY HANDLED.

        THE SWEEP CANNOT HAND THE CMDLET AN ID, and that is not an oversight.
        The calendar log identifies a meeting by its GOID (CleanGlobalObjectId);
        Invoke-ChangeMeetingOrganizer takes EventId -- the AAMk... store id --
        or Subject, and nothing converts one to the other. The documented
        parameter sets are ByEventId and BySubject, so Subject is the only key
        a swept list can supply.

        Subject is a SUBSTRING match, which is the danger: transferring
        'Weekly Sync' would also catch 'Weekly Sync - Engineering'. The cmdlet
        itself is what saves this -- given an ambiguous subject it transfers
        NOTHING and returns the candidates, each with an EventId. So:

            one match      the cmdlet does it, and returns NewMeetingId
            several        pick the candidate whose subject is EXACTLY ours and
                           re-run by EventId; if two share the exact subject,
                           refuse and say so rather than guess

        Returns Status of Moved, Ambiguous, NotMoved or Failed, and a line the
        caller can print. It logs its own outcome either way.
    #>
    param(
        [Parameter(Mandatory)][string]$FromSmtp,
        [Parameter(Mandatory)][string]$ToSmtp,
        [Parameter(Mandatory)][string]$Subject,
        [datetime]$TransferFrom,
        [string]$When = 'from the next instance'
    )
    $res = [PSCustomObject]@{ Subject = $Subject; Status = 'Failed'; Detail = ''; NewId = '' }

    $p = @{ Identity = $FromSmtp; NewOrganizer = $ToSmtp; Subject = $Subject }
    if ($PSBoundParameters.ContainsKey('TransferFrom')) { $p['TransferSeriesStartDate'] = $TransferFrom }

    $moved = {
        param($Result)
        foreach ($r in @($Result)) {
            if ($r -and $r.PSObject.Properties['NewMeetingId'] -and "$($r.NewMeetingId)") { return "$($r.NewMeetingId)" }
        }
        return ''
    }

    try {
        $r1 = Invoke-EXOCommandNoEA -Name 'Invoke-ChangeMeetingOrganizer' -Parameters $p
        $newId = & $moved $r1
        if ($newId) {
            $res.Status = 'Moved'; $res.NewId = $newId
            $res.Detail = "moved $When"
        } else {
            # Candidates came back, or nothing did. Rows without NewMeetingId
            # are the cmdlet declining to act on an ambiguous subject.
            $cands = @($r1 | Where-Object { $_ -and ($_.PSObject.Properties['EventId'] -or $_.PSObject.Properties['MeetingId']) })
            if ($cands.Count -eq 0) {
                $res.Status = 'NotMoved'
                $res.Detail = 'the cmdlet returned nothing and did not confirm a transfer'
            } else {
                $exact = @($cands | Where-Object { "$($_.Subject)" -eq $Subject })
                if ($exact.Count -eq 1 -and "$($exact[0].EventId)") {
                    $p2 = @{ Identity = $FromSmtp; NewOrganizer = $ToSmtp; EventId = "$($exact[0].EventId)" }
                    if ($PSBoundParameters.ContainsKey('TransferFrom')) { $p2['TransferSeriesStartDate'] = $TransferFrom }
                    $r2 = Invoke-EXOCommandNoEA -Name 'Invoke-ChangeMeetingOrganizer' -Parameters $p2
                    $newId2 = & $moved $r2
                    # A quiet return is this cmdlet's success on an id-targeted
                    # transfer; rows back WITHOUT NewMeetingId are a refusal.
                    if ($newId2) {
                        $res.Status = 'Moved'; $res.NewId = $newId2; $res.Detail = "moved $When (by id, subject was shared)"
                    } elseif (@($r2).Count -eq 0) {
                        $res.Status = 'Moved'; $res.Detail = "moved $When (by id; the cmdlet returned no detail)"
                    } else {
                        $res.Status = 'NotMoved'; $res.Detail = 'the cmdlet returned candidates again instead of transferring'
                    }
                } else {
                    $res.Status  = 'Ambiguous'
                    $res.Detail  = "$($cands.Count) meetings match this subject" +
                                   $(if ($exact.Count -gt 1) { " and $($exact.Count) have it exactly" } else { '' }) +
                                   ' -- transfer it by EventId on the main window'
                }
            }
        }
    } catch {
        $res.Status = 'Failed'
        $res.Detail = "$($_.Exception.Message)"
    }

    if ($res.Status -eq 'Moved') {
        Write-CompletedLog -Action 'Configure' -Target $FromSmtp -Result 'Success' `
            -Detail ("Meeting organizer changed to $ToSmtp $When (offboarding sweep; series '$Subject')")
    } else {
        Write-FailedLog -Action 'Configure' -Target $FromSmtp `
            -Detail ("Meeting organizer NOT changed to $ToSmtp (offboarding sweep; series '$Subject') -- $($res.Status): $($res.Detail)")
    }
    return $res
}

function Format-MeetingFacts {
    param($Facts, [string]$Subject)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('ON THE CALENDAR FIRST')
    if (-not $Facts.Ok) {
        [void]$sb.AppendLine("  $($Facts.Note)")
        return $sb.ToString()
    }
    $m = @($Facts.Meetings)
    if ($m.Count -eq 0) {
        [void]$sb.AppendLine("  No calendar item with the exact subject '$Subject' modified in the last 90 days on this mailbox.")
        [void]$sb.AppendLine('  Check the spelling, or whether the meeting lives on a different mailbox.')
        return $sb.ToString()
    }
    [void]$sb.AppendLine("  $($m.Count) calendar item(s) match '$Subject':")
    foreach ($x in $m) {
        $when = $(if ($x.Start) { $x.Start.ToString('ddd yyyy-MM-dd HH:mm') } else { '(no start time)' })
        $kind = $(if ($x.Recurring -eq $true) { 'series' } elseif ($x.Recurring -eq $false) { 'single occurrence' } else { 'series/single not stated in the log' })
        $who  = $(if ($x.IsOrganizer -eq $true) { 'organizer' } elseif ($x.IsOrganizer -eq $false) { 'attendee copy' } else { '' })
        $line = "  - $when  $kind"
        if ($who) { $line += "  ($who)" }
        if ($x.Pattern) { $line += [Environment]::NewLine + "      $($x.Pattern)" }
        if ($x.Cancelled) { $line += '  CANCELLED' }
        [void]$sb.AppendLine($line)
        foreach ($pr in $x.Problems) { [void]$sb.AppendLine("      ! $pr") }
    }
    [void]$sb.AppendLine('  (From the calendar log, for information. The cmdlet''s answer below is the one that counts.)')
    return $sb.ToString()
}

#endregion

#region UI ---------------------------------------------------------------------

function Show-SeriesPicker {
    <#
        The swept list, with a tick box on each row. Returns the ticked series,
        or an empty array when the operator backs out.

        NOTHING IS TICKED WHEN IT OPENS. A handover is a decision per meeting,
        not a default, and a pre-ticked list of twenty series with an OK button
        under it is how somebody transfers a meeting that was supposed to die
        with the role.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Series,
        [Parameter(Mandatory)][string]$FromLabel,
        [Parameter(Mandatory)][string]$ToLabel,
        [string]$Note = ''
    )
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Series organised by this mailbox'
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    $dlg.MinimizeBox = $false
    $dlg.MaximizeBox = $false

    $y = 12
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "$($Series.Count) recurring series on $FromLabel. Tick the ones to hand to $ToLabel."
    $lbl.Location = New-Object System.Drawing.Point(12, $y)
    $lbl.Size = New-Object System.Drawing.Size(960, 18)
    $dlg.Controls.Add($lbl); $y += 22

    if ($Note) {
        $lblNote = New-Object System.Windows.Forms.Label
        $lblNote.Text = $Note
        $lblNote.Location = New-Object System.Drawing.Point(12, $y)
        $lblNote.Size = New-Object System.Drawing.Size(960, 32)
        $lblNote.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
        $dlg.Controls.Add($lblNote); $y += 36
    }

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, $y)
    $grid.Size = New-Object System.Drawing.Size(960, 360)
    $grid.AllowUserToAddRows = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $true
    $grid.AutoSizeColumnsMode = 'None'
    $grid.ScrollBars = 'Both'
    $grid.Font = New-Object System.Drawing.Font('Consolas', 9)
    $colPick = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colPick.Name = 'Pick'; $colPick.HeaderText = ''; $colPick.Width = 34
    [void]$grid.Columns.Add($colPick)
    foreach ($c in @(@('Subject', 300), @('Recurrence', 360), @('First instance', 130), @('Id', 120))) {
        $i = $grid.Columns.Add($c[0], $c[0])
        $grid.Columns[$i].Width = $c[1]
        $grid.Columns[$i].ReadOnly = $true
        $grid.Columns[$i].SortMode = 'Automatic'
    }
    foreach ($m in $Series) {
        $when = $(if ($m.Start) { $m.Start.ToString('yyyy-MM-dd HH:mm') } else { '' })
        # The GOID shortened: it is 100+ characters and the column exists to
        # recognise a row, not to be copied. The full value is in the export.
        $shortId = "$($m.Id)"
        if ($shortId.Length -gt 16) { $shortId = $shortId.Substring($shortId.Length - 16) }
        $r = $grid.Rows.Add($false, "$($m.Subject)", "$($m.Pattern)", $when, $shortId)
        $grid.Rows[$r].Tag = $m
    }
    $dlg.Controls.Add($grid); $y += 370

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = 'Tick all'; $btnAll.Location = New-Object System.Drawing.Point(12, $y)
    $btnAll.Size = New-Object System.Drawing.Size(100, 30)
    $btnAll.Add_Click({ foreach ($r in $grid.Rows) { $r.Cells['Pick'].Value = $true } })
    $dlg.Controls.Add($btnAll)

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = 'Tick none'; $btnNone.Location = New-Object System.Drawing.Point(120, $y)
    $btnNone.Size = New-Object System.Drawing.Size(100, 30)
    $btnNone.Add_Click({ foreach ($r in $grid.Rows) { $r.Cells['Pick'].Value = $false } })
    $dlg.Controls.Add($btnNone)

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Location = New-Object System.Drawing.Point(232, ($y + 7))
    $lblCount.Size = New-Object System.Drawing.Size(400, 18)
    $lblCount.ForeColor = [System.Drawing.Color]::DimGray
    $dlg.Controls.Add($lblCount)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Transfer ticked'; $btnOk.Location = New-Object System.Drawing.Point(712, $y)
    $btnOk.Size = New-Object System.Drawing.Size(130, 30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'; $btnCancel.Location = New-Object System.Drawing.Point(852, $y)
    $btnCancel.Size = New-Object System.Drawing.Size(120, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel

    # COUNT THE TICKS LIVE. CurrentCellDirtyStateChanged then CommitEdit is the
    # WinForms dance a checkbox column needs: without the commit the value only
    # lands when the cell loses focus, so the count -- and OK's enabled state --
    # would lag a click behind what is on screen.
    $refresh = {
        $n = 0
        foreach ($r in $grid.Rows) { if ($r.Cells['Pick'].Value -eq $true) { $n++ } }
        $lblCount.Text = "$n ticked"
        $btnOk.Enabled = ($n -gt 0)
    }
    $grid.Add_CurrentCellDirtyStateChanged({
        if ($grid.IsCurrentCellDirty) { $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
    })
    $grid.Add_CellValueChanged({ & $refresh })
    $btnAll.Add_Click({ & $refresh })
    $btnNone.Add_Click({ & $refresh })
    & $refresh

    $dlg.ClientSize = New-Object System.Drawing.Size(($grid.Right + 12), ($btnOk.Bottom + 12))
    Set-ConsoleTheme -Form $dlg

    $picked = @()
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($r in $grid.Rows) {
            if ($r.Cells['Pick'].Value -eq $true -and $r.Tag) { $picked += $r.Tag }
        }
    }
    $dlg.Dispose()
    return @($picked)
}

function Show-ChangeOrganizerForm {
    # NO -Smtp PARAMETER. $txtFrom below is the address, and it was always
    # editable -- a parameter fixed at build time only ever seeded it.

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Change Meeting Organizer'
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    function New-Lbl { param([string]$T, [int]$Y, [int]$W = 820)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $T; $l.Location = New-Object System.Drawing.Point(12, $Y)
        $l.Size = New-Object System.Drawing.Size($W, 18)
        $form.Controls.Add($l); return $l
    }

    [void](New-Lbl 'Current organizer (the mailbox the meeting is on):' $y); $y += 20
    $txtFrom = New-Object System.Windows.Forms.TextBox
    $txtFrom.Location = New-Object System.Drawing.Point(12, $y)
    $txtFrom.Size = New-Object System.Drawing.Size(400, 24)
    $txtFrom.Text = $script:Target
    $form.Controls.Add($txtFrom); $y += 32

    [void](New-Lbl 'New organizer (address or alias):' $y); $y += 20
    $txtTo = New-Object System.Windows.Forms.TextBox
    $txtTo.Location = New-Object System.Drawing.Point(12, $y)
    $txtTo.Size = New-Object System.Drawing.Size(400, 24)
    $form.Controls.Add($txtTo); $y += 32

    [void](New-Lbl 'Meeting subject:' $y); $y += 20
    $txtSubject = New-Object System.Windows.Forms.TextBox
    $txtSubject.Location = New-Object System.Drawing.Point(12, $y)
    $txtSubject.Size = New-Object System.Drawing.Size(600, 24)
    $form.Controls.Add($txtSubject); $y += 32

    [void](New-Lbl 'MeetingId or EventId -- fill this in only if Find reported several meetings with that subject:' $y); $y += 20
    $txtEventId = New-Object System.Windows.Forms.TextBox
    $txtEventId.Location = New-Object System.Drawing.Point(12, $y)
    $txtEventId.Size = New-Object System.Drawing.Size(820, 24)
    $txtEventId.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtEventId); $y += 34

    $chkDate = New-Object System.Windows.Forms.CheckBox
    $chkDate.Text = 'Transfer from a specific date instead of the next instance'
    $chkDate.Location = New-Object System.Drawing.Point(12, $y)
    $chkDate.Size = New-Object System.Drawing.Size(420, 20)
    $form.Controls.Add($chkDate)

    $dtStart = New-Object System.Windows.Forms.DateTimePicker
    $dtStart.Location = New-Object System.Drawing.Point(440, ($y - 2))
    $dtStart.Size = New-Object System.Drawing.Size(180, 24)
    $dtStart.Format = 'Short'
    $dtStart.MinDate = (Get-Date).Date.AddDays(1)
    $dtStart.Value = (Get-Date).Date.AddDays(1)
    $dtStart.Enabled = $false
    $form.Controls.Add($dtStart)
    $chkDate.Add_CheckedChanged({ $dtStart.Enabled = $chkDate.Checked })
    $y += 30

    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = 'Find (-WhatIf)'; $btnPreview.Location = New-Object System.Drawing.Point(12, $y)
    $btnPreview.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnPreview)

    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = 'Transfer'; $btnGo.Location = New-Object System.Drawing.Point(150, $y)
    $btnGo.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnGo)

    $btnSweep = New-Object System.Windows.Forms.Button
    $btnSweep.Text = 'Everything they organise...'
    $btnSweep.Location = New-Object System.Drawing.Point(288, $y)
    $btnSweep.Size = New-Object System.Drawing.Size(200, 30)
    $form.Controls.Add($btnSweep)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(702, $y)
    $btnClose.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })
    $y += 40

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(12, $y)
    $txtOut.Size = New-Object System.Drawing.Size(820, 330)
    $txtOut.Multiline = $true
    $txtOut.ScrollBars = 'Both'
    $txtOut.WordWrap = $false
    $txtOut.ReadOnly = $true
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)

    # SIZED FROM THE CONTENT, not from a number typed at the top of the
    # function. Every row above is placed by a running $y, so a literal height
    # is stale the moment a row is inserted -- and WinForms does not complain,
    # it just puts the bottom of the window past the bottom of the client area.
    #
    # ClientSize, not Size: the controls are positioned in client coordinates,
    # and Size would silently spend ~39px of this on the title bar and borders.
    $form.ClientSize = New-Object System.Drawing.Size(
        ($txtOut.Right + 12), ($txtOut.Bottom + 12))

    $warn = New-Object System.Text.StringBuilder
    [void]$warn.AppendLine('BEFORE YOU TRANSFER')
    [void]$warn.AppendLine('')
    [void]$warn.AppendLine('  Teams meetings: the Teams online meeting KEEPS THE OLD ORGANIZER.')
    [void]$warn.AppendLine('  The new organizer cannot manage or start it and must replace the')
    [void]$warn.AppendLine('  online meeting details by hand afterwards.')
    [void]$warn.AppendLine('')
    [void]$warn.AppendLine('  Attendees in this tenant  : silently updated, no re-RSVP.')
    [void]$warn.AppendLine('  External / on-premises    : cancellation then a new invitation.')
    [void]$warn.AppendLine('                              They must re-RSVP.')
    [void]$warn.AppendLine('')
    [void]$warn.AppendLine('  ONE WAY, ONCE. The part of the series left on the original calendar')
    [void]$warn.AppendLine('  cannot be transferred again.')
    [void]$warn.AppendLine('')
    [void]$warn.AppendLine('DO NOT TREAT "Find" AS A DRY RUN.')
    [void]$warn.AppendLine('')
    [void]$warn.AppendLine('  It passes -WhatIf, which Microsoft documents as making no changes.')
    [void]$warn.AppendLine('  But -Subject is a SUBSTRING match, not a search: ''abcdefgh'' also matches')
    [void]$warn.AppendLine('  ''new_abcdefgh'' and ''abcdefghij''. With exactly ONE match the cmdlet acts,')
    [void]$warn.AppendLine('  and there are reports that in non-interactive sessions -WhatIf returns')
    [void]$warn.AppendLine('  no message while the transfer still happens. That is unconfirmed and')
    [void]$warn.AppendLine('  untested here, so it is treated as possible.')
    [void]$warn.AppendLine('')
    [void]$warn.AppendLine('  Safe use: give a subject you EXPECT to match several meetings, or one')
    [void]$warn.AppendLine('  you are willing to transfer. Use the returned EventId for the real run.')
    $txtOut.Text = $warn.ToString()

    # ---- shared validation -------------------------------------------------
    $validate = {
        $from = "$($txtFrom.Text)".Trim()
        $to   = "$($txtTo.Text)".Trim()
        $subj = "$($txtSubject.Text)".Trim()
        $eid  = "$($txtEventId.Text)".Trim()

        if (-not $from) { return @{ Ok = $false; Msg = 'The current organizer is required.' } }
        if (-not $to)   { return @{ Ok = $false; Msg = 'The new organizer is required.' } }
        if (-not $subj -and -not $eid) {
            return @{ Ok = $false; Msg = 'Give either a subject or an EventId.' }
        }

        $f = Test-OrganizerMailbox -Smtp $from
        if (-not $f.Ok) { return @{ Ok = $false; Msg = "Current organizer '$from' $($f.Reason)" } }
        $t = Test-OrganizerMailbox -Smtp $to
        if (-not $t.Ok) { return @{ Ok = $false; Msg = "New organizer '$to' $($t.Reason)" } }

        # NO ErrorAction HERE. Invoke-ChangeMeetingOrganizer works bare and
        # fails with 'a server side error has occurred' the moment -ErrorAction
        # is on it, typed or splatted -- proven at a prompt 2026-09-15 after
        # three rounds of chasing sessions. Errors come back through
        # Invoke-EXOCommandNoEA's -ErrorVariable and are thrown from there.
        $p = @{ Identity = $f.Smtp; NewOrganizer = $t.Smtp }

        # THE ID WINS when both are filled in. The two are separate parameter
        # sets and passing both is a binding error -- and the id is the precise
        # one, chosen deliberately after a Find reported several matches.
        #
        # EITHER ID IS ACCEPTED. The cmdlet's own multiple-match output prints
        # a MeetingId (a hex GOID, 040000008200E00074C5B7101A82E008...) AND an
        # EventId (base64, AAMkAGEz...), and its instructions say -MeetingId
        # while Microsoft Learn documents only -EventId. Whichever was pasted
        # is told apart by shape and sent under the matching name; -MeetingId
        # is only used if this module's copy of the cmdlet actually has it.
        if ($eid) {
            $isGoid = ($eid -match '^[0-9A-Fa-f]{40,}$')
            $hasMeetingId = $false
            try { $hasMeetingId = (Get-EXOCommand -Name 'Invoke-ChangeMeetingOrganizer').Parameters.ContainsKey('MeetingId') } catch { }
            if ($isGoid -and $hasMeetingId) { $p['MeetingId'] = $eid }
            elseif ($isGoid) { return @{ Ok = $false; Msg = 'That looks like a MeetingId, and this version of the Exchange Online module has no -MeetingId parameter. Paste the EventId (the long AAMk... value) instead.' } }
            else { $p['EventId'] = $eid }
        } else { $p['Subject'] = $subj }

        if ($chkDate.Checked) {
            # MinDate on the picker already blocks the past, but the cmdlet is
            # the thing that enforces it and this keeps the message ours.
            if ($dtStart.Value.Date -le (Get-Date).Date) {
                return @{ Ok = $false; Msg = 'The transfer date must be in the future.' }
            }
            $p['TransferSeriesStartDate'] = $dtStart.Value.Date
        }
        return @{ Ok = $true; Params = $p; From = $f; To = $t; UsedEventId = [bool]$eid }
    }

    $render = {
        param($Result, [string]$Header)
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine($Header)
        [void]$sb.AppendLine('')
        if ($null -eq $Result -or @($Result).Count -eq 0) {
            [void]$sb.AppendLine('  The cmdlet returned nothing.')
        } else {
            foreach ($r in @($Result)) {
                # -Width 4000: Out-String wraps at the console's 80 columns and
                # split a 120-character MeetingId over two lines, which cannot
                # be copied with one click. The box does not wrap either (below),
                # so an id is always exactly one line: triple-click, copy.
                [void]$sb.AppendLine('  ' + ($r | Format-List | Out-String -Width 4000).Trim().Replace("`n", "`n  "))
                [void]$sb.AppendLine('')
            }
        }
        return $sb.ToString()
    }

    $btnPreview.Add_Click({
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $v = & $validate
            if (-not $v.Ok) { $txtOut.Text = $v.Msg; return }

            # READ THE CALENDAR FIRST and say what is there. By subject only:
            # an EventId came from a previous Find and has already been seen.
            $facts = ''
            if (-not $v.UsedEventId) {
                $txtOut.Text = 'Reading the calendar...'
                $txtOut.Refresh()
                $f = Get-MeetingFacts -Smtp $v.From.Smtp -Subject "$($txtSubject.Text)".Trim()
                $facts = (Format-MeetingFacts -Facts $f -Subject "$($txtSubject.Text)".Trim()) + [Environment]::NewLine
                if ($f.Ok -and @($f.Meetings).Count -eq 0) {
                    # Nothing to ask the cmdlet about; its answer would be the
                    # generic error this pre-check exists to replace.
                    $txtOut.Text = $facts
                    return
                }
            }

            # -WhatIf, because -Subject IS NOT A SEARCH. With exactly one match
            # the cmdlet performs the transfer there and then. A "find the
            # meeting" button that silently transfers it is the worst possible
            # arrangement, so the only safe probe is the one that promises to
            # change nothing.
            $p = $v.Params.Clone()
            $p['WhatIf'] = $true
            $res = Invoke-EXOCommandNoEA -Name 'Invoke-ChangeMeetingOrganizer' -Parameters $p

            $head = ("PREVIEW -- nothing has been changed" + [Environment]::NewLine +
                     "  From : $($v.From.Display) <$($v.From.Smtp)>" + [Environment]::NewLine +
                     "  To   : $($v.To.Display) <$($v.To.Smtp)>" + [Environment]::NewLine +
                     "  By   : " + $(if ($v.UsedEventId) { 'EventId' } else { 'Subject' }))
            if ($null -eq $res -or @($res).Count -eq 0) {
                # NOTHING BACK IS THE GOOD ANSWER. -WhatIf writes 'What if: ...'
                # to the console host, not the pipeline, so exactly one match
                # produces no objects at all. Only the several-matches case
                # returns a list. Saying 'the cmdlet returned nothing' here
                # read as failure and was not.
                $txt = $head + [Environment]::NewLine + [Environment]::NewLine +
                       '  EXACTLY ONE meeting matched. The cmdlet accepted it and would transfer it.' + [Environment]::NewLine +
                       '  Press Transfer to do it for real.'
            } else {
                $txt = & $render $res $head
                $txt += [Environment]::NewLine +
                        'Several meetings matched. Copy the MeetingId or EventId of the right one' +
                        [Environment]::NewLine + 'into the id box above and Find again.'
            }
            $txtOut.Text = $facts + $txt
        } catch {
            # The service returns one generic "server side error" for several
            # different causes, so the message alone tells the operator nothing.
            # These are the ones worth checking first, in rough order of how
            # often they turn out to be it.
            $msg = Format-ErrorDetail -ErrorRecord $_ -Context 'Looking up the meeting'
            if ($facts) { $msg = $facts + [Environment]::NewLine + $msg }
            if ("$($_.Exception.Message)" -match '(?i)server side error') {
                # THE SESSION REPORT COMES FIRST. This message has meant three
                # different things in one week and the report told them apart
                # every time; the checklist below is for when it does not.
                try { $msg += [Environment]::NewLine + [Environment]::NewLine + (Get-EXOSessionReport -CommandName 'Invoke-ChangeMeetingOrganizer') } catch { }
                $msg += [Environment]::NewLine + [Environment]::NewLine +
                    'That message is generic. Things to rule out, in order:' + [Environment]::NewLine +
                    '  1. No meeting matches that subject on that mailbox (default calendar only).' + [Environment]::NewLine +
                    '  0. THE SESSION. If Get-Mailbox works but calendar cmdlets fail with this, the' + [Environment]::NewLine +
                    '     process has expired Exchange Online connections beside the live one,' + [Environment]::NewLine +
                    '     or the cmdlet was given -ErrorAction (this tool no longer does).' + [Environment]::NewLine +
                    '  2. The meeting is in the past, or is a single occurrence rather than a' + [Environment]::NewLine +
                    '     series. Only future instances of a series can move.' + [Environment]::NewLine +
                    '  3. The series has already been transferred once. The remainder on the' + [Environment]::NewLine +
                    '     original calendar cannot be transferred again.' + [Environment]::NewLine +
                    '  4. -WhatIf itself. Try the same lookup from a plain PowerShell prompt to' + [Environment]::NewLine +
                    '     see whether the cmdlet works at all in this tenant.' + [Environment]::NewLine + [Environment]::NewLine +
                    'The organizer does NOT need to be disabled -- the mailbox must be readable.'
            }
            $txtOut.Text = $msg
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $btnGo.Add_Click({
        $v = & $validate
        if (-not $v.Ok) { $txtOut.Text = $v.Msg; return }

        $when = $(if ($v.Params.ContainsKey('TransferSeriesStartDate')) {
            "from $($v.Params['TransferSeriesStartDate'].ToShortDateString())"
        } else { 'from the next instance' })

        # WHICH MEETING. The dialog used to say 'this meeting' and never name
        # it, while -Subject is a SUBSTRING match -- 'abcdef' also matches
        # 'new_abcdef'. With exactly one match the cmdlet acts immediately, so
        # that dialog was the only chance to notice the wrong meeting and it
        # withheld the one fact needed to.
        $what = $(if ($v.Params.ContainsKey('EventId')) {
            "Id      : $($v.Params['EventId'])"
        } elseif ($v.Params.ContainsKey('MeetingId')) {
            "Id      : $($v.Params['MeetingId'])"
        } else {
            "Subject : $($v.Params['Subject'])   (substring match -- anything containing this)"
        })

        $a = Show-ToolDialog(
            ("Transfer this meeting $when" + "`r`n`r`n" +
             $what + "`r`n`r`n" +
             "From : $($v.From.Display)" + "`r`n" +
             "To   : $($v.To.Display)" + "`r`n`r`n" +
             "Teams meetings keep their OLD organizer for the online meeting." + "`r`n" +
             "External attendees will be cancelled and re-invited." + "`r`n" +
             "This cannot be undone for the part of the series that moves." + "`r`n`r`n" +
             "Continue?"),
            'Confirm transfer', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $res = Invoke-EXOCommandNoEA -Name 'Invoke-ChangeMeetingOrganizer' -Parameters $v.Params

            # A LIST BACK MEANS NOTHING MOVED. With several subject matches the
            # cmdlet returns the candidates -- each with a MeetingId/EventId --
            # and performs no transfer. The first version printed TRANSFERRED
            # over that list and logged a success; the meeting was still on the
            # organizer's calendar. Told apart here by what came back.
            # NewMeetingId IS THE DISCRIMINATOR, not the presence of an id. A
            # SUCCESSFUL transfer returns NewMeetingId (it is read below), and
            # it may well carry MeetingId/EventId too -- so testing only for
            # those would report a completed transfer as 'NOT TRANSFERRED YET'
            # and offer to do it again. A candidate list has ids and no
            # NewMeetingId on any row.
            $isList = $false
            foreach ($r in @($res)) {
                if (-not $r) { continue }
                $hasNew = $r.PSObject.Properties['NewMeetingId'] -and "$($r.NewMeetingId)"
                if (-not $hasNew -and ($r.PSObject.Properties['MeetingId'] -or $r.PSObject.Properties['EventId'])) { $isList = $true; break }
            }
            if ($isList) {
                $cands = @($res | Where-Object { $_ -and ($_.PSObject.Properties['MeetingId'] -or $_.PSObject.Properties['EventId']) })
                $txtOut.Text = (& $render $res 'NOT TRANSFERRED YET -- several meetings match that subject')
                $txtOut.Refresh()

                # ONE PER CALL IS THE CMDLET'S RULE, NOT THE OPERATOR'S. 'Move
                # everything called X to the new organizer' is the common ask
                # when somebody leaves, so the list is offered back as a batch:
                # one transfer per meeting, by id, each reported and logged on
                # its own line. The confirmation names every subject, because
                # a substring match ('abcdef') can catch meetings that were
                # not meant ('abcdefghijkl'), and this is the last look.
                $names = @($cands | ForEach-Object { "  - $($_.Subject)" }) -join "`r`n"
                $a2 = Show-ToolDialog(
                    ("Transfer ALL $($cands.Count) of these to $($v.To.Display) $when" + "?`r`n`r`n" + $names + "`r`n`r`n" +
                     "Each is a separate transfer. Untick nothing here -- to leave one out, transfer the others by id instead." + "`r`n`r`n" +
                     "This cannot be undone for the part of each series that moves."),
                    'Transfer all matches', 'YesNo', 'Warning', 'Button2')
                if ($a2 -ne [System.Windows.Forms.DialogResult]::Yes) {
                    Write-CompletedLog -Action 'Configure' -Target $v.From.Smtp -Result 'Failed' `
                        -Detail ("Meeting organizer NOT changed: several meetings match subject '$($txtSubject.Text)'; operator declined the batch.")
                    $txtOut.Text += [Environment]::NewLine + 'Nothing transferred. Copy the MeetingId or EventId of the one you want into' +
                        [Environment]::NewLine + 'the id box and press Transfer again.'
                    return
                }

                $hasMeetingId = $false
                try { $hasMeetingId = (Get-EXOCommand -Name 'Invoke-ChangeMeetingOrganizer').Parameters.ContainsKey('MeetingId') } catch { }
                $sb = New-Object System.Text.StringBuilder
                [void]$sb.AppendLine("TRANSFERRING $($cands.Count) meeting(s) to $($v.To.Display) $when")
                [void]$sb.AppendLine('')
                $okCount = 0
                foreach ($c in $cands) {
                    $pp = @{ Identity = $v.From.Smtp; NewOrganizer = $v.To.Smtp }
                    if ($v.Params.ContainsKey('TransferSeriesStartDate')) { $pp['TransferSeriesStartDate'] = $v.Params['TransferSeriesStartDate'] }
                    # EventId is the documented parameter; MeetingId only where the loaded cmdlet has it.
                    $eidC = "$($c.EventId)"; $midC = "$($c.MeetingId)"
                    if ($eidC) { $pp['EventId'] = $eidC } elseif ($midC -and $hasMeetingId) { $pp['MeetingId'] = $midC }
                    else { [void]$sb.AppendLine("  FAIL  $($c.Subject): no usable id in the list."); continue }
                    $txtOut.Text = $sb.ToString() + "  ...   $($c.Subject)"
                    $txtOut.Refresh()
                    try {
                        # THE SAME CHECK AS THE SINGLE PATH, FOR THE SAME REASON.
                        # This loop used to discard the result and count DONE on
                        # anything that did not throw -- which is exactly the bug
                        # fixed above, still open on this branch. A row with ids
                        # and no NewMeetingId means the cmdlet declined to act.
                        $rb = Invoke-EXOCommandNoEA -Name 'Invoke-ChangeMeetingOrganizer' -Parameters $pp
                        $moved = $false
                        foreach ($r in @($rb)) {
                            if ($r -and $r.PSObject.Properties['NewMeetingId'] -and "$($r.NewMeetingId)") { $moved = $true; break }
                        }
                        # No rows at all is the quiet success this cmdlet gives
                        # on an id-targeted transfer; rows WITHOUT NewMeetingId
                        # are a refusal.
                        if (-not $moved -and @($rb).Count -gt 0) {
                            [void]$sb.AppendLine("  NOT MOVED  $($c.Subject): the cmdlet returned candidates instead of transferring.")
                            Write-FailedLog -Action 'Configure' -Target $v.From.Smtp `
                                -Detail "Change meeting organizer to $($v.To.Smtp) (batch; meeting '$($c.Subject)') -- the cmdlet returned a candidate list and transferred nothing"
                            continue
                        }
                        $okCount++
                        [void]$sb.AppendLine("  DONE  $($c.Subject)   ($($c.StartTime))")
                        Write-CompletedLog -Action 'Configure' -Target $v.From.Smtp -Result 'Success' `
                            -Detail ("Meeting organizer changed to $($v.To.Smtp) $when (batch by subject '$($txtSubject.Text)'; meeting '$($c.Subject)')")
                    } catch {
                        [void]$sb.AppendLine("  FAIL  $($c.Subject): $($_.Exception.Message)")
                        Write-FailedLog -Action 'Configure' -Target $v.From.Smtp -ErrorRecord $_ `
                            -Detail "Change meeting organizer to $($v.To.Smtp) (batch; meeting '$($c.Subject)')"
                    }
                }
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine("$okCount of $($cands.Count) transferred. Instances before the effective date stay on the original calendar.")
                [void]$sb.AppendLine('Remember: a Teams online meeting attached to each series still belongs to the')
                [void]$sb.AppendLine('old organizer. The new organizer must replace it.')
                $txtOut.Text = $sb.ToString()
                return
            }

            Write-CompletedLog -Action 'Configure' -Target $v.From.Smtp -Result 'Success' `
                -Detail ("Meeting organizer changed to $($v.To.Smtp) " + $when +
                         $(if ($v.UsedEventId) { " (id)" } else { " (Subject: $($txtSubject.Text))" }))
            # THE NEW ID IN FULL. The table view truncates it, and it is the one
            # value that finds the new series on the new organizer's calendar
            # (Get-CalendarDiagnosticObjects -MeetingId) or goes to support.
            $newIds = @($res | ForEach-Object { if ($_ -and $_.PSObject.Properties['NewMeetingId']) { "$($_.NewMeetingId)" } } | Where-Object { $_ })
            $idLine = $(if ($newIds.Count -gt 0) { 'New MeetingId on the new organizer''s calendar:' + [Environment]::NewLine + '  ' + ($newIds -join ([Environment]::NewLine + '  ')) + [Environment]::NewLine } else { '' })
            $txtOut.Text = (& $render $res 'TRANSFERRED') + $idLine + [Environment]::NewLine +
                'The series from ' + $(if ($v.Params.ContainsKey('TransferSeriesStartDate')) { $v.Params['TransferSeriesStartDate'].ToShortDateString() } else { 'its next instance' }) +
                " is now on $($v.To.Display)'s calendar. Instances before that stay on the original." + [Environment]::NewLine +
                'Remember: a Teams online meeting attached to this series still belongs to the' +
                [Environment]::NewLine + 'old organizer. The new organizer must replace it.'
        } catch {
            Write-FailedLog -Action 'Configure' -Target $v.From.Smtp -ErrorRecord $_ `
                -Detail "Change meeting organizer to $($v.To.Smtp)"
            $msg = Format-ErrorDetail -ErrorRecord $_ -Context 'Transferring the meeting'
            if ("$($_.Exception.Message)" -match '(?i)transfer meeting action is disabled') {
                $msg += [Environment]::NewLine + [Environment]::NewLine +
                        'That message means the feature is not enabled in this tenant yet. ' +
                        'Nothing was changed.'
            }
            $txtOut.Text = $msg
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $btnSweep.Add_Click({
        # THE OFFBOARDING PATH. Subject box ignored on purpose: this asks the
        # calendar what the person organises rather than making somebody
        # remember. Both mailboxes are still validated first, because a sweep
        # that ends at a shared mailbox has wasted a minute for nothing.
        $from = "$($txtFrom.Text)".Trim()
        $to   = "$($txtTo.Text)".Trim()
        if (-not $from) { $txtOut.Text = 'The current organizer is required.'; return }
        if (-not $to)   { $txtOut.Text = 'The new organizer is required -- the sweep transfers, it does not just list.'; return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $f = Test-OrganizerMailbox -Smtp $from
            if (-not $f.Ok) { $txtOut.Text = "Current organizer '$from' $($f.Reason)"; return }
            $t = Test-OrganizerMailbox -Smtp $to
            if (-not $t.Ok) { $txtOut.Text = "New organizer '$to' $($t.Reason)"; return }

            $txtOut.Text = 'Sweeping the calendar log. Two years, a quarter at a time -- this takes a minute.'
            $txtOut.Refresh()
            $report = { param([string]$M) $txtOut.Text = $M; $txtOut.Refresh() }
            $sw = Get-OrganizedSeries -Smtp $f.Smtp -Progress $report

            if (-not $sw.Ok) { $txtOut.Text = "The sweep failed. $($sw.Note)"; return }

            $head = New-Object System.Text.StringBuilder
            [void]$head.AppendLine("SWEPT $($sw.Slices) quarter(s) of $($f.Display)'s calendar log -- $($sw.Rows) log row(s).")
            [void]$head.AppendLine("$(@($sw.Series).Count) recurring series look transferable.")
            if ($sw.Note) { [void]$head.AppendLine("  $($sw.Note)") }
            [void]$head.AppendLine('')

            if (@($sw.Series).Count -eq 0) {
                [void]$head.AppendLine('Nothing to transfer. That means one of:')
                [void]$head.AppendLine('  - they organise no recurring meetings;')
                [void]$head.AppendLine('  - every series they own was last edited more than two years ago, so the')
                [void]$head.AppendLine('    log window did not reach it (the window is on LAST MODIFIED, not on when')
                [void]$head.AppendLine('    the meeting runs);')
                [void]$head.AppendLine('  - the meetings are on a calendar other than the default one.')
                $txtOut.Text = $head.ToString()
                return
            }

            # THE TRUNCATION WARNING IS RED AND IT IS IN THE PICKER, not only
            # here. A list that might be missing rows is worth acting on, but
            # only by somebody who knows it might be missing rows.
            $note = ''
            if ($sw.Truncated) {
                $note = 'At least one quarter returned the maximum 1000 log rows, so this list may be INCOMPLETE. ' +
                        'Check the calendar directly for anything you expected and do not see.'
            }

            $picked = @(Show-SeriesPicker -Series @($sw.Series) -FromLabel $f.Display -ToLabel $t.Display -Note $note)
            if ($picked.Count -eq 0) {
                $txtOut.Text = $head.ToString() + 'Nothing ticked. Nothing was changed.'
                return
            }

            $when = 'from the next instance'
            $useDate = $chkDate.Checked
            if ($useDate) { $when = "from $($dtStart.Value.ToShortDateString())" }

            $names = @($picked | ForEach-Object { "  - $($_.Subject)" }) -join "`r`n"
            $a = Show-ToolDialog(
                ("Transfer $($picked.Count) series from $($f.Display) to $($t.Display) $when" + "?`r`n`r`n" + $names + "`r`n`r`n" +
                 "Each is a separate transfer, by subject." + "`r`n" +
                 "Teams meetings keep their OLD organizer for the online meeting -- the new" + "`r`n" +
                 "organizer must replace those details by hand." + "`r`n" +
                 "External attendees are cancelled and re-invited." + "`r`n`r`n" +
                 "This cannot be undone for the part of each series that moves."),
                'Transfer these series', 'YesNo', 'Warning', 'Button2')
            if ($a -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtOut.Text = $head.ToString() + 'Cancelled at the confirmation. Nothing was changed.'
                return
            }

            $sb = New-Object System.Text.StringBuilder
            [void]$sb.Append($head.ToString())
            [void]$sb.AppendLine("TRANSFERRING $($picked.Count) series to $($t.Display) $when")
            [void]$sb.AppendLine('')
            $ok = 0; $amb = 0; $bad = 0
            foreach ($m in $picked) {
                $txtOut.Text = $sb.ToString() + "  ...   $($m.Subject)"
                $txtOut.Refresh()
                # NOT $args. That is the automatic variable holding a
                # scriptblock's own arguments, and assigning to it inside a
                # handler is legal, confusing, and the sort of thing that
                # breaks the day somebody adds a param block above.
                $callArgs = @{ FromSmtp = $f.Smtp; ToSmtp = $t.Smtp; Subject = "$($m.Subject)"; When = $when }
                if ($useDate) { $callArgs['TransferFrom'] = $dtStart.Value }
                $r = Invoke-SeriesTransfer @callArgs
                switch ($r.Status) {
                    'Moved'     { $ok++;  [void]$sb.AppendLine("  DONE       $($m.Subject)") }
                    'Ambiguous' { $amb++; [void]$sb.AppendLine("  AMBIGUOUS  $($m.Subject)  --  $($r.Detail)") }
                    default     { $bad++; [void]$sb.AppendLine("  $($r.Status.ToUpper().PadRight(10)) $($m.Subject)  --  $($r.Detail)") }
                }
            }
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("$ok transferred, $amb ambiguous, $bad failed.")
            [void]$sb.AppendLine('Instances before the effective date stay on the original calendar.')
            if ($amb -gt 0) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine('AMBIGUOUS means several meetings share that subject and the cmdlet will not')
                [void]$sb.AppendLine('guess. Put the subject in the box above, press Find, and transfer by EventId.')
            }
            if ($ok -gt 0) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine('Every Teams online meeting attached to these series still belongs to the old')
                [void]$sb.AppendLine('organizer. The new organizer has to replace the online meeting details.')
            }
            $txtOut.Text = $sb.ToString()
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Sweeping the calendar for series to transfer'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Change Meeting Organizer' | Out-Null
    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    [void]$form.ShowDialog()
}

#endregion

Show-ChangeOrganizerForm
