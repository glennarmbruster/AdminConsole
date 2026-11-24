#Requires -Version 5.1
<#
    Search_Purge.ps1
    Content Search across Exchange and Teams: build a search, look at what it
    found, and delete some or all of it.

    THREE THINGS ABOUT PURGE THAT DECIDE THIS WHOLE TOOL'S SHAPE.

    1. PURGE HAS NO ITEM SELECTION. New-ComplianceSearchAction -Purge deletes
       THE ENTIRE RESULT SET of the search it is given. There is no parameter
       for "these five messages". So "purge the ones I ticked" cannot be done
       directly, and any tool that appears to offer it is either lying or
       deleting more than was ticked.

       What this tool does instead: it builds a SECOND, NARROWER SEARCH from the
       ticked rows, runs it, and REFUSES TO PURGE IT unless the item count comes
       back equal to the number of rows ticked. A narrowed query built from
       subject, sender and date can match more than was intended -- the same
       subject sent to a dozen people -- and over-deleting is the one outcome
       that cannot be undone. So the count is the gate, and the gate is closed
       by default.

    2. PREVIEW IS CAPPED, AND THE CAP IS NOT THE RESULT SET. A Preview action
       returns a subset (Microsoft's limit, around a thousand items). A list of
       1,000 rows presented as "what will be purged" when the search matched
       50,000 is the most dangerous screen this tool could draw, so the previewed
       count is always compared with the search's own item count and the
       difference is stated in words.

    3. PURGE IS CAPPED PER MAILBOX PER ACTION. Classic Content Search purge
       removes a limited number of items per mailbox per action -- 10 in the
       documented behaviour -- so one purge of a large result set does NOT empty
       it. The confirmation says so, because "Purge completed" followed by the
       items still being there is how an operator concludes the tool is broken
       when it is Microsoft's limit working as designed.

    REFACTORED FROM a standalone GUI script that connected from inside its own
    modal form. That cannot work here and it is not a style question: B2 -- a
    modal ShowDialog blocks every other window on the thread, so an MFA prompt
    raised from a button draws and then ignores all input. Connecting is the
    console's job (C17), so this tool CHECKS the session and names the button
    that makes one.
#>
# TOOLKIT-HELP-RUNS: Reads Content Search results from the Security and Compliance session, and DELETES mail when told to. Writes CSVs.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Find content across mailboxes and Teams, look at what was found, and delete
# it -- all of it, or only the rows you tick.
#
# IT NEEDS THE SECURITY AND COMPLIANCE SESSION, not Exchange Online. They are
# different endpoints that share cmdlet names. If the Security indicator in the
# console is red, press Connect-IPPSSession first; this window will refuse until
# it is green, and say so rather than failing halfway through a purge.
#
# THE THREE TABS ARE A SEQUENCE:
#
#   1. Create a search   -- what to look for and where.
#   2. Searches          -- run one, watch its status, delete one.
#   3. Preview and purge -- see the items, tick some, delete.
#
# TO FIND EVERYTHING FROM ONE PERSON, put their address in "From (sender)". It
# becomes from:"..." in the query, which is shown in full above the button
# before anything runs -- READ IT. If the sender is not in that query, the
# search is not filtered by sender, however many items it returns.
#
# "Custom start / end" ARE DATES AND NOTHING ELSE. They are greyed out unless
# Date range is "Custom range", because they were once labelled "From / To" and
# an address typed into them was accepted and silently discarded: the search
# then returned every message in the mailbox and looked broken. For cc and bcc
# as well as To, use participants:"..." in the Keywords box.
#
# ABOUT DELETING, AND PLEASE READ THIS ONCE:
#
# "Purge everything this search found" is the operation Microsoft actually
# provides. It deletes the whole result set.
#
# "Purge only the ticked rows" is NOT an operation Microsoft provides. This tool
# fakes it honestly: it builds a second search narrowed to the rows you ticked,
# runs it, and compares the count. If that narrowed search matches MORE items
# than you ticked -- which happens when a subject is not unique -- IT REFUSES,
# because the alternative is deleting somebody else's mail as well. Narrow the
# original search instead and try again.
#
# THE PREVIEW LIST IS NOT ALWAYS THE WHOLE RESULT SET. Microsoft caps preview at
# roughly a thousand items. The window always tells you the previewed count and
# the search's own count, and warns when they differ -- if they do, "purge
# everything" will delete far more than the list in front of you shows.
#
# A SINGLE PURGE DOES NOT NECESSARILY FINISH THE JOB. Microsoft limits how many
# items one purge action removes per mailbox. Re-run it, and check the count
# each time, until the search returns nothing.
#
# Soft delete puts items in Recoverable Items, where a user can get them back
# for the retention window. Hard delete does not. Nothing here recovers them.
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

# CONTENT TYPE TO KQL, IN ONE PLACE AND WITH ITS DOUBTS RECORDED.
#
# The script this was refactored from used 'kind:teams', which is not a value
# the keyword reference lists -- so that option searched for nothing and
# reported no results, which reads as "there are none".
#
# 'microsoftteams' is the documented kind for Teams messages; older material
# says 'im', which also matches Skype chat. VERIFY THESE AGAINST THE TENANT
# before trusting a nil return: a content search that finds nothing because the
# query was wrong looks exactly like a content search that found nothing.
$script:KindMap = [ordered]@{
    'Email messages'             = 'kind:email'
    'Calendar items (meetings)'  = 'kind:meetings'
    'Teams messages'             = 'kind:microsoftteams'
    'Email, meetings and Teams'  = '(kind:email OR kind:meetings OR kind:microsoftteams)'
    'Everything (no kind filter)' = ''
}

$script:Searches    = New-Object System.Collections.ArrayList
$script:PreviewRows = New-Object System.Collections.ArrayList
# Which search the rows on screen came from. A preview list is meaningless
# without it, and purging against a different search than the one previewed is
# the mistake this variable exists to make impossible.
$script:PreviewOf   = ''

#region Session -----------------------------------------------------------------
function Test-CompliancePrereq {
    <#
        Is there a Security and Compliance session in THIS process?

        ASKED BEFORE EVERY OPERATION, not once at startup. A session can expire
        while the window is open, and the failure then lands in the middle of a
        purge -- which is the worst place in this tool for a surprise.

        Test-IppsConnectionState is the shared probe: it calls
        Get-ComplianceSearch, which exists ONLY in this session, so a positive
        answer proves the right endpoint rather than merely that a module loaded.
    #>
    $r = [PSCustomObject]@{ Ok = $false; Reason = '' }
    try {
        $s = Test-IppsConnectionState
        if ($s.Connected) { $r.Ok = $true; $r.Reason = "$($s.Detail)"; return $r }
        $r.Reason = "$($s.Detail)" + [Environment]::NewLine + [Environment]::NewLine +
            'Press Connect-IPPSSession on the console and then Refresh here. This window cannot make the connection itself: a sign-in prompt raised from a modal form draws and then ignores every click (B2), which is why the console owns connecting.'
    } catch {
        $r.Reason = Format-ErrorDetail -ErrorRecord $_ -Context 'Checking the Security and Compliance session'
    }
    return $r
}
#endregion

#region Query building ----------------------------------------------------------
function Get-ContentQuery {
    <#
        The KQL, from the form.

        THE CUSTOM DATE RANGE ACTUALLY WORKS HERE. In the original there was a
        'Custom Range' option, one date box, and no code path that used it: the
        date filter switch returned $null for that choice, so selecting it
        searched ALL TIME while appearing to be restricted. A date filter that
        silently does not apply is how a purge takes years of mail instead of a
        fortnight, so this returns a REASON instead of a query when the dates do
        not parse, and the caller refuses.
    #>
    param(
        [string]$Kind,
        [string]$Keywords,
        [string]$RangeName,
        [string]$CustomStart,
        [string]$CustomEnd,
        [string]$Sender,
        [string]$Recipient
    )

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Query = '' }
    $parts = New-Object System.Collections.ArrayList

    $kindClause = ''
    if ($script:KindMap.Contains($Kind)) { $kindClause = "$($script:KindMap[$Kind])" }
    if ($kindClause) { [void]$parts.Add($kindClause) }

    # SENDER AND RECIPIENT ARE FIRST-CLASS FIELDS, AND THIS IS WHY.
    #
    # 2026-09-01: an operator wanting "everything from Ihtesham" typed the
    # address into the boxes then labelled "From / To" -- which are the custom
    # date range, not a sender filter. The date range was "Last 30 days", so the
    # Custom branch never ran and the address was silently discarded. The search
    # completed with 5,049 items, being every message in the mailbox for a month,
    # and looked like a tool that could not find anything.
    #
    # Nobody misread anything. A mail search window with a box marked "From"
    # that does not mean the sender is a trap, and the fix is not a better label
    # on the date boxes -- it is having the field the operator was looking for.
    #
    # QUOTED, AND INTERNAL QUOTES DOUBLED. An address needs no quoting; a
    # display name typed here ("Baloch, Ihtesham") contains a comma and would
    # otherwise change what the query means.
    if ("$Sender".Trim()) {
        [void]$parts.Add('from:"' + (("$Sender".Trim()) -replace '"', '""') + '"')
    }
    if ("$Recipient".Trim()) {
        [void]$parts.Add('to:"' + (("$Recipient".Trim()) -replace '"', '""') + '"')
    }

    if ("$Keywords".Trim()) {
        # Parenthesised: an unbracketed OR from the operator would otherwise
        # bind loosely against the kind and date clauses and widen the search
        # without looking as though it had.
        [void]$parts.Add("($("$Keywords".Trim()))")
    }

    $start = $null
    $end   = $null
    switch ($RangeName) {
        'Last 7 days'   { $start = (Get-Date).AddDays(-7);   $end = Get-Date }
        'Last 30 days'  { $start = (Get-Date).AddDays(-30);  $end = Get-Date }
        'Last 90 days'  { $start = (Get-Date).AddDays(-90);  $end = Get-Date }
        'Last 12 months'{ $start = (Get-Date).AddYears(-1);  $end = Get-Date }
        'All time'      { }
        'Custom range'  {
            $s = New-Object DateTime
            $e = New-Object DateTime
            if (-not [DateTime]::TryParse("$CustomStart", [ref]$s)) {
                $r.Reason = "The start date '$CustomStart' is not a date this machine understands. Use the format your Windows locale expects -- for US English that is MM/DD/YYYY."
                return $r
            }
            if (-not [DateTime]::TryParse("$CustomEnd", [ref]$e)) {
                $r.Reason = "The end date '$CustomEnd' is not a date this machine understands."
                return $r
            }
            if ($e -lt $s) {
                $r.Reason = "The end date ($($e.ToString('yyyy-MM-dd'))) is before the start date ($($s.ToString('yyyy-MM-dd')))."
                return $r
            }
            $start = $s; $end = $e
        }
    }

    if ($start -and $end) {
        # ISO, always. A date rendered in the machine's locale inside a KQL
        # string is a query that means something different on the next
        # engineer's laptop.
        [void]$parts.Add("received:$($start.ToString('yyyy-MM-dd'))..$($end.ToString('yyyy-MM-dd'))")
    }

    $r.Query = (@($parts) -join ' AND ')
    $r.Ok = $true
    return $r
}

function Get-SearchLocations {
    <#
        Where to look.

        'All' RATHER THAN AN ENUMERATED LIST. The original read every mailbox in
        the tenant with Get-Mailbox -ResultSize Unlimited and passed thousands of
        addresses to -ExchangeLocation. That is slow, it breaks on large
        tenants, and it is unnecessary: the parameter takes All.
    #>
    param([string]$Scope, [string]$Users)

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Locations = @(); Describe = '' }

    if ($Scope -eq 'Every mailbox in the tenant') {
        $r.Locations = @('All')
        $r.Describe  = 'every mailbox in the tenant'
        $r.Ok = $true
        return $r
    }

    $list = @("$Users" -split '[\r\n;,]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (@($list).Count -eq 0) {
        $r.Reason = 'No mailboxes were listed. Type one or more addresses, separated by lines, semicolons or commas.'
        return $r
    }
    $r.Locations = $list
    $r.Describe  = "$(@($list).Count) named mailbox(es)"
    $r.Ok = $true
    return $r
}
#endregion

#region Compliance calls --------------------------------------------------------
function Wait-ForComplianceState {
    <#
        Poll something until it is Completed, or until the patience runs out.

        BOUNDED AND HONEST. The original waited up to an HOUR on the UI thread
        with no cancel and no repaint -- a window that has stopped responding for
        an hour is indistinguishable from a crashed one, and the operator kills
        it. This waits a few minutes, pumps the message queue so the window still
        draws, and RETURNS WHAT IT SAW instead of pretending. A search that is
        still running is a normal state, not a failure; the Searches tab can be
        refreshed later.
    #>
    param(
        [Parameter(Mandatory)][scriptblock]$Probe,
        [int]$TimeoutSeconds = 240,
        [int]$IntervalSeconds = 5,
        [System.Windows.Forms.Label]$Status
    )

    $r = [PSCustomObject]@{ Completed = $false; Object = $null; Waited = 0; Reason = '' }
    $waited = 0
    while ($true) {
        try { $r.Object = & $Probe } catch { $r.Reason = "$($_.Exception.Message)"; return $r }
        $state = ''
        if ($r.Object) { $state = "$($r.Object.Status)" }
        if ($state -eq 'Completed') { $r.Completed = $true; $r.Waited = $waited; return $r }
        if ($waited -ge $TimeoutSeconds) {
            $r.Waited = $waited
            $r.Reason = "Still '$state' after $waited seconds. It has not failed -- it is not finished. Refresh in a minute."
            return $r
        }
        if ($Status) { $Status.Text = "Waiting: $state ($waited s)"; $Status.Refresh() }
        Start-Sleep -Seconds $IntervalSeconds
        [System.Windows.Forms.Application]::DoEvents()
        $waited += $IntervalSeconds
    }
}

function Get-ComplianceSearchList {
    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Rows = @() }
    try {
        $r.Rows = @(Get-ComplianceSearch -ErrorAction Stop |
            Sort-Object -Property CreatedTime -Descending)
        $r.Ok = $true
    } catch {
        $r.Reason = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ComplianceSearch'
    }
    return $r
}

function ConvertFrom-PreviewResults {
    <#
        The Preview action's Results string, turned into rows.

        THIS PARSES A STRING BECAUSE THERE IS NOTHING ELSE TO PARSE.
        Get-ComplianceSearchAction -Details returns the previewed items as one
        long text field of "Key: value; Key: value;" per item. It is not a
        documented, stable contract, and this is the most brittle function in
        the file -- so it is defensive, it never throws, and the caller checks
        whether the number of rows it produced is plausible before showing them.

        NO InternetMessageId IN THAT OUTPUT, which is the whole reason the
        selective purge downstream has to narrow by subject, sender and date and
        then verify the count. If a future version of the service starts
        returning a unique id here, THAT is the day selective purge becomes
        exact, and this is the function to change.
    #>
    param([string]$Results)

    $rows = New-Object System.Collections.ArrayList
    if (-not "$Results".Trim()) { return @($rows.ToArray()) }

    # One item per line in practice; the trailing "Location:" of the next item
    # is the only reliable separator when a subject itself contains newlines.
    $chunks = @("$Results" -split "(?=Location:\s)" | Where-Object { "$_".Trim() })
    foreach ($chunk in $chunks) {
        $h = @{}
        foreach ($m in [regex]::Matches("$chunk", '(?<k>[A-Za-z ]+):\s*(?<v>[^;]*)(;|$)')) {
            $k = "$($m.Groups['k'].Value)".Trim()
            $v = "$($m.Groups['v'].Value)".Trim()
            if ($k) { $h[$k] = $v }
        }
        if (@($h.Keys).Count -eq 0) { continue }

        $subject = ''
        foreach ($n in @('Subject', 'Title')) { if ($h.ContainsKey($n) -and $h[$n]) { $subject = $h[$n]; break } }

        [void]$rows.Add([PSCustomObject]@{
            Location = "$($h['Location'])"
            Sender   = "$($h['Sender'])"
            Subject  = $subject
            Type     = "$($h['Type'])"
            Size     = "$($h['Size'])"
            Received = "$($h['Received Time'])"
        })
    }
    return @($rows.ToArray())
}

function New-NarrowedSearch {
    <#
        A second search matching only the ticked rows -- as closely as KQL can.

        BUILT FROM SUBJECT, SENDER AND THE RECEIVED DAY, because the preview
        gives nothing unique. Each row becomes
        (subject:"..." AND from:"..." AND received:YYYY-MM-DD..YYYY-MM-DD) and
        the rows are ORed together.

        THIS CAN MATCH MORE THAN WAS TICKED. One subject sent to twelve people
        is twelve items. That is not a flaw to be hidden -- it is the reason the
        caller compares the resulting item count with the number of rows ticked
        and refuses when they differ.

        A ROW WITH NO SUBJECT IS REFUSED OUTRIGHT rather than contributing an
        unbounded clause. 'from:x AND received:day' with no subject would match
        every message that person sent that day.
    #>
    param([Parameter(Mandatory)]$Rows, [Parameter(Mandatory)]$Locations, [Parameter(Mandatory)][string]$Name)

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Query = ''; Name = $Name; Skipped = @() }

    $clauses = New-Object System.Collections.ArrayList
    $skipped = New-Object System.Collections.ArrayList
    foreach ($row in @($Rows)) {
        $subject = "$($row.Subject)".Trim()
        $sender  = "$($row.Sender)".Trim()
        if (-not $subject) { [void]$skipped.Add("(no subject) from '$sender' at $($row.Received)"); continue }

        $bits = New-Object System.Collections.ArrayList
        # Quotes doubled: a subject containing a quotation mark would otherwise
        # end the KQL string early and change what the query means.
        [void]$bits.Add('subject:"' + ($subject -replace '"', '""') + '"')
        if ($sender) { [void]$bits.Add('from:"' + ($sender -replace '"', '""') + '"') }

        $d = New-Object DateTime
        if ([DateTime]::TryParse("$($row.Received)", [ref]$d)) {
            $day = $d.ToString('yyyy-MM-dd')
            $nxt = $d.AddDays(1).ToString('yyyy-MM-dd')
            [void]$bits.Add("received:$day..$nxt")
        }
        [void]$clauses.Add('(' + (@($bits) -join ' AND ') + ')')
    }

    $r.Skipped = @($skipped.ToArray())
    if (@($clauses).Count -eq 0) {
        $r.Reason = 'None of the ticked rows can be identified precisely enough to build a narrowed search. A row with no subject would match everything that sender sent that day, so it is refused rather than guessed at.'
        return $r
    }

    $r.Query = (@($clauses) -join ' OR ')
    try {
        # Removed first if a previous attempt left one behind, so a stale query
        # cannot be purged by mistake.
        try { Remove-ComplianceSearch -Identity $Name -Confirm:$false -ErrorAction SilentlyContinue } catch { }
        [void](New-ComplianceSearch -Name $Name -ExchangeLocation @($Locations) -ContentMatchQuery $r.Query -ErrorAction Stop)
        [void](Start-ComplianceSearch -Identity $Name -ErrorAction Stop)
        $r.Ok = $true
    } catch {
        $r.Reason = Format-ErrorDetail -ErrorRecord $_ -Context "Creating the narrowed search '$Name'"
    }
    return $r
}
#endregion

#region Form --------------------------------------------------------------------
# NO Anchor ON ANYTHING, and the tab control is placed rather than docked.
# The anchors were left off for a scaler that surrendered them; it was deleted
# on 2026-09-11 (AGENTS.md rule 8), so this window is simply fixed-layout in a
# sizable frame -- the same as the rest of the toolkit before scaling existed.
# The tab control stays PLACED regardless: a docked, filled one would leave
# Add-ToolHelpButton no gap in the top row, so the '?' would be widened into a
# gutter and then covered by the tab strip. The top row is kept clear for it.
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Content Search and Purge'
$form.ClientSize      = New-Object System.Drawing.Size(1180, 820)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$lblTitle          = New-Object System.Windows.Forms.Label
$lblTitle.Text     = 'Content Search -- Exchange and Teams'
$lblTitle.Location = New-Object System.Drawing.Point(14, 12)
$lblTitle.Size     = New-Object System.Drawing.Size(600, 22)
$lblTitle.Font     = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($lblTitle)

$lblConn          = New-Object System.Windows.Forms.Label
$lblConn.Location = New-Object System.Drawing.Point(14, 38)
$lblConn.Size     = New-Object System.Drawing.Size(1000, 20)
$lblConn.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($lblConn)

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(14, 62)
$tabs.Size     = New-Object System.Drawing.Size(1152, 560)
[void]$form.Controls.Add($tabs)

$tabSearch  = New-Object System.Windows.Forms.TabPage
$tabSearch.Text = '1.  Create a search'
$tabList    = New-Object System.Windows.Forms.TabPage
$tabList.Text = '2.  Searches'
$tabPurge   = New-Object System.Windows.Forms.TabPage
$tabPurge.Text = '3.  Preview and purge'
foreach ($t in @($tabSearch, $tabList, $tabPurge)) { [void]$tabs.TabPages.Add($t) }

# ---- tab 1: create ------------------------------------------------------------
$lblName          = New-Object System.Windows.Forms.Label
$lblName.Text     = 'Search name'
$lblName.Location = New-Object System.Drawing.Point(16, 20)
$lblName.Size     = New-Object System.Drawing.Size(150, 20)
[void]$tabSearch.Controls.Add($lblName)

$txtName          = New-Object System.Windows.Forms.TextBox
$txtName.Location = New-Object System.Drawing.Point(176, 17)
$txtName.Size     = New-Object System.Drawing.Size(380, 22)
$txtName.Text     = "Search-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
[void]$tabSearch.Controls.Add($txtName)

$lblKind          = New-Object System.Windows.Forms.Label
$lblKind.Text     = 'Content type'
$lblKind.Location = New-Object System.Drawing.Point(16, 52)
$lblKind.Size     = New-Object System.Drawing.Size(150, 20)
[void]$tabSearch.Controls.Add($lblKind)

$cboKind               = New-Object System.Windows.Forms.ComboBox
$cboKind.Location      = New-Object System.Drawing.Point(176, 49)
$cboKind.Size          = New-Object System.Drawing.Size(380, 24)
$cboKind.DropDownStyle = 'DropDownList'
foreach ($k in $script:KindMap.Keys) { [void]$cboKind.Items.Add($k) }
$cboKind.SelectedIndex = 0
[void]$tabSearch.Controls.Add($cboKind)

$lblScope          = New-Object System.Windows.Forms.Label
$lblScope.Text     = 'Where to search'
$lblScope.Location = New-Object System.Drawing.Point(16, 84)
$lblScope.Size     = New-Object System.Drawing.Size(150, 20)
[void]$tabSearch.Controls.Add($lblScope)

$cboScope               = New-Object System.Windows.Forms.ComboBox
$cboScope.Location      = New-Object System.Drawing.Point(176, 81)
$cboScope.Size          = New-Object System.Drawing.Size(380, 24)
$cboScope.DropDownStyle = 'DropDownList'
[void]$cboScope.Items.AddRange(@('Named mailboxes', 'Every mailbox in the tenant'))
$cboScope.SelectedIndex = 0
[void]$tabSearch.Controls.Add($cboScope)

$lblUsers          = New-Object System.Windows.Forms.Label
$lblUsers.Text     = 'Mailboxes'
$lblUsers.Location = New-Object System.Drawing.Point(16, 116)
$lblUsers.Size     = New-Object System.Drawing.Size(150, 20)
[void]$tabSearch.Controls.Add($lblUsers)

$txtUsers            = New-Object System.Windows.Forms.TextBox
$txtUsers.Location   = New-Object System.Drawing.Point(176, 113)
$txtUsers.Size       = New-Object System.Drawing.Size(380, 66)
$txtUsers.Multiline  = $true
$txtUsers.ScrollBars = 'Vertical'
$txtUsers.WordWrap   = $false
[void]$tabSearch.Controls.Add($txtUsers)

$lblUsersHint          = New-Object System.Windows.Forms.Label
$lblUsersHint.Text     = 'One per line, or separated by ; or , -- full addresses.'
$lblUsersHint.Location = New-Object System.Drawing.Point(566, 116)
$lblUsersHint.Size     = New-Object System.Drawing.Size(420, 40)
[void]$tabSearch.Controls.Add($lblUsersHint)

$lblRange          = New-Object System.Windows.Forms.Label
$lblRange.Text     = 'Date range'
$lblRange.Location = New-Object System.Drawing.Point(16, 192)
$lblRange.Size     = New-Object System.Drawing.Size(150, 20)
[void]$tabSearch.Controls.Add($lblRange)

$cboRange               = New-Object System.Windows.Forms.ComboBox
$cboRange.Location      = New-Object System.Drawing.Point(176, 189)
$cboRange.Size          = New-Object System.Drawing.Size(380, 24)
$cboRange.DropDownStyle = 'DropDownList'
[void]$cboRange.Items.AddRange(@('Last 7 days', 'Last 30 days', 'Last 90 days', 'Last 12 months', 'All time', 'Custom range'))
$cboRange.SelectedIndex = 1
[void]$tabSearch.Controls.Add($cboRange)

# NAMED FOR WHAT THEY ARE, AND DISABLED UNTIL THEY MEAN ANYTHING.
#
# These were labelled 'From / To' and were live at all times, so an address
# typed into them in the belief that it filtered the sender was accepted and
# then thrown away -- see the note in Get-ContentQuery. A box that cannot
# affect the query is greyed out, which is the same rule the console applies to
# a tool it cannot run.
$lblCustom          = New-Object System.Windows.Forms.Label
$lblCustom.Text     = 'Custom start / end'
$lblCustom.Location = New-Object System.Drawing.Point(16, 224)
$lblCustom.Size     = New-Object System.Drawing.Size(150, 20)
[void]$tabSearch.Controls.Add($lblCustom)

$txtCustomStart          = New-Object System.Windows.Forms.TextBox
$txtCustomStart.Location = New-Object System.Drawing.Point(176, 221)
$txtCustomStart.Size     = New-Object System.Drawing.Size(180, 22)
[void]$tabSearch.Controls.Add($txtCustomStart)

$txtCustomEnd          = New-Object System.Windows.Forms.TextBox
$txtCustomEnd.Location = New-Object System.Drawing.Point(376, 221)
$txtCustomEnd.Size     = New-Object System.Drawing.Size(180, 22)
[void]$tabSearch.Controls.Add($txtCustomEnd)

$lblDateHint          = New-Object System.Windows.Forms.Label
$lblDateHint.Text     = 'DATES, not people. Only used when Date range is "Custom range", and greyed out otherwise.'
$lblDateHint.Location = New-Object System.Drawing.Point(566, 224)
$lblDateHint.Size     = New-Object System.Drawing.Size(420, 34)
[void]$tabSearch.Controls.Add($lblDateHint)

# THE FIELDS THAT SHOULD HAVE BEEN HERE ALL ALONG.
$lblSender          = New-Object System.Windows.Forms.Label
$lblSender.Text     = 'From (sender)'
$lblSender.Location = New-Object System.Drawing.Point(16, 260)
$lblSender.Size     = New-Object System.Drawing.Size(150, 20)
[void]$tabSearch.Controls.Add($lblSender)

$txtSender          = New-Object System.Windows.Forms.TextBox
$txtSender.Location = New-Object System.Drawing.Point(176, 257)
$txtSender.Size     = New-Object System.Drawing.Size(380, 22)
[void]$tabSearch.Controls.Add($txtSender)

$lblSenderHint          = New-Object System.Windows.Forms.Label
$lblSenderHint.Text     = 'Address or display name. Becomes from:"..." in the query.'
$lblSenderHint.Location = New-Object System.Drawing.Point(566, 260)
$lblSenderHint.Size     = New-Object System.Drawing.Size(420, 20)
[void]$tabSearch.Controls.Add($lblSenderHint)

$lblRecip          = New-Object System.Windows.Forms.Label
$lblRecip.Text     = 'To (recipient)'
$lblRecip.Location = New-Object System.Drawing.Point(16, 292)
$lblRecip.Size     = New-Object System.Drawing.Size(150, 20)
[void]$tabSearch.Controls.Add($lblRecip)

$txtRecip          = New-Object System.Windows.Forms.TextBox
$txtRecip.Location = New-Object System.Drawing.Point(176, 289)
$txtRecip.Size     = New-Object System.Drawing.Size(380, 22)
[void]$tabSearch.Controls.Add($txtRecip)

$lblRecipHint          = New-Object System.Windows.Forms.Label
$lblRecipHint.Text     = 'Becomes to:"..." -- for cc and bcc as well, use participants: in Keywords.'
$lblRecipHint.Location = New-Object System.Drawing.Point(566, 292)
$lblRecipHint.Size     = New-Object System.Drawing.Size(420, 34)
[void]$tabSearch.Controls.Add($lblRecipHint)

$lblKeys          = New-Object System.Windows.Forms.Label
$lblKeys.Text     = 'Keywords (KQL)'
$lblKeys.Location = New-Object System.Drawing.Point(16, 328)
$lblKeys.Size     = New-Object System.Drawing.Size(150, 20)
[void]$tabSearch.Controls.Add($lblKeys)

$txtKeys            = New-Object System.Windows.Forms.TextBox
$txtKeys.Location   = New-Object System.Drawing.Point(176, 325)
$txtKeys.Size       = New-Object System.Drawing.Size(380, 60)
$txtKeys.Multiline  = $true
$txtKeys.ScrollBars = 'Vertical'
[void]$tabSearch.Controls.Add($txtKeys)

$lblQueryCap          = New-Object System.Windows.Forms.Label
$lblQueryCap.Text     = 'The query that will be sent'
$lblQueryCap.Location = New-Object System.Drawing.Point(16, 396)
$lblQueryCap.Size     = New-Object System.Drawing.Size(300, 20)
$lblQueryCap.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$tabSearch.Controls.Add($lblQueryCap)

# THE QUERY IS SHOWN BEFORE IT IS RUN. A content search is a question whose
# wrong answers are indistinguishable from right ones, and the only way to
# check it is to read it.
$txtQuery            = New-Object System.Windows.Forms.TextBox
$txtQuery.Location   = New-Object System.Drawing.Point(16, 418)
$txtQuery.Size       = New-Object System.Drawing.Size(970, 46)
$txtQuery.Multiline  = $true
$txtQuery.ReadOnly   = $true
$txtQuery.WordWrap   = $true
$txtQuery.Font       = New-Object System.Drawing.Font('Consolas', 9)
[void]$tabSearch.Controls.Add($txtQuery)

$btnCreate           = New-Object System.Windows.Forms.Button
$btnCreate.Text      = 'Create and start search'
$btnCreate.Location  = New-Object System.Drawing.Point(16, 474)
$btnCreate.Size      = New-Object System.Drawing.Size(220, 30)
$btnCreate.FlatStyle = 'Flat'
[void]$tabSearch.Controls.Add($btnCreate)

$lblCreateState          = New-Object System.Windows.Forms.Label
$lblCreateState.Location = New-Object System.Drawing.Point(246, 480)
$lblCreateState.Size     = New-Object System.Drawing.Size(740, 20)
[void]$tabSearch.Controls.Add($lblCreateState)

# ---- tab 2: searches ---------------------------------------------------------
$gridSearches                     = New-Object System.Windows.Forms.DataGridView
$gridSearches.Location            = New-Object System.Drawing.Point(16, 16)
$gridSearches.Size                = New-Object System.Drawing.Size(1110, 420)
$gridSearches.ReadOnly            = $true
$gridSearches.AllowUserToAddRows  = $false
$gridSearches.AutoSizeColumnsMode = 'Fill'
$gridSearches.SelectionMode       = 'FullRowSelect'
$gridSearches.MultiSelect         = $false
$gridSearches.RowHeadersVisible   = $false
[void]$tabList.Controls.Add($gridSearches)

$btnRefreshList           = New-Object System.Windows.Forms.Button
$btnRefreshList.Text      = 'Refresh'
$btnRefreshList.Location  = New-Object System.Drawing.Point(16, 448)
$btnRefreshList.Size      = New-Object System.Drawing.Size(120, 30)
$btnRefreshList.FlatStyle = 'Flat'
[void]$tabList.Controls.Add($btnRefreshList)

$btnRunSearch           = New-Object System.Windows.Forms.Button
$btnRunSearch.Text      = 'Run selected'
$btnRunSearch.Location  = New-Object System.Drawing.Point(146, 448)
$btnRunSearch.Size      = New-Object System.Drawing.Size(140, 30)
$btnRunSearch.FlatStyle = 'Flat'
[void]$tabList.Controls.Add($btnRunSearch)

$btnDeleteSearch           = New-Object System.Windows.Forms.Button
$btnDeleteSearch.Text      = 'Delete selected search'
$btnDeleteSearch.Location  = New-Object System.Drawing.Point(296, 448)
$btnDeleteSearch.Size      = New-Object System.Drawing.Size(200, 30)
$btnDeleteSearch.FlatStyle = 'Flat'
[void]$tabList.Controls.Add($btnDeleteSearch)

$lblListNote          = New-Object System.Windows.Forms.Label
$lblListNote.Text     = 'Deleting a SEARCH deletes the definition, not any mail. Purging is on the next tab.'
$lblListNote.Location = New-Object System.Drawing.Point(510, 454)
$lblListNote.Size     = New-Object System.Drawing.Size(616, 20)
[void]$tabList.Controls.Add($lblListNote)

# ---- tab 3: preview and purge ------------------------------------------------
$lblPvSearch          = New-Object System.Windows.Forms.Label
$lblPvSearch.Text     = 'Search'
$lblPvSearch.Location = New-Object System.Drawing.Point(16, 20)
$lblPvSearch.Size     = New-Object System.Drawing.Size(60, 20)
[void]$tabPurge.Controls.Add($lblPvSearch)

$cboPvSearch               = New-Object System.Windows.Forms.ComboBox
$cboPvSearch.Location      = New-Object System.Drawing.Point(80, 17)
$cboPvSearch.Size          = New-Object System.Drawing.Size(420, 24)
$cboPvSearch.DropDownStyle = 'DropDownList'
[void]$tabPurge.Controls.Add($cboPvSearch)

$btnPreview           = New-Object System.Windows.Forms.Button
$btnPreview.Text      = 'Preview items'
$btnPreview.Location  = New-Object System.Drawing.Point(516, 16)
$btnPreview.Size      = New-Object System.Drawing.Size(150, 26)
$btnPreview.FlatStyle = 'Flat'
[void]$tabPurge.Controls.Add($btnPreview)

$btnTickAll           = New-Object System.Windows.Forms.Button
$btnTickAll.Text      = 'Tick all'
$btnTickAll.Location  = New-Object System.Drawing.Point(676, 16)
$btnTickAll.Size      = New-Object System.Drawing.Size(100, 26)
$btnTickAll.FlatStyle = 'Flat'
[void]$tabPurge.Controls.Add($btnTickAll)

$btnTickNone           = New-Object System.Windows.Forms.Button
$btnTickNone.Text      = 'Tick none'
$btnTickNone.Location  = New-Object System.Drawing.Point(786, 16)
$btnTickNone.Size      = New-Object System.Drawing.Size(100, 26)
$btnTickNone.FlatStyle = 'Flat'
[void]$tabPurge.Controls.Add($btnTickNone)

$btnExportPv           = New-Object System.Windows.Forms.Button
$btnExportPv.Text      = 'Export list to CSV'
$btnExportPv.Location  = New-Object System.Drawing.Point(896, 16)
$btnExportPv.Size      = New-Object System.Drawing.Size(170, 26)
$btnExportPv.FlatStyle = 'Flat'
[void]$tabPurge.Controls.Add($btnExportPv)

# THE COUNT COMPARISON, ON ITS OWN LINE AND IN BOLD. Whether this list is the
# whole result set is the single most important fact on this tab.
$lblPvCount          = New-Object System.Windows.Forms.Label
$lblPvCount.Location = New-Object System.Drawing.Point(16, 48)
$lblPvCount.Size     = New-Object System.Drawing.Size(1090, 34)
$lblPvCount.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$tabPurge.Controls.Add($lblPvCount)

$gridPreview                     = New-Object System.Windows.Forms.DataGridView
$gridPreview.Location            = New-Object System.Drawing.Point(16, 86)
$gridPreview.Size                = New-Object System.Drawing.Size(1110, 306)
$gridPreview.AllowUserToAddRows  = $false
$gridPreview.AutoSizeColumnsMode = 'Fill'
$gridPreview.SelectionMode       = 'FullRowSelect'
$gridPreview.MultiSelect         = $false
$gridPreview.RowHeadersVisible   = $false
[void]$tabPurge.Controls.Add($gridPreview)

$colTick = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colTick.HeaderText   = 'Purge'
$colTick.Name         = 'Tick'
$colTick.AutoSizeMode = 'None'
$colTick.Width        = 60
[void]$gridPreview.Columns.Add($colTick)
foreach ($c in @(
    @{ H = 'Location'; W = 22 }, @{ H = 'Sender'; W = 20 }, @{ H = 'Subject'; W = 34 },
    @{ H = 'Type'; W = 8 }, @{ H = 'Size'; W = 8 }, @{ H = 'Received'; W = 16 })) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $c.H
    $col.FillWeight = $c.W
    $col.ReadOnly   = $true
    [void]$gridPreview.Columns.Add($col)
}

$lblPurgeType          = New-Object System.Windows.Forms.Label
$lblPurgeType.Text     = 'Delete type'
$lblPurgeType.Location = New-Object System.Drawing.Point(16, 404)
$lblPurgeType.Size     = New-Object System.Drawing.Size(90, 20)
[void]$tabPurge.Controls.Add($lblPurgeType)

$cboPurgeType               = New-Object System.Windows.Forms.ComboBox
$cboPurgeType.Location      = New-Object System.Drawing.Point(110, 401)
$cboPurgeType.Size          = New-Object System.Drawing.Size(390, 24)
$cboPurgeType.DropDownStyle = 'DropDownList'
[void]$cboPurgeType.Items.AddRange(@(
    'Soft delete -- recoverable by the user from Recoverable Items',
    'HARD DELETE -- not recoverable by anyone'))
$cboPurgeType.SelectedIndex = 0
[void]$tabPurge.Controls.Add($cboPurgeType)

$lblConfirm          = New-Object System.Windows.Forms.Label
$lblConfirm.Text     = "Type PURGE to arm"
$lblConfirm.Location = New-Object System.Drawing.Point(516, 404)
$lblConfirm.Size     = New-Object System.Drawing.Size(140, 20)
[void]$tabPurge.Controls.Add($lblConfirm)

$txtConfirm          = New-Object System.Windows.Forms.TextBox
$txtConfirm.Location = New-Object System.Drawing.Point(660, 401)
$txtConfirm.Size     = New-Object System.Drawing.Size(120, 22)
[void]$tabPurge.Controls.Add($txtConfirm)

$btnPurgeSelected           = New-Object System.Windows.Forms.Button
$btnPurgeSelected.Text      = 'Purge ticked rows only'
$btnPurgeSelected.Location  = New-Object System.Drawing.Point(16, 434)
$btnPurgeSelected.Size      = New-Object System.Drawing.Size(230, 30)
$btnPurgeSelected.FlatStyle = 'Flat'
[void]$tabPurge.Controls.Add($btnPurgeSelected)

$btnPurgeAll           = New-Object System.Windows.Forms.Button
$btnPurgeAll.Text      = 'Purge EVERYTHING this search found'
$btnPurgeAll.Location  = New-Object System.Drawing.Point(256, 434)
$btnPurgeAll.Size      = New-Object System.Drawing.Size(300, 30)
$btnPurgeAll.FlatStyle = 'Flat'
[void]$tabPurge.Controls.Add($btnPurgeAll)

$lblPurgeNote          = New-Object System.Windows.Forms.Label
$lblPurgeNote.Text     = 'Microsoft limits how many items one purge removes per mailbox, so a large result set needs the purge repeated. Check the count after each run.'
$lblPurgeNote.Location = New-Object System.Drawing.Point(566, 434)
$lblPurgeNote.Size     = New-Object System.Drawing.Size(560, 34)
[void]$tabPurge.Controls.Add($lblPurgeNote)

# ---- output ------------------------------------------------------------------
$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(14, 634)
$txtOut.Size       = New-Object System.Drawing.Size(1152, 172)
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
# B6.
$txtOut.WordWrap   = $true
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
[void]$form.Controls.Add($txtOut)
#endregion

#region View --------------------------------------------------------------------
function Show-Message {
    param([string]$Text, [switch]$Keep)
    if ($Keep) { $txtOut.Text = "$Text$([Environment]::NewLine)$([Environment]::NewLine)$($txtOut.Text)" }
    else { $txtOut.Text = "$Text" }
}

function Update-ConnectionLine {
    $p = Test-CompliancePrereq
    if ($p.Ok) {
        $lblConn.Text = 'Security and Compliance: connected.'
        $lblConn.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
    } else {
        $lblConn.Text = 'Security and Compliance: NOT CONNECTED -- press Connect-IPPSSession on the console, then Refresh.'
        $lblConn.ForeColor = [System.Drawing.Color]::FromArgb(235, 110, 110)
    }
    # B3: a semantic colour must survive theming.
    $lblConn.Tag = 'SemanticColour'
    return $p.Ok
}

function Update-QueryPreview {
    $q = Get-ContentQuery -Kind "$($cboKind.SelectedItem)" -Keywords $txtKeys.Text `
        -RangeName "$($cboRange.SelectedItem)" -CustomStart $txtFrom.Text -CustomEnd $txtTo.Text
    if (-not $q.Ok) { $txtQuery.Text = "(cannot be built) $($q.Reason)"; return }
    if (-not $q.Query) { $txtQuery.Text = '(empty -- every item in the selected mailboxes will match)'; return }
    $txtQuery.Text = $q.Query
}

function Update-SearchGrid {
    $script:Searches.Clear()
    $gridSearches.DataSource = $null

    if (-not (Update-ConnectionLine)) { Show-Message (Test-CompliancePrereq).Reason; return }

    $res = Get-ComplianceSearchList
    if (-not $res.Ok) { Show-Message $res.Reason; return }

    $table = New-Object System.Data.DataTable
    foreach ($c in @('Name', 'Status', 'Items', 'Size', 'Created')) { [void]$table.Columns.Add($c) }
    foreach ($s in @($res.Rows)) {
        [void]$script:Searches.Add($s)
        $created = ''
        try { $created = ([datetime]$s.CreatedTime).ToString('yyyy-MM-dd HH:mm') } catch { $created = "$($s.CreatedTime)" }
        [void]$table.Rows.Add("$($s.Name)", "$($s.Status)", "$($s.Items)", "$($s.Size)", $created)
    }
    $gridSearches.DataSource = $table

    # The purge tab's picker is filled from the same read, so the two tabs
    # cannot disagree about which searches exist.
    $keep = "$($cboPvSearch.SelectedItem)"
    $cboPvSearch.Items.Clear()
    foreach ($s in @($res.Rows)) { [void]$cboPvSearch.Items.Add("$($s.Name)") }
    if ($keep -and $cboPvSearch.Items.Contains($keep)) { $cboPvSearch.SelectedItem = $keep }
    elseif ($cboPvSearch.Items.Count -gt 0) { $cboPvSearch.SelectedIndex = 0 }

    Show-Message "$(@($res.Rows).Count) search(es) in this tenant."
}

function Get-SelectedSearchRow {
    if (-not $gridSearches.CurrentRow) { return $null }
    $i = $gridSearches.CurrentRow.Index
    if ($i -lt 0 -or $i -ge $script:Searches.Count) { return $null }
    return $script:Searches[$i]
}

function Update-PreviewGrid {
    $gridPreview.Rows.Clear()
    foreach ($r in $script:PreviewRows) {
        [void]$gridPreview.Rows.Add($false, $r.Location, $r.Sender, $r.Subject, $r.Type, $r.Size, $r.Received)
    }
}

function Get-TickedPreviewRows {
    $out = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $gridPreview.Rows.Count; $i++) {
        if ($i -ge $script:PreviewRows.Count) { break }
        if ([bool]$gridPreview.Rows[$i].Cells[0].Value) { [void]$out.Add($script:PreviewRows[$i]) }
    }
    return @($out.ToArray())
}
#endregion

#region Actions -----------------------------------------------------------------
function Start-NewSearch {
    if (-not (Update-ConnectionLine)) { Show-Message (Test-CompliancePrereq).Reason; return }

    $name = "$($txtName.Text)".Trim()
    if (-not $name) { Show-Message 'Give the search a name.'; return }

    $q = Get-ContentQuery -Kind "$($cboKind.SelectedItem)" -Keywords $txtKeys.Text `
        -RangeName "$($cboRange.SelectedItem)" -CustomStart $txtFrom.Text -CustomEnd $txtTo.Text
    if (-not $q.Ok) { Show-Message "The search was NOT created: $($q.Reason)"; return }

    $loc = Get-SearchLocations -Scope "$($cboScope.SelectedItem)" -Users $txtUsers.Text
    if (-not $loc.Ok) { Show-Message "The search was NOT created: $($loc.Reason)"; return }

    $existing = $null
    try { $existing = Get-ComplianceSearch -Identity $name -ErrorAction SilentlyContinue } catch { }
    if ($existing) {
        # THE ORIGINAL DELETED THIS SILENTLY. A search name is the only handle an
        # operator has on a previous result set, and removing one without asking
        # discards whatever it found.
        $r = Show-ToolDialog(
            "A search called '$name' already exists (status $($existing.Status), $($existing.Items) item(s))." + [Environment]::NewLine + [Environment]::NewLine +
            'Replacing it deletes that search and whatever it found. Its results cannot be recovered afterwards.' + [Environment]::NewLine + [Environment]::NewLine +
            'Replace it?', 'Search already exists',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { Show-Message 'Nothing was changed.'; return }
        try { Remove-ComplianceSearch -Identity $name -Confirm:$false -ErrorAction Stop } catch {
            Show-Message (Format-ErrorDetail -ErrorRecord $_ -Context "Removing the existing search '$name'")
            return
        }
    }

    $detail = "New-ComplianceSearch '$name' over $($loc.Describe). Query: $($q.Query)"
    $old = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnCreate.Enabled = $false
    try {
        [void](New-ComplianceSearch -Name $name -ExchangeLocation @($loc.Locations) -ContentMatchQuery $q.Query -ErrorAction Stop)
        [void](Start-ComplianceSearch -Identity $name -ErrorAction Stop)

        $w = Wait-ForComplianceState -Probe { Get-ComplianceSearch -Identity $name -ErrorAction Stop } -Status $lblCreateState
        $s = $w.Object

        $L = New-Object System.Collections.ArrayList
        if ($w.Completed) {
            [void]$L.Add("SEARCH COMPLETED: '$name' found $($s.Items) item(s), $($s.Size) byte(s).")
        } else {
            [void]$L.Add("SEARCH STARTED: '$name'. $($w.Reason)")
        }
        [void]$L.Add("  Scope : $($loc.Describe)")
        [void]$L.Add("  Query : $($q.Query)")
        [void]$L.Add('')
        [void]$L.Add('Nothing has been deleted. Go to "Preview and purge" to see what it found.')
        Show-Message ($L -join [Environment]::NewLine)
        $lblCreateState.Text = if ($w.Completed) { "Completed -- $($s.Items) item(s)" } else { 'Still running' }

        Write-CompletedLog -Action 'Read' -Target (Get-LogTarget $name 'Content Search') -Result 'Success' `
            -Detail "$detail Result: status $($s.Status), $($s.Items) item(s)." -Environment $Environment
        Update-SearchGrid
        $cboPvSearch.SelectedItem = $name
    } catch {
        Show-Message (Format-ErrorDetail -ErrorRecord $_ -Context "Creating search '$name'")
        Write-FailedLog -Action 'Read' -Target (Get-LogTarget $name 'Content Search') `
            -Detail $detail -ErrorRecord $_ -Environment $Environment
    } finally {
        $btnCreate.Enabled = $true
        $form.Cursor = $old
    }
}

function Start-PreviewOfSearch {
    if (-not (Update-ConnectionLine)) { Show-Message (Test-CompliancePrereq).Reason; return }

    $name = "$($cboPvSearch.SelectedItem)"
    if (-not $name) { Show-Message 'Pick a search to preview.'; return }

    $script:PreviewRows.Clear()
    $script:PreviewOf = ''
    Update-PreviewGrid
    $lblPvCount.Text = ''

    $old = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnPreview.Enabled = $false
    try {
        $search = Get-ComplianceSearch -Identity $name -ErrorAction Stop
        if ("$($search.Status)" -ne 'Completed') {
            Show-Message "'$name' has not finished running (status $($search.Status)), so there is nothing to preview yet. Run it on the Searches tab and wait for Completed."
            return
        }
        $script:PreviewItemsReported = [int]$search.Items
        if ([int]$search.Items -eq 0) {
            Show-Message "'$name' found nothing. There is nothing to preview and nothing to purge."
            $lblPvCount.Text = 'The search found 0 items.'
            return
        }

        # A Preview action per search, reused if one already exists -- creating a
        # second raises an error rather than replacing the first.
        $actionName = "$($name)_Preview"
        $action = $null
        try { $action = Get-ComplianceSearchAction -Identity $actionName -ErrorAction SilentlyContinue } catch { }
        if (-not $action) {
            [void](New-ComplianceSearchAction -SearchName $name -Preview -ErrorAction Stop)
        }

        $w = Wait-ForComplianceState -Probe { Get-ComplianceSearchAction -Identity $actionName -Details -ErrorAction Stop } -Status $lblPvCount
        $act = $w.Object
        if (-not $w.Completed) {
            Show-Message "The preview has not finished. $($w.Reason)"
            return
        }

        $rows = @(ConvertFrom-PreviewResults -Results "$($act.Results)")
        foreach ($r in $rows) { [void]$script:PreviewRows.Add($r) }
        $script:PreviewOf = $name
        Update-PreviewGrid

        # THE COMPARISON THAT MAKES THIS SCREEN HONEST.
        $shown = @($rows).Count
        $total = [int]$search.Items
        $L = New-Object System.Collections.ArrayList
        if ($shown -eq 0) {
            $lblPvCount.Text = "The search reports $total item(s) but the preview returned nothing this window could read."
            [void]$L.Add("PREVIEW PARSED NOTHING. The search reports $total item(s) and the preview action completed, but the text it returned did not parse into rows.")
            [void]$L.Add('That is this tool''s parser, not the service: the preview comes back as one long text field with no documented format. Export the list from the compliance portal instead, and do not purge from this window until the list here matches what you expect.')
        } elseif ($shown -lt $total) {
            $lblPvCount.Text = "SHOWING $shown OF $total ITEMS. Preview is capped -- this list is NOT everything the search found."
            [void]$L.Add("SHOWING $shown OF $total ITEMS.")
            [void]$L.Add('Microsoft caps preview results, so the rows below are a subset. "Purge EVERYTHING this search found" would delete all ' + "$total" + ' items, not the ' + "$shown" + ' you can see.')
            [void]$L.Add('If you need to see everything before deleting anything, narrow the search until the two numbers agree.')
        } else {
            $lblPvCount.Text = "Showing all $shown item(s) the search found."
            [void]$L.Add("Showing all $shown item(s). The previewed count matches the search's own count, so this list is the whole result set.")
        }
        [void]$L.Add('')
        [void]$L.Add('Nothing has been deleted. Tick rows and use "Purge ticked rows only", or purge the whole result set.')
        Show-Message ($L -join [Environment]::NewLine)

        Write-CompletedLog -Action 'Read' -Target (Get-LogTarget $name 'Content Search') -Result 'Success' `
            -Detail "Previewed '$name': $shown row(s) parsed, search reports $total item(s)." -Environment $Environment
    } catch {
        Show-Message (Format-ErrorDetail -ErrorRecord $_ -Context "Previewing '$name'")
        Write-FailedLog -Action 'Read' -Target (Get-LogTarget $name 'Content Search') `
            -Detail "Preview of '$name' failed." -ErrorRecord $_ -Environment $Environment
    } finally {
        $btnPreview.Enabled = $true
        $form.Cursor = $old
    }
}

function Get-PurgeTypeName {
    if ($cboPurgeType.SelectedIndex -eq 1) { return 'HardDelete' }
    return 'SoftDelete'
}

function Test-PurgeArmed {
    <#
        The typed word, checked case-sensitively.

        A CHECKBOX WOULD NOT DO. This deletes other people's mail, and the whole
        value of a typed confirmation is that it cannot be produced by a stray
        click on a window somebody thought was something else.
    #>
    if ("$($txtConfirm.Text)" -cne 'PURGE') {
        Show-Message 'Type PURGE in the box, in capitals, to arm the delete buttons. Nothing was done.'
        return $false
    }
    return $true
}

function Invoke-PurgeOfSearch {
    <#
        Purge a whole search, whatever it contains.

        $Rationale is carried into the audit row so that a purge of a NARROWED
        search records what it was narrowed to. A log line saying
        'purged Search-20260901-101500_Selected' answers nothing on its own.
    #>
    param(
        [Parameter(Mandatory)][string]$SearchName,
        [Parameter(Mandatory)][string]$PurgeType,
        [Parameter(Mandatory)][string]$Rationale
    )

    $detail = "New-ComplianceSearchAction -Purge -PurgeType $PurgeType on '$SearchName'. $Rationale"
    try {
        [void](New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType $PurgeType -Confirm:$false -ErrorAction Stop)
        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("PURGE SUBMITTED against '$SearchName' as $PurgeType.")
        [void]$L.Add($Rationale)
        [void]$L.Add('')
        [void]$L.Add('It runs in the service, not here, so "submitted" is not "finished". Microsoft also limits how many items one purge removes per mailbox -- re-run the search and look at the count before assuming the result set is empty.')
        if ($PurgeType -eq 'HardDelete') {
            [void]$L.Add('')
            [void]$L.Add('THIS WAS A HARD DELETE. Nothing in this toolkit, and nothing in the portal, brings those items back.')
        }
        Show-Message ($L -join [Environment]::NewLine) -Keep
        Write-CompletedLog -Action 'Delete' -Target (Get-LogTarget $SearchName 'Content Search') -Result 'Success' `
            -Detail $detail -Environment $Environment
        $txtConfirm.Clear()
        return $true
    } catch {
        Show-Message (Format-ErrorDetail -ErrorRecord $_ -Context "Purging '$SearchName'") -Keep
        Write-FailedLog -Action 'Delete' -Target (Get-LogTarget $SearchName 'Content Search') `
            -Detail $detail -ErrorRecord $_ -Environment $Environment
        return $false
    }
}

function Start-PurgeEverything {
    if (-not (Update-ConnectionLine)) { Show-Message (Test-CompliancePrereq).Reason; return }
    if (-not (Test-PurgeArmed)) { return }

    $name = "$($cboPvSearch.SelectedItem)"
    if (-not $name) { Show-Message 'Pick a search first.'; return }

    $search = $null
    try { $search = Get-ComplianceSearch -Identity $name -ErrorAction Stop } catch {
        Show-Message (Format-ErrorDetail -ErrorRecord $_ -Context "Reading '$name'")
        return
    }
    if ([int]$search.Items -eq 0) { Show-Message "'$name' found nothing, so there is nothing to purge."; return }

    $type = Get-PurgeTypeName
    $shown = @($script:PreviewRows).Count
    $warn = ''
    if ($script:PreviewOf -ne $name) {
        $warn = 'THE LIST ON SCREEN IS NOT THIS SEARCH. Nothing has been previewed for it in this window.' + [Environment]::NewLine + [Environment]::NewLine
    } elseif ($shown -lt [int]$search.Items) {
        $warn = "YOU HAVE SEEN $shown OF THESE $($search.Items) ITEMS. The rest were never displayed." + [Environment]::NewLine + [Environment]::NewLine
    }

    $msg = "PURGE EVERYTHING '$name' FOUND." + [Environment]::NewLine + [Environment]::NewLine +
        $warn +
        "Items to be deleted : $($search.Items)" + [Environment]::NewLine +
        "Delete type         : $type" + [Environment]::NewLine + [Environment]::NewLine +
        'This deletes the entire result set, not a selection. It cannot be limited to the rows you can see.' + [Environment]::NewLine + [Environment]::NewLine
    if ($type -eq 'HardDelete') {
        $msg += 'HARD DELETE IS NOT RECOVERABLE. Not by the user, not by an administrator, not from Recoverable Items.' + [Environment]::NewLine + [Environment]::NewLine
    } else {
        $msg += 'Soft delete moves items to Recoverable Items, where the user can restore them until the retention window ends.' + [Environment]::NewLine + [Environment]::NewLine
    }
    $msg += 'Proceed?'

    $r = Show-ToolDialog($msg, 'Purge the whole result set',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Stop,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { Show-Message 'Nothing was deleted.'; return }

    [void](Invoke-PurgeOfSearch -SearchName $name -PurgeType $type `
        -Rationale "Whole result set of '$name' -- $($search.Items) item(s) reported at submission.")
}

function Start-PurgeSelected {
    <#
        The ticked rows, and only the ticked rows -- or nothing at all.

        THE COUNT IS THE GATE. A narrowed query built from subject, sender and
        date is a good filter and not an identity, so this creates the narrowed
        search, runs it, and compares its item count with the number of rows
        ticked. Equal: purge it. Greater: REFUSE, and say by how much. There is
        no third branch, because "probably the right items" is not a standard
        that survives deleting somebody's mail.
    #>
    if (-not (Update-ConnectionLine)) { Show-Message (Test-CompliancePrereq).Reason; return }
    if (-not (Test-PurgeArmed)) { return }

    $name = "$($cboPvSearch.SelectedItem)"
    if (-not $name) { Show-Message 'Pick a search first.'; return }
    if ($script:PreviewOf -ne $name) {
        Show-Message "The rows on screen came from '$script:PreviewOf', not '$name'. Press Preview items for this search before purging from it."
        return
    }

    $ticked = @(Get-TickedPreviewRows)
    if (@($ticked).Count -eq 0) { Show-Message 'No rows are ticked.'; return }

    $type = Get-PurgeTypeName
    $narrowName = "$($name)_Selected"

    $old = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnPurgeSelected.Enabled = $false
    try {
        # The narrowed search must look in the same places as the original, or it
        # would either miss items or reach into mailboxes nobody chose.
        $orig = Get-ComplianceSearch -Identity $name -ErrorAction Stop
        $locations = @($orig.ExchangeLocation)
        if (@($locations).Count -eq 0) { $locations = @('All') }

        $built = New-NarrowedSearch -Rows $ticked -Locations $locations -Name $narrowName
        if (-not $built.Ok) { Show-Message "Nothing was deleted. $($built.Reason)"; return }

        $w = Wait-ForComplianceState -Probe { Get-ComplianceSearch -Identity $narrowName -ErrorAction Stop } -Status $lblPvCount
        if (-not $w.Completed) {
            Show-Message "Nothing was deleted. The narrowed search has not finished: $($w.Reason) Its name is '$narrowName' -- check it on the Searches tab and try again."
            return
        }

        $found  = [int]$w.Object.Items
        $wanted = @($ticked).Count
        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("Ticked rows      : $wanted")
        [void]$L.Add("Narrowed search  : '$narrowName' matched $found item(s)")
        if (@($built.Skipped).Count -gt 0) {
            [void]$L.Add("Rows refused     : $(@($built.Skipped).Count) had no subject and could not be identified precisely:")
            foreach ($s in $built.Skipped) { [void]$L.Add("    $s") }
        }

        if ($found -gt $wanted) {
            [void]$L.Add('')
            [void]$L.Add("REFUSED. The narrowed search matches $found items but you ticked $wanted, so purging it would delete $($found - $wanted) item(s) you did not choose -- almost certainly the same subject sent to other people.")
            [void]$L.Add('')
            [void]$L.Add('Content Search cannot delete individual messages: it deletes whatever a search matches. To remove exactly these, make the original search narrower -- add the sender, tighten the dates, add a distinctive keyword -- until it matches only what you want, then purge the whole thing.')
            [void]$L.Add("The narrowed search '$narrowName' has been left in place so you can inspect it. Nothing was deleted.")
            Show-Message ($L -join [Environment]::NewLine)
            Write-FailedLog -Action 'Delete' -Target (Get-LogTarget $narrowName 'Content Search') `
                -Detail "Selective purge REFUSED: narrowed search matched $found item(s) for $wanted ticked row(s). Nothing deleted." -Environment $Environment
            return
        }

        if ($found -lt $wanted) {
            [void]$L.Add('')
            [void]$L.Add("The narrowed search matches FEWER items than you ticked. Purging it would delete $found of the $wanted rows -- the others could not be matched by subject, sender and date.")
        }

        $msg = "PURGE $found ITEM(S) matched from the $wanted row(s) you ticked." + [Environment]::NewLine + [Environment]::NewLine +
            "This is done by purging a narrowed search, '$narrowName', because Content Search has no way to delete individual messages." + [Environment]::NewLine + [Environment]::NewLine +
            "Delete type : $type" + [Environment]::NewLine + [Environment]::NewLine
        if ($type -eq 'HardDelete') {
            $msg += 'HARD DELETE IS NOT RECOVERABLE by anyone.' + [Environment]::NewLine + [Environment]::NewLine
        }
        $msg += 'Proceed?'

        $r = Show-ToolDialog($msg, 'Purge the ticked rows',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Stop,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) {
            [void]$L.Add('')
            [void]$L.Add('Nothing was deleted.')
            Show-Message ($L -join [Environment]::NewLine)
            return
        }

        Show-Message ($L -join [Environment]::NewLine)
        $subjects = @($ticked | ForEach-Object { "$($_.Sender) / $($_.Subject)" })
        [void](Invoke-PurgeOfSearch -SearchName $narrowName -PurgeType $type `
            -Rationale "Narrowed from '$name' to $wanted ticked row(s); the narrowed search matched $found item(s). Rows: $(@($subjects) -join ' | ')")
    } catch {
        Show-Message (Format-ErrorDetail -ErrorRecord $_ -Context 'Purging the ticked rows')
        Write-FailedLog -Action 'Delete' -Target (Get-LogTarget $narrowName 'Content Search') `
            -Detail 'Selective purge failed before submission.' -ErrorRecord $_ -Environment $Environment
    } finally {
        $btnPurgeSelected.Enabled = $true
        $form.Cursor = $old
    }
}

function Export-PreviewCsv {
    if (@($script:PreviewRows).Count -eq 0) { Show-Message 'There is nothing to export. Preview a search first.'; return }

    $stem = ("$($script:PreviewOf)" -replace '[^A-Za-z0-9._-]', '').ToLowerInvariant()
    if (-not $stem) { $stem = 'contentsearch' }
    $path = Join-Path (Get-ReportPath) "$stem-preview.csv"

    try {
        $rows = @($script:PreviewRows | ForEach-Object {
            [PSCustomObject][ordered]@{
                SearchName = "$($script:PreviewOf)"
                Location   = "$($_.Location)"
                Sender     = "$($_.Sender)"
                Subject    = "$($_.Subject)"
                Type       = "$($_.Type)"
                Size       = "$($_.Size)"
                Received   = "$($_.Received)"
            }
        })
        $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
        Show-Message ("EXPORTED $(@($rows).Count) previewed row(s)." + [Environment]::NewLine + "  $path" + [Environment]::NewLine +
            'This is the previewed subset, not necessarily the whole result set -- see the count above the list.')
        Write-CompletedLog -Action 'Report' -Target (Get-LogTarget "$($script:PreviewOf)" 'Content Search') -Result 'Success' `
            -Detail "Exported $(@($rows).Count) previewed row(s) to $path." -Environment $Environment
    } catch {
        Show-Message (Format-ErrorDetail -ErrorRecord $_ -Context "Writing $path")
    }
}
#endregion

#region Handlers ----------------------------------------------------------------
# B1: every handler wrapped. An exception escaping one is caught by .NET rather
# than PowerShell, and the operator gets a CLR dialog with no script name.
$btnCreate.Add_Click({ try { Start-NewSearch } catch { Show-HandlerError -ErrorRecord $_ -Context 'Create search' } })
$btnRefreshList.Add_Click({ try { Update-SearchGrid } catch { Show-HandlerError -ErrorRecord $_ -Context 'Refresh' } })

foreach ($c in @($cboKind, $cboRange)) {
    $c.Add_SelectedIndexChanged({ try { Update-QueryPreview } catch { Show-HandlerError -ErrorRecord $_ -Context 'Query' } })
}
foreach ($c in @($txtKeys, $txtFrom, $txtTo)) {
    $c.Add_TextChanged({ try { Update-QueryPreview } catch { Show-HandlerError -ErrorRecord $_ -Context 'Query' } })
}
$cboScope.Add_SelectedIndexChanged({
    try {
        $named = ("$($cboScope.SelectedItem)" -eq 'Named mailboxes')
        $txtUsers.Enabled = $named
        $lblUsers.Enabled = $named
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Scope' }
})

function Update-CustomDateEnabled {
    <#
        The custom date boxes are live only when they can affect the query.

        A BOX THAT ACCEPTS TYPING AND THEN DISCARDS IT IS THE BUG THIS EXISTS TO
        PREVENT. They used to be enabled always, and labelled 'From / To' --
        which is why an address landed in one and vanished. Greyed out, the
        question does not arise.

        Anything already typed is CLEARED when the range stops being custom,
        rather than left sitting in a disabled box looking like part of the
        search. A stale date is a worse lie than an empty one.
    #>
    $custom = ("$($cboRange.SelectedItem)" -eq 'Custom range')
    $lblCustom.Enabled      = $custom
    $txtCustomStart.Enabled = $custom
    $txtCustomEnd.Enabled   = $custom
    if (-not $custom) {
        if ("$($txtCustomStart.Text)".Trim() -or "$($txtCustomEnd.Text)".Trim()) {
            $txtCustomStart.Text = ''
            $txtCustomEnd.Text   = ''
        }
    }
}

$cboRange.Add_SelectedIndexChanged({
    try { Update-CustomDateEnabled } catch { Show-HandlerError -ErrorRecord $_ -Context 'Date range' }
})

$btnRunSearch.Add_Click({
    try {
        $row = Get-SelectedSearchRow
        if (-not $row) { Show-Message 'Select a search in the list.'; return }
        [void](Start-ComplianceSearch -Identity "$($row.Name)" -ErrorAction Stop)
        Show-Message "Started '$($row.Name)'. Refresh to watch its status."
        Write-CompletedLog -Action 'Read' -Target (Get-LogTarget "$($row.Name)" 'Content Search') -Result 'Success' `
            -Detail "Start-ComplianceSearch on '$($row.Name)'." -Environment $Environment
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Run search' }
})

$btnDeleteSearch.Add_Click({
    try {
        $row = Get-SelectedSearchRow
        if (-not $row) { Show-Message 'Select a search in the list.'; return }
        $r = Show-ToolDialog(
            "Delete the search '$($row.Name)'?" + [Environment]::NewLine + [Environment]::NewLine +
            'This removes the search and its results. IT DOES NOT DELETE ANY MAIL -- purging is on the third tab.',
            'Delete search',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        Remove-ComplianceSearch -Identity "$($row.Name)" -Confirm:$false -ErrorAction Stop
        Show-Message "Deleted the search '$($row.Name)'. No mail was affected."
        Write-CompletedLog -Action 'Delete' -Target (Get-LogTarget "$($row.Name)" 'Content Search') -Result 'Success' `
            -Detail "Remove-ComplianceSearch '$($row.Name)' (search definition only, no mail)." -Environment $Environment
        Update-SearchGrid
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Delete search' }
})

$btnPreview.Add_Click({ try { Start-PreviewOfSearch } catch { Show-HandlerError -ErrorRecord $_ -Context 'Preview' } })
$btnExportPv.Add_Click({ try { Export-PreviewCsv } catch { Show-HandlerError -ErrorRecord $_ -Context 'Export' } })

# A checkbox cell does not commit until it loses focus. Without this, ticking a
# row and pressing Purge immediately reads that last tick as unticked -- and in
# this tool that means deleting a different set than the one on screen.
$gridPreview.Add_CurrentCellDirtyStateChanged({
    try {
        if ($gridPreview.IsCurrentCellDirty) {
            $gridPreview.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    } catch { }
})

$btnTickAll.Add_Click({
    try { for ($i = 0; $i -lt $gridPreview.Rows.Count; $i++) { $gridPreview.Rows[$i].Cells[0].Value = $true } }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Tick all' }
})
$btnTickNone.Add_Click({
    try { for ($i = 0; $i -lt $gridPreview.Rows.Count; $i++) { $gridPreview.Rows[$i].Cells[0].Value = $false } }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Tick none' }
})

$btnPurgeSelected.Add_Click({ try { Start-PurgeSelected } catch { Show-HandlerError -ErrorRecord $_ -Context 'Purge ticked' } })
$btnPurgeAll.Add_Click({ try { Start-PurgeEverything } catch { Show-HandlerError -ErrorRecord $_ -Context 'Purge all' } })
#endregion

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Content Search and Purge' | Out-Null

# NOTHING IS QUERIED DURING CONSTRUCTION. The original called PerformClick on
# three Refresh buttons while building its tabs, so it fired three
# "please connect first" message boxes before the window had even appeared.
# The first read happens here, once, after the form exists.
$form.Add_Shown({
    try {
        $txtUsers.Enabled = ("$($cboScope.SelectedItem)" -eq 'Named mailboxes')
        Update-CustomDateEnabled
        Update-QueryPreview
        if (Update-ConnectionLine) {
            Update-SearchGrid
        } else {
            Show-Message ((Test-CompliancePrereq).Reason)
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Content Search' }
})

[void]$form.ShowDialog()
