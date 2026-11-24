#Requires -Version 5.1
<#
    Unified-Trace.ps1
    MOCKUP. One trace across every system a message crosses -- Exchange Online,
    on-premises Exchange, both Proofpoint pods and SendGrid.

    THIS TOOL INVENTS ITS DATA. Nothing here contacts Proofpoint or SendGrid,
    because there is no API access to either. It exists to answer "what would a
    consolidated trace look like, and is it worth building" before anybody
    spends money finding out.

    The Exchange halves are real in HybridMessageTrace.ps1. This is deliberately
    NOT wired to them -- see the note above Get-MockRows.
#>
# TOOLKIT-HELP-RUNS: NOTHING. Every row is generated locally. No system is contacted.
# TOOLKIT-VARIANT: Mock
# TOOLKIT-HELP-BEGIN
# A MOCKUP of a single message trace spanning every system mail crosses here.
#
# EVERY ROW IS INVENTED. This tool does not connect to Exchange, Proofpoint or
# SendGrid -- there is no API access to the last two, which is the whole reason
# it exists. It is a working sketch of what a consolidated trace COULD look
# like, so the idea can be judged before anybody buys anything.
#
# FIVE SOURCES, and the point is seeing them on one timeline:
#
#   ExO         Exchange Online
#   OnPrem      on-premises Exchange transport
#   Proofpoint  the primary pod -- inbound filtering
#   SER         Proofpoint Secure Email Relay -- application-generated mail
#   SendGrid    transactional send, and the suppression lists that silently
#               stop it
#
# WHY THIS IS WORTH BUILDING FOR REAL. "They never got the email" is currently
# four separate investigations in four consoles, and the answer is usually in
# the one you check last. The two SendGrid outcomes below are the clearest
# example: a message that was never sent at all because the address is on a
# suppression list leaves NO trace in Exchange, so the Exchange-only answer is
# "no such message" -- which is true and useless.
#
# THE REMOVE BUTTONS DO NOTHING. They show the confirmation and report what the
# real call would have been. They cannot change anything, because there is
# nothing connected to change.
#
# Try bob.smith@gcearth.com as a recipient. The mock data rotates through
# several scenarios, so searching twice gives a different story each time.
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

$script:AllRows    = @()
$script:Suppress   = @{}
$script:MockDir    = Join-Path $PSScriptRoot 'MockData\UnifiedTrace'
$script:DataFile   = ''
$script:Legacy     = @()

# WHICH MESSAGES ARE OPEN. Collapsed is the default and it is the important
# one: somebody who does not know the topology should see ONE line per message
# carrying the outcome that actually matters, not four rows of plumbing they
# have to read in order to work out which one decided it.
$script:SearchCancelled = $false
$script:Expanded   = @{}
$script:RowTrace   = @()

# EVERY ADDRESS IN THIS FILE IS FICTIONAL, and gcearth.com is not a domain this
# organisation owns. That is deliberate: mock data carrying a real address is
# one copy-paste away from a real ticket.
$script:MockRecipients = @(
    'bob.smith@gcearth.com'
    'alice.jones@gcearth.com'
    'carl.reed@gcearth.com'
    'dana.patel@gcearth.com'
)

#region Mock data ---------------------------------------------------------------
<#
    WHY THIS IS NOT WIRED TO THE REAL EXCHANGE TRACE.

    HybridMessageTrace.ps1 genuinely queries Exchange Online and on-premises
    transport, and it would have been easy to call it for the ExO and OnPrem
    rows and invent only the other three.

    THAT WOULD BE THE WORST POSSIBLE OUTCOME. A grid holding real Exchange rows
    beside invented Proofpoint rows, in one table, sorted together, is data that
    LOOKS entirely real and is half fiction. Somebody would export it into an
    incident ticket and nobody downstream could tell which half was which.

    So this tool touches nothing. Every row is fake, the banner says so, the
    grid says so per row, and the export filename says so. When the APIs exist,
    the mock generator is deleted and the real fetchers take its place -- the
    row shape and the merge do not change, which is most of what a mockup is
    for.
#>

function Get-MockMessageTable {
    <#
        THE CORPUS AS MESSAGES, NOT EVENTS.

        It used to be a flat list of rows and the UI guessed which belonged
        together by matching subjects. That is backwards: a message is the
        thing, and a hop is something that happened to it. Modelling it the
        right way round is what lets the grid collapse to one row per message
        and expand to the path it actually took.

        THE HOP ORDER IS THE TOPOLOGY, AND IT ONLY GOES ONE WAY. gcearth.com is
        a CUSTOMER domain, so everything here is outbound:

            OnPrem  ->  ExO  ->  Proofpoint  ->  gcearth MX
                        ExO  ->  Proofpoint  ->  gcearth MX
                                 SER         ->  gcearth MX
            SendGrid                         ->  gcearth MX

        Four rules fall out of that, and Test-HopOrder enforces every one:

          - OnPrem, when present, is FIRST. It is where the message originates.
          - ExO is NEVER LAST. It always hands off to an egress.
          - Proofpoint, SER and SendGrid are TERMINAL -- the last hop we own.
          - SendGrid never shares a message with Proofpoint or SER. It is a
            separate path out, not a stage in the same one.

        A CHAIN STOPS WHERE IT FAILS. A message quarantined at ExO has no
        Proofpoint hop, because it never got there. That is the whole reason
        the collapsed row shows the LAST hop -- it is the one that decided the
        outcome.
    #>
    $now = (Get-Date).ToUniversalTime()
    $t = { param($mins) $now.AddMinutes(-1 * $mins).ToString('yyyy-MM-dd HH:mm:ss') }

    return @(
        # ============================================================ bob.smith
        @{ Id='b01'; F='no-reply@northwind-demo.test'; R='bob.smith@gcearth.com'; Sub='Your statement is ready'; Sz=48210
           Hops=@( @{S='SendGrid'; M=6; O='Blocked'; Srv='sendgrid-api'; D='Spam Suppression -- the address is on the spam suppression list from a report on 2026-07-02. Nothing was sent.'} ) }

        @{ Id='b02'; F='payroll@northwind-demo.test'; R='bob.smith@gcearth.com'; Sub='Payslip 2026-08'; Sz=88320
           Hops=@( @{S='ExO'; M=45; O='Delivered'; Srv=''; D='Accepted from the mailbox and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=44; O='Delivered'; Srv='pod-a-mx01'; D='Clean. Relayed to the recipient domain.'} ) }

        @{ Id='b03'; F='newsletter@marketing-demo.test'; R='bob.smith@gcearth.com'; Sub='August roundup'; Sz=210556
           Hops=@( @{S='ExO'; M=121; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=120; O='Quarantined'; Srv='pod-a-mx02'; D='Bulk mail held in the pod quarantine. Releasable for 30 days.'; Q='Spam'} ) }

        @{ Id='b04'; F='invoice@sketchy-demo.test'; R='bob.smith@gcearth.com'; Sub='Invoice 7741 (attached)'; Sz=418200
           Hops=@( @{S='ExO'; M=134; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=133; O='Quarantined'; Srv='pod-a-mx03'; D='MALWARE: attachment matched a known dropper signature. Held in the pod malware quarantine.'; Q='Malware'} ) }

        @{ Id='b05'; F='promotions@bulk-demo.test'; R='bob.smith@gcearth.com'; Sub='Final notice: renew today'; Sz=64100
           Hops=@( @{S='ExO'; M=151; O='Quarantined'; Srv=''; D='EOP outbound spam quarantine, high confidence spam. Never reached the egress.'; Q='Spam'} ) }

        @{ Id='b06'; F='hr-docs@spoof-demo.test'; R='bob.smith@gcearth.com'; Sub='Updated policy.docm'; Sz=302880
           Hops=@( @{S='ExO'; M=166; O='Quarantined'; Srv=''; D='EOP MALWARE quarantine. Common Attachment Filter matched a macro-enabled document.'; Q='Malware'} ) }

        @{ Id='b07'; F='noreply@app-demo.test'; R='bob.smith@gcearth.com'; Sub='Ticket #44812 updated'; Sz=9880
           Hops=@( @{S='SER'; M=180; O='Delivered'; Srv='ser-relay-03'; D='Relayed by SER on behalf of the ticketing application.'} ) }

        @{ Id='b08'; F='j.wilson@partner-demo.test'; R='bob.smith@gcearth.com'; Sub='RE: Q3 contract review'; Sz=31900
           Hops=@( @{S='ExO'; M=206; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=205; O='Delivered'; Srv='pod-a-mx01'; D='Clean. Known correspondent.'} ) }

        @{ Id='b09'; F='teams@northwind-demo.test'; R='bob.smith@gcearth.com'; Sub='You have 3 missed messages'; Sz=14220
           Hops=@( @{S='ExO'; M=241; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=240; O='Delivered'; Srv='pod-a-mx02'; D='Clean.'} ) }

        @{ Id='b10'; F='alerts@northwind-demo.test'; R='bob.smith@gcearth.com'; Sub='Sign-in from a new device'; Sz=7640
           Hops=@( @{S='SendGrid'; M=262; O='Delivered'; Srv='sendgrid-api'; D='Accepted by the receiving server, 250 OK.'} ) }

        @{ Id='b11'; F='no-reply@travel-demo.test'; R='bob.smith@gcearth.com'; Sub='Itinerary confirmation LHR-JFK'; Sz=52110
           Hops=@( @{S='ExO'; M=291; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=290; O='Delivered'; Srv='pod-a-mx02'; D='Clean.'} ) }

        @{ Id='b12'; F='reports@app-demo.test'; R='bob.smith@gcearth.com'; Sub='Nightly reconciliation report'; Sz=128400
           Hops=@( @{S='SER'; M=330; O='Delivered'; Srv='ser-relay-01'; D='Relayed by SER.'} ) }

        @{ Id='b13'; F='m.chen@northwind-demo.test'; R='bob.smith@gcearth.com'; Sub='Lunch?'; Sz=3180
           Hops=@( @{S='ExO'; M=401; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=400; O='Delivered'; Srv='pod-a-mx01'; D='Clean.'} ) }

        # ========================================================== alice.jones
        # THE FULL THREE-HOP PATH. Her correspondents send from mailboxes that
        # are still on-premises, so every one starts at OnPrem.
        @{ Id='a01'; F='supplier@vendor-demo.test'; R='alice.jones@gcearth.com'; Sub='Invoice 99321 attached'; Sz=246118
           Hops=@( @{S='OnPrem'; M=10; O='Delivered'; Srv='EXCH-HUB-02'; D='RECEIVE from the mailbox -> SEND to the hybrid connector.'}
                   @{S='ExO'; M=9; O='Delivered'; Srv=''; D='Received across the hybrid connector and routed outbound.'}
                   @{S='Proofpoint'; M=8; O='Delivered'; Srv='pod-a-mx01'; D='Clean after attachment scan (1 attachment, 240 KB).'} ) }

        @{ Id='a02'; F='statements@app-demo.test'; R='alice.jones@gcearth.com'; Sub='Monthly statement'; Sz=54990
           Hops=@( @{S='SER'; M=52; O='Delivered'; Srv='ser-relay-01'; D='Relayed by SER.'} ) }

        @{ Id='a03'; F='b.taylor@partner-demo.test'; R='alice.jones@gcearth.com'; Sub='Signed NDA returned'; Sz=188220
           Hops=@( @{S='OnPrem'; M=142; O='Delivered'; Srv='EXCH-HUB-02'; D='RECEIVE -> SEND to the hybrid connector.'}
                   @{S='ExO'; M=141; O='Delivered'; Srv=''; D='Received across the hybrid connector and routed outbound.'}
                   @{S='Proofpoint'; M=140; O='Delivered'; Srv='pod-a-mx03'; D='Clean.'} ) }

        @{ Id='a04'; F='no-reply@northwind-demo.test'; R='alice.jones@gcearth.com'; Sub='Welcome to the portal'; Sz=15330
           Hops=@( @{S='SendGrid'; M=300; O='Delivered'; Srv='sendgrid-api'; D='Accepted by the receiving server, 250 OK.'} ) }

        @{ Id='a05'; F='k.reilly@northwind-demo.test'; R='alice.jones@gcearth.com'; Sub='Revised drawings v4'; Sz=4402100
           Hops=@( @{S='OnPrem'; M=190; O='Delivered'; Srv='EXCH-HUB-01'; D='RECEIVE -> SEND to the hybrid connector.'}
                   @{S='ExO'; M=189; O='Delivered'; Srv=''; D='Received across the hybrid connector and routed outbound.'}
                   @{S='Proofpoint'; M=188; O='Delivered'; Srv='pod-a-mx02'; D='Clean after attachment scan (3 attachments).'} ) }

        # ============================================================ carl.reed
        @{ Id='c01'; F='alerts@northwind-demo.test'; R='carl.reed@gcearth.com'; Sub='Unusual sign-in'; Sz=8120
           Hops=@( @{S='SendGrid'; M=4; O='Deferred'; Srv='sendgrid-api'; D='421 4.7.0 temporarily deferred by the receiving server. SendGrid will retry for 72 hours.'} ) }

        @{ Id='c02'; F='security@phish-demo.test'; R='carl.reed@gcearth.com'; Sub='URGENT: verify your account'; Sz=18440
           Hops=@( @{S='ExO'; M=26; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=25; O='Rejected'; Srv='pod-a-mx03'; D='Rejected at the pod: sender reputation and URL verdict both failed. Not delivered anywhere.'} ) }

        @{ Id='c03'; F='promo@marketing-demo.test'; R='carl.reed@gcearth.com'; Sub='50% off everything'; Sz=99210
           Hops=@( @{S='ExO'; M=62; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=61; O='Quarantined'; Srv='pod-a-mx01'; D='Bulk mail held in the pod quarantine. Releasable for 30 days.'; Q='Spam'} ) }

        @{ Id='c04'; F='billing@vendor-demo.test'; R='carl.reed@gcearth.com'; Sub='Receipt for order 55120'; Sz=27700
           Hops=@( @{S='ExO'; M=78; O='Quarantined'; Srv=''; D='EOP outbound spam quarantine, bulk complaint level exceeded. Never reached the egress.'; Q='Spam'} ) }

        @{ Id='c05'; F='internal.sender@northwind-demo.test'; R='carl.reed@gcearth.com'; Sub='Team meeting moved'; Sz=6120
           Hops=@( @{S='ExO'; M=96; O='FilteredAsSpam'; Srv=''; D='Marked as spam on the way out and delivered to the recipient Junk folder by the receiving side.'}
                   @{S='Proofpoint'; M=95; O='Delivered'; Srv='pod-a-mx02'; D='Relayed with a spam header already set.'} ) }

        @{ Id='c06'; F='scanner@northwind-demo.test'; R='carl.reed@gcearth.com'; Sub='Scanned document'; Sz=1204880
           Hops=@( @{S='OnPrem'; M=140; O='Failed'; Srv='EXCH-HUB-01'; D='550 5.4.1 Relay access denied -- the multifunction device is not in the receive connector scope.'} ) }

        @{ Id='c07'; F='noreply@app-demo.test'; R='carl.reed@gcearth.com'; Sub='Ticket #44980 closed'; Sz=7740
           Hops=@( @{S='SER'; M=210; O='Delivered'; Srv='ser-relay-02'; D='Relayed by SER.'} ) }

        @{ Id='c08'; F='k.osei@partner-demo.test'; R='carl.reed@gcearth.com'; Sub='Site survey photos'; Sz=8842000
           Hops=@( @{S='ExO'; M=256; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=255; O='Delivered'; Srv='pod-a-mx02'; D='Clean after attachment scan (6 attachments).'} ) }

        @{ Id='c09'; F='facilities@northwind-demo.test'; R='carl.reed@gcearth.com'; Sub='Car park closed Thursday'; Sz=4980
           Hops=@( @{S='ExO'; M=321; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=320; O='Delivered'; Srv='pod-a-mx01'; D='Clean.'} ) }

        # =========================================================== dana.patel
        @{ Id='d01'; F='no-reply@northwind-demo.test'; R='dana.patel@gcearth.com'; Sub='Your order has shipped'; Sz=22110
           Hops=@( @{S='SendGrid'; M=12; O='Blocked'; Srv='sendgrid-api'; D='Previous Bounce -- hard bounce on 2026-05-30 (550 5.1.10 recipient not found). Nothing was sent.'} ) }

        @{ Id='d02'; F='alerts@northwind-demo.test'; R='dana.patel@gcearth.com'; Sub='Security alert'; Sz=9010
           Hops=@( @{S='SendGrid'; M=70; O='Blocked'; Srv='sendgrid-api'; D='Previous Bounce -- suppressed by the same bounce record.'} ) }

        @{ Id='d03'; F='billing@northwind-demo.test'; R='dana.patel@gcearth.com'; Sub='Payment receipt'; Sz=13300
           Hops=@( @{S='SendGrid'; M=96; O='Blocked'; Srv='sendgrid-api'; D='Previous Bounce -- suppressed by the same bounce record.'} ) }

        @{ Id='d04'; F='statements@bank-demo.test'; R='dana.patel@gcearth.com'; Sub='Statement available'; Sz=61200
           Hops=@( @{S='ExO'; M=119; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=118; O='Quarantined'; Srv='pod-a-mx01'; D='Bulk mail held in the pod quarantine. Releasable for 30 days.'; Q='Spam'} ) }

        @{ Id='d05'; F='update@driver-demo.test'; R='dana.patel@gcearth.com'; Sub='Critical update required.js'; Sz=91400
           Hops=@( @{S='ExO'; M=133; O='Quarantined'; Srv=''; D='EOP MALWARE quarantine. Script attachment blocked by the Common Attachment Filter.'; Q='Malware'} ) }

        @{ Id='d06'; F='deals@marketing-demo.test'; R='dana.patel@gcearth.com'; Sub='Weekend flash sale'; Sz=74300
           Hops=@( @{S='ExO'; M=148; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=147; O='Quarantined'; Srv='pod-a-mx02'; D='Bulk mail held in the pod quarantine.'; Q='Spam'} ) }

        @{ Id='d07'; F='ceo@lookalike-demo.test'; R='dana.patel@gcearth.com'; Sub='Are you at your desk?'; Sz=4100
           Hops=@( @{S='ExO'; M=163; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=162; O='Rejected'; Srv='pod-a-mx03'; D='Rejected at the pod: display-name impersonation of an internal executive.'} ) }

        @{ Id='d08'; F='resume@applicant-demo.test'; R='dana.patel@gcearth.com'; Sub='Application - R Fielding.docx'; Sz=214600
           Hops=@( @{S='ExO'; M=178; O='Quarantined'; Srv=''; D='EOP outbound spam quarantine, high confidence spam. Never reached the egress.'; Q='Spam'} ) }

        @{ Id='d09'; F='hr@northwind-demo.test'; R='dana.patel@gcearth.com'; Sub='Benefits enrolment closes Friday'; Sz=41220
           Hops=@( @{S='ExO'; M=201; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=200; O='Delivered'; Srv='pod-a-mx02'; D='Clean.'} ) }

        @{ Id='d10'; F='invoices@app-demo.test'; R='dana.patel@gcearth.com'; Sub='Invoice batch 2026-08-24'; Sz=302400
           Hops=@( @{S='SER'; M=226; O='Delivered'; Srv='ser-relay-01'; D='Relayed by SER.'} ) }

        @{ Id='d11'; F='noreply@app-demo.test'; R='dana.patel@gcearth.com'; Sub='Your request was approved'; Sz=11400
           Hops=@( @{S='SER'; M=248; O='Delivered'; Srv='ser-relay-02'; D='Relayed by SER.'} ) }

        @{ Id='d12'; F='no-reply@survey-demo.test'; R='dana.patel@gcearth.com'; Sub='How did we do?'; Sz=6600
           Hops=@( @{S='SendGrid'; M=272; O='Deferred'; Srv='sendgrid-api'; D='451 4.3.0 greylisted by the receiving server. Retrying.'} ) }

        @{ Id='d13'; F='t.iqbal@partner-demo.test'; R='dana.patel@gcearth.com'; Sub='RE: Site visit next week'; Sz=28800
           Hops=@( @{S='OnPrem'; M=292; O='Delivered'; Srv='EXCH-HUB-02'; D='RECEIVE -> SEND to the hybrid connector.'}
                   @{S='ExO'; M=291; O='Delivered'; Srv=''; D='Received across the hybrid connector and routed outbound.'}
                   @{S='Proofpoint'; M=290; O='Delivered'; Srv='pod-a-mx01'; D='Clean. Known correspondent.'} ) }

        @{ Id='d14'; F='r.novak@northwind-demo.test'; R='dana.patel@gcearth.com'; Sub='Slides for Monday'; Sz=2204100
           Hops=@( @{S='ExO'; M=311; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=310; O='Delivered'; Srv='pod-a-mx03'; D='Clean.'} ) }

        @{ Id='d15'; F='it-support@northwind-demo.test'; R='dana.patel@gcearth.com'; Sub='Your laptop is ready'; Sz=5240
           Hops=@( @{S='ExO'; M=345; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=344; O='Delivered'; Srv='pod-a-mx01'; D='Clean.'} ) }

        @{ Id='d16'; F='no-reply@events-demo.test'; R='dana.patel@gcearth.com'; Sub='Your ticket for the summit'; Sz=38800
           Hops=@( @{S='ExO'; M=381; O='Delivered'; Srv=''; D='Accepted and routed to the outbound connector.'}
                   @{S='Proofpoint'; M=380; O='Delivered'; Srv='pod-a-mx03'; D='Clean. Relayed to the recipient domain.'} ) }
    ) | ForEach-Object {
        $m = $_
        [PSCustomObject]@{
            Id = $m.Id; Sender = $m.F; Recipient = $m.R; Subject = $m.Sub; SizeBytes = $m.Sz
            Hops = @($m.Hops | ForEach-Object {
                [PSCustomObject]@{
                    Source = $_.S; Outcome = $_.O; Server = $_.Srv; Detail = $_.D
                    TimestampUtc = (& $t $_.M)
                    Quarantine = $(if ($_.ContainsKey('Q')) { "$($_.Q)" } else { '' })
                }
            })
        }
    }
}

function Get-HopRank {
    <#
        Where a source sits in the outbound path. Lower is earlier.

        SendGrid and SER rank alongside Proofpoint because they are alternative
        EGRESSES, not later stages -- a message leaves by exactly one of them.
    #>
    param([string]$Source)
    switch ("$Source") {
        'OnPrem'     { return 1 }
        'ExO'        { return 2 }
        'Proofpoint' { return 3 }
        'SER'        { return 3 }
        'SendGrid'   { return 3 }
        default      { return 9 }
    }
}

function Test-HopOrder {
    <#
        Check the corpus obeys the topology, and say so loudly if it does not.

        THE RULES ARE IN THE DATA, SO THEY CAN BE BROKEN BY EDITING THE DATA --
        and the CSV is deliberately hand-editable. A message whose hops are out
        of order would collapse to the wrong final row and quietly show the
        wrong outcome, which is the one failure a trace tool must not have.

        Returns a list of complaints, empty when the corpus is sound.
    #>
    param($Messages)

    $bad = New-Object System.Collections.ArrayList
    foreach ($m in @($Messages)) {
        $hops = @($m.Hops)
        if (@($hops).Count -eq 0) { [void]$bad.Add("$($m.Id): no hops"); continue }

        $srcs = @($hops | ForEach-Object { "$($_.Source)" })
        $last = $srcs[-1]

        if ($srcs -contains 'OnPrem' -and $srcs[0] -ne 'OnPrem') {
            [void]$bad.Add("$($m.Id): OnPrem is present but not first ($($srcs -join ' -> '))")
        }
        if ($last -eq 'ExO' -and @($hops).Count -gt 0) {
            # ExO last is only legitimate when ExO itself stopped the message.
            if ("$($hops[-1].Outcome)" -eq 'Delivered') {
                [void]$bad.Add("$($m.Id): ExO is the last hop and reports Delivered -- it must hand off to an egress ($($srcs -join ' -> '))")
            }
        }
        if (($srcs -contains 'SendGrid') -and (($srcs -contains 'Proofpoint') -or ($srcs -contains 'SER'))) {
            [void]$bad.Add("$($m.Id): SendGrid shares a message with another egress ($($srcs -join ' -> '))")
        }
        for ($i = 1; $i -lt @($srcs).Count; $i++) {
            if ((Get-HopRank -Source $srcs[$i]) -lt (Get-HopRank -Source $srcs[$i - 1])) {
                [void]$bad.Add("$($m.Id): hop order goes backwards ($($srcs -join ' -> '))")
                break
            }
        }
    }
    return @($bad.ToArray())
}

function Get-MockDataTable {
    <#
        Flatten the messages into one row per hop, ordered, carrying the
        TraceId that binds them back together.
    #>
    $out = New-Object System.Collections.ArrayList
    foreach ($m in (Get-MockMessageTable)) {
        $i = 0
        foreach ($h in @($m.Hops)) {
            $i++
            [void]$out.Add([PSCustomObject]@{
                TraceId      = $m.Id
                HopIndex     = $i
                HopCount     = @($m.Hops).Count
                TimestampUtc = $h.TimestampUtc
                Source       = $h.Source
                Sender       = $m.Sender
                Recipient    = $m.Recipient
                Subject      = $m.Subject
                Outcome      = $h.Outcome
                Detail       = $h.Detail
                Server       = $h.Server
                SizeBytes    = $m.SizeBytes
                MessageId    = ('<{0}@northwind-demo.test>' -f $m.Id)
                Quarantine   = $h.Quarantine
            })
        }
    }
    return @($out.ToArray())
}

function Initialize-MockData {
    <#
        Write the corpus out as ONE CSV and return its path.

        ONE FILE, NOT FOUR. The four rotating scenario files are gone: the
        product being sketched searches a single corpus, and rotation made it
        impossible to show a wildcard returning everything at once.

        REWRITTEN IF THE SCHEMA MOVED, otherwise left alone, so hand-edits
        survive. A file written before a column existed would have Import-Csv
        fill that column with nothing rather than failing -- every quarantine
        row would quietly vanish and the release buttons would never enable,
        with no error anywhere.
    #>
    if (-not (Test-Path -LiteralPath $script:MockDir)) {
        [void](New-Item -ItemType Directory -Path $script:MockDir -Force)
    }

    $f = Join-Path $script:MockDir 'unified-trace-mock.csv'

    $stale = $true
    if (Test-Path -LiteralPath $f) {
        $stale = $false
        try {
            $head = @(Get-Content -LiteralPath $f -TotalCount 1)
            if (@($head).Count -eq 0 -or "$($head[0])" -notmatch '(?i)traceid') { $stale = $true }
        } catch { $stale = $true }
    }
    if ($stale) { (Get-MockDataTable) | Export-Csv -LiteralPath $f -NoTypeInformation -Encoding UTF8 }

    # THE OLD ROTATION FILES ARE REPORTED, NOT DELETED. They may have been
    # hand-edited, and silently ignoring somebody's edits is worse than saying
    # they are no longer read.
    $script:Legacy = @(Get-ChildItem -LiteralPath $script:MockDir -Filter 'scenario-*.csv' -ErrorAction SilentlyContinue |
                       ForEach-Object { $_.Name })
    return $f
}

function Get-MockRows {
    <#
        The whole corpus, every time. No rotation.
    #>
    if (-not $script:DataFile -or -not (Test-Path -LiteralPath $script:DataFile)) { return @() }

    $raw = @()
    try { $raw = @(Import-Csv -LiteralPath $script:DataFile) } catch { return @() }

    $out = New-Object System.Collections.ArrayList
    foreach ($r in $raw) {
        $utc = $null
        try {
            $utc = [datetime]::Parse("$($r.TimestampUtc)",
                [System.Globalization.CultureInfo]::InvariantCulture,
                ([System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal))
        } catch { $utc = (Get-Date).ToUniversalTime() }

        $sz = 0; $hi = 0; $hc = 0
        [void][long]::TryParse("$($r.SizeBytes)", [ref]$sz)
        [void][int]::TryParse("$($r.HopIndex)", [ref]$hi)
        [void][int]::TryParse("$($r.HopCount)", [ref]$hc)

        [void]$out.Add([PSCustomObject]@{
            # THE FIELDS THAT BIND HOPS INTO A MESSAGE. Losing these is not a
            # cosmetic fault: every row collapses into one group and the grid
            # shows a single line claiming seventy hops.
            TraceId      = "$($r.TraceId)"
            HopIndex     = $hi
            HopCount     = $hc
            TimestampUtc = $utc
            Source       = "$($r.Source)"
            Sender       = "$($r.Sender)"
            Recipient    = "$($r.Recipient)"
            Subject      = "$($r.Subject)"
            Outcome      = "$($r.Outcome)"
            Detail       = "$($r.Detail)"
            Server       = "$($r.Server)"
            SizeBytes    = $sz
            MessageId    = "$($r.MessageId)"
            Quarantine   = "$($r.Quarantine)"
            # KEPT ON THE OBJECT THOUGH IT IS NOT A GRID COLUMN. The EXPORT is
            # what outlives this window and gets forwarded without the banner,
            # so it carries a MOCK_DATA column and a filename that says so.
            Mock         = $true
        })
    }
    return @($out.ToArray())
}

function Test-TraceMatch {
    <#
        Does this value match what the operator typed?

        THREE BEHAVIOURS FROM ONE BOX, decided by what is in it:

            empty            matches everything -- no filter
            contains * or ?  treated as a wildcard, so *@gcearth.com finds every
                             recipient in the domain
            anything else    substring, so 'bob' finds bob.smith without
                             anybody having to type the wildcards

        SUBSTRING RATHER THAN EXACT for the plain case, because a trace box is
        used with half an address far more often than with all of it. An exact
        match would make 'bob' return nothing, which reads as "no such
        messages".
    #>
    param([string]$Value, [string]$Pattern)

    $p = "$Pattern".Trim()
    if (-not $p) { return $true }
    $v = "$Value"
    if ($p -match '[\*\?]') { return ($v -like $p) }
    return ($v -like "*$p*")
}

function Get-QuarantineAction {
    <#
        Can this row be released, and if not, why not.

        SPAM AND MALWARE ARE NOT THE SAME DECISION, which is the whole reason
        the button greys rather than being absent.

        A spam verdict is a judgement call and wrong often enough that releasing
        is routine. A malware verdict means something matched a signature, and
        in EOP an end user cannot release it at all -- the item is admin-only by
        default and most tenants block admin release too. Proofpoint's malware
        quarantine behaves the same way for the same reason.

        A BUTTON THAT IS ABSENT LOOKS LIKE A MISSING FEATURE. One that is greyed
        with a reason on hover teaches the difference, which is the entire point
        of putting it in a mockup.
    #>
    param($Row)

    $out = [PSCustomObject]@{ Quarantined = $false; Source = ''; Kind = ''; CanRelease = $false; Reason = '' }
    if (-not $Row) { $out.Reason = 'No row selected.'; return $out }
    if ("$($Row.Outcome)" -ne 'Quarantined') { $out.Reason = 'The selected row is not quarantined.'; return $out }

    $out.Quarantined = $true
    $out.Source = "$($Row.Source)"
    $out.Kind   = "$($Row.Quarantine)"

    if ($out.Kind -eq 'Malware') {
        $out.CanRelease = $false
        $out.Reason = 'MALWARE quarantine cannot be released from here. EOP makes malware admin-only and most tenants block release entirely; Proofpoint treats its malware quarantine the same way. Releasing known-malicious mail to a mailbox is a security decision, not a mail-flow one -- it goes through the security team.'
        return $out
    }

    $out.CanRelease = $true
    $out.Reason = "Spam quarantine on $($out.Source). Releasable."
    return $out
}

function Get-SuppressionState {
    <#
        The SendGrid suppression lists, derived from the rows on screen.

        DERIVED, NOT STORED, so the panel cannot disagree with the grid. In a
        real build this is a GET against the suppression API; here the blocked
        rows already say which list caught them, and re-deriving is closer to
        what the real thing returns than a second table would be.
    #>
    param($Rows)

    $state = @{}
    foreach ($r in @($Rows)) {
        if ("$($r.Source)" -ne 'SendGrid') { continue }
        if ("$($r.Outcome)" -ne 'Blocked') { continue }

        $who = "$($r.Recipient)".ToLowerInvariant()
        if (-not $state.ContainsKey($who)) {
            $state[$who] = [PSCustomObject]@{ Recipient = "$($r.Recipient)"; Spam = $false; Bounce = $false; Why = New-Object System.Collections.ArrayList }
        }
        if ("$($r.Detail)" -match '(?i)spam suppression') {
            $state[$who].Spam = $true
            [void]$state[$who].Why.Add("Spam: $($r.Detail)")
        }
        if ("$($r.Detail)" -match '(?i)previous bounce|bounce list|hard.bounce') {
            $state[$who].Bounce = $true
            [void]$state[$who].Why.Add("Bounce: $($r.Detail)")
        }
    }
    return $state
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Unified Trace  --  MOCKUP'
$form.ClientSize      = New-Object System.Drawing.Size(1280, 880)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'
$form.AutoScroll      = $true

# THE BANNER IS THE MOST IMPORTANT CONTROL ON THE FORM.
#
# Everything below it is plausible, formatted like a real trace, and completely
# invented. One person pasting a screenshot of this into an incident bridge
# would do more damage than the tool is worth, so it says so at the top, in the
# title bar, on every row, and in the export filename.
$lblMock           = New-Object System.Windows.Forms.Label
$lblMock.Text      = 'MOCKUP -- This shows what is Possible with API access to ProofPoint & SendGrid. Every row is invented; nothing is queried. Do not paste these results into a ticket.'
$lblMock.Location  = New-Object System.Drawing.Point(14, 10)
$lblMock.Size      = New-Object System.Drawing.Size(1210, 34)
$lblMock.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblMock.ForeColor = [System.Drawing.Color]::FromArgb(255, 105, 97)
$lblMock.Tag       = 'SemanticColour'
$lblMock.Anchor    = 'Top,Left,Right'
[void]$form.Controls.Add($lblMock)

$grpFind          = New-Object System.Windows.Forms.GroupBox
$grpFind.Text     = 'Search'
$grpFind.Location = New-Object System.Drawing.Point(14, 48)
$grpFind.Size     = New-Object System.Drawing.Size(1210, 138)
$grpFind.Anchor   = 'Top,Left,Right'
[void]$form.Controls.Add($grpFind)

$lblTo          = New-Object System.Windows.Forms.Label
$lblTo.Text     = 'Recipient'
$lblTo.Location = New-Object System.Drawing.Point(12, 28)
$lblTo.Size     = New-Object System.Drawing.Size(70, 20)
[void]$grpFind.Controls.Add($lblTo)

$cboTo               = New-Object System.Windows.Forms.ComboBox
$cboTo.Location      = New-Object System.Drawing.Point(88, 25)
$cboTo.Size          = New-Object System.Drawing.Size(260, 22)
$cboTo.DropDownStyle = 'DropDown'
# THE WILDCARD IS THE FIRST ENTRY, because it is the one that demonstrates the
# point -- everything for the domain, from five systems, on one timeline.
[void]$cboTo.Items.Add('*@gcearth.com')
foreach ($r in $script:MockRecipients) { [void]$cboTo.Items.Add($r) }
[void]$grpFind.Controls.Add($cboTo)

$lblFrom          = New-Object System.Windows.Forms.Label
$lblFrom.Text     = 'Sender'
$lblFrom.Location = New-Object System.Drawing.Point(362, 28)
$lblFrom.Size     = New-Object System.Drawing.Size(60, 20)
[void]$grpFind.Controls.Add($lblFrom)

$txtFrom          = New-Object System.Windows.Forms.TextBox
$txtFrom.Location = New-Object System.Drawing.Point(428, 25)
$txtFrom.Size     = New-Object System.Drawing.Size(240, 22)
[void]$grpFind.Controls.Add($txtFrom)

$lblSubj          = New-Object System.Windows.Forms.Label
$lblSubj.Text     = 'Subject has'
$lblSubj.Location = New-Object System.Drawing.Point(682, 28)
$lblSubj.Size     = New-Object System.Drawing.Size(76, 20)
[void]$grpFind.Controls.Add($lblSubj)

$txtSubj          = New-Object System.Windows.Forms.TextBox
$txtSubj.Location = New-Object System.Drawing.Point(764, 25)
$txtSubj.Size     = New-Object System.Drawing.Size(240, 22)
[void]$grpFind.Controls.Add($txtSubj)

$btnGo           = New-Object System.Windows.Forms.Button
$btnGo.Text      = 'Trace'
$btnGo.Location  = New-Object System.Drawing.Point(1020, 23)
$btnGo.Size      = New-Object System.Drawing.Size(170, 28)
$btnGo.FlatStyle = 'Flat'
[void]$grpFind.Controls.Add($btnGo)

# THE SOURCE RADIOS. Six, in a panel of their own so the group owns the
# mutual exclusion rather than the form -- every RadioButton on one parent is
# one group, and mixing them with anything else on the form is how an unrelated
# radio elsewhere starts clearing these.
$pnlSrc          = New-Object System.Windows.Forms.Panel
$pnlSrc.Location = New-Object System.Drawing.Point(88, 56)
$pnlSrc.Size     = New-Object System.Drawing.Size(920, 28)
[void]$grpFind.Controls.Add($pnlSrc)

$lblSrc          = New-Object System.Windows.Forms.Label
$lblSrc.Text     = 'Source'
$lblSrc.Location = New-Object System.Drawing.Point(12, 60)
$lblSrc.Size     = New-Object System.Drawing.Size(70, 20)
[void]$grpFind.Controls.Add($lblSrc)

$script:Radios = @()
$x = 0
foreach ($spec in @(
    @{ N = 'ExO';        W = 70;  T = 'Exchange Online' }
    @{ N = 'OnPrem';     W = 86;  T = 'On-premises Exchange transport' }
    @{ N = 'Proofpoint'; W = 104; T = 'Proofpoint -- primary pod' }
    @{ N = 'SER';        W = 66;  T = 'Proofpoint Secure Email Relay pod' }
    @{ N = 'SendGrid';   W = 96;  T = 'SendGrid transactional send' }
    @{ N = 'All';        W = 60;  T = 'Every source on one timeline' })) {

    $rb          = New-Object System.Windows.Forms.RadioButton
    $rb.Text     = $spec.N
    $rb.Location = New-Object System.Drawing.Point($x, 4)
    $rb.Size     = New-Object System.Drawing.Size($spec.W, 20)
    # B5: the value rides on .Tag, never captured in a closure.
    $rb.Tag      = $spec.N
    if ($spec.N -eq 'All') { $rb.Checked = $true }
    [void]$pnlSrc.Controls.Add($rb)
    $script:Radios += $rb
    $x += $spec.W
}

$tips = New-Object System.Windows.Forms.ToolTip
foreach ($rb in $script:Radios) {
    switch ("$($rb.Tag)") {
        'ExO'        { $tips.SetToolTip($rb, 'Exchange Online') }
        'OnPrem'     { $tips.SetToolTip($rb, 'On-premises Exchange transport logs') }
        'Proofpoint' { $tips.SetToolTip($rb, 'Proofpoint primary pod -- inbound filtering, quarantine and reject') }
        'SER'        { $tips.SetToolTip($rb, 'Proofpoint Secure Email Relay -- application-generated mail') }
        'SendGrid'   { $tips.SetToolTip($rb, 'SendGrid transactional send, including suppression blocks that never reach Exchange') }
        'All'        { $tips.SetToolTip($rb, 'Every source, merged onto one timeline') }
    }
}

$btnStop           = New-Object System.Windows.Forms.Button
$btnStop.Text      = 'Stop'
$btnStop.Location  = New-Object System.Drawing.Point(1020, 55)
$btnStop.Size      = New-Object System.Drawing.Size(170, 26)
$btnStop.FlatStyle = 'Flat'
$btnStop.Visible   = $false
[void]$grpFind.Controls.Add($btnStop)

# THE BAR LIVES UNDER THE SEARCH ROW and is hidden when idle. A progress bar
# that is always on screen at zero reads as something being stuck.
$pbSearch          = New-Object System.Windows.Forms.ProgressBar
$pbSearch.Location = New-Object System.Drawing.Point(88, 120)
$pbSearch.Size     = New-Object System.Drawing.Size(916, 10)
$pbSearch.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
$pbSearch.Visible  = $false
[void]$grpFind.Controls.Add($pbSearch)
$lblTz          = New-Object System.Windows.Forms.Label
$lblTz.Text     = 'Show times in'
$lblTz.Location = New-Object System.Drawing.Point(12, 96)
$lblTz.Size     = New-Object System.Drawing.Size(80, 20)
[void]$grpFind.Controls.Add($lblTz)

$cboTz               = New-Object System.Windows.Forms.ComboBox
$cboTz.Location      = New-Object System.Drawing.Point(98, 93)
$cboTz.Size          = New-Object System.Drawing.Size(160, 22)
$cboTz.DropDownStyle = 'DropDownList'
foreach ($z in @('Local', 'UTC')) { [void]$cboTz.Items.Add($z) }
$cboTz.SelectedIndex = 0
[void]$grpFind.Controls.Add($cboTz)

$lblOut          = New-Object System.Windows.Forms.Label
$lblOut.Text     = 'Outcome'
$lblOut.Location = New-Object System.Drawing.Point(272, 96)
$lblOut.Size     = New-Object System.Drawing.Size(60, 20)
[void]$grpFind.Controls.Add($lblOut)

$cboOutcome               = New-Object System.Windows.Forms.ComboBox
$cboOutcome.Location      = New-Object System.Drawing.Point(338, 93)
$cboOutcome.Size          = New-Object System.Drawing.Size(180, 22)
$cboOutcome.DropDownStyle = 'DropDownList'
[void]$grpFind.Controls.Add($cboOutcome)

$lblScenario           = New-Object System.Windows.Forms.Label
$lblScenario.Location  = New-Object System.Drawing.Point(536, 96)
$lblScenario.Size      = New-Object System.Drawing.Size(290, 20)
$lblScenario.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grpFind.Controls.Add($lblScenario)

$btnExpandAll           = New-Object System.Windows.Forms.Button
$btnExpandAll.Text      = 'Expand all'
$btnExpandAll.Location  = New-Object System.Drawing.Point(836, 91)
$btnExpandAll.Size      = New-Object System.Drawing.Size(88, 26)
$btnExpandAll.FlatStyle = 'Flat'
[void]$grpFind.Controls.Add($btnExpandAll)

$btnCollapseAll           = New-Object System.Windows.Forms.Button
$btnCollapseAll.Text      = 'Collapse all'
$btnCollapseAll.Location  = New-Object System.Drawing.Point(928, 91)
# 88, ITS NEIGHBOUR'S WIDTH, AND PUT BACK 2026-09-11. It was widened to 94
# earlier the same day because 'Collapse all' measured 89px at a 9.75pt bold
# floor. That floor came back down to 9pt, where the caption needs 77px and 88
# is comfortable.
#
# WIDENING IT WAS ALSO WRONG, not merely unnecessary: btnExport starts at 1020
# and 928+94 is 1022, so the 94 overlapped it by 2px. The claim that grpFind
# being 1210 wide made this safe ignored the button sitting immediately to the
# right, and the layout check that should have caught it was pointed at a stale
# local copy of the toolkit. Both fixed.
$btnCollapseAll.Size      = New-Object System.Drawing.Size(88, 26)
$btnCollapseAll.FlatStyle = 'Flat'
[void]$grpFind.Controls.Add($btnCollapseAll)
$btnExport           = New-Object System.Windows.Forms.Button
$btnExport.Text      = 'Export CSV'
$btnExport.Location  = New-Object System.Drawing.Point(1020, 91)
$btnExport.Size      = New-Object System.Drawing.Size(170, 26)
$btnExport.FlatStyle = 'Flat'
$btnExport.Enabled   = $false
[void]$grpFind.Controls.Add($btnExport)

$grid                     = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(14, 196)
$grid.Size                = New-Object System.Drawing.Size(1252, 330)
$grid.ReadOnly            = $true
$grid.AllowUserToAddRows  = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode       = 'FullRowSelect'
$grid.MultiSelect         = $false
$grid.Anchor              = 'Top,Left,Right,Bottom'
$grid.Font                = New-Object System.Drawing.Font('Consolas', 8)
[void]$form.Controls.Add($grid)

# SENDGRID SUPPRESSION -- the thing an Exchange-only trace can never show.
$grpSup          = New-Object System.Windows.Forms.GroupBox
$grpSup.Text     = 'SendGrid suppression  --  why a message was never sent at all'
$grpSup.Location = New-Object System.Drawing.Point(14, 534)
$grpSup.Size     = New-Object System.Drawing.Size(620, 116)
$grpSup.Anchor   = 'Left,Bottom'
[void]$form.Controls.Add($grpSup)

$lblSupWho          = New-Object System.Windows.Forms.Label
$lblSupWho.Location = New-Object System.Drawing.Point(12, 22)
$lblSupWho.Size     = New-Object System.Drawing.Size(396, 20)
$lblSupWho.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$grpSup.Controls.Add($lblSupWho)

$lblSupDetail            = New-Object System.Windows.Forms.Label
$lblSupDetail.Location   = New-Object System.Drawing.Point(12, 44)
$lblSupDetail.Size       = New-Object System.Drawing.Size(396, 62)
$lblSupDetail.Font       = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grpSup.Controls.Add($lblSupDetail)

$btnRmSpam           = New-Object System.Windows.Forms.Button
$btnRmSpam.Text      = 'Remove Spam Suppression'
$btnRmSpam.Location  = New-Object System.Drawing.Point(414, 26)
$btnRmSpam.Size      = New-Object System.Drawing.Size(196, 30)
$btnRmSpam.FlatStyle = 'Flat'
$btnRmSpam.Enabled   = $false
[void]$grpSup.Controls.Add($btnRmSpam)

$btnRmBounce           = New-Object System.Windows.Forms.Button
$btnRmBounce.Text      = 'Remove Bounce'
$btnRmBounce.Location  = New-Object System.Drawing.Point(414, 62)
$btnRmBounce.Size      = New-Object System.Drawing.Size(196, 30)
$btnRmBounce.FlatStyle = 'Flat'
$btnRmBounce.Enabled   = $false
[void]$grpSup.Controls.Add($btnRmBounce)

# QUARANTINE RELEASE -- driven by the SELECTED ROW, not by the recipient.
#
# Suppression is a property of the address; quarantine is a property of one
# message. Two different scopes, so two panels rather than one row of buttons
# that mean different things depending on what is highlighted.
$grpQuar          = New-Object System.Windows.Forms.GroupBox
$grpQuar.Text     = 'Quarantine  --  acts on the row selected above'
$grpQuar.Location = New-Object System.Drawing.Point(646, 534)
$grpQuar.Size     = New-Object System.Drawing.Size(620, 116)
$grpQuar.Anchor   = 'Left,Right,Bottom'
[void]$form.Controls.Add($grpQuar)

$lblQuarWho          = New-Object System.Windows.Forms.Label
$lblQuarWho.Location = New-Object System.Drawing.Point(12, 22)
$lblQuarWho.Size     = New-Object System.Drawing.Size(396, 20)
$lblQuarWho.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$grpQuar.Controls.Add($lblQuarWho)

$lblQuarDetail          = New-Object System.Windows.Forms.Label
$lblQuarDetail.Location = New-Object System.Drawing.Point(12, 44)
$lblQuarDetail.Size     = New-Object System.Drawing.Size(396, 62)
$lblQuarDetail.Font     = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grpQuar.Controls.Add($lblQuarDetail)

$btnRelEop           = New-Object System.Windows.Forms.Button
$btnRelEop.Text      = 'Release from EOP'
$btnRelEop.Location  = New-Object System.Drawing.Point(414, 26)
$btnRelEop.Size      = New-Object System.Drawing.Size(196, 30)
$btnRelEop.FlatStyle = 'Flat'
$btnRelEop.Enabled   = $false
[void]$grpQuar.Controls.Add($btnRelEop)

$btnRelPfpt           = New-Object System.Windows.Forms.Button
$btnRelPfpt.Text      = 'Release from Proofpoint'
$btnRelPfpt.Location  = New-Object System.Drawing.Point(414, 62)
$btnRelPfpt.Size      = New-Object System.Drawing.Size(196, 30)
$btnRelPfpt.FlatStyle = 'Flat'
$btnRelPfpt.Enabled   = $false
[void]$grpQuar.Controls.Add($btnRelPfpt)

$quarTips = New-Object System.Windows.Forms.ToolTip
$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(14, 658)
$txtOut.Size       = New-Object System.Drawing.Size(1252, 208)
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
# B6.
$txtOut.WordWrap   = $false
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)
#endregion

#region Behaviour ---------------------------------------------------------------
function Get-SelectedSource {
    foreach ($rb in $script:Radios) { if ($rb.Checked) { return "$($rb.Tag)" } }
    return 'All'
}

function ConvertTo-DisplayTime {
    param($Utc)
    if (-not $Utc) { return $null }
    if ("$($cboTz.SelectedItem)" -eq 'UTC') { return $Utc }
    return $Utc.ToLocalTime()
}

function Get-FinalHop {
    <#
        The hop that decided the outcome -- the last one in the chain.

        NOT THE LATEST BY TIMESTAMP. Hops within a message can share a second,
        and on a busy relay the clocks are close enough that sorting by time
        picks an arbitrary one. HopIndex is the order the message actually
        travelled in, which is what the collapsed row must report.
    #>
    param($Hops)
    return @($Hops | Sort-Object HopIndex)[-1]
}

function Update-Grid {
    <#
        One row per MESSAGE, collapsed, showing the hop that decided it.

        COLLAPSED IS THE DEFAULT AND THAT IS THE POINT OF THIS VERSION. The
        flat view showed 'Your ticket for the summit' twice -- once for ExO and
        once for Proofpoint -- which is accurate, useless, and actively
        misleading to anybody who does not already know the topology. They see
        two messages where there is one.

        Expanding shows the hops in the order the message travelled, plus the
        handoff to the recipient's MX, which is the hop nobody can see today.
    #>
    $want = Get-SelectedSource
    $rows = @($script:AllRows)

    # SOURCE FILTERING KEEPS THE WHOLE MESSAGE, not just the matching hop.
    # Filtering to Proofpoint and then showing a bare Proofpoint row would hide
    # that the message came through ExO first -- the filter is there to find
    # messages that touched a system, not to amputate their history.
    if ($want -ne 'All') {
        $keep = @($rows | Where-Object { "$($_.Source)" -eq $want } | ForEach-Object { $_.TraceId } | Sort-Object -Unique)
        $rows = @($rows | Where-Object { $keep -contains $_.TraceId })
    }

    $oc = "$($cboOutcome.SelectedItem)"
    if ($oc -and $oc -ne '(any outcome)') {
        $keep = @($rows | Where-Object { "$($_.Outcome)" -eq $oc } | ForEach-Object { $_.TraceId } | Sort-Object -Unique)
        $rows = @($rows | Where-Object { $keep -contains $_.TraceId })
    }

    $groups = @($rows | Group-Object TraceId)
    # Newest message first, by the hop that finished it.
    $groups = @($groups | Sort-Object { (Get-FinalHop -Hops $_.Group).TimestampUtc } -Descending)

    $t = New-Object System.Data.DataTable
    foreach ($c in @(' ', 'Time', 'Source', 'Sender', 'Recipient', 'Subject', 'Outcome', 'Server', 'Size')) {
        [void]$t.Columns.Add($c)
    }

    $flat    = New-Object System.Collections.ArrayList
    $traceOf = New-Object System.Collections.ArrayList
    foreach ($g in $groups) {
        $hops  = @($g.Group | Sort-Object HopIndex)
        $final = Get-FinalHop -Hops $hops
        # ContainsKey ALONE IS NOT ENOUGH. A message can be in the set from an
        # Expand All that ran before a filter reduced it to one hop, and it
        # would then render children with no way to close them.
        $open  = ($script:Expanded.ContainsKey("$($g.Name)") -and @($hops).Count -gt 1)

        # ONLY A MULTI-HOP MESSAGE EXPANDS.
        #
        # A one-hop message has nothing underneath worth a second row -- the
        # collapsed line already IS the whole path, and an expander that opens
        # to restate it is a control that does nothing.
        #
        # I had this the other way round briefly. The reason was real: with the
        # marker conditional, Expand All still opened single-hop rows and they
        # could then not be clicked closed. The fix is to make expansion
        # multi-hop everywhere -- marker, click AND Expand All -- rather than to
        # give every row a marker. Both halves of that rule are enforced below
        # and in Set-AllExpanded; changing one without the other brings the
        # stuck-open row straight back.
        $canExpand = (@($hops).Count -gt 1)
        $marker = if (-not $canExpand) { '   ' } elseif ($open) { '[-]' } else { '[+]' }
        $when   = ConvertTo-DisplayTime -Utc $final.TimestampUtc
        $srcTxt = "$($final.Source)"
        if (@($hops).Count -gt 1) { $srcTxt = "$srcTxt  ($(@($hops).Count) hops)" }

        [void]$t.Rows.Add(
            $marker,
            $(if ($when) { $when.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }),
            $srcTxt, $final.Sender, $final.Recipient, $final.Subject, $final.Outcome, $final.Server,
            $(if ($final.SizeBytes) { '{0:N0}' -f $final.SizeBytes } else { '' }))
        [void]$flat.Add($final)
            [void]$traceOf.Add("$($g.Name)")

        if (-not $open) { continue }

        foreach ($h in $hops) {
            $hw = ConvertTo-DisplayTime -Utc $h.TimestampUtc
            [void]$t.Rows.Add(
                '',
                $(if ($hw) { '   ' + $hw.ToString('HH:mm:ss') } else { '' }),
                "   $($h.HopIndex). $($h.Source)", '', '', '', $h.Outcome, $h.Server, '')
            [void]$flat.Add($h)
            [void]$traceOf.Add("$($g.Name)")
        }

        # THE HANDOFF TO THE RECIPIENT'S MX -- the hop that exists in reality and
        # in none of the four consoles. Synthesised, and only when the message
        # actually left: a quarantined or rejected message never reached it.
        if ("$($final.Outcome)" -eq 'Delivered') {
            $dom = "$($final.Recipient)"
            if ($dom -match '@') { $dom = ($dom -split '@')[-1] }
            $mw = ConvertTo-DisplayTime -Utc $final.TimestampUtc
            [void]$t.Rows.Add(
                '',
                $(if ($mw) { '   ' + $mw.ToString('HH:mm:ss') } else { '' }),
                "   $(@($hops).Count + 1). $dom MX", '', '', '', 'Accepted', "mx.$dom", '')
            [void]$flat.Add($final)
            [void]$traceOf.Add("$($g.Name)")
        }
    }

    # Scroll position survives the rebuild -- assigning DataSource resets it,
    # which on a long trace loses the row the operator was reading.
    $firstVisible = 0
    if ($grid.FirstDisplayedScrollingRowIndex -gt 0) { $firstVisible = $grid.FirstDisplayedScrollingRowIndex }
    $grid.DataSource = $t
    if ($firstVisible -gt 0 -and $grid.Rows.Count -gt $firstVisible) {
        try { $grid.FirstDisplayedScrollingRowIndex = $firstVisible } catch { }
    }
    $grid.Tag = @($flat.ToArray())
    # WHICH MESSAGE EACH DISPLAYED LINE BELONGS TO. DataGridView has no second
    # Tag, and parsing it back out of the row text would break the moment a
    # subject contained the separator. A parallel list built in the same loop
    # cannot drift from the grid.
    $script:RowTrace = @($traceOf.ToArray())

    foreach ($gr in $grid.Rows) {
        $o = "$($gr.Cells['Outcome'].Value)"
        if ($o -eq 'Blocked' -or $o -eq 'Rejected' -or $o -eq 'Failed') {
            $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick
        } elseif ($o -eq 'Quarantined' -or $o -eq 'Deferred' -or $o -eq 'FilteredAsSpam') {
            $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGoldenrod
        }
        if ("$($gr.Cells[' '].Value)" -eq '' -and "$($gr.Cells['Sender'].Value)" -eq '') {
            $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
        }
    }

    $grid.Columns[' '].Width = 34
    return @($groups).Count
}

function Update-SuppressionPanel {
    <#
        DRIVEN BY THE SELECTED ROW, falling back to the typed address.

        IT USED TO KEY OFF THE SEARCH BOX ALONE, which broke the moment the box
        could hold *@gcearth.com: a wildcard is not an address, so it matched no
        suppression entry and the panel went blank on the very search most
        likely to contain one.

        Selecting a row now names the recipient, exactly as the quarantine panel
        does -- one rule for both panels rather than two that behave differently
        depending on what is in the search box.
    #>
    $st  = Get-SuppressionState -Rows $script:AllRows
    $row = Get-SelectedTraceRow

    $who = ''
    if ($row) { $who = "$($row.Recipient)".Trim() }
    if (-not $who) {
        $typed = "$($cboTo.Text)".Trim()
        if ($typed -and $typed -notmatch '[\*\?]') { $who = $typed }
    }

    $btnRmSpam.Enabled   = $false
    $btnRmBounce.Enabled = $false
    $lblSupDetail.Text   = ''

    if (-not $who) {
        # A WILDCARD SEARCH WITH NOTHING SELECTED still has something useful to
        # say: how many of the returned recipients are suppressed at all.
        $n = @($st.Keys).Count
        $lblSupWho.ForeColor = if ($n) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::Gray }
        $lblSupWho.Text = if ($n) { "$n recipient(s) in these results are suppressed." } else { 'Nobody in these results is suppressed.' }
        if ($n) {
            $lblSupDetail.Text = (@($st.Keys | ForEach-Object { $st[$_].Recipient }) -join ', ') +
                [Environment]::NewLine + 'Select one of their rows above to act on it.'
        }
        return
    }

    $key = $who.ToLowerInvariant()
    if (-not $st.ContainsKey($key)) {
        $lblSupWho.ForeColor = [System.Drawing.Color]::Gray
        $lblSupWho.Text = "$who is not on any SendGrid suppression list."
        $lblSupDetail.Text = 'Nothing is being silently dropped before it is sent. If a message is missing, it was sent and something downstream did not deliver it -- look at the Proofpoint and Exchange rows above.'
        return
    }

    $e = $st[$key]
    $lists = New-Object System.Collections.ArrayList
    if ($e.Spam)   { [void]$lists.Add('SPAM SUPPRESSION') }
    if ($e.Bounce) { [void]$lists.Add('PREVIOUS BOUNCE') }

    $lblSupWho.ForeColor = [System.Drawing.Color]::Firebrick
    $lblSupWho.Text = "$($e.Recipient) -- $(@($lists) -join '  and  ')"
    $lblSupDetail.Text = (@($e.Why) -join [Environment]::NewLine) + [Environment]::NewLine +
        'While an address is suppressed, SendGrid accepts the API call and sends nothing. No bounce, no NDR, no trace in Exchange.'

    $btnRmSpam.Enabled   = $e.Spam
    $btnRmBounce.Enabled = $e.Bounce
}

function Get-SourceSearchPlan {
    <#
        WHAT EACH SOURCE WOULD COST TO SEARCH, and why.

        A mockup that returns instantly teaches the wrong thing. The real value
        of a consolidated trace is partly that it hides how uneven these systems
        are -- one of them is a fan-out across every transport server and one is
        a single API call -- and an operator deciding whether to fund this should
        see that the cost is not evenly spread.

        THE FIGURES ARE PLAUSIBLE, NOT MEASURED, and the tool says so. They are
        modelled on what each query actually is:

          OnPrem      the slowest by far. Get-MessageTrackingLog has no
                      organisation-wide view, so it is one query PER TRANSPORT
                      SERVER, serially, and any server that is down contributes
                      a timeout rather than an answer.
          ExO         one Get-MessageTraceV2 call, but a throttled service.
          Proofpoint  a REST call per pod, and the primary pod is several
                      machines behind one name.
          SER         one REST call to a smaller estate.
          SendGrid    an API built for this. Fastest of the five.
    #>
    return @(
        [PSCustomObject]@{ Source = 'OnPrem';     Label = 'Querying on-premises transport logs, server by server'; Ms = 2400; Steps = 6 }
        [PSCustomObject]@{ Source = 'ExO';        Label = 'Querying Exchange Online message trace';                Ms = 1100; Steps = 4 }
        [PSCustomObject]@{ Source = 'Proofpoint'; Label = 'Querying the Proofpoint primary pod';                    Ms =  900; Steps = 3 }
        [PSCustomObject]@{ Source = 'SER';        Label = 'Querying the Proofpoint SER pod';                        Ms =  600; Steps = 2 }
        [PSCustomObject]@{ Source = 'SendGrid';   Label = 'Querying the SendGrid activity API';                     Ms =  400; Steps = 2 }
    )
}

function Invoke-SimulatedSearch {
    <#
        Walk the progress bar through the sources this search would touch.

        ONLY THE SOURCES ACTUALLY SELECTED. Picking SendGrid and waiting the
        on-premises fan-out would be a lie in the operator's favour -- it would
        make the tool look slower than it is and hide that the expensive source
        is the one they did not choose.

        DoEvents ON EVERY STEP, because this runs on the UI thread. Without it
        the bar does not paint, the window greys out, and the simulation looks
        like a hang -- which is the opposite of the point.

        CANCELLABLE, for the same reason the real thing has to be: a search you
        cannot stop is a search people avoid running.
    #>
    param([string]$Want)

    $plan = @(Get-SourceSearchPlan)
    if ($Want -ne 'All') { $plan = @($plan | Where-Object { $_.Source -eq $Want }) }
    if (@($plan).Count -eq 0) { return $true }

    $total = 0
    foreach ($s in $plan) { $total += $s.Steps }

    $pbSearch.Maximum = [Math]::Max(1, $total)
    $pbSearch.Value   = 0
    $pbSearch.Visible = $true
    $script:SearchCancelled = $false
    $btnStop.Visible  = $true
    $btnGo.Enabled    = $false

    try {
        foreach ($s in $plan) {
            $per = [int]([Math]::Max(1, $s.Ms / [Math]::Max(1, $s.Steps)))
            for ($i = 1; $i -le $s.Steps; $i++) {
                if ($script:SearchCancelled) {
                    $lblScenario.ForeColor = [System.Drawing.Color]::Goldenrod
                    $lblScenario.Text = 'Search stopped. Nothing was queried -- this is a mockup.'
                    return $false
                }
                $lblScenario.ForeColor = [System.Drawing.Color]::Gray
                $lblScenario.Text = "$($s.Label)  ($i of $($s.Steps))"
                $pbSearch.Value = [Math]::Min($pbSearch.Maximum, $pbSearch.Value + 1)
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds $per
            }
        }
        return $true
    } finally {
        $pbSearch.Visible = $false
        $btnStop.Visible  = $false
        $btnGo.Enabled    = $true
    }
}
function Invoke-Trace {
    $who = "$($cboTo.Text)".Trim()
    if (-not $who) {
        $txtOut.Text = 'Enter a recipient, or *@gcearth.com for everything. Known: ' + ($script:MockRecipients -join ', ')
        return
    }

    # THE SIMULATION RUNS BEFORE THE DATA IS TOUCHED, and what it costs depends
    # on which sources were selected -- see Get-SourceSearchPlan. Returns false
    # if the operator pressed Stop, in which case nothing is loaded or shown.
    if (-not (Invoke-SimulatedSearch -Want (Get-SelectedSource))) { return }

    $rows = @(Get-MockRows)
    $rows = @($rows | Where-Object { Test-TraceMatch -Value "$($_.Recipient)" -Pattern $who })

    $sf = "$($txtFrom.Text)".Trim()
    if ($sf) { $rows = @($rows | Where-Object { Test-TraceMatch -Value "$($_.Sender)" -Pattern $sf }) }
    $sj = "$($txtSubj.Text)".Trim()
    if ($sj) { $rows = @($rows | Where-Object { Test-TraceMatch -Value "$($_.Subject)" -Pattern $sj }) }

    $script:AllRows = @($rows)

    $keep = "$($cboOutcome.SelectedItem)"
    $cboOutcome.Items.Clear()
    [void]$cboOutcome.Items.Add('(any outcome)')
    foreach ($o in @($rows | ForEach-Object { $_.Outcome } | Sort-Object -Unique)) { [void]$cboOutcome.Items.Add("$o") }
    $cboOutcome.SelectedIndex = 0
    if ($keep -and $cboOutcome.Items.Contains($keep)) { $cboOutcome.SelectedItem = $keep }

    $shown = Update-Grid
    Update-SuppressionPanel
    $btnExport.Enabled = (@($script:AllRows).Count -gt 0)

    $people = @($rows | ForEach-Object { $_.Recipient } | Sort-Object -Unique)
    $lblScenario.ForeColor = [System.Drawing.Color]::Gray
    $lblScenario.Text = "$(@($rows).Count) row(s) across $(@($people).Count) recipient(s)"

    $L = New-Object System.Collections.ArrayList
    [void]$L.Add("MOCK DATA. $shown row(s) shown of $(@($script:AllRows).Count) matching '$who'.")
    [void]$L.Add('')

    if (@($script:AllRows).Count -eq 0) {
        [void]$L.Add("Nothing matched '$who'. The corpus knows about:")
        foreach ($r in $script:MockRecipients) { [void]$L.Add("    $r") }
        [void]$L.Add('')
        [void]$L.Add('Wildcards work: *@gcearth.com returns every message for the domain.')
    } else {
        if (@($people).Count -gt 1) {
            [void]$L.Add('BY RECIPIENT')
            foreach ($r in $people) {
                $mine = @($script:AllRows | Where-Object { $_.Recipient -eq $r })
                $bad  = @($mine | Where-Object { $_.Outcome -in @('Blocked', 'Rejected', 'Failed') }).Count
                $q    = @($mine | Where-Object { $_.Outcome -eq 'Quarantined' }).Count
                [void]$L.Add(("    {0,-26} {1,2} row(s){2}{3}" -f $r, @($mine).Count,
                    $(if ($q)   { "   $q quarantined" } else { '' }),
                    $(if ($bad) { "   $bad never arrived" } else { '' })))
            }
            [void]$L.Add('')
        }

        [void]$L.Add('BY SOURCE')
        foreach ($g in @($script:AllRows | Group-Object Source | Sort-Object Name)) {
            $bad = @($g.Group | Where-Object { $_.Outcome -in @('Blocked', 'Rejected', 'Failed') }).Count
            [void]$L.Add(("    {0,-12} {1,2} row(s){2}" -f $g.Name, $g.Count, $(if ($bad) { "   $bad did not deliver" } else { '' })))
        }
        [void]$L.Add('')

        $blocked = @($script:AllRows | Where-Object { $_.Outcome -eq 'Blocked' })
        if (@($blocked).Count -gt 0) {
            [void]$L.Add('NEVER SENT -- these have no counterpart anywhere in Exchange:')
            foreach ($b in $blocked) { [void]$L.Add("    $($b.Recipient)  --  $($b.Subject)") }
            [void]$L.Add('')
        }
        [void]$L.Add('Select a row to act on it. Double-click for the full path across sources.')
    }
    $txtOut.Text = ($L -join [Environment]::NewLine)
}

function Get-TraceGuidance {
    <#
        WHAT TO DO ABOUT THIS ROW, AND WHAT TO SAY TO THE PERSON WHO RAISED IT.

        A trace that reports an outcome and stops leaves the actual work
        undone. 'Rejected at the pod: sender reputation and URL verdict both
        failed' is precise, correct, and does not tell an operator whether to
        release it, escalate it, or close the ticket -- and it certainly does
        not tell them what to put in the reply.

        TWO AUDIENCES, TWO REGISTERS. The operator text names cmdlets, consoles
        and who to escalate to. The customer text is written to be pasted into
        a reply: no jargon, no verdict names, no blame.

        THE CUSTOMER TEXT DELIBERATELY NEVER PROMISES DELIVERY. Several of these
        outcomes end with the message staying blocked, and a reply that implies
        otherwise generates the follow-up ticket.
    #>
    param([Parameter(Mandatory)]$Row)

    $o    = "$($Row.Outcome)"
    $src  = "$($Row.Source)"
    $q    = "$($Row.Quarantine)"
    $det  = "$($Row.Detail)"
    $from = "$($Row.Sender)"

    $g = [PSCustomObject]@{ You = New-Object System.Collections.ArrayList; Them = '' }

    switch -Regex ($o) {

        '^Delivered$' {
            [void]$g.You.Add('Nothing is wrong with mail flow. The message reached the mailbox, so if they cannot see it the cause is client-side.')
            [void]$g.You.Add('Check, in this order: Junk Email, an inbox rule moving it, Focused/Other, a mobile client that has already filed it, and Deleted Items.')
            [void]$g.You.Add('Inbox rules are the usual answer and the one nobody checks -- Unified Lookup > Inbox rules on this mailbox.')
            $g.Them = "Good news -- that message did reach your mailbox at $($Row.TimestampUtc.ToLocalTime().ToString('HH:mm')) on $($Row.TimestampUtc.ToLocalTime().ToString('d MMMM')). It has not been blocked or lost. It is most likely in your Junk Email folder, or an Outlook rule has moved it somewhere else. Could you check Junk Email first, and search your whole mailbox rather than just the Inbox? If it is genuinely not there, let me know and I will look at your rules with you."
        }

        '^Rejected$' {
            [void]$g.You.Add('THIS IS THE FILTER WORKING. The message was refused at the edge and never entered the estate -- there is no copy to release, anywhere.')
            [void]$g.You.Add('Do NOT try to route around it. If the sender is genuinely legitimate, the fix is a Proofpoint submission so the verdict is corrected centrally, not an exception for one recipient.')
            if ($det -match '(?i)impersonat|lookalike|display.name') {
                [void]$g.You.Add('This one is display-name impersonation of an internal person. If the recipient ACTED on a similar message that did get through, that is an incident -- escalate to security now, do not just close the ticket.')
            } else {
                [void]$g.You.Add('Reputation plus URL verdict together is a strong signal. Treat a request to release this as a request to be phished.')
            }
            [void]$g.You.Add("If it is a false positive: submit the sample to Proofpoint and ask the recipient to have the sender resend AFTER the verdict changes -- resending now will be rejected identically.")
            $g.Them = "That message was stopped by our email security filter before it reached the network, because it matched several signs of a phishing attempt. It has not been delivered and I am not able to release it -- there is no copy held anywhere. If you were expecting something from $from and you are confident it is genuine, the safest thing is to contact them by phone using a number you already have, not by replying to any email. Let me know and I can ask our security team to review the block."
        }

        '^Blocked$' {
            if ($det -match '(?i)spam suppression') {
                [void]$g.You.Add('SENDGRID NEVER SENT THIS. The recipient is on the spam suppression list, so the API call was accepted and silently discarded. This is why Exchange shows nothing -- there is no message to find.')
                [void]$g.You.Add('The address is there because somebody at that address pressed "this is spam" on a previous message from us.')
                [void]$g.You.Add('BEFORE REMOVING IT: confirm the recipient actually wants this mail. Removing a spam suppression re-sends to somebody who complained, which is a deliverability and compliance decision -- get it in writing on the ticket.')
                [void]$g.You.Add('Then use Remove Spam Suppression below and ask the sending application to resend. Removing the suppression does not resend anything by itself.')
                $g.Them = "I have found the problem. Our sending platform had your address on a do-not-send list, because at some point a message of this type was reported as junk from your address. That is why nothing arrived and why there was no bounce message either -- it was never actually sent. Can you confirm you do want to receive these? Once you confirm, I will clear the block and ask for it to be sent again."
            }
            elseif ($det -match '(?i)bounce') {
                [void]$g.You.Add('SENDGRID NEVER SENT THIS. The address is on the bounce suppression list after a previous hard bounce, so the send was discarded before leaving. Nothing reaches Exchange, which is why an Exchange-only trace finds nothing.')
                [void]$g.You.Add('CHECK THE ADDRESS IS ACTUALLY VALID FIRST. A hard bounce usually means it genuinely did not exist at the time -- a typo in the application, a mailbox that had not been created yet, or a since-renamed user.')
                [void]$g.You.Add('If the mailbox exists now, use Remove Bounce below, then have the application resend. If you clear it without checking, the next send hard-bounces again and repeated hard bounces damage the sending reputation for everybody.')
                [void]$g.You.Add('If the address changed, fix it in the SOURCE application -- clearing the suppression on a dead address fixes nothing.')
                $g.Them = "I have traced this. Our sending platform stopped sending to your address because an earlier message bounced back as undeliverable, which puts the address on hold automatically. Your mailbox looks fine now, so I will clear that hold and have the message sent again. You should not need to do anything -- but if you have recently had your email address changed, please let me know, because the system may still be using the old one."
            }
            else {
                [void]$g.You.Add('Blocked before sending. Read the detail line above for which list caught it, then use the matching Remove button.')
                $g.Them = 'That message was stopped by our sending platform before it went out. I am looking into why and will come back to you.'
            }
        }

        '^Quarantined$' {
            if ($q -eq 'Malware') {
                [void]$g.You.Add('DO NOT RELEASE. Malware quarantine means something matched a signature -- the release button is greyed for that reason, not because of a permissions problem.')
                [void]$g.You.Add('Releasing known-malicious mail to a mailbox is a security decision. If the recipient insists they need it, that goes to the security team with the message ID above, not through a helpdesk release.')
                [void]$g.You.Add('If the sender is a known partner, the file almost certainly needs to come another way -- a file share or a portal, not email.')
                [void]$g.You.Add('WORTH CHECKING: whether the same sender has landed anything else in this mailbox recently. One malware attempt is noise; a pattern against one person is targeting.')
                $g.Them = "That message was blocked because the attachment matched a known malware signature. I am not able to release it -- for anything flagged as malware that decision sits with our security team rather than with me. If you are expecting a genuine file from $from, the safest route is to ask them to send it another way, for example a shared link, and I can help set that up. If you would like the block reviewed I can raise it with security for you."
            }
            elseif ($src -eq 'ExO') {
                [void]$g.You.Add('EOP spam quarantine. Releasable -- use Release from EOP below.')
                [void]$g.You.Add('RELEASE ALONE TEACHES THE FILTER NOTHING. It delivers this one message and the next one from the same sender is quarantined identically. Submit it as a false positive at the same time, or you will see this ticket again.')
                [void]$g.You.Add('Check whether it should have been released to ALL recipients or just this one -- they are different actions and the sender usually mailed several people.')
                [void]$g.You.Add('If this sender is repeatedly quarantined and is genuinely wanted, the durable fix is a tenant allow entry, not repeated releases.')
                $g.Them = "That message was held in our spam quarantine rather than being delivered. I have looked at it and it appears legitimate, so I will release it to your Inbox now -- it should arrive within a few minutes. I will also report it as wrongly filtered so that future messages from $from come through normally. If anything else from them goes missing in the next few days, tell me and I will set up a permanent exception."
            }
            else {
                [void]$g.You.Add('Proofpoint pod quarantine. Releasable -- use Release from Proofpoint below.')
                [void]$g.You.Add('A POD RELEASE RE-INJECTS THE MESSAGE DOWNSTREAM, so it then passes through EOP as well and can be quarantined a second time. Re-run this trace after releasing; if it stops at EOP, release it there too.')
                [void]$g.You.Add('Release is per-pod. Releasing on the primary pod does nothing for a copy held on SER.')
                [void]$g.You.Add('Bulk mail quarantined here is usually correctly classified. If the recipient wants it routinely, a safe-sender entry is better than releasing one at a time.')
                $g.Them = "That message was held by our email filter as bulk or promotional mail rather than being blocked outright. I can release it to you now. If you would like messages from $from to come through normally in future rather than being held each time, tell me and I will add them to the allowed list."
            }
        }

        '^FilteredAsSpam$' {
            [void]$g.You.Add('DELIVERED, not blocked -- it is in the Junk Email folder. Nothing to release and nothing to fix in mail flow.')
            [void]$g.You.Add('Ask them to mark it Not Junk from within Outlook, which trains the filter for their mailbox. Dragging it to the Inbox does not.')
            [void]$g.You.Add('If it keeps happening from an internal sender, that is worth looking at properly -- internal mail being junked usually points at a broken SPF or DKIM record on the sending domain.')
            $g.Them = "That message did arrive -- it was delivered to your Junk Email folder rather than your Inbox. Please have a look there. When you find it, use the Not Junk option in Outlook rather than dragging it out, as that teaches the filter to leave future messages from $from alone. If it keeps happening let me know, because we can fix it at our end."
        }

        '^Deferred$' {
            [void]$g.You.Add('IN PROGRESS, NOT FAILED. The receiving side asked the sender to try later -- greylisting or throttling. Retries continue automatically.')
            [void]$g.You.Add('DO NOTHING YET. Resending now usually makes it worse: a second copy gets deferred too and both eventually arrive.')
            [void]$g.You.Add('SendGrid retries for up to 72 hours. Re-run this trace in an hour; if it is still deferred after several hours the receiving domain may be having a genuine problem.')
            $g.Them = "That message is still on its way -- the receiving mail server asked us to try again shortly, which is a normal anti-spam measure and usually clears itself within an hour. Nothing has failed and nothing is lost. Please do not resend it for now, as that tends to slow things down further. I will keep an eye on it and let you know if it has not gone through by the end of the day."
        }

        '^Failed$' {
            [void]$g.You.Add('PERMANENT FAILURE. It will not arrive and it will not retry.')
            if ($det -match '(?i)relay access denied|5\.4\.1') {
                [void]$g.You.Add('550 5.4.1 Relay access denied from an internal device means the sending IP is not in scope on any receive connector. This is a configuration fix, not a mail problem.')
                [void]$g.You.Add("Find the device's IP and add it to the anonymous relay connector, then have them resend. Scanners and multifunction printers are the usual culprits, especially after one is replaced or re-addressed.")
                [void]$g.You.Add('Check whether other devices on the same subnet are also failing -- a replaced network segment breaks them all at once.')
                $g.Them = "That did not go through. The device sending it -- it looks like a scanner or printer -- is not currently allowed to send email through our mail server, which is a configuration setting on our side rather than anything you did wrong. I am fixing that now. Once it is done I will let you know and you can scan the document again."
            } else {
                [void]$g.You.Add('Read the exact status code in the detail line above -- it names the cause. A 5.x.x is permanent; the sender should have received a non-delivery report.')
                $g.Them = "Unfortunately that message could not be delivered and it will not retry automatically. I am looking at exactly why and will come back to you shortly with what needs to happen."
            }
        }

        default {
            [void]$g.You.Add('No specific guidance for this outcome yet. Read the detail line above and the full path across sources below.')
            $g.Them = ''
        }
    }
    return $g
}
function Show-RowDetail {
    param([Parameter(Mandatory)]$Row)
    $L = New-Object System.Collections.ArrayList
    [void]$L.Add('MOCK ROW -- invented, not retrieved.')
    [void]$L.Add('')
    [void]$L.Add("Source      : $($Row.Source)")
    [void]$L.Add("Time (UTC)  : $($Row.TimestampUtc.ToString('yyyy-MM-dd HH:mm:ss'))")
    [void]$L.Add("Sender      : $($Row.Sender)")
    [void]$L.Add("Recipient   : $($Row.Recipient)")
    [void]$L.Add("Subject     : $($Row.Subject)")
    [void]$L.Add("Outcome     : $($Row.Outcome)")
    if ($Row.Server) { [void]$L.Add("Server      : $($Row.Server)") }
    [void]$L.Add("Size        : $('{0:N0}' -f $Row.SizeBytes) bytes")
    [void]$L.Add("Message ID  : $($Row.MessageId)")
    [void]$L.Add('')
    [void]$L.Add('What the source reported:')
    [void]$L.Add("    $($Row.Detail)")
    [void]$L.Add('')

    # THE TWO SECTIONS THAT MAKE THIS A TOOL RATHER THAN A LOG VIEWER.
    $g = Get-TraceGuidance -Row $Row
    [void]$L.Add('------------------------------------------------------------------------')
    [void]$L.Add('WHAT YOU SHOULD DO')
    foreach ($line in @($g.You)) {
        # Wrapped by hand: the pane does not word-wrap, because a wrapped
        # message ID or DN is unreadable.
        $s = "$line"
        while ($s.Length -gt 86) {
            $cut = $s.LastIndexOf(' ', 86)
            if ($cut -lt 40) { $cut = 86 }
            [void]$L.Add("    $($s.Substring(0, $cut))")
            $s = $s.Substring($cut).TrimStart()
        }
        [void]$L.Add("    $s")
        [void]$L.Add('')
    }

    if ($g.Them) {
        [void]$L.Add('------------------------------------------------------------------------')
        [void]$L.Add('WHAT TO TELL THEM  --  written to be pasted into a reply')
        [void]$L.Add('')
        $s = "$($g.Them)"
        while ($s.Length -gt 86) {
            $cut = $s.LastIndexOf(' ', 86)
            if ($cut -lt 40) { $cut = 86 }
            [void]$L.Add("    $($s.Substring(0, $cut))")
            $s = $s.Substring($cut).TrimStart()
        }
        [void]$L.Add("    $s")
        [void]$L.Add('')
    }
    [void]$L.Add('------------------------------------------------------------------------')
    [void]$L.Add('')

    # THE CROSS-SOURCE HOP LIST -- the single most useful thing a unified trace
    # would do, and the reason it is in the mockup at all. Correlated on subject
    # here because the mock sources do not share a Message ID; a real build
    # would correlate on Message ID, which is why the field is carried.
    $sibs = @($script:AllRows | Where-Object { "$($_.Subject)" -eq "$($Row.Subject)" } | Sort-Object TimestampUtc)
    if (@($sibs).Count -gt 1) {
        [void]$L.Add("PATH ACROSS SOURCES ($(@($sibs).Count) hops)")
        foreach ($s in $sibs) {
            [void]$L.Add(("    {0}  {1,-11} {2,-14} {3}" -f `
                $s.TimestampUtc.ToString('HH:mm:ss'), $s.Source, $s.Outcome, $s.Server))
        }
        [void]$L.Add('')
        [void]$L.Add('In a real build these hops correlate on Message ID. Following this today means four consoles and four different search syntaxes.')
    } else {
        [void]$L.Add('Only one source saw this message.')
        if ("$($Row.Outcome)" -eq 'Blocked') {
            [void]$L.Add('It was blocked before sending, so no other system ever saw it -- there is nothing else to find.')
        }
    }
    $txtOut.Text = ($L -join [Environment]::NewLine)
}

function Get-RowTrace {
    param([int]$Index)
    $m = @($script:RowTrace)
    if ($Index -lt 0 -or $Index -ge @($m).Count) { return '' }
    return "$($m[$Index])"
}

function Toggle-TraceRow {
    <#
        Open or close the message on the clicked line.

        CLICKING ANYWHERE ON A PARENT ROW TOGGLES IT, not just the marker.
        A 34px target is a fiddly thing to hit, and the alternative -- clicking
        the row and nothing happening -- reads as the tool being unresponsive.
        Child rows do not toggle, because there is nothing under them.
    #>
    param([int]$Index)

    $id = Get-RowTrace -Index $Index
    if (-not $id) { return }

    # Only a PARENT line toggles. A child line has an empty marker cell.
    $marker = "$($grid.Rows[$Index].Cells[' '].Value)".Trim()
    if (-not $marker) { return }

    if ($script:Expanded.ContainsKey($id)) { $script:Expanded.Remove($id) }
    else { $script:Expanded[$id] = $true }

    # THE SELECTION IS PUT BACK ON THE SAME MESSAGE after the repaint. Without
    # this the grid jumps to row 0 on every expand, which loses the row the
    # action panels are pointed at -- and the release button silently starts
    # meaning a different message.
    [void](Update-Grid)
    for ($i = 0; $i -lt @($script:RowTrace).Count; $i++) {
        if ((Get-RowTrace -Index $i) -eq $id -and "$($grid.Rows[$i].Cells[' '].Value)".Trim()) {
            $grid.ClearSelection()
            $grid.Rows[$i].Selected = $true
            $grid.CurrentCell = $grid.Rows[$i].Cells['Subject']
            break
        }
    }
}

function Set-AllExpanded {
    param([bool]$Open)
    $script:Expanded = @{}
    if ($Open) {
        # MULTI-HOP ONLY. Opening a one-hop message would render a child row
        # under a line with no marker to close it -- see Update-Grid.
        foreach ($g in @($script:AllRows | Group-Object TraceId)) {
            if (@($g.Group).Count -le 1) { continue }
            $script:Expanded["$($g.Name)"] = $true
        }
    }
    [void](Update-Grid)
}
function Get-SelectedTraceRow {
    if (-not $grid.CurrentRow) { return $null }
    $rows = @($grid.Tag)
    $i = $grid.CurrentRow.Index
    if ($i -lt 0 -or $i -ge @($rows).Count) { return $null }
    return $rows[$i]
}

function Update-QuarantinePanel {
    <#
        Repaint the release panel for whatever row is selected.

        BOTH BUTTONS ARE ALWAYS VISIBLE and only their enabled state moves. A
        button that appears and disappears with the selection reads as the tool
        being unsure what it can do; one that greys with a reason on hover reads
        as the tool knowing exactly what it cannot do and why.
    #>
    $row = Get-SelectedTraceRow
    $q   = Get-QuarantineAction -Row $row

    $btnRelEop.Enabled  = $false
    $btnRelPfpt.Enabled = $false
    $quarTips.SetToolTip($btnRelEop,  'Select a quarantined row above.')
    $quarTips.SetToolTip($btnRelPfpt, 'Select a quarantined row above.')

    if (-not $q.Quarantined) {
        $lblQuarWho.ForeColor = [System.Drawing.Color]::Gray
        $lblQuarWho.Text      = 'Nothing quarantined selected.'
        $lblQuarDetail.Text   = $q.Reason + ' Pick a row whose Outcome is Quarantined -- this scenario has four, two held by EOP and two by Proofpoint.'
        return
    }

    $kind = if ($q.Kind) { $q.Kind.ToUpperInvariant() } else { 'UNCLASSIFIED' }
    $lblQuarWho.ForeColor = if ($q.CanRelease) { [System.Drawing.Color]::DarkGoldenrod } else { [System.Drawing.Color]::Firebrick }
    $lblQuarWho.Text      = "$($q.Source) $kind quarantine  --  $($row.Subject)"
    $lblQuarDetail.Text   = $q.Reason

    if (-not $q.CanRelease) {
        # GREYED, WITH THE REASON ON BOTH BUTTONS. Whichever one the operator
        # reaches for, the hover explains it.
        $quarTips.SetToolTip($btnRelEop,  $q.Reason)
        $quarTips.SetToolTip($btnRelPfpt, $q.Reason)
        return
    }

    # ONLY THE BUTTON FOR THE SYSTEM HOLDING IT. Releasing an EOP-quarantined
    # message from Proofpoint is not a thing, and offering both would invite
    # exactly that click.
    if ($q.Source -eq 'ExO') {
        $btnRelEop.Enabled = $true
        $quarTips.SetToolTip($btnRelEop, "Release '$($row.Subject)' from the EOP spam quarantine.")
        $quarTips.SetToolTip($btnRelPfpt, 'This message is held by EOP, not Proofpoint.')
    } elseif ($q.Source -eq 'Proofpoint' -or $q.Source -eq 'SER') {
        $btnRelPfpt.Enabled = $true
        $quarTips.SetToolTip($btnRelPfpt, "Release '$($row.Subject)' from the Proofpoint pod quarantine.")
        $quarTips.SetToolTip($btnRelEop, 'This message is held by Proofpoint, not EOP.')
    }
}

function Invoke-FakeRelease {
    <#
        Confirm, then say plainly that nothing happened -- same contract as the
        suppression buttons. See Invoke-FakeRemoval.
    #>
    param([Parameter(Mandatory)][ValidateSet('EOP', 'Proofpoint')][string]$System)

    $row = Get-SelectedTraceRow
    $q   = Get-QuarantineAction -Row $row
    if (-not $q.CanRelease) { return }

    $ans = Show-ToolDialog(
        "Release this message from $System quarantine?" + [Environment]::NewLine + [Environment]::NewLine +
        "  $($row.Subject)" + [Environment]::NewLine +
        "  from $($row.Sender)" + [Environment]::NewLine +
        "  to   $($row.Recipient)",
        "Release from $System",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
        $txtOut.Text = 'Cancelled. Nothing would have happened anyway -- this is a mockup.'
        return
    }

    $call = if ($System -eq 'EOP') {
        "Release-QuarantineMessage -Identity <id> -User '$($row.Recipient)'"
    } else {
        "POST /v2/quarantine/release   { folder: 'Quarantine', messageId: '$($row.MessageId)' }"
    }

    $L = New-Object System.Collections.ArrayList
    [void]$L.Add("NOTHING HAPPENED. This is a mockup with no connection to $System.")
    [void]$L.Add('')
    [void]$L.Add('A real build would have called:')
    [void]$L.Add("    $call")
    [void]$L.Add('')
    [void]$L.Add('It would then re-read the quarantine and show the message as released, or report why the release was refused.')
    [void]$L.Add('')
    [void]$L.Add('WORTH DECIDING BEFORE THIS IS BUILT FOR REAL:')
    if ($System -eq 'EOP') {
        [void]$L.Add('  - Release delivers to the mailbox and does NOT report the verdict as wrong.')
        [void]$L.Add('    Without a submission the filter learns nothing and the next one is quarantined')
        [void]$L.Add('    too. The real tool should offer "release and report as false positive".')
        [void]$L.Add('  - Releasing to ALL recipients versus one is a different call and a different')
        [void]$L.Add('    parameter. The button must say which it is doing.')
    } else {
        [void]$L.Add('  - A pod release re-injects the message downstream, so it is then subject to')
        [void]$L.Add('    EOP as well. It can be quarantined a second time, which looks like the')
        [void]$L.Add('    release failed. The real tool should follow it and say so.')
        [void]$L.Add('  - Proofpoint release is per-pod. Releasing on the primary pod does nothing')
        [void]$L.Add('    for a copy held on SER, and the tool should name which pod it acted on.')
    }
    [void]$L.Add('')
    [void]$L.Add('The grid above is unchanged, deliberately -- pretending the row released is how a')
    [void]$L.Add('mockup gets mistaken for a working tool.')
    $txtOut.Text = ($L -join [Environment]::NewLine)

    Write-CompletedLog -Action 'MockReleaseQuarantine' -Target (Get-LogTarget $row.Recipient 'Unified Trace mock') -Result 'Success' `
        -Detail "MOCKUP -- no system contacted. $System, $($q.Kind) quarantine, subject '$($row.Subject)'. Would have called: $call" `
        -Environment 'Mock'
}
function Invoke-FakeRemoval {
    <#
        The confirmation, and then honesty.

        IT WOULD HAVE BEEN EASY TO POP A SUCCESS MESSAGE and let the demo look
        complete. That is precisely the wrong thing to build: somebody would
        click it during a real incident, believe the suppression was cleared,
        and wait for mail that is still being dropped. It confirms, then says
        plainly that nothing happened and shows the call it would have made.
    #>
    param([Parameter(Mandatory)][ValidateSet('Spam', 'Bounce')][string]$Kind)

    $who = "$($cboTo.Text)".Trim()
    if (-not $who) { return }

    $label = if ($Kind -eq 'Spam') { 'Spam Suppression' } else { 'Bounce' }
    $ans = Show-ToolDialog(
        "Remove $label for $who?",
        "Remove $label",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
        $txtOut.Text = 'Cancelled. Nothing would have happened anyway -- this is a mockup.'
        return
    }

    $endpoint = if ($Kind -eq 'Spam') { "DELETE /v3/suppression/spam_reports/$who" }
                else                  { "DELETE /v3/suppression/bounces/$who" }

    $L = New-Object System.Collections.ArrayList
    [void]$L.Add('NOTHING HAPPENED. This is a mockup and there is no SendGrid connection.')
    [void]$L.Add('')
    [void]$L.Add("A real build would have called:")
    [void]$L.Add("    $endpoint")
    [void]$L.Add('    Authorization: Bearer <SendGrid API key from the secret store>')
    [void]$L.Add('')
    [void]$L.Add("It would then re-read the suppression lists and show $who as clear, or report why the removal was refused.")
    [void]$L.Add('')
    [void]$L.Add('WORTH DECIDING BEFORE THIS IS BUILT FOR REAL:')
    [void]$L.Add('  - Removing a bounce suppression on an address that genuinely does not exist')
    [void]$L.Add('    means the next send hard-bounces again and hurts sending reputation. The real')
    [void]$L.Add('    tool should confirm the address resolves before offering the button.')
    [void]$L.Add('  - Removing a SPAM suppression re-sends to somebody who pressed "this is spam".')
    [void]$L.Add('    That is a deliverability and a compliance decision, not a helpdesk one, and it')
    [void]$L.Add('    should be logged with who approved it.')
    [void]$L.Add('')
    [void]$L.Add('The grid above is unchanged, deliberately -- pretending the row cleared is how a')
    [void]$L.Add('mockup gets mistaken for a working tool.')
    $txtOut.Text = ($L -join [Environment]::NewLine)

    # LOGGED AS A MOCK ACTION. It changed nothing, but "who clicked the button
    # that does not work yet" is exactly the evidence that says whether this is
    # worth building.
    Write-CompletedLog -Action 'MockRemoveSuppression' -Target (Get-LogTarget $who 'Unified Trace mock') -Result 'Success' `
        -Detail "MOCKUP -- no system contacted. Would have called: $endpoint" -Environment 'Mock'
}

function Export-Rows {
    if (@($script:AllRows).Count -eq 0) { return }
    $rows = @($grid.Tag)
    if (@($rows).Count -eq 0) { $rows = @($script:AllRows) }

    $out = @($rows | Select-Object `
        @{ N = 'MOCK_DATA'; E = { 'YES -- INVENTED, NOT RETRIEVED' } },
        @{ N = 'TimeUtc';   E = { $_.TimestampUtc.ToString('yyyy-MM-dd HH:mm:ss') } },
        @{ N = 'TimeLocal'; E = { $_.TimestampUtc.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss') } },
        Source, Sender, Recipient, Subject, Outcome, Server, SizeBytes, MessageId, Detail)

    # THE FILENAME CARRIES THE WARNING TOO. An exported CSV outlives the window
    # it came from and gets forwarded without the banner.
    $p = Write-ReportCsv -Directory (Get-ToolkitReportDirectory) -Prefix 'MOCKUP_UnifiedTrace_NOT_REAL_DATA' -Rows $out
    $txtOut.Text = "Exported $(@($out).Count) MOCK row(s) to:" + [Environment]::NewLine + $p +
        [Environment]::NewLine + [Environment]::NewLine +
        'Every row in that file is invented. The filename and the first column both say so.'
}
#endregion

$btnGo.Add_Click({ try { Invoke-Trace } catch { Show-HandlerError -ErrorRecord $_ -Context 'Trace' } })
$btnStop.Add_Click({ try { $script:SearchCancelled = $true } catch { } })
$btnExport.Add_Click({ try { Export-Rows } catch { Show-HandlerError -ErrorRecord $_ -Context 'Export' } })
$btnRmSpam.Add_Click({ try { Invoke-FakeRemoval -Kind 'Spam' } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove spam suppression' } })
$btnRmBounce.Add_Click({ try { Invoke-FakeRemoval -Kind 'Bounce' } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove bounce' } })
$cboTz.Add_SelectedIndexChanged({ try { if (@($script:AllRows).Count) { [void](Update-Grid) } } catch { } })
$cboOutcome.Add_SelectedIndexChanged({ try { if (@($script:AllRows).Count) { [void](Update-Grid) } } catch { } })
$cboTo.Add_TextChanged({ try { Update-SuppressionPanel } catch { } })
foreach ($rb in $script:Radios) {
    $rb.Add_CheckedChanged({
        try { if ($this.Checked -and @($script:AllRows).Count) { [void](Update-Grid) } }
        catch { Show-HandlerError -ErrorRecord $_ -Context 'Source filter' }
    })
}
$btnRelEop.Add_Click({ try { Invoke-FakeRelease -System 'EOP' } catch { Show-HandlerError -ErrorRecord $_ -Context 'Release from EOP' } })
$btnRelPfpt.Add_Click({ try { Invoke-FakeRelease -System 'Proofpoint' } catch { Show-HandlerError -ErrorRecord $_ -Context 'Release from Proofpoint' } })
$grid.Add_SelectionChanged({
    try { Update-QuarantinePanel; Update-SuppressionPanel } catch { }
})
$btnExpandAll.Add_Click({ try { Set-AllExpanded -Open $true } catch { Show-HandlerError -ErrorRecord $_ -Context 'Expand all' } })
$btnCollapseAll.Add_Click({ try { Set-AllExpanded -Open $false } catch { Show-HandlerError -ErrorRecord $_ -Context 'Collapse all' } })
# THE MARKER COLUMN TOGGLES. NOTHING ELSE DOES.
#
# This fired on a click anywhere in the row, so selecting a message to read it
# expanded or collapsed it instead -- and because expanding rebuilds the grid,
# the click also threw away the selection. Selecting is the common action;
# toggling is the rare one.
$grid.Add_CellClick({
    param($sender, $e)
    try {
        if ($e.RowIndex -lt 0) { return }
        if ($e.ColumnIndex -ne 0) { return }
        Toggle-TraceRow -Index $e.RowIndex
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Expand' }
})
$grid.Add_CellDoubleClick({
    param($sender, $e)
    try {
        if ($e.RowIndex -lt 0) { return }
        $rows = @($grid.Tag)
        if ($e.RowIndex -ge @($rows).Count) { return }
        Show-RowDetail -Row $rows[$e.RowIndex]
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Row detail' }
})
$form.AcceptButton = $btnGo

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Unified Trace' | Out-Null

$form.Add_Shown({
    try {
        $script:DataFile = Initialize-MockData

        $seed = "$SmtpAddress".Trim()
        # STARTS EMPTY, and the wildcard is the first item in the dropdown
        # rather than the default. A box that arrives pre-filled invites Trace
        # to be pressed before anybody has read what the tool does, and the
        # first thing they see is 60 rows of somebody else's mail.
        $cboTo.Text = "$seed"

        $all = @(Get-MockRows)
        $who = @($all | ForEach-Object { $_.Recipient } | Sort-Object -Unique)

        $L = New-Object System.Collections.ArrayList
        [void]$L.Add('UNIFIED TRACE -- MOCKUP')
        [void]$L.Add('')
        [void]$L.Add('Nothing here is real. No system is contacted, because there is no API access to Proofpoint or SendGrid -- which is the point: this shows what a consolidated trace WOULD look like, so it can be judged before anybody buys anything.')
        [void]$L.Add('')
        [void]$L.Add("$(@($all).Count) message events across $(@($who).Count) recipients, in one file:")
        [void]$L.Add("    $script:DataFile")
        [void]$L.Add('Plain CSV and safe to hand-edit. Delete it and it is written again on the next launch.')
        if (@($script:Legacy).Count -gt 0) {
            [void]$L.Add('')
            [void]$L.Add("NOTE: $(@($script:Legacy).Count) old scenario-*.csv file(s) are still in that folder and are NO LONGER READ -- the rotating scenarios were replaced by the single corpus above. They are left alone in case you edited them.")
        }
        [void]$L.Add('')
        [void]$L.Add('SEARCH')
        [void]$L.Add('    *@gcearth.com     everything, all five sources, one timeline')
        [void]$L.Add('    bob               substring -- finds bob.smith without typing the rest')
        [void]$L.Add('    *@partner*        wildcards work on Sender and Subject too')
        [void]$L.Add('')
        foreach ($r in $who) {
            $mine = @($all | Where-Object { $_.Recipient -eq $r })
            [void]$L.Add(("    {0,-26} {1,2} events" -f $r, @($mine).Count))
        }
        [void]$L.Add('')
        [void]$L.Add('bob.smith is held by a SPAM suppression, dana.patel by a BOUNCE -- one each, so either button can be shown on its own. Neither leaves any trace in Exchange, which is exactly the gap this would close.')
        $txtOut.Text = ($L -join [Environment]::NewLine)

        Update-SuppressionPanel
        Update-QuarantinePanel
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Unified Trace' }
})

[void]$form.ShowDialog()
