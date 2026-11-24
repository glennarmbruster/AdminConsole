#Requires -Version 5.1
<#
    NewUser.ps1
    Enables on-prem Remote Mailboxes (with archive) in Exchange Hybrid, with
    automatic directory detection.

    Author   : Glenn Armbruster
    Revised  : 2026-10-09  (Retry Queue tab, like Term Requests)
#>
# TOOLKIT-HELP-BEGIN
# Creates the mailbox for a new starter in the hybrid setup.
#
# It enables a remote mailbox on-premises -- the on-premises object that tells
# the directory this person's mailbox lives in the cloud -- along with their
# archive, and works out which directory the account is in on its own.
#
# The mailbox itself appears in the cloud after the next directory sync, not
# immediately. That gap is normal and is the usual reason a brand-new account
# looks broken for a while.
#
# THE RETRY QUEUE (second tab, 2026-10-09). Anything that does not finish
# cleanly lands there instead of being lost: an account that does not exist
# yet (the commonest -- the AD account is created after the ticket), an
# enable that failed, a mailbox whose archive did not enable, or one that
# could not be verified. The row says which step failed and why. Fix the
# cause (usually: wait for the account to be created), tick the rows, Retry
# Selected. A retry resumes -- a mailbox that exists only gets its archive --
# and a row disappears when it fully succeeds. The queue is a CSV
# (newuser-retry.csv) in the same folder as the Term Requests queue, so it
# survives closing the tool.
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

#region Configuration --------------------------------------------------------
$RemoteRoutingDomain = "$(((Get-TenantConfig).OnMicrosoftDomain) -replace '\.onmicrosoft\.com$','.mail.onmicrosoft.com')"
# The directory this tool pins a controller for: the default one, or the
# first that actually names controllers. A tenant that pins none gets an empty
# list and the tool lets AD locate a server -- which is the preferred answer.
$script:PinDirectory = $null
foreach ($d in @(Get-TenantDirectory)) {
    if (@($d.DomainControllers).Count -gt 0) {
        if (-not $script:PinDirectory -or $d.IsDefault) { $script:PinDirectory = $d }
    }
}
$script:PinDirectoryId = if ($script:PinDirectory) { "$($script:PinDirectory.Id)" } else { '' }

# Licence lookups within one batch share a cache.
#
# Not for the repeated ACCOUNT -- every account in a batch is different, so
# there is nothing to hit there. It is for the failure path: when an account is
# not directly in a licence group, Get-UserLicenseState resolves each configured
# group to a DN before testing for nested membership, and those groups are the
# same for every account in the run. Thirty unlicensed starters would otherwise
# re-resolve the same two groups thirty times.
$script:LicenceCache = @{}

# A REAL controller, not a domain name.
#
# This was Get-TenantDirectoryServer(), which deliberately returns the DOMAIN
# when a directory names several controllers (ARCHITECTURE 14f -- do not pin).
# That is right for the AD module's -Server, and WRONG for BOTH of Exchange's
# server parameters: -PreferredServer and -DomainController each validate the
# argument and reject a domain with "Active Directory error 0x52 ...
# suitability of server". The second half of that sentence was learned the
# hard way on 2026-08-12 -- see the resolve loop.
$DefaultFranchiseDC = ''
if ($script:PinDirectoryId) {
    $DefaultFranchiseDC = @(Get-TenantDomainController -DirectoryId $script:PinDirectoryId)[0]
}
#endregion

#region Retry queue ----------------------------------------------------------
# Same folder and the same shape as Disable-User's queue, one CSV per tool.
$script:RetryCsvPath = Join-Path (Get-RetryQueuePath) 'newuser-retry.csv'
$script:RetryColumns = @('Account', 'Environment', 'FirstFailed', 'LastAttempt', 'Attempts', 'FailedStep', 'StepsDone', 'LastError')

function Import-NuQueue {
    if (-not (Test-Path -LiteralPath $script:RetryCsvPath)) { return @() }
    try { $rows = @(Import-Csv -LiteralPath $script:RetryCsvPath -ErrorAction Stop) }
    catch { Write-Warning "Could not read the retry queue: $($_.Exception.Message)"; return @() }
    $out = New-Object System.Collections.ArrayList
    foreach ($row in $rows) {
        if (-not "$($row.Account)".Trim()) { continue }
        $n = 0; [void][int]::TryParse("$($row.Attempts)", [ref]$n)
        [void]$out.Add([PSCustomObject]@{
            Account = "$($row.Account)".Trim(); Environment = "$($row.Environment)".Trim()
            FirstFailed = "$($row.FirstFailed)"; LastAttempt = "$($row.LastAttempt)"; Attempts = $n
            FailedStep = "$($row.FailedStep)"; StepsDone = @("$($row.StepsDone)".Split(';') | Where-Object { $_ })
            LastError = "$($row.LastError)"
        })
    }
    return @($out)
}

function Export-NuQueue {
    param($Records)
    $q = Test-RetryQueueConfigured
    if (-not $q.Exists) {
        if ($q.MayBeCreated) { New-Item -Path $q.Path -ItemType Directory -Force | Out-Null }
        else { throw "Retry queue folder '$($q.Path)' does not exist. It is set in ToolManager -> Settings and is not created automatically. $(@($Records).Count) row(s) were NOT saved." }
    }
    if (@($Records).Count -eq 0) { ($script:RetryColumns -join ',') | Set-Content -LiteralPath $script:RetryCsvPath -Encoding UTF8; return }
    $rows = foreach ($r in $Records) {
        [PSCustomObject]@{
            Account = $r.Account; Environment = $r.Environment; FirstFailed = $r.FirstFailed; LastAttempt = $r.LastAttempt
            Attempts = $r.Attempts; FailedStep = $r.FailedStep; StepsDone = (@($r.StepsDone) -join ';'); LastError = $r.LastError
        }
    }
    $rows | Export-Csv -LiteralPath $script:RetryCsvPath -NoTypeInformation -Encoding UTF8
}

function Update-NuQueue {
    <#
        One account's row: replaced on a failure, removed on success. Attempts
        and FirstFailed carry over from the row being replaced, so the queue
        says how long something has been stuck.
    #>
    param(
        [Parameter(Mandatory)][string]$Account,
        [string]$Environment,
        [Parameter(Mandatory)][ValidateSet('Success', 'Partial', 'Failed')][string]$Outcome,
        [string]$FailedStep, [string]$LastError, [string[]]$StepsDone
    )
    $queue = New-Object System.Collections.ArrayList
    $prior = $null
    foreach ($q in (Import-NuQueue)) {
        if ($q.Account -ieq $Account) { $prior = $q; continue }
        [void]$queue.Add($q)
    }
    if ($Outcome -ne 'Success') {
        $now = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        [void]$queue.Add([PSCustomObject]@{
            Account = $Account; Environment = $(if ($Environment) { $Environment } elseif ($prior) { $prior.Environment } else { '' })
            FirstFailed = $(if ($prior -and $prior.FirstFailed) { $prior.FirstFailed } else { $now }); LastAttempt = $now
            Attempts = $(if ($prior) { [int]$prior.Attempts + 1 } else { 1 })
            FailedStep = $FailedStep; StepsDone = @($StepsDone); LastError = $LastError
        })
    }
    Export-NuQueue -Records @($queue)
}
#endregion

#region Parsing --------------------------------------------------------------
function Get-AccountList {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return @() }
    return $Raw -split "[`r`n,;]+" |
           ForEach-Object { $_.Trim() } |
           Where-Object   { $_ -ne '' } |
           Select-Object  -Unique
}
#endregion

#region Actions --------------------------------------------------------------
function Get-OnPremState {
    param([string]$Id, [string]$DC)
    $params = @{ Identity = $Id; ErrorAction = 'Stop' }
    if ($DC) { $params['DomainController'] = $DC }
    Get-RemoteMailbox @params |
        Select-Object Name, PrimarySmtpAddress, RemoteRoutingAddress, ArchiveStatus, ArchiveGuid
}

function Format-ArchiveState {
    param($State)
    $hasGuid = ($State.ArchiveGuid -and $State.ArchiveGuid -ne [Guid]::Empty)
    $status  = if ($State.ArchiveStatus) { $State.ArchiveStatus } else { 'None' }

    if (-not $hasGuid) {
        return ("      Archive              : Not enabled  (ArchiveStatus: {0})" -f $status)
    }
    if ($status -eq 'None') {
        return "      Archive              : Enabled -- provisioning; may not display immediately"
    }
    return ("      Archive              : Enabled  (ArchiveStatus: {0})" -f $status)
}

function Format-OnPremState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("    --- $Heading ---")
    [void]$sb.AppendLine(("      Name                 : {0}" -f $State.Name))
    [void]$sb.AppendLine(("      PrimarySmtpAddress   : {0}" -f $State.PrimarySmtpAddress))
    [void]$sb.AppendLine(("      RemoteRoutingAddress : {0}" -f $State.RemoteRoutingAddress))
    [void]$sb.AppendLine((Format-ArchiveState -State $State))
    return $sb.ToString()
}

function Invoke-MailboxEnable {
    <#
        Returns Text (what used to be the whole return), plus Outcome
        (Success / Partial / Failed), FailedStep, LastError and StepsDone for
        the retry queue. -RetryArchive: a mailbox that already exists is not
        'SKIPPED' but gets its archive if it has none -- the retry for a row
        whose first run enabled the mailbox and lost the archive.
    #>
    param(
        [string]$Account,
        [string]$DC,
        [string]$Mode,
        [string]$DirectoryId,
        [switch]$RetryArchive
    )
    $out = New-Object System.Text.StringBuilder
    $mb  = $Account.Trim()
    $rra = "$mb@$RemoteRoutingDomain"
    $res = [PSCustomObject]@{ Text = ''; Outcome = 'Success'; FailedStep = ''; LastError = ''; StepsDone = @() }
    $finish = { param([string]$Outcome, [string]$Step, [string]$Err)
        $res.Outcome = $Outcome; $res.FailedStep = $Step; $res.LastError = $Err; $res.Text = $out.ToString(); return $res }

    [void]$out.AppendLine("----- $mb  [$Mode] -----")

    # EVERY CALL BELOW NAMES ITS OWN DIRECTORY. That is the fix for an HQ
    # account being verified against a franchise controller on 2026-08-12.
    #
    # $DC used to be '' for HQ, so the enable and the verify carried no
    # -DomainController and fell through to whatever the SESSION happened to be
    # scoped to. Session scope is set by Set-ADServerSettings, is process-wide,
    # and outlives the tool that set it: any earlier franchise operation in the
    # same PowerShell window leaves a preferred server behind. There is no
    # documented way to clear one -- -PreferredServer is a MANDATORY Fqdn, and
    # -ViewEntireForest $true is documented to clear RecipientViewRoot and
    # nothing else -- so the old '-PreferredServer $null' rested on behaviour
    # Microsoft does not promise.
    #
    # A per-cmdlet -DomainController overrides the session preference, so once
    # every call names a controller the ambient scope stops mattering. $DC is
    # a real controller from Tenant.json for this account's directory, or ''
    # -- never a domain name, which Exchange rejects with 0x52.
    if (-not "$DC".Trim()) {
        [void]$out.AppendLine("  NOTE: no controller could be resolved for the '$DirectoryId' directory -- none is " +
                              "named in Tenant Manager and Active Directory discovery found none either. This runs " +
                              "on whichever controller the session is scoped to, which may not be this directory's.")
    }
    try {
        # Forest-wide first, pin second, never a recovery call. -DirectoryId is
        # passed so a candidate is validated against THIS account's directory:
        # without it, a franchise controller validates happily for an HQ
        # operation, which is the same failure arriving by a different door.
        $scope = Set-ADSessionScope -Candidate $DC -DirectoryId $DirectoryId
        if ($scope.Fallback) {
            [void]$out.AppendLine("  NOTE: $($scope.Detail)")
        }
    } catch {
        [void]$out.AppendLine("  Could not set AD server scope: $($_.Exception.Message)")
        return (& $finish 'Failed' 'Scope' "$($_.Exception.Message)")
    }
    [void]$out.AppendLine("  Directory: $Mode  |  server: $(if ($DC) { $DC } else { '(AD locates one)' })")

    # BEFORE THE FIRST EARLY RETURN, and that placement is the whole point.
    #
    # This sat at the end, next to the verify step, so it only ever printed on
    # the path where a mailbox had just been created. Every other outcome --
    # SKIPPED because a remote mailbox already exists, SKIPPED as not eligible,
    # FAILED to enable -- returns before reaching it, and those are exactly the
    # accounts somebody is looking at when they want to know why mail is not
    # working. The one case that already printed it was the case least likely to
    # need it.
    #
    # REPORTED, NOT ASSIGNED. This tool enables a mailbox; it does not put
    # anybody in a licence group, and adding that here would be a write hiding
    # inside a line of output.
    #
    # It is worth SAYING because the two are done by different people. A starter
    # whose mailbox exists and whose licence group is empty looks completely
    # finished from this window, and the gap is found later by the user, when
    # mail does not arrive.
    #
    # Its own try: this is informational, and a licence lookup that throws must
    # not turn a successful mailbox enable into a failure.
    try {
        $lic = Get-UserLicenseState -SamAccountName $mb -DirectoryId $DirectoryId -Server $DC -Cache $script:LicenceCache
        [void]$out.AppendLine("  Licence group: $(Format-UserLicenseState -State $lic)")
    } catch {
        [void]$out.AppendLine("  Licence group: could not be read -- $($_.Exception.Message)")
    }

    $before = $null
    try { $before = Get-OnPremState -Id $mb -DC $DC } catch { $before = $null }
    if ($before) {
        $hasArchive = ($before.ArchiveGuid -and $before.ArchiveGuid -ne [Guid]::Empty)
        if ($RetryArchive -and -not $hasArchive) {
            # THE RETRY FOR A LOST ARCHIVE. The mailbox is there; only the
            # second call is outstanding.
            [void]$out.AppendLine("  Remote mailbox already exists -- enabling the missing archive.")
            $res.StepsDone = @('Enable')
            $ap = @{ Identity = $mb; Archive = $true; ErrorAction = 'Stop' }
            if ($DC) { $ap['DomainController'] = $DC }
            try {
                Enable-RemoteMailbox @ap | Out-Null
                [void]$out.AppendLine("  Archive enabled.")
                $res.StepsDone = @('Enable', 'Archive')
                Write-CompletedLog -Action 'EnableRemoteMailboxArchive' -Target (Get-LogTarget $before.PrimarySmtpAddress $mb) -Result 'Success' -Detail "Archive enabled on retry ($Mode)"
                return (& $finish 'Success' '' '')
            } catch {
                [void]$out.AppendLine("  FAILED to enable the archive: $($_.Exception.Message)")
                Write-FailedLog -Action 'EnableRemoteMailboxArchive' -Target $mb -ErrorRecord $_ -Detail 'retry from queue'
                return (& $finish 'Partial' 'Archive' "$($_.Exception.Message)")
            }
        }
        [void]$out.AppendLine("  SKIPPED -- already has a remote mailbox.")
        [void]$out.Append((Format-OnPremState -Heading 'EXISTING (on-prem)' -State $before))
        return (& $finish 'Success' '' '')
    }

    $enableParams = @{
        Identity             = $mb
        Alias                = $mb
        RemoteRoutingAddress = $rra
        Confirm              = $false
        ErrorAction          = 'Stop'
    }
    if ($DC) { $enableParams['DomainController'] = $DC }
    try {
        Enable-RemoteMailbox @enableParams | Out-Null
        [void]$out.AppendLine("  Mailbox enabled.  ($rra)")
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'does not support recipients of this type' -or $msg -match 'already') {
            [void]$out.AppendLine("  SKIPPED -- already enabled, or not eligible for a remote mailbox.")
            return (& $finish 'Failed' 'Enable' "Not eligible for a remote mailbox, or already enabled: $msg")
        }
        [void]$out.AppendLine("  FAILED to enable: $msg")
        Write-FailedLog -Action 'EnableRemoteMailbox' -Target $mb -ErrorRecord $_ -Detail "remote routing: $rra"
        return (& $finish 'Failed' 'Enable' $msg)
    }
    $res.StepsDone = @('Enable')

    $archiveParams = @{ Identity = $mb; Archive = $true; ErrorAction = 'Stop' }
    if ($DC) { $archiveParams['DomainController'] = $DC }
    $archiveOk = $false
    try {
        Enable-RemoteMailbox @archiveParams | Out-Null
        $archiveOk = $true
        $res.StepsDone = @('Enable', 'Archive')
        [void]$out.AppendLine("  Archive enabled.")
    } catch {
        $res.LastError = "$($_.Exception.Message)"
        [void]$out.AppendLine("  WARNING -- mailbox enabled but archive failed: $($_.Exception.Message)")
        Write-FailedLog -Action 'EnableRemoteMailboxArchive' -Target $mb -ErrorRecord $_ -Detail 'mailbox enabled, archive did not'
    }

    try {
        $after = Get-OnPremState -Id $mb -DC $DC
        [void]$out.AppendLine("  Verified.")
        [void]$out.Append((Format-OnPremState -Heading 'AFTER (on-prem)' -State $after))
        # TARGET FALLS BACK TO THE LOGON NAME, ON PURPOSE.
        #
        # A new starter has no mailbox in the cloud yet -- it appears after the
        # next directory sync -- so PrimarySmtpAddress is frequently empty at the
        # moment this row is written. The v1 -Task call could not express that
        # and left Target BLANK, which is the one thing an audit row must never
        # be: a row nobody can search for is a row that may as well not exist.
        #
        # Get-LogTarget takes candidates in preference order, so this is the
        # address when there is one and the sAMAccountName when there is not.
        # Its fallback exists for exactly this case.
        $logDetail = if ($archiveOk) { "Enabled RemoteMailbox $mb ($Mode)" }
                     else            { "Enabled RemoteMailbox $mb ($Mode) -- ARCHIVE FAILED" }

        # PARTIAL, NOT SUCCESS, WHEN THE ARCHIVE DID NOT ENABLE. The mailbox is
        # usable and the archive is not, and the text alone said so while the
        # Result column said Success -- so a count of failures taken from the log
        # missed every one of these.
        Write-CompletedLog -Action 'EnableRemoteMailbox' `
            -Target (Get-LogTarget $after.PrimarySmtpAddress $mb) `
            -Result $(if ($archiveOk) { 'Success' } else { 'Partial' }) `
            -Detail $logDetail
        [void]$out.AppendLine("  Logged to CompletedTasks.csv")
    } catch {
        [void]$out.AppendLine("  Could not verify after enable: $($_.Exception.Message)")
        [void]$out.AppendLine("  NOT logged (verification failed).")
        return (& $finish 'Partial' 'Verify' "$($_.Exception.Message)")
    }
    if ($archiveOk) { return (& $finish 'Success' '' '') }
    return (& $finish 'Partial' 'Archive' $res.LastError)
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Enable Remote Mailbox'
$form.Size            = New-Object System.Drawing.Size(760, 700)
$form.StartPosition   = 'CenterScreen'
# FormBorderStyle/MaximizeBox/MinimizeBox are set by Set-FormResizable at the
# bottom of this file, not here. They used to be FixedDialog + both boxes off,
# which meant an operator pasting thirty account names could not make the list
# box any taller. The size above stays the size it OPENS at; it is now the
# minimum rather than the only.
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Text      = 'Connect to on-prem Exchange before running.'
$lblWarn.Location  = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size      = New-Object System.Drawing.Size(430, 20)
$lblWarn.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

# The "Directory: <server>" label that used to sit here has been removed.
#
# It answered a question nobody asks. Every account is resolved to its own
# directory individually -- the field label two lines down says so -- and this
# showed the DEFAULT, which is not necessarily where any given account was
# found. So it was decorative at best and misleading at worst.
#
# It was also a tenant server name printed into a window that ships to every
# tenant, and it sat under the '?' button, covering the text it was competing
# with. Deleting it fixes all three at once. The variable was created and never
# referenced again, so nothing else depended on it.

$lblDC          = New-Object System.Windows.Forms.Label
$lblDC.Text     = 'Domain controller (used only where an account needs a specific one):'
$lblDC.Location = New-Object System.Drawing.Point(12, 40)
$lblDC.AutoSize = $true
$form.Controls.Add($lblDC)

$cmbDC = New-Object System.Windows.Forms.ComboBox
$cmbDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbDC.Location = New-Object System.Drawing.Point(360, 37)
$cmbDC.Size     = New-Object System.Drawing.Size(372, 22)
# THIS DIRECTORY'S controllers only.
#
# Unfiltered, this listed every directory's servers and the other directory's
# entry sat at index 0. The default below then failed to match, SelectedIndex
# fell back to 0, and a write meant for one directory was pointed at another --
# silently, because a plausible server name is indistinguishable from the right
# one.
@(Get-TenantDomainController -DirectoryId $script:PinDirectoryId) |
    ForEach-Object { [void]$cmbDC.Items.Add($_) }
# Select the configured default, or the first entry -- but only when the list
# is non-empty. SelectedIndex = 0 on an empty list throws, and on a MIXED list
# it was the bug above.
if ($cmbDC.Items.Count -gt 0) {
    $idx = $cmbDC.Items.IndexOf($DefaultFranchiseDC)
    $cmbDC.SelectedIndex = if ($idx -ge 0) { $idx } else { 0 }
}
$lblDC.Visible = $false
$cmbDC.Visible = $false
# DC selector hidden 2026-08-03: the toolkit binds to the default franchise
# controller from Common.ps1 and nobody was choosing another one. The control
# is kept rather than deleted so it still supplies that default to the code
# below -- .Visible is the whole change, and reversing it is one word.
$form.Controls.Add($cmbDC)

# TWO TABS (2026-10-09), the same shape as Term Requests: New Requests holds
# what this window always had; Retry Queue holds what did not finish.
$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, 66)
$tabs.Size     = New-Object System.Drawing.Size(720, 236)
$tabs.Anchor   = 'Top,Left,Right'
$form.Controls.Add($tabs)

$tabNew = New-Object System.Windows.Forms.TabPage
$tabNew.Text = 'New Requests'
$tabs.TabPages.Add($tabNew)

$tabRetry = New-Object System.Windows.Forms.TabPage
$tabRetry.Text = 'Retry Queue'
$tabs.TabPages.Add($tabRetry)

$lblAcc          = New-Object System.Windows.Forms.Label
$lblAcc.Text     = 'Account name(s) -- one per line (sAMAccountName). The directory is detected automatically for each.'
$lblAcc.Location = New-Object System.Drawing.Point(8, 8)
$lblAcc.AutoSize = $true
$tabNew.Controls.Add($lblAcc)

$txtAcc               = New-Object System.Windows.Forms.TextBox
$txtAcc.Multiline     = $true
$txtAcc.ScrollBars    = 'Vertical'
$txtAcc.AcceptsReturn = $true
$txtAcc.Location      = New-Object System.Drawing.Point(8, 28)
$txtAcc.Size          = New-Object System.Drawing.Size(696, 140)
$txtAcc.Font          = New-Object System.Drawing.Font('Consolas', 9)
# Widens with the window, keeps its height. Extra height goes to the OUTPUT
# box below, because that is the one you scroll through afterwards.
$txtAcc.Anchor        = 'Top,Left,Right'
$tabNew.Controls.Add($txtAcc)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run'
$btnRun.Location = New-Object System.Drawing.Point(8, 176)
$btnRun.Size     = New-Object System.Drawing.Size(100, 28)
$btnRun.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$tabNew.Controls.Add($btnRun)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text     = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(116, 176)
$btnClear.Size     = New-Object System.Drawing.Size(90, 28)
$tabNew.Controls.Add($btnClear)

# ---- Retry Queue tab ----
$btnRetryPlan = New-Object System.Windows.Forms.Button
$btnRetryPlan.Text = 'Retry Selected'; $btnRetryPlan.Location = New-Object System.Drawing.Point(8, 8); $btnRetryPlan.Size = New-Object System.Drawing.Size(120, 26)
$tabRetry.Controls.Add($btnRetryPlan)
$btnRetryRefresh = New-Object System.Windows.Forms.Button
$btnRetryRefresh.Text = 'Refresh'; $btnRetryRefresh.Location = New-Object System.Drawing.Point(136, 8); $btnRetryRefresh.Size = New-Object System.Drawing.Size(80, 26)
$tabRetry.Controls.Add($btnRetryRefresh)
$btnRetryDrop = New-Object System.Windows.Forms.Button
$btnRetryDrop.Text = 'Remove Selected'; $btnRetryDrop.Location = New-Object System.Drawing.Point(224, 8); $btnRetryDrop.Size = New-Object System.Drawing.Size(125, 26)
$tabRetry.Controls.Add($btnRetryDrop)
$btnRetryAll = New-Object System.Windows.Forms.Button
$btnRetryAll.Text = 'Select All'; $btnRetryAll.Location = New-Object System.Drawing.Point(357, 8); $btnRetryAll.Size = New-Object System.Drawing.Size(85, 26)
$tabRetry.Controls.Add($btnRetryAll)
$lblRetryNote = New-Object System.Windows.Forms.Label
$lblRetryNote.Text = 'Fix the cause first (usually: the account has been created now), tick rows, press Retry Selected. A retry resumes where it stopped.'
$lblRetryNote.Location = New-Object System.Drawing.Point(8, 40); $lblRetryNote.Size = New-Object System.Drawing.Size(696, 16)
$lblRetryNote.Anchor = 'Top,Left,Right'
$tabRetry.Controls.Add($lblRetryNote)
$gridRetry = New-Object System.Windows.Forms.DataGridView
$gridRetry.Location = New-Object System.Drawing.Point(8, 60); $gridRetry.Size = New-Object System.Drawing.Size(696, 144)
$gridRetry.Anchor = 'Top,Left,Right'
$gridRetry.AllowUserToAddRows = $false; $gridRetry.AllowUserToDeleteRows = $false; $gridRetry.RowHeadersVisible = $false
$gridRetry.SelectionMode = 'FullRowSelect'; $gridRetry.MultiSelect = $false
$tabRetry.Controls.Add($gridRetry)
$cSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$cSel.Name = 'Sel'; $cSel.HeaderText = ''; $cSel.Width = 28
[void]$gridRetry.Columns.Add($cSel)
foreach ($c in @(@{N='Account';W=120}, @{N='Environment';W=80}, @{N='Attempts';W=60}, @{N='FailedStep';W=80}, @{N='LastAttempt';W=110}, @{N='LastError';W=360})) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.N; $col.HeaderText = $c.N; $col.Width = $c.W; $col.ReadOnly = $true
    [void]$gridRetry.Columns.Add($col)
}

function Update-NuRetryGrid {
    $gridRetry.Rows.Clear()
    foreach ($r in (Import-NuQueue)) {
        [void]$gridRetry.Rows.Add($false, $r.Account, $r.Environment, $r.Attempts, $r.FailedStep, $r.LastAttempt, $r.LastError)
    }
    $tabRetry.Text = if ($gridRetry.Rows.Count -gt 0) { "Retry Queue ($($gridRetry.Rows.Count))" } else { 'Retry Queue' }
}

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(642, 308)
$btnClose.Size     = New-Object System.Drawing.Size(90, 28)
# Follows the right edge. Run and Clear stay pinned left, where they were.
$btnClose.Anchor   = 'Top,Right'
$form.Controls.Add($btnClose)

$lblResult          = New-Object System.Windows.Forms.Label
$lblResult.Text     = 'Result:'
$lblResult.Location = New-Object System.Drawing.Point(12, 314)
$lblResult.Size     = New-Object System.Drawing.Size(720, 20)
$lblResult.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblResult.Anchor   = 'Top,Left,Right'
$form.Controls.Add($lblResult)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(12, 336)
$txtOut.Size       = New-Object System.Drawing.Size(720, 312)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
# All four edges: this is the control that absorbs every extra pixel, in both
# directions. Resizing a window is nearly always a request for more of ONE
# thing, and here that thing is the run output -- a twelve-account batch prints
# well past 330px and you should not have to scroll a box you just enlarged the
# window to avoid scrolling.
$txtOut.Anchor     = 'Top,Left,Right,Bottom'
$txtOut.BackColor  = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($txtOut)

function Invoke-NuRun {
    <#
        The run, for both buttons. -FromQueue: the accounts came from the
        Retry Queue, so an existing mailbox gets its archive rather than
        SKIPPED, and the result dialog says so.
    #>
    param([string[]]$Accounts, [switch]$FromQueue)
    $locked = @($btnRun, $btnClear, $btnClose, $btnRetryPlan, $btnRetryRefresh, $btnRetryDrop, $btnRetryAll, $tabs)
    foreach ($c in $locked) { $c.Enabled = $false }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $accounts = @($Accounts)
        if ($accounts.Count -eq 0) {
            $null = Show-ToolDialog(
                'No account names entered.', 'Nothing to do',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        $franDC = $cmbDC.SelectedItem

        $lblResult.Text = 'Resolving accounts...'
        $form.Refresh()

        $resolved    = @()
        $missing     = @()
        $probeErrors = @()
        foreach ($a in $accounts) {
            $envInfo = Find-UserEnvironmentBySam -SamAccountName $a
            if ($envInfo.Errors.Count -gt 0) { $probeErrors += $envInfo.Errors }
            if ($envInfo.Environment) {
                # A REAL CONTROLLER FOR THIS ACCOUNT'S DIRECTORY, OR NOTHING.
                #
                # Not $envInfo.ADServer. That is what the LOOKUP used, and the
                # lookup is Get-ADUser -Server, which happily accepts a domain
                # name. Exchange's -DomainController does not: like
                # -PreferredServer it runs a suitability check and rejects a
                # domain with "Active Directory error 0x52 ... A local error
                # occurred". Passing the lookup's server straight to
                # Enable-RemoteMailbox failed exactly that way on 2026-08-12.
                #
                # Three parameters, three different appetites:
                #   Get-ADUser -Server            domain OR server
                #   Exchange -DomainController    server ONLY (0x52 otherwise)
                #   Set-ADServerSettings
                #     -PreferredServer            server ONLY (0x52 otherwise)
                #
                # So the Exchange calls get a real controller for THIS
                # directory: one Tenant.json names, or failing that one
                # DCLocator discovers in that directory's domain. Only if both
                # come up empty does $writeDC stay '' -- and then the tool says
                # so rather than quietly riding on session scope.
                $writeDC = Resolve-DirectoryDomainController -DirectoryId $envInfo.Environment
                if ($envInfo.Environment -eq 'Franchise' -and "$franDC".Trim()) {
                    $writeDC = "$franDC"
                }
                $resolved += [PSCustomObject]@{
                    Account     = $a
                    Environment = $envInfo.Environment
                    DC          = $writeDC
                }
            } else {
                $missing += $a
            }
        }

        # NOT FOUND GOES TO THE QUEUE, not just a dialog. It is the commonest
        # outcome -- the AD account is usually created after the ticket -- and
        # the row is what brings it back once the account exists.
        foreach ($m in $missing) {
            try { Update-NuQueue -Account $m -Outcome 'Failed' -FailedStep 'Resolve' -LastError 'Not found in any configured directory -- create the account, then retry.' }
            catch { $probeErrors += "Could not queue '$m': $($_.Exception.Message)" }
        }
        if ($missing.Count -gt 0) {
            $extra = "`r`n`r`nThey are in the Retry Queue tab; tick them and press Retry Selected once the account exists."
            if ($resolved.Count -gt 0) {
                $extra += "`r`n`r`nThe other $($resolved.Count) account(s) can still be processed."
            }
            $null = Show-ToolDialog(
                ("Account does not yet exist:`r`n`r`n{0}`r`n`r`nNot found in any configured directory.{1}" -f `
                    ($missing -join "`r`n"), $extra),
                'Account does not yet exist',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        }

        if ($resolved.Count -eq 0) {
            $lblResult.Text = 'Nothing to do -- no account resolved.'
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('No account was found in any configured directory:')
            foreach ($m in $missing) { [void]$sb.AppendLine("  $m") }
            if ($probeErrors.Count -gt 0) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine('--- Lookup errors ---')
                foreach ($e in ($probeErrors | Select-Object -Unique)) { [void]$sb.AppendLine($e) }
            }
            $txtOut.Text = $sb.ToString()
            return
        }

        $hqN   = @($resolved | Where-Object { $_.Environment -eq 'HQ' }).Count
        $franN = @($resolved | Where-Object { $_.Environment -eq 'Franchise' }).Count
        $bits  = @()
        if ($hqN)   { $bits += "$hqN HQ" }
        if ($franN) { $bits += "$franN Franchise" }

        # The preview names the server per account, not just the environment.
        # Confirming "YXVaghani -> HQ" tells you nothing about where it is
        # about to be written; naming the server makes a wrong one refusable
        # before anything happens rather than diagnosable afterwards.
        $preview = ($resolved | ForEach-Object {
            $where = if ("$($_.DC)".Trim()) { "$($_.DC)" } else { 'AD locates a controller' }
            "  {0}  ->  {1}  ({2})" -f $_.Account, $_.Environment, $where
        }) -join "`r`n"
        $confirm = Show-ToolDialog(
            ("Enable remote mailbox + archive for {0} account(s)?`r`n`r`n{1}`r`n`r`n{2}" -f `
                $resolved.Count, $preview, ($bits -join '   |   ')),
            'Confirm',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $lblResult.Text = 'Cancelled.'
            $txtOut.Text = 'Cancelled. No changes made.'
            return
        }

        $out = New-Object System.Text.StringBuilder
        [void]$out.AppendLine("===== $(Get-Date -Format 'yyyy-MM-dd HH:mm') =====")
        if ($missing.Count -gt 0) {
            [void]$out.AppendLine('')
            [void]$out.AppendLine('DOES NOT EXIST YET (skipped, nothing attempted):')
            foreach ($m in $missing) { [void]$out.AppendLine("  $m") }
        }

        $n = 0
        $summary = @{ Success = 0; Partial = 0; Failed = 0 }
        foreach ($grp in @('HQ', 'Franchise')) {
            $set = @($resolved | Where-Object { $_.Environment -eq $grp })
            if ($set.Count -eq 0) { continue }
            # The header names the server the batch will actually use, for
            # every environment. It used to name one only for Franchise, so an
            # HQ batch printed no server at all -- which is precisely the
            # information that would have shown, at a glance, that HQ was about
            # to run against whatever the session felt like.
            $servers = @($set | ForEach-Object { "$($_.DC)" } | Where-Object { $_ } | Select-Object -Unique)
            $onDC = if ($servers.Count -gt 0) { ' on ' + ($servers -join ', ') } else { ' (AD locates a controller)' }
            [void]$out.AppendLine('')
            [void]$out.AppendLine("===== $grp batch ($($set.Count) account(s))$onDC =====")
            foreach ($r in $set) {
                $n++
                $lblResult.Text = "Processing $n of $($resolved.Count): $($r.Account) [$($r.Environment)] ..."
                $form.Refresh()
                $one = Invoke-MailboxEnable -Account $r.Account -DC $r.DC `
                           -Mode $r.Environment -DirectoryId $r.Environment -RetryArchive:$FromQueue
                [void]$out.Append($one.Text)
                [void]$out.AppendLine('')
                $summary[$one.Outcome] = [int]$summary[$one.Outcome] + 1
                try {
                    Update-NuQueue -Account $r.Account -Environment $r.Environment -Outcome $one.Outcome `
                        -FailedStep $one.FailedStep -LastError $one.LastError -StepsDone $one.StepsDone
                } catch { [void]$out.AppendLine("  (retry queue not updated: $($_.Exception.Message))") }
            }
        }

        if ($probeErrors.Count -gt 0) {
            [void]$out.AppendLine('--- Lookup errors (the accounts above were still processed) ---')
            foreach ($e in ($probeErrors | Select-Object -Unique)) { [void]$out.AppendLine($e) }
        }

        $lblResult.Text = "Done -- $($summary.Success) complete, $($summary.Partial) partial, $($summary.Failed) failed, $($missing.Count) not found."
        $txtOut.Text = $out.ToString()
        if ($summary.Partial -or $summary.Failed -or $missing.Count) {
            $null = Show-ToolDialog(
                ("$($summary.Partial) partial, $($summary.Failed) failed, $($missing.Count) not found." + "`r`n`r`n" +
                 'They are in the Retry Queue tab with the reason. Fix the cause, tick the rows, press Retry Selected.'),
                'Finished with problems', 'OK', 'Warning')
        }
    }
    catch {
        $lblResult.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Enable Remote Mailbox'
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        foreach ($c in $locked) { if ($c -and -not $c.IsDisposed) { $c.Enabled = $true } }
        try { Update-NuRetryGrid } catch { }
    }
}

$btnRun.Add_Click({ Invoke-NuRun -Accounts (Get-AccountList -Raw $txtAcc.Text) })

$btnRetryPlan.Add_Click({
    try {
        $picked = @()
        for ($i = 0; $i -lt $gridRetry.Rows.Count; $i++) {
            if ($gridRetry.Rows[$i].Cells['Sel'].Value -eq $true) { $picked += "$($gridRetry.Rows[$i].Cells['Account'].Value)" }
        }
        if ($picked.Count -eq 0) { $null = Show-ToolDialog('No rows ticked.', 'Retry', 'OK', 'Information'); return }
        Invoke-NuRun -Accounts $picked -FromQueue
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Retry Selected' }
})

$btnRetryRefresh.Add_Click({ try { Update-NuRetryGrid } catch { Show-HandlerError -ErrorRecord $_ -Context 'Refresh Retry Queue' } })

$btnRetryAll.Add_Click({
    try {
        $any = $false
        for ($i = 0; $i -lt $gridRetry.Rows.Count; $i++) { if ($gridRetry.Rows[$i].Cells['Sel'].Value -ne $true) { $any = $true; break } }
        for ($i = 0; $i -lt $gridRetry.Rows.Count; $i++) { $gridRetry.Rows[$i].Cells['Sel'].Value = $any }
        $btnRetryAll.Text = if ($any) { 'Select None' } else { 'Select All' }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Select All' }
})

$btnRetryDrop.Add_Click({
    try {
        $drop = @()
        for ($i = 0; $i -lt $gridRetry.Rows.Count; $i++) {
            if ($gridRetry.Rows[$i].Cells['Sel'].Value -eq $true) { $drop += "$($gridRetry.Rows[$i].Cells['Account'].Value)".ToLower() }
        }
        if ($drop.Count -eq 0) { $null = Show-ToolDialog('No rows ticked.', 'Remove', 'OK', 'Information'); return }
        $ans = Show-ToolDialog("Remove $($drop.Count) row(s) from the retry queue?`r`n`r`nThis does not undo anything -- it only stops tracking them.", 'Remove from queue', 'YesNo', 'Question', 'Button2')
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $keep = @(Import-NuQueue | Where-Object { $drop -notcontains "$($_.Account)".ToLower() })
        Export-NuQueue -Records $keep
        Update-NuRetryGrid
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove From Queue' }
})

$btnClear.Add_Click({
    $txtAcc.Text = ''
    $txtOut.Text = ''
    $lblResult.Text = 'Result:'
})

$btnClose.Add_Click({ $form.Close() })

Update-NuRetryGrid

Set-ConsoleTheme -Form $form
# AFTER the size is final, because it pins MinimumSize to whatever Size is now.
Set-FormResizable -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'New User' | Out-Null

[void]$form.ShowDialog()
$form.Dispose()
#endregion
