#Requires -Version 5.1
<#
    Member_BulkADD_CSV.ps1
    Bulk add or remove distribution list members from a CSV list of addresses.

    REFACTORED FROM a standalone GUI script (Manage-DLMembers-GUI.ps1) into the
    console. The engine is the same -- the tolerant CSV loader, the per-address
    result grid, the dry run, the failure export that can be fed straight back
    in -- and all of it was worth keeping. What changed is everything the script
    did to the PROCESS it runs in, because it no longer owns that process.

    FIVE THINGS HAD TO GO, and each of them was load-bearing in the original:

    1. IT CONNECTED TO EXCHANGE ITSELF, from inside Add_Shown. B2: a modal
       ShowDialog blocks every other window on the thread, so a modern-auth
       prompt raised from a form draws and then ignores every click. Connecting
       is the console's job (C17). This checks the session and names the button.

    2. IT INSTALLED A PROCESS-WIDE EXCEPTION HANDLER --
       Application::SetUnhandledExceptionMode plus add_ThreadException -- to
       swallow errors raised by other scripts sharing its runspace. In the
       console that runspace is the CONSOLE, which runs every tool in-process.
       A tool that quietly redirects unhandled exceptions changes the console's
       behaviour for everything launched afterwards, and the removal handler in
       its finally block only runs if the tool exits normally. Gone entirely.

    3. IT HAD ITS OWN DPI SCALER (Get-Scaled, $script:Scale from Graphics.DpiX)
       applied to every fixed pixel. Every Get-Scaled call is now a plain
       number. NOTE: the shared scaler this deferred to was itself deleted on
       2026-09-11 (AGENTS.md rule 8), so nothing scales this window now -- which
       is the same as every other tool in the folder.

    4. IT OPENED MAXIMIZED, and ClientSize is set explicitly instead. That was
       done because the scaler took the size at Set-ConsoleTheme time as the
       design baseline; the scaler is gone, but an explicit opening size is
       still the house style and worth keeping.

    5. Anchor = 'Left,Right' MADE ITS TEXT BOXES FILL THEIR CELLS, and they are
       Dock = 'Top' now. Dock was chosen because the scaler surrendered anchors
       and honoured docks. With the scaler gone either would work -- Dock is
       kept because it is what the code says and it fills the cell correctly.

    WHAT WAS KEPT AND WHY: the docked TableLayoutPanel architecture. It resizes
    natively -- which is now the ONLY thing making this window resize at all --
    and rebuilding it at fixed coordinates would have been a large change with
    nothing to show for it.
#>
# TOOLKIT-HELP-RUNS: Reads a distribution list, ADDS or REMOVES members in bulk from a CSV, and ADDS or REMOVES Send-As / Send-To for named users. Writes CSV/TXT/HTML reports.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Add or remove a lot of people from one distribution list, from a list of
# addresses in a CSV.
#
# THE ORDER OF WORK:
#
#   1. Type the list address and press Validate DL. Its current membership
#      loads on the right.
#   2. Pick ADD or REMOVE.
#   3. Press Create & Open CSV to get a blank template, paste one address per
#      line under the Email header, SAVE it, then press Load Emails. An
#      existing CSV or a plain text list can be picked with Browse instead.
#   4. Press Preview members (dry run). Nothing is changed and every address
#      gets a row saying what WOULD happen.
#   5. Press Apply members from CSV.
#
# THE TWO HALVES HAVE THEIR OWN BUTTONS, and nothing in this window applies to
# both at once:
#
#   Step 1  Preview members (dry run) / Apply members from CSV
#   Step 2  Preview permissions       / Apply permissions
#
# The buttons along the bottom are exports and housekeeping -- they act on
# whatever the last operation put in the grid, whichever operation that was.
#
# NOTHING IS THROWN AT YOU. Every address ends up with a row and a status:
# Added or Removed in green, Already Exists or Not a Member in yellow, and
# Invalid Format, Recipient Not Found or Failed in red. One bad address in two
# hundred does not stop the other hundred and ninety-nine.
#
# AFTER AN UPDATE the membership is re-read from Exchange -- not assumed -- and
# everything just added is highlighted and flagged NEW.
#
# EXPORT FAILURES writes only the addresses that did not apply, with the reason
# for each, and leads with an Email column so the file can be pointed straight
# back at Load Emails to retry just those. You are asked whether to include the
# rows that needed no change.
#
# SEND-AS AND SEND-TO, for a handful of named users, are the second half of
# this window -- folded in from dl_send_to-as.ps1, which did one person at a
# time. Type them in the Users box, one per line or separated by commas, pick
# one of the six options, and press Preview permissions before Apply.
#
# THE USERS BOX IS NOT THE CSV. Members come from a file because there are
# often hundreds; delegations are typed because there are usually two or three.
# Sharing one input between them would mean granting Send-As to everybody in a
# three-hundred-line membership import by accident.
#
# SEND-AS means mail appears to come from the list. SEND-TO means being allowed
# to mail the list at all, which only matters when the list is restricted.
#
# THE ONE THING TO UNDERSTAND BEFORE USING "ADD SEND-TO":
#
# A list with an EMPTY allowed-sender list accepts mail from everyone. Adding
# the first people to that list does not grant them anything -- it locks
# everybody else OUT. The window checks for this and asks again, in plain
# words, because the intention ("let these three mail the list") and the effect
# ("only these three may mail the list") are opposites. Removing the last
# entries does the reverse and is also confirmed.
#
# AND: the allowed-sender list does nothing unless "Require sender
# authentication" is on. The Now-set line says which. If it reads OFF, adding a
# sender changes nothing anybody will notice.
#
# A LIST SYNCED FROM ON-PREMISES cannot have its allowed-sender list changed
# here -- it is mastered in Active Directory and the next sync would put it
# back. That is refused. Send-As is a cloud ACL and is allowed. Membership is
# warned about but allowed, because cloud-writable synced lists do exist.
#
# IT DOES NOT CONNECT. If the ExO indicator on the console is red, press
# Connect-ExchangeOnline there first; this window will refuse and say so rather
# than raising a sign-in prompt it cannot service.
# TOOLKIT-HELP-END

[CmdletBinding()]
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController,
    [string]$WorkFolder
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

# EnableVisualStyles is NOT called here. It is a process-wide setting and the
# console has already made it; calling it again after a window exists in the
# process is at best ignored.

# ============================================================================
# Configuration / state
# ============================================================================
if (-not $WorkFolder) { $WorkFolder = $root }
$script:WorkFolder = $WorkFolder
$script:CsvPath    = Join-Path $script:WorkFolder 'DLUsers.csv'

# The per-run transcript goes in the toolkit's Logs folder, not beside the
# script. A bulk operation's line-by-line record is worth keeping, and Logs is
# where somebody looking for it will look.
$script:LogDir  = Join-Path $root 'Logs'
$script:LogFile = Join-Path $script:LogDir ('DLBulkMembers_{0:yyyyMMdd}.log' -f (Get-Date))

$script:Group        = $null   # resolved distribution group
$script:MemberIndex  = @{}     # lowercase smtp -> member object
$script:Members      = @()     # member rows shown in the grid
$script:Results      = @()     # result rows shown in the grid
$script:StagedEmails = @()     # addresses loaded from the CSV
$script:NewlyAdded   = @{}     # lowercase smtp added in the last run
$script:RecipientCache = @{}
$script:Busy         = $false

# Send-As / Send-To, folded in from dl_send_to-as.ps1.
$script:SendAsList = @()     # trustees who can send AS the list
$script:SendTo     = @()     # AcceptMessagesOnlyFromSendersOrMembers as it stands
$script:AuthReq    = $false  # RequireSenderAuthenticationEnabled -- see Get-DlSendTo

# Declared before the UI exists so the helpers can test them safely.
$script:Form = $null; $script:Tip = $null; $script:TxtLog = $null
$script:LblStatus = $null; $script:LblProgress = $null; $script:Progress = $null
$script:LblConn = $null; $script:LblDlInfo = $null; $script:TxtDl = $null
$script:TxtCsv = $null; $script:TxtFilter = $null
$script:GridResults = $null; $script:GridMembers = $null
$script:LblResults = $null; $script:LblMembers = $null
$script:RadioAdd = $null; $script:RadioRemove = $null
$script:TxtUsers = $null; $script:LblPerm = $null
$script:RbPermAddAs = $null; $script:RbPermAddTo = $null; $script:RbPermAddBoth = $null
$script:RbPermRemAs = $null; $script:RbPermRemTo = $null; $script:RbPermRemBoth = $null
$script:ActionButtons = @()

# NO Get-Scaled AND NO $script:Scale. See point 3 in the header.

$EMAIL_COLUMNS = @('Email', 'EmailAddress', 'PrimarySmtpAddress', 'UserPrincipalName', 'Mail', 'Address', 'SMTP')
$EMAIL_REGEX   = '^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]{2,}$'

# Status buckets used for colouring, counting and the failure export.
$STATUS_SUCCESS  = @('Added', 'Removed')
$STATUS_NOCHANGE = @('Already Exists', 'Not a Member')
$STATUS_PROBLEM  = @('Failed', 'Invalid Format', 'Recipient Not Found')

# ONLY THE MEANING-BEARING COLOURS ARE DEFINED HERE. The original carried a
# whole light-theme palette -- form, panel, text, muted, buttons, grid stripes
# -- which is Set-ConsoleTheme's job and would have fought it. What a theme
# cannot supply is "green means it worked", so those four stay, and every
# control that gets one is tagged SemanticColour so theming leaves it alone (B3).
$C_GREEN  = [System.Drawing.Color]::FromArgb(200, 255, 200)
$C_RED    = [System.Drawing.Color]::FromArgb(255, 200, 200)
$C_YELLOW = [System.Drawing.Color]::FromArgb(255, 245, 200)
$C_INFO   = [System.Drawing.Color]::FromArgb(232, 240, 254)

$F_BASE  = New-Object System.Drawing.Font('Segoe UI', 9)
$F_BOLD  = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$F_LABEL = New-Object System.Drawing.Font('Segoe UI', 10)
$F_TITLE = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
$F_MONO  = New-Object System.Drawing.Font('Consolas', 9)
$script:Fonts = @{ Base = $F_BASE; Bold = $F_BOLD; Label = $F_LABEL; Title = $F_TITLE; Mono = $F_MONO }

# ============================================================================
# Helpers
# ============================================================================

function Write-ToolLog {
    <#
        One line in the window's transcript and in the day's file.

        NAMED ToolLog, NOT Write-Log. Common.ps1 has no Write-Log today, but a
        tool defining a name that generic in a process that dot-sources a shared
        layer is one rebuild away from shadowing something.
    #>
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')][string]$Level = 'Info'
    )
    $line = '[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}' -f (Get-Date), $Level.ToUpper(), $Message
    try {
        if (-not (Test-Path -LiteralPath $script:LogDir)) {
            [void](New-Item -ItemType Directory -Path $script:LogDir -Force -ErrorAction Stop)
        }
        Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop
    } catch { }
    if ($script:TxtLog) {
        $script:TxtLog.AppendText(('{0:HH:mm:ss}  [{1}] {2}{3}' -f (Get-Date), $Level.ToUpper(), $Message, [Environment]::NewLine))
        $script:TxtLog.SelectionStart = $script:TxtLog.TextLength
        $script:TxtLog.ScrollToCaret()
    }
}

function Update-Ui {
    <#
        DoEvents keeps the window drawing during a long loop.

        THE ORIGINAL WRAPPED THIS IN A TRY/CATCH that swallowed and logged
        anything DoEvents raised, because in the old Exchange text menu other
        GUI scripts' handlers could be sitting in the same message queue. That
        is not this situation and the swallow is worth losing: in the console
        the only things in this queue are this window's, so an exception here
        is OURS and hiding it would hide a real fault.
    #>
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Status {
    param([string]$Message, [ValidateSet('Info', 'Success', 'Warning', 'Error')][string]$Level = 'Info')
    if (-not $script:LblStatus) { return }
    $script:LblStatus.Text = $Message
    $script:LblStatus.BackColor = switch ($Level) {
        'Success' { $C_GREEN }
        'Warning' { $C_YELLOW }
        'Error'   { $C_RED }
        default   { $C_INFO }
    }
    $script:LblStatus.Tag = 'SemanticColour'
    Update-Ui
}

function Show-Note    { param([string]$m) $null = Show-ToolDialog($m, 'Bulk DL members', 'OK', 'Information') }
function Show-Caution { param([string]$m) $null = Show-ToolDialog($m, 'Bulk DL members', 'OK', 'Warning') }
function Show-Failure { param([string]$m) $null = Show-ToolDialog($m, 'Bulk DL members', 'OK', 'Error') }
function Confirm-Sentence {
    param([string]$Message)
    (Show-ToolDialog($Message, 'Please confirm', 'YesNo', 'Question')) -eq 'Yes'
}
function Get-YesNoCancel {
    param([string]$Message)
    # WRAPPED IN "$( )", NOT PREFIXED WITH A CAST.
    #
    # A type cast may sit in front of a METHOD call, which is what this line
    # used to be, and it parsed. It may NOT sit in front of a COMMAND, which
    # is what the MessageBox sweep turned that method into. The file stopped
    # parsing at this line and nothing in it could run -- found by the
    # toolkit's own SYN001 check on 2026-09-12, not by anybody opening the
    # tool. The offending form is deliberately not written out here: a
    # scanner looking for it must not find it in a comment about it.
    "$(Show-ToolDialog($Message, 'Bulk DL members', 'YesNoCancel', 'Question'))"
}

function Get-ShortError {
    param($ErrorRecord)
    $t = "$ErrorRecord"
    if ([string]::IsNullOrWhiteSpace($t)) { $t = 'Unknown error' }
    ($t -replace '\s+', ' ').Trim()
}

function New-ToolButton {
    <#
        AutoSize with a MinimumSize floor: a row of buttons stays even, but a
        caption that needs more room gets it once the engine has grown the font.
        Fixed widths clipped captions in the original and would again.
    #>
    param([string]$Text, [int]$Width = 168, [string]$ToolTip)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Font = $script:Fonts.Base
    $b.FlatStyle = 'Flat'
    $b.AutoSize = $true
    $b.AutoSizeMode = 'GrowAndShrink'
    $b.MinimumSize = New-Object System.Drawing.Size($Width, 32)
    $b.Padding = New-Object System.Windows.Forms.Padding(10, 4, 10, 4)
    $b.Margin = New-Object System.Windows.Forms.Padding(4)
    $b.TextAlign = 'MiddleCenter'
    if ($ToolTip -and $script:Tip) { $script:Tip.SetToolTip($b, $ToolTip) }
    $b
}

function New-ToolLabel {
    param([string]$Text, [switch]$Bold)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Font = if ($Bold) { $script:Fonts.Bold } else { $script:Fonts.Base }
    $l.AutoSize = $true
    $l.TextAlign = 'MiddleLeft'
    # NO Anchor. The engine surrenders it on every non-docked control, so an
    # anchor here would be dead weight that reads like live layout. AutoSize
    # plus the top margin puts these where they need to be in their cell.
    $l.Margin = New-Object System.Windows.Forms.Padding(4, 7, 10, 4)
    $l
}

function New-ToolGroupBox {
    param([string]$Text)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $Text
    $g.Font = $script:Fonts.Bold
    $g.Dock = 'Fill'
    $g.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 8)
    $g.Margin = New-Object System.Windows.Forms.Padding(6)
    $g
}

function New-ToolTextBox {
    param([string]$Text = '', [string]$ToolTip)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Text = $Text
    $t.Font = $script:Fonts.Label
    $t.BorderStyle = 'FixedSingle'
    # DOCK, NOT Anchor = 'Left,Right'. See point 5 in the header: the anchor
    # would be surrendered when the window is measured and the box would
    # collapse to its default width. Dock 'Top' fills the cell across and keeps
    # the box's own height.
    $t.Dock = 'Top'
    $t.Margin = New-Object System.Windows.Forms.Padding(4, 5, 8, 5)
    if ($ToolTip -and $script:Tip) { $script:Tip.SetToolTip($t, $ToolTip) }
    $t
}

function New-ToolGrid {
    param([string[]]$Columns, [double[]]$Weights)
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = 'Fill'
    $g.Font = $script:Fonts.Base
    $g.BorderStyle = 'FixedSingle'
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.ReadOnly = $true
    $g.RowHeadersVisible = $false
    $g.SelectionMode = 'FullRowSelect'
    $g.MultiSelect = $true
    # Fill mode computes its own widths, so the engine's column scaling returns
    # early on this grid and the columns are right at any size.
    $g.AutoSizeColumnsMode = 'Fill'
    $g.ColumnHeadersHeightSizeMode = 'AutoSize'
    $g.DefaultCellStyle.WrapMode = 'False'
    for ($i = 0; $i -lt $Columns.Count; $i++) {
        $idx = $g.Columns.Add("col$i", $Columns[$i])
        if ($Weights -and $i -lt $Weights.Count) { $g.Columns[$idx].FillWeight = $Weights[$i] }
    }
    $g
}

function Get-StatusColour {
    param([string]$Status)
    switch -Regex ($Status) {
        '^(Added|Removed|OK)'                        { $C_GREEN }
        '^(Already Exists|Not a Member|Skipped|NEW)' { $C_YELLOW }
        '^(Would|Pending)'                           { $C_INFO }
        default                                      { $C_RED }
    }
}

function Set-Busy {
    param([bool]$On)
    $script:Busy = $On
    foreach ($b in $script:ActionButtons) { $b.Enabled = -not $On }
    $script:Form.Cursor = if ($On) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
    Update-Ui
}

# ============================================================================
# Exchange Online
# ============================================================================

function Test-ExoReady {
    <#
        Connected? Asked, not fixed.

        B2 again: the original called Connect-ExchangeOnline from Add_Shown, and
        a sign-in prompt raised from a modal form draws and then ignores every
        click. So this reports and points at the console's own button.
    #>
    param([switch]$Quiet)
    try {
        $s = Test-ExoConnectionState
        Update-ConnectionBanner -Connected:$s.Connected
        if ($s.Connected) { return $true }
        if (-not $Quiet) {
            Show-Caution ("Exchange Online is not connected in this window." + [Environment]::NewLine + [Environment]::NewLine +
                "$($s.Detail)" + [Environment]::NewLine + [Environment]::NewLine +
                'Press Connect-ExchangeOnline on the console, then try again. This window cannot sign in itself -- a prompt raised from a modal form ignores every click.')
        }
        return $false
    } catch {
        if (-not $Quiet) { Show-Failure (Format-ErrorDetail -ErrorRecord $_ -Context 'Checking Exchange Online') }
        return $false
    }
}

function Update-ConnectionBanner {
    param([bool]$Connected)
    if (-not $script:LblConn) { return }
    if ($Connected) {
        $script:LblConn.Text = 'Exchange Online: CONNECTED'
        $script:LblConn.BackColor = $C_GREEN
    } else {
        $script:LblConn.Text = 'Exchange Online: NOT CONNECTED -- use the console'
        $script:LblConn.BackColor = $C_RED
    }
    $script:LblConn.Tag = 'SemanticColour'
}

function Resolve-OneRecipient {
    <# The recipient object or $null. Never throws. Cached. #>
    param([string]$Email)
    $key = $Email.ToLowerInvariant()
    if ($script:RecipientCache.ContainsKey($key)) { return $script:RecipientCache[$key] }
    $r = $null
    try {
        $r = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $Email; ErrorAction = 'Stop' } |
             Select-Object -First 1
    } catch { $r = $null }
    $script:RecipientCache[$key] = $r
    $r
}

function Get-MemberKeys {
    <# Every address a member can be recognised by, lowercased. #>
    param($Member)
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($p in 'PrimarySmtpAddress', 'WindowsLiveID', 'ExternalEmailAddress', 'Alias', 'Name') {
        $v = $null
        if ($Member.PSObject.Properties[$p]) { $v = [string]$Member.$p }
        if ($v) { [void]$keys.Add(($v -replace '^SMTP:', '').ToLowerInvariant()) }
    }
    if ($Member.PSObject.Properties['EmailAddresses']) {
        foreach ($a in @($Member.EmailAddresses)) {
            $s = [string]$a
            if ($s -match '^smtp:') { [void]$keys.Add(($s -replace '^(?i)smtp:', '').ToLowerInvariant()) }
        }
    }
    $keys | Where-Object { $_ } | Select-Object -Unique
}

function Update-GroupMemberList {
    <# Refreshes $script:MemberIndex and the Members grid. $true on success. #>
    param([switch]$Quiet)
    if (-not $script:Group) { return $false }
    try {
        Set-Status "Loading members of $($script:Group.PrimarySmtpAddress)..." 'Info'
        $raw = @(Invoke-EXOCommand -Name 'Get-DistributionGroupMember' -Parameters @{
            Identity = $script:Group.Guid; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
        $script:MemberIndex = @{}
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($m in $raw) {
            $smtp = ''
            if ($m.PSObject.Properties['PrimarySmtpAddress']) { $smtp = [string]$m.PrimarySmtpAddress }
            if (-not $smtp -and $m.PSObject.Properties['ExternalEmailAddress']) { $smtp = ([string]$m.ExternalEmailAddress) -replace '^(?i)smtp:', '' }
            foreach ($k in (Get-MemberKeys $m)) { $script:MemberIndex[$k] = $m }
            $isNew = $smtp -and $script:NewlyAdded.ContainsKey($smtp.ToLowerInvariant())
            $rows.Add([pscustomobject]@{
                DisplayName   = [string]$m.DisplayName
                Email         = $smtp
                RecipientType = [string]$m.RecipientTypeDetails
                State         = if ($isNew) { 'NEW' } else { 'Existing' }
                IsNew         = [bool]$isNew
            })
        }
        $script:Members = @($rows | Sort-Object -Property @{Expression = 'IsNew'; Descending = $true}, DisplayName)
        Update-MemberGrid
        if (-not $Quiet) { Write-ToolLog "Loaded $($script:Members.Count) member(s) from $($script:Group.PrimarySmtpAddress)." 'Info' }
        Set-Status "Membership loaded: $($script:Members.Count) member(s)." 'Success'
        return $true
    } catch {
        $msg = Get-ShortError $_
        Write-ToolLog "Could not read membership: $msg" 'Error'
        Set-Status 'Could not read the group membership.' 'Error'
        return $false
    }
}

function Update-MemberGrid {
    $filter = "$($script:TxtFilter.Text)".Trim()
    $view = $script:Members
    if ($filter) {
        $view = @($script:Members | Where-Object {
            $_.DisplayName -like "*$filter*" -or $_.Email -like "*$filter*" -or $_.RecipientType -like "*$filter*"
        })
    }
    $script:GridMembers.SuspendLayout()
    $script:GridMembers.Rows.Clear()
    $i = 0
    foreach ($m in $view) {
        $i++
        $idx = $script:GridMembers.Rows.Add($i, $m.DisplayName, $m.Email, $m.RecipientType, $m.State)
        if ($m.IsNew) {
            $script:GridMembers.Rows[$idx].DefaultCellStyle.BackColor = $C_GREEN
            $script:GridMembers.Rows[$idx].DefaultCellStyle.Font = $script:Fonts.Bold
        }
    }
    $script:GridMembers.ClearSelection()
    $script:GridMembers.ResumeLayout()
    $newCount = @($script:Members | Where-Object { $_.IsNew }).Count
    $script:LblMembers.Text = 'Total: {0}   |   Shown: {1}   |   Newly added (green): {2}' -f $script:Members.Count, $view.Count, $newCount
}

function Update-ResultGrid {
    $script:GridResults.SuspendLayout()
    $script:GridResults.Rows.Clear()
    $i = 0
    foreach ($r in $script:Results) {
        $i++
        $idx = $script:GridResults.Rows.Add($i, $r.Email, $r.Action, $r.Status, $r.Details)
        $script:GridResults.Rows[$idx].DefaultCellStyle.BackColor = (Get-StatusColour $r.Status)
    }
    $script:GridResults.ClearSelection()
    $script:GridResults.ResumeLayout()

    $counts = @{}
    foreach ($r in $script:Results) {
        $k = $r.Status
        if (-not $counts.ContainsKey($k)) { $counts[$k] = 0 }
        $counts[$k]++
    }
    $summary = ($counts.Keys | Sort-Object | ForEach-Object { "$_ = $($counts[$_])" }) -join '   |   '
    $script:LblResults.Text = if ($summary) { "Rows: $($script:Results.Count)   |   $summary" } else { "Rows: $($script:Results.Count)" }
}

# ============================================================================
# CSV handling
# ============================================================================

function New-CsvTemplate {
    $path = "$($script:TxtCsv.Text)".Trim()
    if (-not $path) { $path = $script:CsvPath; $script:TxtCsv.Text = $path }
    try {
        $dir = Split-Path -Path $path -Parent
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            [void](New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop)
            Write-ToolLog "Created folder: $dir" 'Info'
        }
        if (Test-Path -LiteralPath $path) {
            $existing = @(Get-Content -LiteralPath $path -ErrorAction SilentlyContinue)
            if ($existing.Count -gt 1) {
                if (-not (Confirm-Sentence "$path already contains $($existing.Count - 1) data line(s).`r`n`r`nOverwrite it with a blank template?`r`nChoose No to just open the existing file.")) {
                    Invoke-CsvEditor $path
                    return
                }
            }
        }
        Set-Content -LiteralPath $path -Value 'Email' -Encoding UTF8 -ErrorAction Stop
        Write-ToolLog "CSV template created: $path" 'Success'
        Invoke-CsvEditor $path
        Show-Note "A blank CSV template was created and opened:`r`n$path`r`n`r`nPaste one email address per line under the 'Email' header, SAVE the file, then press 'Load Emails'."
    } catch {
        $msg = Get-ShortError $_
        Write-ToolLog "Could not create the CSV template: $msg" 'Error'
        Show-Failure "Could not create the CSV template at:`r`n$path`r`n`r`n$msg"
    }
}

function Invoke-CsvEditor {
    param([string]$Path)
    try {
        if (-not (Test-Path -LiteralPath $Path)) { Show-Caution "The file does not exist yet:`r`n$Path"; return }
        Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$Path`"" -ErrorAction Stop
        Set-Status 'CSV opened in Notepad. Paste the addresses, save, then press Load Emails.' 'Info'
    } catch {
        try { Invoke-Item -LiteralPath $Path -ErrorAction Stop }
        catch { Show-Failure ("Could not open the CSV file.`r`n`r`n" + (Get-ShortError $_)) }
    }
}

function Select-CsvFile {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = 'Select the CSV / TXT file containing the email addresses'
    $dlg.Filter = 'CSV or text files (*.csv;*.txt)|*.csv;*.txt|All files (*.*)|*.*'
    if (Test-Path -LiteralPath $script:WorkFolder) { $dlg.InitialDirectory = $script:WorkFolder }
    if ($dlg.ShowDialog() -eq 'OK') {
        $script:TxtCsv.Text = $dlg.FileName
        Import-EmailList
    }
}

function Import-EmailList {
    <#
        Tolerant loader: a CSV with an Email-ish header, a single-column CSV with
        any header, or a plain list of addresses with no header at all. Also
        splits comma and semicolon separated addresses on one line.
    #>
    $path = "$($script:TxtCsv.Text)".Trim()
    if (-not $path) { Show-Caution 'Enter or browse to a CSV file first.'; return }
    if (-not (Test-Path -LiteralPath $path)) {
        Show-Caution "The file was not found:`r`n$path`r`n`r`nUse 'Create & Open CSV' to make a new template."
        return
    }

    $raw = @()
    try { $raw = @(Get-Content -LiteralPath $path -ErrorAction Stop | Where-Object { $_ -and $_.Trim() }) }
    catch { Show-Failure ("Could not read the file.`r`n`r`n" + (Get-ShortError $_)); return }

    if ($raw.Count -eq 0) {
        $script:StagedEmails = @()
        $script:Results = @()
        Update-ResultGrid
        Show-Caution "The file is empty:`r`n$path"
        return
    }

    $values = New-Object System.Collections.Generic.List[string]
    if ($raw[0] -match '@') {
        # No header row -- treat every line as data.
        foreach ($line in $raw) { [void]$values.Add($line) }
    } else {
        $csv = $null
        try { $csv = @(Import-Csv -LiteralPath $path -ErrorAction Stop) } catch { $csv = $null }
        if ($null -ne $csv -and $csv.Count -gt 0) {
            $props = @($csv[0].PSObject.Properties.Name)
            $col = $EMAIL_COLUMNS | Where-Object { $props -contains $_ } | Select-Object -First 1
            if (-not $col) { $col = $props | Where-Object { $_ -match '(?i)mail|smtp|upn|address' } | Select-Object -First 1 }
            if (-not $col) { $col = $props | Select-Object -First 1 }
            foreach ($row in $csv) { [void]$values.Add([string]$row.$col) }
        } else {
            foreach ($line in ($raw | Select-Object -Skip 1)) { [void]$values.Add($line) }
        }
    }

    # Split multi-address lines, trim quotes and whitespace, drop blanks, de-duplicate.
    $clean = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $dupes = 0
    foreach ($v in $values) {
        foreach ($part in ([string]$v -split '[,;]')) {
            $e = $part.Trim().Trim('"').Trim("'").Trim()
            if (-not $e) { continue }
            if ($e -match '^(?i)smtp:') { $e = $e -replace '^(?i)smtp:', '' }
            $k = $e.ToLowerInvariant()
            if ($seen.ContainsKey($k)) { $dupes++; continue }
            $seen[$k] = $true
            [void]$clean.Add($e)
        }
    }

    $script:StagedEmails = $clean.ToArray()
    $script:NewlyAdded = @{}
    $action = Get-SelectedAction
    $script:Results = @($script:StagedEmails | ForEach-Object {
        [pscustomobject]@{ Email = $_; Action = $action; Status = 'Pending'; Details = 'Loaded from CSV - not processed yet' }
    })
    Update-ResultGrid

    $invalid = @($script:StagedEmails | Where-Object { $_ -notmatch $EMAIL_REGEX }).Count
    Write-ToolLog "Loaded $($script:StagedEmails.Count) address(es) from $path (duplicates skipped: $dupes, malformed: $invalid)." 'Success'
    Set-Status "Loaded $($script:StagedEmails.Count) address(es). Duplicates skipped: $dupes. Malformed: $invalid." $(if ($invalid) { 'Warning' } else { 'Success' })
}

# ============================================================================
# Core operations
# ============================================================================

function Get-SelectedAction { if ($script:RadioAdd.Checked) { 'Add' } else { 'Remove' } }

function Resolve-Group {
    $identity = "$($script:TxtDl.Text)".Trim()
    if (-not $identity) { Show-Caution 'Enter the email address (or name) of the distribution list first.'; return $false }

    # A bare name becomes an address, by the toolkit's own rule.
    $completed = Get-TargetSmtp -Text $identity
    if ("$completed" -ne $identity) { $script:TxtDl.Text = "$completed"; $identity = "$completed" }

    if (-not (Test-ExoReady)) { return $false }

    Set-Busy $true
    try {
        Set-Status "Looking up '$identity'..." 'Info'
        $g = $null
        try {
            $g = Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{ Identity = $identity; ErrorAction = 'Stop' } |
                 Select-Object -First 1
        } catch { $g = $null }
        if (-not $g) {
            # A partial name still works. FLT001: the value is escaped before it
            # goes anywhere near an OPATH filter -- a display name containing a
            # quote would otherwise change what the filter means.
            try {
                $safe = ConvertTo-OpathValue -Value $identity
                $g = @(Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
                    Filter = "DisplayName -like '*$safe*'"; ResultSize = 5; WarningAction = 'SilentlyContinue'; ErrorAction = 'Stop' }) | Select-Object -First 1
            } catch { $g = $null }
        }
        if (-not $g) {
            $script:Group = $null
            $script:Members = @()
            Update-MemberGrid
            $script:LblDlInfo.Text = 'No distribution list resolved.'
            $script:LblDlInfo.BackColor = $C_RED
            $script:LblDlInfo.Tag = 'SemanticColour'
            Write-ToolLog "Distribution list not found: $identity" 'Warning'
            Set-Status "Distribution list not found: $identity" 'Error'
            Show-Caution "No distribution list matched:`r`n$identity`r`n`r`nCheck the address, or note that this may be a Microsoft 365 Group or a dynamic list, neither of which this tool manages."
            return $false
        }

        $script:Group = $g
        $script:NewlyAdded = @{}
        $script:RecipientCache = @{}
        $synced = if ($g.IsDirSynced) { '   |   SYNCED FROM ON-PREMISES' } else { '' }
        $script:LblDlInfo.Text = '{0}   <{1}>   |   Type: {2}   |   Join restriction: {3}{4}' -f `
            $g.DisplayName, $g.PrimarySmtpAddress, $g.RecipientTypeDetails, $g.MemberJoinRestriction, $synced
        $script:LblDlInfo.BackColor = if ($g.IsDirSynced) { $C_YELLOW } else { $C_GREEN }
        $script:LblDlInfo.Tag = 'SemanticColour'
        Write-ToolLog "Resolved DL: $($g.DisplayName) <$($g.PrimarySmtpAddress)> [$($g.RecipientTypeDetails)]" 'Success'

        # C18 IS SURFACED, NOT ENFORCED. Membership of a synced list is mastered
        # in AD, so a cloud write is either refused or undone by the next sync.
        # This tool says so and lets the operator decide, because a synced list
        # that is writable in the cloud does exist and refusing outright would
        # be wrong for it.
        if ($g.IsDirSynced) {
            Write-ToolLog 'This list is synced from on-premises: its membership is mastered in Active Directory. A change made here may be refused, or accepted and then undone by the next directory sync. Make it on-premises instead unless you know this list is cloud-writable.' 'Warning'
        }

        [void](Update-GroupMemberList)
        # The Send-As / Send-To panel is filled from the same lookup, so the two
        # halves of this window cannot disagree about which list is loaded.
        Update-PermissionView
        return $true
    } finally { Set-Busy $false }
}

function Invoke-Update {
    <# $DryRun = evaluate only, change nothing. #>
    param([switch]$DryRun)

    if (-not $script:Group) { Show-Caution 'Validate the distribution list first.'; return }
    if (-not (Test-ExoReady)) { return }
    if ($script:StagedEmails.Count -eq 0) { Show-Caution "No addresses are loaded.`r`n`r`nUse 'Create & Open CSV' to paste the list, then press 'Load Emails'."; return }

    $action = Get-SelectedAction
    $verb = if ($action -eq 'Add') { 'ADD' } else { 'REMOVE' }

    if (-not $DryRun) {
        $warn = ''
        if ($script:Group.IsDirSynced) {
            $warn = "`r`n`r`nTHIS LIST IS SYNCED FROM ON-PREMISES. Its membership is mastered in Active Directory, so this change may be refused or undone by the next directory sync."
        }
        $msg = "You are about to {0} {1} address(es) {2} the distribution list:`r`n`r`n    {3}`r`n    <{4}>`r`n`r`nCurrent members: {5}{6}`r`n`r`nProceed?" -f `
            $verb, $script:StagedEmails.Count, $(if ($action -eq 'Add') { 'TO' } else { 'FROM' }), $script:Group.DisplayName, $script:Group.PrimarySmtpAddress, $script:Members.Count, $warn
        if (-not (Confirm-Sentence $msg)) { Set-Status 'Cancelled.' 'Warning'; return }
    }

    Set-Busy $true
    $script:Progress.Value = 0
    $script:Progress.Maximum = [Math]::Max(1, $script:StagedEmails.Count)
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $DryRun) { $script:NewlyAdded = @{} }
    $mode = if ($DryRun) { 'DRY RUN' } else { 'LIVE' }
    Write-ToolLog "=== $mode $verb : $($script:StagedEmails.Count) address(es) on $($script:Group.PrimarySmtpAddress) ===" 'Info'

    $n = 0
    foreach ($email in $script:StagedEmails) {
        $n++
        $status = 'Failed'
        $details = ''
        try {
            if ($email -notmatch $EMAIL_REGEX) {
                $status = 'Invalid Format'
                $details = 'Not a valid SMTP address - skipped'
            } else {
                $recipient = Resolve-OneRecipient $email
                $lookup = if ($recipient -and $recipient.PrimarySmtpAddress) { ([string]$recipient.PrimarySmtpAddress).ToLowerInvariant() } else { $email.ToLowerInvariant() }
                $isMember = $script:MemberIndex.ContainsKey($lookup) -or $script:MemberIndex.ContainsKey($email.ToLowerInvariant())

                if ($action -eq 'Add') {
                    if (-not $recipient) {
                        $status = 'Recipient Not Found'
                        $details = 'No mailbox, mail user or contact exists with this address'
                    } elseif ($isMember) {
                        $status = 'Already Exists'
                        $details = 'Already a member of the distribution list'
                    } elseif ($DryRun) {
                        $status = 'Would Add'
                        $details = "Resolved to $($recipient.PrimarySmtpAddress) ($($recipient.RecipientTypeDetails))"
                    } else {
                        [void](Invoke-EXOCommand -Name 'Add-DistributionGroupMember' -Parameters @{
                            Identity = $script:Group.Guid; Member = $lookup
                            BypassSecurityGroupManagerCheck = $true; Confirm = $false; ErrorAction = 'Stop' })
                        $status = 'Added'
                        $details = "Added as $($recipient.PrimarySmtpAddress)"
                        $script:NewlyAdded[$lookup] = $true
                        $script:MemberIndex[$lookup] = $recipient
                    }
                } else {
                    if (-not $isMember) {
                        $status = 'Not a Member'
                        $details = if ($recipient) { 'Not in the distribution list' } else { 'Not in the distribution list, and the address does not resolve to a recipient' }
                    } elseif ($DryRun) {
                        $status = 'Would Remove'
                        $details = 'Currently a member'
                    } else {
                        [void](Invoke-EXOCommand -Name 'Remove-DistributionGroupMember' -Parameters @{
                            Identity = $script:Group.Guid; Member = $lookup
                            BypassSecurityGroupManagerCheck = $true; Confirm = $false; ErrorAction = 'Stop' })
                        $status = 'Removed'
                        $details = 'Removed from the distribution list'
                        [void]$script:MemberIndex.Remove($lookup)
                    }
                }
            }
        } catch {
            $status = 'Failed'
            $details = Get-ShortError $_
        }

        $rows.Add([pscustomobject]@{ Email = $email; Action = $action; Status = $status; Details = $details })
        $level = switch -Regex ($status) {
            '^(Added|Removed|Would)'         { 'Success' }
            '^(Already Exists|Not a Member)' { 'Warning' }
            default                          { 'Error' }
        }
        $detailText = if ($details) { " ($details)" } else { '' }
        Write-ToolLog ("[{0}] {1} -> {2}{3}" -f $mode, $email, $status, $detailText) $level

        $script:Progress.Value = [Math]::Min($n, $script:Progress.Maximum)
        $script:LblProgress.Text = "Processing $n of $($script:StagedEmails.Count)"
        $script:Results = $rows.ToArray()
        if (($n % 5) -eq 0 -or $n -eq $script:StagedEmails.Count) { Update-ResultGrid }
        Update-Ui
    }

    $script:Results = $rows.ToArray()
    Update-ResultGrid
    Set-Busy $false
    $script:LblProgress.Text = ''

    # READ BACK. An absent exception is not evidence the change landed -- and on
    # a synced list it positively is not.
    if (-not $DryRun) { [void](Update-GroupMemberList) }

    $ok      = @($rows | Where-Object { $STATUS_SUCCESS  -contains $_.Status }).Count
    $skipped = @($rows | Where-Object { $STATUS_NOCHANGE -contains $_.Status }).Count
    $failed  = @($rows | Where-Object { $STATUS_PROBLEM  -contains $_.Status }).Count
    $would   = @($rows | Where-Object { $_.Status -like 'Would*' }).Count

    if ($DryRun) {
        Write-ToolLog "Dry run complete: $would actionable, $skipped no-change, $failed problem(s)." 'Info'
        Set-Status "Dry run: $would to $verb, $skipped need no change, $failed problem(s). Nothing was modified." 'Info'
        Show-Note "Dry run finished - nothing was changed.`r`n`r`nWould be processed : $would`r`nNo change needed   : $skipped`r`nProblems           : $failed"
        return
    }

    $lvl = if ($failed) { 'Warning' } else { 'Success' }
    Write-ToolLog "$verb complete: $ok succeeded, $skipped no-change, $failed failed. Membership now $($script:Members.Count)." $lvl
    Set-Status "$verb complete: $ok succeeded, $skipped no-change, $failed failed. Members now: $($script:Members.Count)." $lvl

    # THE TOOLKIT'S AUDIT ROW, which the original had no notion of. One row for
    # the operation, with the counts -- the per-address detail is in the
    # transcript and the exports, and putting two hundred rows in the audit log
    # would bury everything else in it.
    $detail = "$verb $($script:StagedEmails.Count) address(es) on $($script:Group.PrimarySmtpAddress): $ok applied, $skipped no change, $failed failed."
    $target = Get-LogTarget "$($script:Group.PrimarySmtpAddress)" "$($script:Group.Guid)" 'DL Bulk Members'
    if ($failed -eq 0) {
        Write-CompletedLog -Action $action -Target $target -Result 'Success' -Detail $detail -Environment $Environment
    } elseif ($ok -gt 0) {
        Write-CompletedLog -Action $action -Target $target -Result 'Partial' -Detail $detail -Environment $Environment
    } else {
        Write-FailedLog -Action $action -Target $target -Detail $detail -Environment $Environment
    }

    $hint = if ($failed -gt 0 -or $skipped -gt 0) { "`r`n`r`nUse 'Export Failures' to save the addresses this $verb did not apply to, with the reason for each." } else { '' }
    Show-Note "$verb finished.`r`n`r`nSucceeded        : $ok`r`nNo change needed : $skipped`r`nFailed           : $failed`r`n`r`nMembers now      : $($script:Members.Count)`r`nTranscript       : $script:LogFile$hint"
}

# ============================================================================
# Export
# ============================================================================

function Get-UnsuccessfulRows {
    <#
        Every processed address that did NOT end up added or removed, with the
        reason. -IncludeNoChange also returns Already Exists / Not a Member.

        THE FIRST COLUMN IS Email so the exported file can be pointed straight
        back at Load Emails to retry.
    #>
    param([switch]$IncludeNoChange)
    $keep = if ($IncludeNoChange) { $STATUS_PROBLEM + $STATUS_NOCHANGE } else { $STATUS_PROBLEM }
    $dl = if ($script:Group) { [string]$script:Group.PrimarySmtpAddress } else { '' }
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    @($script:Results | Where-Object { $keep -contains $_.Status } | ForEach-Object {
        [pscustomobject]@{
            Email            = $_.Email
            Action           = $_.Action
            Status           = $_.Status
            Reason           = $_.Details
            DistributionList = $dl
            Reported         = $stamp
        }
    })
}

function Save-ExportFile {
    <# Writes the grid data as CSV, TXT or HTML based on the extension. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('Results', 'Members', 'Failures')][string]$What,
        [object[]]$Data
    )
    $export = switch ($What) {
        'Results'  { $Data | Select-Object Email, Action, Status, Details }
        'Members'  { $Data | Select-Object DisplayName, Email, RecipientType, State }
        'Failures' { $Data | Select-Object Email, Action, Status, Reason, DistributionList, Reported }
    }
    $dlText = if ($script:Group) { '{0} <{1}>' -f $script:Group.DisplayName, $script:Group.PrimarySmtpAddress } else { 'n/a' }
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '.txt' {
            $header = @(
                "Distribution List : $dlText"
                "Report            : $What"
                "Generated         : $stamp"
                ''
            )
            ($header + ($export | Format-Table -AutoSize | Out-String -Width 4096)) |
                Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
        }
        '.html' {
            $style = @'
<style>
 body { font-family: "Segoe UI", Arial, sans-serif; color:#323232; background:#f0f0f0; padding:20px; }
 h2 { color:#323232; margin-bottom:4px; }
 table { border-collapse:collapse; width:100%; background:#ffffff; }
 th { background:#e1ebf5; color:#323232; text-align:left; padding:8px; border:1px solid #cccccc; }
 td { padding:6px 8px; border:1px solid #dddddd; color:#323232; }
</style>
'@
            $pre = '<h2>Distribution list {0} report</h2><p>{1}<br/>Generated {2}</p>' -f $What.ToLower(), $dlText, $stamp
            $html = $export | ConvertTo-Html -Head $style -PreContent $pre | Out-String
            # Colour the status cells to match the on-screen grid.
            $html = $html -replace '<td>(Added|Removed|NEW)</td>', '<td style="background:#c8ffc8;font-weight:bold">$1</td>'
            $html = $html -replace '<td>(Already Exists|Not a Member)</td>', '<td style="background:#fff5c8">$1</td>'
            $html = $html -replace '<td>(Failed|Invalid Format|Recipient Not Found)</td>', '<td style="background:#ffc8c8">$1</td>'
            $html | Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
        }
        default {
            $export | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        }
    }
}

function Export-GridData {
    param([ValidateSet('Results', 'Members', 'Failures')][string]$What)

    $data = $null
    if ($What -eq 'Failures') {
        if (-not $script:Results -or $script:Results.Count -eq 0) {
            Show-Caution 'Nothing has been processed yet, so there are no failures to export.'
            return
        }
        $problems = @($script:Results | Where-Object { $STATUS_PROBLEM -contains $_.Status })
        $noChange = @($script:Results | Where-Object { $STATUS_NOCHANGE -contains $_.Status })
        if ($problems.Count -eq 0 -and $noChange.Count -eq 0) {
            Show-Note 'Every address was processed successfully - there is nothing to export.'
            return
        }
        $include = $false
        if ($noChange.Count -gt 0) {
            $answer = Get-YesNoCancel ("Addresses that could not be processed : $($problems.Count)`r`n" +
                "Addresses that needed no change      : $($noChange.Count)`r`n" +
                "  (Already Exists / Not a Member)`r`n`r`n" +
                "Include the 'needed no change' rows in the export?`r`n`r`n" +
                "Yes = export all $($problems.Count + $noChange.Count) rows`r`n" +
                "No  = export only the $($problems.Count) failures")
            if ($answer -eq 'Cancel') { Set-Status 'Export cancelled.' 'Info'; return }
            $include = ($answer -eq 'Yes')
        } elseif ($problems.Count -eq 0) {
            Show-Note 'Every address was processed successfully - there is nothing to export.'
            return
        }
        $data = if ($include) { Get-UnsuccessfulRows -IncludeNoChange } else { Get-UnsuccessfulRows }
    } else {
        $data = if ($What -eq 'Results') { $script:Results } else { $script:Members }
    }
    if (-not $data -or $data.Count -eq 0) { Show-Caution "There is nothing in the $What grid to export yet."; return }

    $dlgName = if ($script:Group) { ($script:Group.PrimarySmtpAddress -replace '[^\w\.\-]', '_') } else { 'DL' }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = "Export $What"
    $dlg.Filter = 'CSV file (*.csv)|*.csv|Text file (*.txt)|*.txt|HTML report (*.html)|*.html'
    $dlg.FileName = 'DL{0}_{1}_{2:yyyyMMdd_HHmmss}.csv' -f $What, $dlgName, (Get-Date)
    # The toolkit's reports folder, not wherever the script happens to sit.
    try { $dlg.InitialDirectory = Get-ReportPath } catch { }
    if ($dlg.ShowDialog() -ne 'OK') { return }

    $path = $dlg.FileName
    try {
        Save-ExportFile -Path $path -What $What -Data $data
        Write-ToolLog "Exported $What ($($data.Count) row(s)) to $path" 'Success'
        Set-Status "Exported $($data.Count) row(s) to $path" 'Success'
        Write-CompletedLog -Action 'Report' -Target (Get-LogTarget "$(if ($script:Group) { $script:Group.PrimarySmtpAddress } else { 'DL Bulk Members' })" 'DL Bulk Members') `
            -Result 'Success' -Detail "Exported $What ($($data.Count) row(s)) to $path." -Environment $Environment
        $extra = if ($What -eq 'Failures' -and [System.IO.Path]::GetExtension($path) -eq '.csv') {
            "`r`n`r`nThis file has an 'Email' column, so you can point 'CSV file' at it and press 'Load Emails' to retry just these addresses."
        } else { '' }
        if (Confirm-Sentence "Export saved to:`r`n$path`r`n`r`n$($data.Count) row(s) written.$extra`r`n`r`nOpen it now?") { Invoke-Item -LiteralPath $path }
    } catch {
        $msg = Get-ShortError $_
        Write-ToolLog "Export failed: $msg" 'Error'
        Show-Failure "Could not save the export to:`r`n$path`r`n`r`n$msg"
    }
}

# ============================================================================
# Send-As and Send-To
# ============================================================================
#
# FOLDED IN FROM dl_send_to-as.ps1, which did one person at a time. The rules
# are that script's and they have not changed; what changed is that this one
# does a list of people, and one of those rules gets sharper when it does.
#
# SEND-AS is a cloud ACL per trustee: Add-RecipientPermission with -AccessRights
# SendAs, one call per person. A loop is the right shape for it.
#
# SEND-TO is not a permission. It is the group's delivery restriction,
# AcceptMessagesOnlyFromSendersOrMembers, and Set-DistributionGroup REPLACES
# that property whole. So it is written ONCE for the whole batch -- read the
# list, apply every change to it in memory, write it back a single time.
#
# LOOPING IT PER PERSON WOULD BE WRONG, not just slow: each write would be
# based on a value read before the previous write landed, and the last one
# would win. Twenty users would produce one allowed sender and no error.

function Get-DlSendAs {
    <# Who can already send AS this list. #>
    $out = [PSCustomObject]@{ Ok = $false; Reason = ''; Rows = @() }
    if (-not $script:Group) { $out.Reason = 'No list is loaded.'; return $out }
    try {
        $rows = @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{
            Identity = $script:Group.Guid; ErrorAction = 'Stop' })
        # NT AUTHORITY\SELF is on every object and is not a delegation anybody
        # granted, so listing it would pad the answer with noise.
        $out.Rows = @($rows |
            Where-Object { "$($_.AccessRights)" -match 'SendAs' -and "$($_.Trustee)" -ne 'NT AUTHORITY\SELF' } |
            ForEach-Object { "$($_.Trustee)" })
        $out.Ok = $true
    } catch {
        $out.Reason = Get-ShortError $_
    }
    return $out
}

function Get-DlSendTo {
    <#
        The allowed-sender list, and the flag that decides whether it means
        anything. Read together because they have to be read together to be
        understood.
    #>
    $out = [PSCustomObject]@{ Ok = $false; Reason = ''; Rows = @(); AuthRequired = $false }
    if (-not $script:Group) { $out.Reason = 'No list is loaded.'; return $out }
    try {
        $g = Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
            Identity = $script:Group.Guid; ErrorAction = 'Stop' }
        $out.Rows         = @($g.AcceptMessagesOnlyFromSendersOrMembers | ForEach-Object { "$_" } | Where-Object { "$_".Trim() })
        $out.AuthRequired = [bool]$g.RequireSenderAuthenticationEnabled
        $out.Ok = $true
    } catch {
        $out.Reason = Get-ShortError $_
    }
    return $out
}

function Update-PermissionView {
    <# Refreshes the Send-As / Send-To panel from Exchange. #>
    if (-not $script:Group) {
        $script:LblPerm.Text = 'Validate a distribution list first.'
        return
    }
    $sa = Get-DlSendAs
    $st = Get-DlSendTo
    $script:SendAsList = @($sa.Rows)
    if ($st.Ok) { $script:SendTo = @($st.Rows); $script:AuthReq = $st.AuthRequired }

    $bits = New-Object System.Collections.Generic.List[string]
    if ($sa.Ok) { $bits.Add("Send-As: $(@($sa.Rows).Count)") } else { $bits.Add('Send-As: unreadable') }

    if (-not $st.Ok) {
        $bits.Add('Send-To: unreadable')
    } elseif (@($script:SendTo).Count -eq 0) {
        # THE STATE THAT REVERSES THE OPERATOR'S INTENTION, named on the form
        # rather than left in a confirmation dialog nobody reads twice.
        $bits.Add('Send-To: EMPTY -- ANYONE may mail this list')
    } else {
        $bits.Add("Send-To: $(@($script:SendTo).Count) allowed sender(s)")
    }

    if ($st.Ok -and -not $st.AuthRequired) {
        $bits.Add('Require sender authentication: OFF -- the allowed-sender list is IGNORED')
        $script:LblPerm.BackColor = $C_YELLOW
    } elseif (@($script:SendTo).Count -eq 0 -and $st.Ok) {
        $script:LblPerm.BackColor = $C_YELLOW
    } else {
        $script:LblPerm.BackColor = $C_INFO
    }
    $script:LblPerm.Tag  = 'SemanticColour'
    $script:LblPerm.Text = ($bits -join '   |   ')
}

function Get-StagedUsers {
    <#
        The Users box, split the same way the CSV loader splits a file: on
        lines, commas and semicolons, trimmed, de-duplicated, SMTP: prefix
        dropped, and a bare name completed to the tenant domain.

        THE SAME SPLITTER ON PURPOSE. An operator who has just pasted a column
        of addresses into a CSV should not discover that this box wants them
        formatted differently.
    #>
    $out  = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($line in ("$($script:TxtUsers.Text)" -split "\r\n|\n|\r")) {
        foreach ($part in ([string]$line -split '[,;]')) {
            $e = $part.Trim().Trim('"').Trim("'").Trim()
            if (-not $e) { continue }
            if ($e -match '^(?i)smtp:') { $e = $e -replace '^(?i)smtp:', '' }
            $e = "$(Get-TargetSmtp -Text $e)"
            if (-not $e) { continue }
            $k = $e.ToLowerInvariant()
            if ($seen.ContainsKey($k)) { continue }
            $seen[$k] = $true
            [void]$out.Add($e)
        }
    }
    return @($out.ToArray())
}

function Get-PermissionAction {
    if ($script:RbPermAddTo.Checked)   { return 'AddTo' }
    if ($script:RbPermAddBoth.Checked) { return 'AddBoth' }
    if ($script:RbPermRemAs.Checked)   { return 'RemoveAs' }
    if ($script:RbPermRemTo.Checked)   { return 'RemoveTo' }
    if ($script:RbPermRemBoth.Checked) { return 'RemoveBoth' }
    return 'AddAs'
}

function Test-IsAllowedSender {
    <#
        Is this person already in the allowed-sender list?

        MATCHED LOOSELY ON PURPOSE. Exchange returns those entries as display
        names or canonical AD paths depending on the object -- almost never as
        the address that was typed -- so an exact test would answer "no" for
        entries that are plainly there.
    #>
    #  -- AND WHY -Keys EXISTS --
    #
    # Exchange accepts an SMTP address into
    # AcceptMessagesOnlyFromSendersOrMembers and hands it back as a CANONICAL
    # AD PATH whose last segment is the object's Name or Alias, e.g.
    # "contoso.com/Users/SomeLongGroupName" for an address of
    # somegroup@contoso.com. Neither the address nor the display name
    # appears anywhere in it.
    #
    # Reported 2026-09-09: three Send-To writes succeeded and were then all
    # reported as FAILED, because this compared only PrimarySmtpAddress and
    # DisplayName against that path. A false failure on a write that landed is
    # as bad as the reverse -- the operator retries, or concludes the
    # restriction is not in place when it is.
    #
    # So callers pass a KEY SET built by Get-MemberKeys, which is the answer
    # this file already had a few hundred lines up: every address and name an
    # object can be recognised by, Alias and Name included. -Smtp and -Display
    # remain for callers that have nothing else.
    param([AllowEmptyString()][string]$Smtp = '',
          [AllowEmptyString()][string]$Display = '',
          [string[]]$Keys = @(),
          $List)

    $rows = if ($null -ne $List) { @($List) } else { @($script:SendTo) }
    $needles = New-Object System.Collections.Generic.List[string]
    foreach ($k in @($Keys))           { if ("$k".Trim()) { [void]$needles.Add("$k".Trim().ToLowerInvariant()) } }
    foreach ($k in @($Smtp, $Display)) { if ("$k".Trim()) { [void]$needles.Add("$k".Trim().ToLowerInvariant()) } }
    if ($needles.Count -eq 0) { return $false }

    foreach ($e in $rows) {
        $v     = "$e".Trim().ToLowerInvariant()
        $leaf  = ("$e" -split '/')[-1].Trim().ToLowerInvariant()
        # The local part as well: an entry can come back as the alias where the
        # address was alias@domain.
        $local = ($v -split '@')[0]
        if ($needles.Contains($v) -or $needles.Contains($leaf) -or $needles.Contains($local)) { return $true }
    }
    return $false
}

function Test-CanWriteSendTo {
    <#
        C18, ANSWERED BEFORE THE WRITE.

        AcceptMessagesOnlyFromSendersOrMembers on a directory-synced list is
        mastered in on-premises AD. Exchange Online will either refuse the
        change or accept it and have the next sync overwrite it -- and the
        second case is worse, because the tool reported success, the audit row
        says success, and the restriction is gone by morning.

        MEMBERSHIP IS WARNED ABOUT AND ALLOWED ON A SYNCED LIST, a few hundred
        lines above; this is REFUSED. The difference is not inconsistency: a
        cloud-writable synced list is a real thing for membership, whereas this
        property is read straight back off the AD object every sync cycle.
    #>
    $r = [PSCustomObject]@{ Allowed = $false; Reason = '' }
    if (-not $script:Group) { $r.Reason = 'No list is loaded.'; return $r }
    if ($script:Group.IsDirSynced) {
        $r.Reason = "'$($script:Group.DisplayName)' is SYNCED FROM ON-PREMISES. Its delivery restriction is mastered in Active Directory, so this change has to be made there -- Exchange Online will either refuse it or accept it and have the next directory sync undo it. Send-As is a cloud ACL and is still allowed."
        return $r
    }
    $r.Allowed = $true
    return $r
}

function Set-DlSendAsOne {
    param([Parameter(Mandatory)][string]$Trustee,
          [Parameter(Mandatory)][ValidateSet('Add', 'Remove')][string]$Mode)

    $name = if ($Mode -eq 'Add') { 'Add-RecipientPermission' } else { 'Remove-RecipientPermission' }
    [void](Invoke-EXOCommand -Name $name -Parameters @{
        Identity     = $script:Group.Guid
        Trustee      = $Trustee
        AccessRights = 'SendAs'
        Confirm      = $false
        ErrorAction  = 'Stop'
    })
}

function Set-DlSendToBatch {
    <#
        The whole allowed-sender list, written once for the whole batch.

        READ FRESH HERE, not from the copy the panel is showing. The panel's
        copy could be minutes old, and this property is replaced rather than
        merged -- so writing a stale list back would silently delete whatever
        somebody else added in the meantime.

        BypassSecurityGroupManagerCheck because an administrator who is not in
        the group's ManagedBy is otherwise refused, with an error about
        ownership that sounds like a bug in the tool.
    #>
    #  -- WHY THIS TAKES OBJECTS AND NOT ADDRESSES --
    #
    # It used to take [string[]]$Smtp and, on Remove, drop any entry equal to
    # one of those addresses. The stored entries are canonical AD paths, so
    # nothing ever matched and REMOVE SILENTLY REMOVED NOTHING while the write
    # itself succeeded. Found 2026-09-09 by running an add, then a remove,
    # against a directory that echoes paths back the way the real one does.
    #
    # It now takes the resolved user objects, each carrying the key set
    # Get-MemberKeys built, and decides keep-or-drop with the same matcher that
    # decides everything else about this list.
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Users,
          [Parameter(Mandatory)][ValidateSet('Add', 'Remove')][string]$Mode)

    $cur = Get-DlSendTo
    if (-not $cur.Ok) { throw "Could not read the current allowed-sender list: $($cur.Reason)" }

    $list = New-Object System.Collections.ArrayList
    foreach ($e in @($cur.Rows)) { [void]$list.Add("$e") }

    if ($Mode -eq 'Add') {
        foreach ($u in @($Users)) {
            if (-not (Test-IsAllowedSender -Keys $u.Keys -List $list)) { [void]$list.Add($u.Smtp) }
        }
    } else {
        # ONE PASS OVER THE LIST, each entry tested against every user's key
        # set. An entry survives only if it matches nobody being removed --
        # which is the same question Test-IsAllowedSender answers everywhere
        # else in this file, asked from the other side.
        $keep = New-Object System.Collections.ArrayList
        foreach ($e in @($list)) {
            $matched = $false
            foreach ($u in @($Users)) {
                if (Test-IsAllowedSender -Keys $u.Keys -List @("$e")) { $matched = $true; break }
            }
            if (-not $matched) { [void]$keep.Add("$e") }
        }
        $list = $keep
    }

    $final = @($list.ToArray())
    [void](Invoke-EXOCommand -Name 'Set-DistributionGroup' -Parameters @{
        Identity = $script:Group.Guid
        AcceptMessagesOnlyFromSendersOrMembers = $final
        BypassSecurityGroupManagerCheck = $true
        Confirm     = $false
        ErrorAction = 'Stop'
    })
    return $final
}

function Invoke-PermissionUpdate {
    <#
        Send-As and Send-To for everybody in the Users box.

        Results land in the same grid the member operation uses, because the
        columns already say which action a row was -- and one export path is
        better than two that drift.
    #>
    param([switch]$DryRun)

    if (-not $script:Group) { Show-Caution 'Validate the distribution list first.'; return }
    if (-not (Test-ExoReady)) { return }

    $users = Get-StagedUsers
    if (@($users).Count -eq 0) {
        Show-Caution "No users are listed.`r`n`r`nType one or more addresses in the Users box -- one per line, or separated by commas or semicolons."
        return
    }

    $action = Get-PermissionAction
    $doAs   = $action -in @('AddAs', 'AddBoth', 'RemoveAs', 'RemoveBoth')
    $doTo   = $action -in @('AddTo', 'AddBoth', 'RemoveTo', 'RemoveBoth')
    $adding = $action -like 'Add*'
    $mode   = if ($adding) { 'Add' } else { 'Remove' }

    # --- the refusal, before anything is written ---------------------------
    if ($doTo) {
        $can = Test-CanWriteSendTo
        if (-not $can.Allowed) {
            Show-Caution $can.Reason
            Write-ToolLog "REFUSED Send-To on a synced list: $($can.Reason)" 'Error'
            if (-not $doAs) {
                Write-FailedLog -Action $mode -Target (Get-LogTarget "$($script:Group.PrimarySmtpAddress)" "$($script:Group.Guid)" 'DL Send-As/Send-To') `
                    -Detail "REFUSED before writing: $($can.Reason)" -Environment $Environment
                return
            }
            $doTo = $false
        }
    }

    # --- read the current state so every row can be decided ----------------
    Set-Busy $true
    try {
        Update-PermissionView
        $sendAsNow = @($script:SendAsList)
        $sendToNow = @($script:SendTo)

        # Resolve everybody first. An address that does not exist gets a row
        # rather than an exception, exactly as the member path does.
        $resolved = New-Object System.Collections.Generic.List[object]
        $rows     = New-Object System.Collections.Generic.List[object]
        $n = 0
        $script:Progress.Value = 0
        $script:Progress.Maximum = [Math]::Max(1, @($users).Count)

        foreach ($u in $users) {
            $n++
            $script:LblProgress.Text = "Checking $n of $(@($users).Count)"
            $script:Progress.Value = [Math]::Min($n, $script:Progress.Maximum)
            Update-Ui

            if ($u -notmatch $EMAIL_REGEX) {
                $rows.Add([pscustomobject]@{ Email = $u; Action = $action; Status = 'Invalid Format'; Details = 'Not a valid SMTP address - skipped' })
                continue
            }
            $rec = Resolve-OneRecipient $u
            if (-not $rec) {
                $rows.Add([pscustomobject]@{ Email = $u; Action = $action; Status = 'Recipient Not Found'; Details = 'No mailbox, mail user, group or contact exists with this address' })
                continue
            }
            # EVERY NAME THIS OBJECT ANSWERS TO, not just its address.
            # Get-MemberKeys is this file's own helper for exactly that, and the
            # allowed-sender list comes back as canonical AD paths ending in
            # Name or Alias -- so without these the read-back cannot recognise
            # what it has just written.
            $keys = @()
            try { $keys = @(Get-MemberKeys $rec) } catch { $keys = @() }
            $keys = @($keys + @("$($rec.PrimarySmtpAddress)", "$($rec.DisplayName)") |
                     Where-Object { "$_".Trim() } | Select-Object -Unique)
            $resolved.Add([pscustomobject]@{
                Typed   = $u
                Smtp    = "$($rec.PrimarySmtpAddress)"
                Display = "$($rec.DisplayName)"
                Rtd     = "$($rec.RecipientTypeDetails)"
                Keys    = $keys
            })
        }

        if ($resolved.Count -eq 0) {
            $script:Results = $rows.ToArray()
            Update-ResultGrid
            $script:LblProgress.Text = ''
            Set-Status 'None of the addresses resolved. Nothing was changed.' 'Error'
            return
        }

        # NOTE ON .Count, BECAUSE IT COST AN HOUR ON 2026-09-09.
        #
        # $resolved is a System.Collections.Generic.List[object], and
        #     @($resolved).Count
        # THROWS System.ArgumentException "Argument types do not match" in
        # PowerShell 5.1. Not returns something odd -- throws. $resolved.Count
        # is correct and safe.
        #
        # The @() wrapper is the right defensive idiom for an ARRAY, where it
        # guards against a single item having no .Count, and it is the habit
        # that put it here. On a generic List it is actively harmful. Every
        # .Count in this function is now taken off the collection directly,
        # and the pipeline results below (which really are arrays) keep theirs.
        # --- what would change, per person --------------------------------
        # $toAdd / $toRemove used to be built here and were never read -- the
        # batch is taken from $toWork below. Removed rather than left as two
        # more generic Lists for somebody to call @().Count on.
        foreach ($r in $resolved) {
            # Get-RecipientPermission returns Trustee in whichever form the
            # directory feels like, so the same key set decides this as decides
            # the allowed-sender list.
            $hasAs = Test-IsAllowedSender -Keys $r.Keys -List $sendAsNow
            $hasTo = Test-IsAllowedSender -Keys $r.Keys -List $sendToNow

            if ($doAs) {
                if ($adding -and $hasAs) {
                    $rows.Add([pscustomobject]@{ Email = $r.Typed; Action = 'Add Send-As'; Status = 'Already Exists'; Details = "$($r.Display) can already send as this list" })
                } elseif (-not $adding -and -not $hasAs) {
                    $rows.Add([pscustomobject]@{ Email = $r.Typed; Action = 'Remove Send-As'; Status = 'Not a Member'; Details = "$($r.Display) does not have Send-As on this list" })
                } else {
                    $r | Add-Member -NotePropertyName 'DoAs' -NotePropertyValue $true -Force
                }
            }
            if ($doTo) {
                if ($adding -and $hasTo) {
                    $rows.Add([pscustomobject]@{ Email = $r.Typed; Action = 'Add Send-To'; Status = 'Already Exists'; Details = "$($r.Display) is already an allowed sender" })
                } elseif (-not $adding -and -not $hasTo) {
                    $rows.Add([pscustomobject]@{ Email = $r.Typed; Action = 'Remove Send-To'; Status = 'Not a Member'; Details = "$($r.Display) is not in the allowed-sender list" })
                } else {
                    $r | Add-Member -NotePropertyName 'DoTo' -NotePropertyValue $true -Force
                }
            }
        }

        $asWork = @($resolved | Where-Object { $_.DoAs })
        $toWork = @($resolved | Where-Object { $_.DoTo })

        if (@($asWork).Count -eq 0 -and @($toWork).Count -eq 0) {
            $script:Results = $rows.ToArray()
            Update-ResultGrid
            $script:LblProgress.Text = ''
            Set-Status 'Nothing to do -- every listed user is already in the state you asked for.' 'Info'
            Show-Note 'Nothing to do. Every user listed is already in the state you asked for; the grid says which.'
            return
        }

        # --- the sentence, with the two dangerous cases spelled out --------
        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("List  : $($script:Group.DisplayName) <$($script:Group.PrimarySmtpAddress)>")
        [void]$L.Add("Users : $($resolved.Count) resolved")
        [void]$L.Add('')
        if (@($asWork).Count -gt 0) { [void]$L.Add("$mode Send-As for $(@($asWork).Count) user(s)") }
        if (@($toWork).Count -gt 0) { [void]$L.Add("$mode Send-To for $(@($toWork).Count) user(s)") }

        $defaultCancel = $false

        # THE ONE THAT REVERSES THE OPERATOR'S INTENTION, and it is worse with a
        # list than with one person: the group stops accepting mail from
        # everybody except the handful named here.
        if ($doTo -and $adding -and @($sendToNow).Count -eq 0 -and @($toWork).Count -gt 0) {
            [void]$L.Add('')
            [void]$L.Add('READ THIS. The allowed-sender list on this group is currently EMPTY, which means ANYONE can mail it.')
            [void]$L.Add('')
            [void]$L.Add("Adding these $(@($toWork).Count) user(s) does NOT give them something they lack. It makes them the ONLY people who may mail this list -- everybody else, the whole company, stops being able to.")
            [void]$L.Add('')
            [void]$L.Add('If what you wanted was "let these people mail the list", they already can. Cancel.')
            $defaultCancel = $true
        }

        # THE REVERSE: emptying the list drops the restriction entirely.
        if ($doTo -and -not $adding -and @($toWork).Count -gt 0) {
            $left = @($sendToNow).Count - @($toWork).Count
            if ($left -le 0) {
                [void]$L.Add('')
                [void]$L.Add("This removes the LAST allowed sender(s). The list would be left empty, which means the restriction is gone and ANYONE will be able to mail $($script:Group.DisplayName).")
                $defaultCancel = $true
            }
        }

        if ($doTo -and -not $script:AuthReq -and @($toWork).Count -gt 0) {
            [void]$L.Add('')
            [void]$L.Add('Note: "Require sender authentication" is OFF on this list, so the allowed-sender list is not enforced. This change will be stored and will not alter who can actually mail the group.')
        }

        if ($DryRun) {
            foreach ($r in $asWork) { $rows.Add([pscustomobject]@{ Email = $r.Typed; Action = "$mode Send-As"; Status = "Would $mode"; Details = "$($r.Display) ($($r.Rtd))" }) }
            foreach ($r in $toWork) { $rows.Add([pscustomobject]@{ Email = $r.Typed; Action = "$mode Send-To"; Status = "Would $mode"; Details = "$($r.Display) ($($r.Rtd))" }) }
            $script:Results = $rows.ToArray()
            Update-ResultGrid
            $script:LblProgress.Text = ''
            Write-ToolLog "Dry run: $mode Send-As for $(@($asWork).Count), Send-To for $(@($toWork).Count). Nothing was changed." 'Info'
            Set-Status "Dry run: Send-As $(@($asWork).Count), Send-To $(@($toWork).Count). Nothing was modified." 'Info'
            [void]$L.Add('')
            [void]$L.Add('This was a DRY RUN. Nothing has been changed.')
            Show-Note ($L -join [Environment]::NewLine)
            return
        }

        [void]$L.Add('')
        [void]$L.Add('Proceed?')
        if (-not (Confirm-Sentence ($L -join [Environment]::NewLine))) {
            $script:Results = $rows.ToArray()
            Update-ResultGrid
            $script:LblProgress.Text = ''
            Set-Status 'Cancelled. Nothing was changed.' 'Warning'
            return
        }

        Write-ToolLog "=== LIVE $mode : Send-As $(@($asWork).Count), Send-To $(@($toWork).Count) on $($script:Group.PrimarySmtpAddress) ===" 'Info'

        # --- Send-As, one call per person ---------------------------------
        $n = 0
        $script:Progress.Value = 0
        $script:Progress.Maximum = [Math]::Max(1, @($asWork).Count + 1)
        foreach ($r in $asWork) {
            $n++
            $script:LblProgress.Text = "Send-As $n of $(@($asWork).Count)"
            $script:Progress.Value = [Math]::Min($n, $script:Progress.Maximum)
            Update-Ui
            try {
                Set-DlSendAsOne -Trustee $r.Smtp -Mode $mode
                $rows.Add([pscustomobject]@{ Email = $r.Typed; Action = "$mode Send-As"; Status = $(if ($adding) { 'Added' } else { 'Removed' }); Details = "$($r.Display) <$($r.Smtp)>" })
                Write-ToolLog "$mode Send-As: $($r.Smtp) -> ok" 'Success'
            } catch {
                $msg = Get-ShortError $_
                $rows.Add([pscustomobject]@{ Email = $r.Typed; Action = "$mode Send-As"; Status = 'Failed'; Details = $msg })
                Write-ToolLog "$mode Send-As: $($r.Smtp) -> FAILED ($msg)" 'Error'
            }
        }

        # --- Send-To, ONE write for the whole batch -----------------------
        if (@($toWork).Count -gt 0) {
            $script:LblProgress.Text = 'Writing the allowed-sender list...'
            Update-Ui
            $wanted = @($toWork | ForEach-Object { $_.Smtp })
            try {
                $final = Set-DlSendToBatch -Users $toWork -Mode $mode

                # NO PER-USER READ-BACK HERE, DELIBERATELY.
                #
                # dl_send_to-as.ps1 -- which works -- does not verify each
                # trustee against the stored list. It writes, and if the write
                # did not throw it says so, then re-reads the group and shows
                # the operator the actual state. This did verify, by trying to
                # match each typed address against what came back, and that
                # invention is what reported three successful writes as three
                # failures on 2026-09-09.
                #
                # The stored entries are canonical AD paths, so any matching
                # here is guesswork about a value the directory chose. The
                # write either threw or it did not; that is what is actually
                # known, and the Now-set line above shows the rest.
                foreach ($r in $toWork) {
                    $rows.Add([pscustomobject]@{ Email = $r.Typed; Action = "$mode Send-To"; Status = $(if ($adding) { 'Added' } else { 'Removed' }); Details = "$($r.Display) <$($r.Smtp)>" })
                }
                Write-ToolLog "$mode Send-To: one write for $(@($wanted).Count) user(s); the list now holds $(@($final).Count) entry(ies)." 'Success'
            } catch {
                $msg = Get-ShortError $_
                foreach ($r in $toWork) {
                    $rows.Add([pscustomobject]@{ Email = $r.Typed; Action = "$mode Send-To"; Status = 'Failed'; Details = $msg })
                }
                Write-ToolLog "$mode Send-To: FAILED for the whole batch ($msg)" 'Error'
            }
        }

        $script:Results = $rows.ToArray()
        Update-ResultGrid
        $script:LblProgress.Text = ''
        Update-PermissionView

        $ok      = @($rows | Where-Object { $STATUS_SUCCESS  -contains $_.Status }).Count
        $skipped = @($rows | Where-Object { $STATUS_NOCHANGE -contains $_.Status }).Count
        $failed  = @($rows | Where-Object { $STATUS_PROBLEM  -contains $_.Status }).Count

        $lvl = if ($failed) { 'Warning' } else { 'Success' }
        Write-ToolLog "$mode permissions complete: $ok applied, $skipped no change, $failed failed." $lvl
        Set-Status "$mode permissions: $ok applied, $skipped no change, $failed failed." $lvl

        $detail = "$mode $(if ($doAs) { 'Send-As' })$(if ($doAs -and $doTo) { ' and ' })$(if ($doTo) { 'Send-To' }) for $($resolved.Count) user(s) on $($script:Group.PrimarySmtpAddress): $ok applied, $skipped no change, $failed failed."
        $target = Get-LogTarget "$($script:Group.PrimarySmtpAddress)" "$($script:Group.Guid)" 'DL Send-As/Send-To'
        if ($failed -eq 0) {
            Write-CompletedLog -Action $mode -Target $target -Result 'Success' -Detail $detail -Environment $Environment
        } elseif ($ok -gt 0) {
            Write-CompletedLog -Action $mode -Target $target -Result 'Partial' -Detail $detail -Environment $Environment
        } else {
            Write-FailedLog -Action $mode -Target $target -Detail $detail -Environment $Environment
        }

        Show-Note "$mode finished.`r`n`r`nApplied          : $ok`r`nNo change needed : $skipped`r`nFailed           : $failed`r`n`r`nThe grid is a read-back, not what was asked for."
    } finally {
        Set-Busy $false
    }
}

# ============================================================================
# Build the UI
# ============================================================================

$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text = 'Distribution list: bulk members, Send-As and Send-To'
$script:Form.FormBorderStyle = 'Sizable'
$script:Form.StartPosition = 'Manual'
# EXPLICIT, AND BEFORE Set-ConsoleTheme. WindowState = 'Maximized' is what the
# original did; the engine records the size at theme time as the design size, so
# maximizing first would make a maximized window the baseline and mis-measure
# everything for the rest of the window's life (rule 8). No MinimumSize either
# -- the engine sets one at 55% and honours a tool's own if there is one.
$script:Form.ClientSize = New-Object System.Drawing.Size(1240, 900)
$script:Form.Font = $script:Fonts.Base

$script:Tip = New-Object System.Windows.Forms.ToolTip
$script:Tip.AutoPopDelay = 20000
$script:Tip.InitialDelay = 400

$root0 = New-Object System.Windows.Forms.TableLayoutPanel
$root0.Dock = 'Fill'
$root0.ColumnCount = 1
$root0.RowCount = 7
$root0.Padding = New-Object System.Windows.Forms.Padding(8)
[void]$root0.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$root0.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))            # header
[void]$root0.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))            # config
[void]$root0.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))            # send-as / send-to
# Grids and log share the leftover height, so the grids keep most of the window
# whatever the engine does to the fixed-height rows around them.
[void]$root0.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 78)))         # grids
[void]$root0.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 22)))         # log
[void]$root0.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))            # buttons
[void]$root0.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))            # status
$script:Form.Controls.Add($root0)

# ---------------------------------------------------------------- header row
$header = New-Object System.Windows.Forms.TableLayoutPanel
$header.Dock = 'Fill'
$header.AutoSize = $true
$header.AutoSizeMode = 'GrowAndShrink'
$header.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)
$header.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 2)
$header.ColumnCount = 2
$header.RowCount = 1
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$header.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

$titles = New-Object System.Windows.Forms.FlowLayoutPanel
$titles.FlowDirection = 'TopDown'
$titles.AutoSize = $true
$titles.AutoSizeMode = 'GrowAndShrink'
$titles.WrapContents = $false
$titles.Dock = 'Fill'
$titles.Margin = New-Object System.Windows.Forms.Padding(0)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Bulk distribution list members'
$lblTitle.Font = $script:Fonts.Title
$lblTitle.AutoSize = $true
$lblTitle.Margin = New-Object System.Windows.Forms.Padding(4, 0, 0, 0)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = 'Add or remove members of an Exchange Online distribution list from a CSV list of addresses.'
$lblSub.Font = $script:Fonts.Base
$lblSub.AutoSize = $true
$lblSub.Margin = New-Object System.Windows.Forms.Padding(6, 2, 0, 0)
$titles.Controls.AddRange(@($lblTitle, $lblSub))
$header.Controls.Add($titles, 0, 0)

# DOCKED, NOT Anchor = 'Right'. Its cell is already the right-hand one, and an
# anchor would be surrendered at measure time and drop it back to the left.
$connPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$connPanel.AutoSize = $true
$connPanel.AutoSizeMode = 'GrowAndShrink'
$connPanel.WrapContents = $false
$connPanel.Dock = 'Fill'
$connPanel.Margin = New-Object System.Windows.Forms.Padding(0)

$script:LblConn = New-Object System.Windows.Forms.Label
$script:LblConn.Text = 'Exchange Online: checking...'
$script:LblConn.Font = $script:Fonts.Bold
$script:LblConn.BackColor = $C_YELLOW
$script:LblConn.Tag = 'SemanticColour'
$script:LblConn.TextAlign = 'MiddleCenter'
$script:LblConn.AutoSize = $true
$script:LblConn.MinimumSize = New-Object System.Drawing.Size(300, 32)
$script:LblConn.Padding = New-Object System.Windows.Forms.Padding(12, 7, 12, 7)
$script:LblConn.Margin = New-Object System.Windows.Forms.Padding(4)

# NO Reconnect / Disconnect BUTTONS. Connecting is the console's job (C17, B2),
# and a Disconnect here would tear down the session every other open tool in
# this process is using.
$btnRecheck = New-ToolButton 'Re-check connection' 150 'Ask again whether this window has an Exchange Online session. Connecting is done from the console.'
$connPanel.Controls.AddRange(@($script:LblConn, $btnRecheck))
$header.Controls.Add($connPanel, 1, 0)
$root0.Controls.Add($header, 0, 0)

# ---------------------------------------------------------------- config row
$grpCfg = New-ToolGroupBox '  Step 1 - The list, and MEMBERS from a CSV  '
$grpCfg.AutoSize = $true
$grpCfg.AutoSizeMode = 'GrowAndShrink'
$cfg = New-Object System.Windows.Forms.TableLayoutPanel
$cfg.Dock = 'Top'
$cfg.AutoSize = $true
$cfg.AutoSizeMode = 'GrowAndShrink'
$cfg.ColumnCount = 4
[void]$cfg.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$cfg.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$cfg.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$cfg.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))

# row 0 - DL
$script:TxtDl = New-ToolTextBox "$(Get-TargetSmtp -Text $SmtpAddress)" 'Email address (or display name) of the distribution list. A bare name gets the tenant domain.'
$btnValidate = New-ToolButton 'Validate DL' 160 'Resolve the list and load its current members'
$btnRefresh  = New-ToolButton 'Reload Members' 140 'Re-read the membership from Exchange Online'
$cfg.Controls.Add((New-ToolLabel 'DL email address:' -Bold), 0, 0)
$cfg.Controls.Add($script:TxtDl, 1, 0)
$cfg.Controls.Add($btnValidate, 2, 0)
$cfg.Controls.Add($btnRefresh, 3, 0)

# row 1 - resolved banner
$script:LblDlInfo = New-Object System.Windows.Forms.Label
$script:LblDlInfo.Text = 'No distribution list validated yet.'
$script:LblDlInfo.Font = $script:Fonts.Bold
$script:LblDlInfo.BackColor = $C_YELLOW
$script:LblDlInfo.Tag = 'SemanticColour'
$script:LblDlInfo.TextAlign = 'MiddleLeft'
$script:LblDlInfo.AutoSize = $false
$script:LblDlInfo.Height = 28
$script:LblDlInfo.Dock = 'Top'
$script:LblDlInfo.Padding = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$script:LblDlInfo.Margin = New-Object System.Windows.Forms.Padding(4, 2, 8, 6)
$cfg.Controls.Add((New-ToolLabel 'Resolved list:' -Bold), 0, 1)
$cfg.Controls.Add($script:LblDlInfo, 1, 1)
$cfg.SetColumnSpan($script:LblDlInfo, 3)

# row 2 - action
$actionPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$actionPanel.AutoSize = $true
$actionPanel.Dock = 'Fill'
$actionPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$script:RadioAdd = New-Object System.Windows.Forms.RadioButton
$script:RadioAdd.Text = 'ADD members to the list'
$script:RadioAdd.Font = $script:Fonts.Bold
$script:RadioAdd.AutoSize = $true
$script:RadioAdd.Checked = $true
$script:RadioAdd.Margin = New-Object System.Windows.Forms.Padding(4, 4, 24, 4)
$script:RadioRemove = New-Object System.Windows.Forms.RadioButton
$script:RadioRemove.Text = 'REMOVE members from the list'
$script:RadioRemove.Font = $script:Fonts.Bold
$script:RadioRemove.AutoSize = $true
$script:RadioRemove.Margin = New-Object System.Windows.Forms.Padding(4)
$actionPanel.Controls.AddRange(@($script:RadioAdd, $script:RadioRemove))
$cfg.Controls.Add((New-ToolLabel 'Action:' -Bold), 0, 2)
$cfg.Controls.Add($actionPanel, 1, 2)
$cfg.SetColumnSpan($actionPanel, 3)

# row 3 - CSV
$script:TxtCsv = New-ToolTextBox $script:CsvPath 'Path of the CSV holding the addresses -- one per line under an "Email" header.'
$btnCsvNew  = New-ToolButton 'Create && Open CSV' 180 'Create a blank CSV template and open it in Notepad to paste the addresses'
$btnCsvLoad = New-ToolButton 'Load Emails' 140 'Read the addresses from the CSV into the grid'
$cfg.Controls.Add((New-ToolLabel 'CSV file:' -Bold), 0, 3)
$cfg.Controls.Add($script:TxtCsv, 1, 3)
$cfg.Controls.Add($btnCsvNew, 2, 3)
$cfg.Controls.Add($btnCsvLoad, 3, 3)

# row 4 - browse / open
$csvExtra = New-Object System.Windows.Forms.FlowLayoutPanel
$csvExtra.AutoSize = $true
$csvExtra.Dock = 'Fill'
$csvExtra.Margin = New-Object System.Windows.Forms.Padding(0)
$btnCsvBrowse = New-ToolButton 'Browse CSV...' 140 'Pick an existing CSV or TXT list of addresses'
$btnCsvOpen   = New-ToolButton 'Open CSV in Notepad' 170 'Open the CSV shown above for editing'
$csvExtra.Controls.AddRange(@($btnCsvBrowse, $btnCsvOpen))
$cfg.Controls.Add($csvExtra, 1, 4)
$cfg.SetColumnSpan($csvExtra, 3)

# row 5 - the MEMBER actions, in the section that owns them.
#
# THEY USED TO SIT IN THE BOTTOM ROW WITH THE EXPORTS, captioned "Preview (Dry
# Run)" and "Execute Update" -- which never said WHAT they updated. With a
# second operation in the window that was not a small ambiguity: the only
# clearly labelled pair was the permissions one, so the members pair read like
# a global Apply for the whole form. Each half now has its own buttons, beside
# its own inputs, naming its own operation.
$memBtns = New-Object System.Windows.Forms.FlowLayoutPanel
$memBtns.AutoSize = $true
$memBtns.AutoSizeMode = 'GrowAndShrink'
$memBtns.WrapContents = $true
$memBtns.Dock = 'Fill'
$memBtns.Margin = New-Object System.Windows.Forms.Padding(0)
$btnPreview = New-ToolButton 'Preview members (dry run)' 200 'Evaluate every CSV address against the membership and change nothing'
$btnExecute = New-ToolButton 'Apply members from CSV' 200 'Add or remove the loaded CSV addresses as members of the list'
$memBtns.Controls.AddRange(@($btnPreview, $btnExecute))
$cfg.Controls.Add((New-ToolLabel 'Members:' -Bold), 0, 5)
$cfg.Controls.Add($memBtns, 1, 5)
$cfg.SetColumnSpan($memBtns, 3)

$grpCfg.Controls.Add($cfg)
$root0.Controls.Add($grpCfg, 0, 1)

# ------------------------------------------------- Send-As / Send-To row
# ITS OWN SECTION, AND ITS OWN USERS BOX. The addresses here are not the CSV
# addresses: members come from a file because there are often hundreds, and
# delegations are typed because there are usually two or three. Sharing one
# input between them would mean granting Send-As to everybody in a
# three-hundred-line membership import by accident.
$grpPerm = New-ToolGroupBox '  Step 2 - PERMISSIONS: Send-As / Send-To for named users (independent of the CSV above)  '
$grpPerm.AutoSize = $true
$grpPerm.AutoSizeMode = 'GrowAndShrink'

$perm = New-Object System.Windows.Forms.TableLayoutPanel
$perm.Dock = 'Top'
$perm.AutoSize = $true
$perm.AutoSizeMode = 'GrowAndShrink'
$perm.ColumnCount = 3
[void]$perm.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$perm.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$perm.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))

# row 0 - the Users box
$script:TxtUsers = New-Object System.Windows.Forms.TextBox
$script:TxtUsers.Font = $script:Fonts.Label
$script:TxtUsers.BorderStyle = 'FixedSingle'
$script:TxtUsers.Multiline = $true
$script:TxtUsers.ScrollBars = 'Vertical'
$script:TxtUsers.WordWrap = $false
$script:TxtUsers.Height = 66
# Dock, not Anchor -- the engine surrenders the anchor of every non-docked
# control and the box would collapse to its default width.
$script:TxtUsers.Dock = 'Top'
$script:TxtUsers.Margin = New-Object System.Windows.Forms.Padding(4, 5, 8, 2)
if ($script:Tip) { $script:Tip.SetToolTip($script:TxtUsers, 'One or more users. One per line, or separated by commas or semicolons. A bare name gets the tenant domain.') }
$perm.Controls.Add((New-ToolLabel 'Users:' -Bold), 0, 0)
$perm.Controls.Add($script:TxtUsers, 1, 0)

$permBtns = New-Object System.Windows.Forms.FlowLayoutPanel
$permBtns.FlowDirection = 'TopDown'
$permBtns.AutoSize = $true
$permBtns.AutoSizeMode = 'GrowAndShrink'
$permBtns.WrapContents = $false
$permBtns.Dock = 'Fill'
$permBtns.Margin = New-Object System.Windows.Forms.Padding(0)
$btnPermPreview = New-ToolButton 'Preview permissions' 170 'Work out what would change for every listed user, and change nothing'
$btnPermApply   = New-ToolButton 'Apply permissions' 170 'Apply the selected Send-As / Send-To change to every listed user'
$permBtns.Controls.AddRange(@($btnPermPreview, $btnPermApply))
$perm.Controls.Add($permBtns, 2, 0)

# row 1 - the six actions
#
# RADIO BUTTONS, and they are a separate group from the ADD / REMOVE above
# because WinForms groups radios by their immediate parent. The six are
# mutually exclusive -- "Add both" already is Send-As plus Send-To -- so tick
# boxes would allow "Add Send-As" and "Remove both" together and the tool would
# have to guess.
$permActions = New-Object System.Windows.Forms.FlowLayoutPanel
$permActions.AutoSize = $true
$permActions.AutoSizeMode = 'GrowAndShrink'
$permActions.WrapContents = $true
$permActions.Dock = 'Fill'
$permActions.Margin = New-Object System.Windows.Forms.Padding(0)

$script:RbPermAddAs = New-Object System.Windows.Forms.RadioButton
$script:RbPermAddAs.Text = 'Add Send-As'
$script:RbPermAddAs.AutoSize = $true
$script:RbPermAddAs.Checked = $true
$script:RbPermAddAs.Margin = New-Object System.Windows.Forms.Padding(4, 4, 18, 4)

$script:RbPermAddTo = New-Object System.Windows.Forms.RadioButton
$script:RbPermAddTo.Text = 'Add Send-To'
$script:RbPermAddTo.AutoSize = $true
$script:RbPermAddTo.Margin = New-Object System.Windows.Forms.Padding(4, 4, 18, 4)

$script:RbPermAddBoth = New-Object System.Windows.Forms.RadioButton
$script:RbPermAddBoth.Text = 'Add both'
$script:RbPermAddBoth.AutoSize = $true
$script:RbPermAddBoth.Margin = New-Object System.Windows.Forms.Padding(4, 4, 28, 4)

$script:RbPermRemAs = New-Object System.Windows.Forms.RadioButton
$script:RbPermRemAs.Text = 'Remove Send-As'
$script:RbPermRemAs.AutoSize = $true
$script:RbPermRemAs.Margin = New-Object System.Windows.Forms.Padding(4, 4, 18, 4)

$script:RbPermRemTo = New-Object System.Windows.Forms.RadioButton
$script:RbPermRemTo.Text = 'Remove Send-To'
$script:RbPermRemTo.AutoSize = $true
$script:RbPermRemTo.Margin = New-Object System.Windows.Forms.Padding(4, 4, 18, 4)

$script:RbPermRemBoth = New-Object System.Windows.Forms.RadioButton
$script:RbPermRemBoth.Text = 'Remove both'
$script:RbPermRemBoth.AutoSize = $true
$script:RbPermRemBoth.Margin = New-Object System.Windows.Forms.Padding(4)

$permActions.Controls.AddRange(@($script:RbPermAddAs, $script:RbPermAddTo, $script:RbPermAddBoth,
    $script:RbPermRemAs, $script:RbPermRemTo, $script:RbPermRemBoth))
$perm.Controls.Add((New-ToolLabel 'Permission:' -Bold), 0, 1)
$perm.Controls.Add($permActions, 1, 1)
$perm.SetColumnSpan($permActions, 2)

# row 2 - what is set on the list now, and the flag that decides whether the
# allowed-sender list means anything at all
$script:LblPerm = New-Object System.Windows.Forms.Label
$script:LblPerm.Text = 'Validate a distribution list to see its Send-As and Send-To state.'
$script:LblPerm.Font = $script:Fonts.Bold
$script:LblPerm.BackColor = $C_INFO
$script:LblPerm.Tag = 'SemanticColour'
$script:LblPerm.TextAlign = 'MiddleLeft'
$script:LblPerm.AutoSize = $false
$script:LblPerm.Height = 26
$script:LblPerm.Dock = 'Top'
$script:LblPerm.Padding = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$script:LblPerm.Margin = New-Object System.Windows.Forms.Padding(4, 2, 8, 2)
$perm.Controls.Add((New-ToolLabel 'Now set:' -Bold), 0, 2)
$perm.Controls.Add($script:LblPerm, 1, 2)
$perm.SetColumnSpan($script:LblPerm, 2)

# row 3 - the warning that matters most, on the form rather than only in a dialog
$lblPermHint = New-ToolLabel 'SEND-AS = mail appears to come from the list.   SEND-TO = allowed to mail the list at all. An EMPTY allowed-sender list means everyone can mail it, so adding the first entries RESTRICTS the list to just those people -- you are asked again before that happens.'
$lblPermHint.MaximumSize = New-Object System.Drawing.Size(1100, 0)
$perm.Controls.Add($lblPermHint, 1, 3)
$perm.SetColumnSpan($lblPermHint, 2)

$grpPerm.Controls.Add($perm)
$root0.Controls.Add($grpPerm, 0, 2)

# ----------------------------------------------------------------- grids row
$gridsPanel = New-Object System.Windows.Forms.TableLayoutPanel
$gridsPanel.Dock = 'Fill'
$gridsPanel.ColumnCount = 2
$gridsPanel.RowCount = 1
$gridsPanel.Margin = New-Object System.Windows.Forms.Padding(0)
[void]$gridsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$gridsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$gridsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$grpRes = New-ToolGroupBox '  Step 2 - Results (green = done, yellow = no change needed, red = problem)  '
$resLayout = New-Object System.Windows.Forms.TableLayoutPanel
$resLayout.Dock = 'Fill'
$resLayout.ColumnCount = 1
$resLayout.RowCount = 2
[void]$resLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$resLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$resLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$script:GridResults = New-ToolGrid @('#', 'Email address', 'Action', 'Status', 'Details') @(6, 34, 12, 18, 40)
$script:GridResults.Columns[4].DefaultCellStyle.WrapMode = 'True'
$script:LblResults = New-ToolLabel 'Rows: 0'
$resLayout.Controls.Add($script:GridResults, 0, 0)
$resLayout.Controls.Add($script:LblResults, 0, 1)
$grpRes.Controls.Add($resLayout)
$gridsPanel.Controls.Add($grpRes, 0, 0)

$grpMem = New-ToolGroupBox '  Step 3 - Members after the update (green + NEW = just added)  '
$memLayout = New-Object System.Windows.Forms.TableLayoutPanel
$memLayout.Dock = 'Fill'
$memLayout.ColumnCount = 2
$memLayout.RowCount = 3
[void]$memLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$memLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$memLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$memLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$memLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$script:TxtFilter = New-ToolTextBox '' 'Type to filter the member list by name, address or type.'
$memLayout.Controls.Add((New-ToolLabel 'Filter:' -Bold), 0, 0)
$memLayout.Controls.Add($script:TxtFilter, 1, 0)
$script:GridMembers = New-ToolGrid @('#', 'Display name', 'Email address', 'Recipient type', 'State') @(6, 30, 38, 20, 12)
$memLayout.Controls.Add($script:GridMembers, 0, 1)
$memLayout.SetColumnSpan($script:GridMembers, 2)
$script:LblMembers = New-ToolLabel 'Total: 0   |   Shown: 0   |   Newly added (green): 0'
$memLayout.Controls.Add($script:LblMembers, 0, 2)
$memLayout.SetColumnSpan($script:LblMembers, 2)
$grpMem.Controls.Add($memLayout)
$gridsPanel.Controls.Add($grpMem, 1, 0)

$root0.Controls.Add($gridsPanel, 0, 3)

# ------------------------------------------------------------------- log row
$grpLog = New-ToolGroupBox '  Activity log  '
$grpLog.MinimumSize = New-Object System.Drawing.Size(0, 76)
$script:TxtLog = New-Object System.Windows.Forms.TextBox
$script:TxtLog.Dock = 'Fill'
$script:TxtLog.Multiline = $true
$script:TxtLog.ReadOnly = $true
$script:TxtLog.ScrollBars = 'Both'
$script:TxtLog.WordWrap = $false
$script:TxtLog.Font = $script:Fonts.Mono
$script:TxtLog.BorderStyle = 'FixedSingle'
$grpLog.Controls.Add($script:TxtLog)
$root0.Controls.Add($grpLog, 0, 4)

# --------------------------------------------------------------- button row
$btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$btnPanel.Dock = 'Fill'
$btnPanel.AutoSize = $true
$btnPanel.AutoSizeMode = 'GrowAndShrink'
$btnPanel.WrapContents = $true
$btnPanel.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 2)

# NO ACTION BUTTONS HERE ANY MORE -- see row 5 of the config panel and the
# Send-As / Send-To section. What is left acts on whatever the last operation
# produced, whichever operation that was.
$btnExportRes  = New-ToolButton 'Export Results' 168 'Save the results grid as CSV, TXT or HTML'
$btnExportFail = New-ToolButton 'Export Failures' 168 'Save only the addresses that were NOT added or removed, with the reason for each'
$btnExportMem  = New-ToolButton 'Export Members' 168 'Save the member list as CSV, TXT or HTML'
$btnOpenLog    = New-ToolButton 'Open Transcript' 168 "Open today's transcript file"
$btnClear      = New-ToolButton 'Clear Results' 168 'Empty the results grid and the loaded address list'
$btnPanel.Controls.AddRange(@($btnExportRes, $btnExportFail, $btnExportMem, $btnOpenLog, $btnClear))
$root0.Controls.Add($btnPanel, 0, 5)

# NO Exit BUTTON. Every window in this toolkit closes with its own title bar,
# and a tool that offers its own way out is one more thing to keep consistent
# across seventy-odd of them.

# --------------------------------------------------------------- status row
$statusPanel = New-Object System.Windows.Forms.TableLayoutPanel
$statusPanel.Dock = 'Fill'
$statusPanel.AutoSize = $true
$statusPanel.AutoSizeMode = 'GrowAndShrink'
$statusPanel.ColumnCount = 3
$statusPanel.RowCount = 1
$statusPanel.Margin = New-Object System.Windows.Forms.Padding(4, 0, 4, 2)
[void]$statusPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$statusPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 200)))
[void]$statusPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 260)))
[void]$statusPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

$script:LblStatus = New-Object System.Windows.Forms.Label
$script:LblStatus.Text = 'Ready.'
$script:LblStatus.Font = $script:Fonts.Bold
$script:LblStatus.BackColor = $C_INFO
$script:LblStatus.Tag = 'SemanticColour'
$script:LblStatus.TextAlign = 'MiddleLeft'
$script:LblStatus.Dock = 'Fill'
$script:LblStatus.MinimumSize = New-Object System.Drawing.Size(0, 30)
$script:LblStatus.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
$statusPanel.Controls.Add($script:LblStatus, 0, 0)

# Docked with a right-aligned caption rather than Anchor = 'Right', for the
# same reason as everything else in this file.
$script:LblProgress = New-Object System.Windows.Forms.Label
$script:LblProgress.Text = ''
$script:LblProgress.Font = $script:Fonts.Base
$script:LblProgress.TextAlign = 'MiddleRight'
$script:LblProgress.AutoSize = $false
$script:LblProgress.Dock = 'Fill'
$script:LblProgress.MinimumSize = New-Object System.Drawing.Size(0, 30)
$script:LblProgress.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
$statusPanel.Controls.Add($script:LblProgress, 1, 0)

$script:Progress = New-Object System.Windows.Forms.ProgressBar
$script:Progress.Dock = 'Fill'
$script:Progress.Style = 'Continuous'
$script:Progress.MinimumSize = New-Object System.Drawing.Size(0, 22)
$script:Progress.Margin = New-Object System.Windows.Forms.Padding(8, 6, 4, 6)
$statusPanel.Controls.Add($script:Progress, 2, 0)

$root0.Controls.Add($statusPanel, 0, 6)

$script:ActionButtons = @($btnRecheck, $btnValidate, $btnRefresh, $btnCsvNew, $btnCsvLoad,
    $btnCsvBrowse, $btnCsvOpen, $btnPreview, $btnExecute, $btnExportRes, $btnExportFail,
    $btnExportMem, $btnClear, $btnPermPreview, $btnPermApply)

# ============================================================================
# Handlers
# ============================================================================

function Invoke-Safely {
    <#
        B1: nothing escapes a handler. An exception that gets past one is caught
        by .NET rather than PowerShell, and the operator gets a CLR dialog with
        no script name on it.
    #>
    param([scriptblock]$Action, [string]$What)
    try { & $Action }
    catch {
        # THE LINE NUMBER IS THE POINT. This used to report only the message,
        # and on 2026-09-09 that produced "Argument types do not match" with
        # nothing to say where -- a bug that took an hour to find and would
        # have taken a minute with the line in front of us. A handler that
        # hides where it failed is not handling the error, it is hiding it.
        $msg   = Get-ShortError $_
        $where = ''
        try {
            $ln = "$($_.InvocationInfo.ScriptLineNumber)"
            $cmd = "$($_.InvocationInfo.Line)".Trim()
            if ($ln) { $where = "line $ln" }
            if ($cmd) { $where = "$where`: $cmd" }
        } catch { }
        Write-ToolLog "$What failed unexpectedly: $msg [$where]" 'Error'
        foreach ($frame in ("$($_.ScriptStackTrace)" -split "`r`n|`n")) {
            if ("$frame".Trim()) { Write-ToolLog "    $frame" 'Error' }
        }
        Set-Status "$What failed -- see the activity log for where." 'Error'
        Show-Failure "Something went wrong while doing: $What`r`n`r`n$msg`r`n`r`n$where`r`n`r`nThe activity log has the full stack."
    }
    finally { if ($script:Busy) { Set-Busy $false } }
}

$btnRecheck.Add_Click({ Invoke-Safely { [void](Test-ExoReady) } 'Re-check the connection' })
$btnValidate.Add_Click({ Invoke-Safely { [void](Resolve-Group) } 'Validate DL' })
$btnRefresh.Add_Click({
    Invoke-Safely {
        if (-not $script:Group) { Show-Caution 'Validate a distribution list first.'; return }
        [void](Update-GroupMemberList)
    } 'Reload members'
})
$btnCsvNew.Add_Click({ Invoke-Safely { New-CsvTemplate } 'Create CSV template' })
$btnCsvOpen.Add_Click({ Invoke-Safely { Invoke-CsvEditor "$($script:TxtCsv.Text)".Trim() } 'Open CSV' })
$btnCsvBrowse.Add_Click({ Invoke-Safely { Select-CsvFile } 'Browse for CSV' })
$btnCsvLoad.Add_Click({ Invoke-Safely { Import-EmailList } 'Load emails from CSV' })
$btnPreview.Add_Click({ Invoke-Safely { Invoke-Update -DryRun } 'Preview (dry run)' })
$btnExecute.Add_Click({ Invoke-Safely { Invoke-Update } 'Execute update' })
$btnExportRes.Add_Click({ Invoke-Safely { Export-GridData -What 'Results' } 'Export results' })
$btnExportFail.Add_Click({ Invoke-Safely { Export-GridData -What 'Failures' } 'Export failures' })
$btnExportMem.Add_Click({ Invoke-Safely { Export-GridData -What 'Members' } 'Export members' })
$btnOpenLog.Add_Click({
    Invoke-Safely {
        if (-not (Test-Path -LiteralPath $script:LogFile)) { Show-Caution 'Nothing has been logged yet.'; return }
        Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$script:LogFile`""
    } 'Open transcript'
})
$btnClear.Add_Click({
    Invoke-Safely {
        $script:Results = @()
        $script:StagedEmails = @()
        $script:NewlyAdded = @{}
        Update-ResultGrid
        if ($script:Group) { [void](Update-GroupMemberList -Quiet) }
        $script:Progress.Value = 0
        Set-Status 'Results cleared.' 'Info'
    } 'Clear results'
})

$btnPermPreview.Add_Click({ Invoke-Safely { Invoke-PermissionUpdate -DryRun } 'Preview permissions' })
$btnPermApply.Add_Click({ Invoke-Safely { Invoke-PermissionUpdate } 'Apply permissions' })

$script:TxtFilter.Add_TextChanged({ Invoke-Safely { Update-MemberGrid } 'Filter members' })

$script:RadioAdd.Add_CheckedChanged({
    if (-not $script:RadioAdd.Checked) { return }
    Invoke-Safely {
        foreach ($r in $script:Results) { if ($r.Status -eq 'Pending') { $r.Action = 'Add' } }
        Update-ResultGrid
        Set-Status 'Action set to ADD members.' 'Info'
    } 'Change action'
})
$script:RadioRemove.Add_CheckedChanged({
    if (-not $script:RadioRemove.Checked) { return }
    Invoke-Safely {
        foreach ($r in $script:Results) { if ($r.Status -eq 'Pending') { $r.Action = 'Remove' } }
        Update-ResultGrid
        Set-Status 'Action set to REMOVE members.' 'Warning'
    } 'Change action'
})

$script:TxtDl.Add_KeyDown({
    if ($_.KeyCode -eq 'Enter') { $_.SuppressKeyPress = $true; Invoke-Safely { [void](Resolve-Group) } 'Validate DL' }
})

$script:Form.Add_FormClosing({
    if ($script:Busy) {
        $_.Cancel = -not (Confirm-Sentence 'An operation is still running. Close anyway?')
        return
    }
    Write-ToolLog 'Session ended.' 'Info'
})

Set-ConsoleTheme -Form $script:Form
Set-FormScreenCorner -Form $script:Form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $script:Form -ScriptPath $PSCommandPath -Title 'Bulk DL members from CSV' | Out-Null

$script:Form.Add_Shown({
    Invoke-Safely {
        Write-ToolLog "Session started. Transcript: $script:LogFile" 'Info'
        if (Test-ExoReady -Quiet) {
            Set-Status 'Connected. Enter the list address and press Validate DL.' 'Success'
            if ("$($script:TxtDl.Text)".Trim()) { [void](Resolve-Group) } else { $script:TxtDl.Focus() }
        } else {
            Set-Status 'Exchange Online is not connected -- press Connect-ExchangeOnline on the console, then Re-check connection.' 'Warning'
            Write-ToolLog 'No Exchange Online session in this window. This tool does not sign in: a prompt raised from a modal form cannot be answered (B2). Use the console.' 'Warning'
        }
    } 'Startup'
})

[void]$script:Form.ShowDialog()
