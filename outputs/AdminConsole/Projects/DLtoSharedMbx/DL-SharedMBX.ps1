#Requires -Version 5.1
# TOOLKIT-VARIANT: Dual
#
# RUNS AGAINST ON-PREMISES for steps 1 and 2 and against EXCHANGE ONLINE for
# steps 3, 4 and 5, from ONE window in the console's own process. Nothing here
# is called unqualified: every on-premises call goes through
# Invoke-OnPremCommand and every cloud call through Invoke-EXOCommand. The
# cmdlet names in this file -- Get-DistributionGroup, Set-DistributionGroup,
# New-DistributionGroup, Get-Recipient -- exist in BOTH Exchanges and mean
# different directories, which is exactly why not one of them is bare. Connect
# both before starting; the window says which is missing.
<#
    DL-SharedMBX.ps1
    A BATCH of distribution lists become shared mailboxes, in five numbered
    steps, driven by an input CSV and recorded in one report.

    This is the many-at-once form of AdminConsole\DLtoSharedMB.ps1 (the
    one-off) and it replaces the four numbered scripts that used to live in
    this folder: 1-ONPREM-Export, 2-ONPREM-Decomm, 3-CLOUD-CheckSync and
    4-CLOUD-create. Their logic is here; their separate report files and
    hand-off-by-CSV are not.

    Author   : Glenn Armbruster
    Revised  : 2026-09-10
#>
# TOOLKIT-HELP-RUNS: steps 1, 3 and 5 read. Step 1 also hides kept groups. Steps 2 and 4 change things and both ask first.
# TOOLKIT-HELP-BEGIN
# "Turn these fourteen distribution lists into shared mailboxes."
#
# The address has to move from a group to a mailbox and the two cannot hold it
# at once. So for every row: capture everything, decommission the group
# on-premises, wait for a sync so the cloud releases the address, then build
# the mailbox and its access group in the cloud.
#
# THE INPUT FILE is DL-Input.csv next to this script. Two columns, and every
# row needs the comma:
#
#     Name,NewName
#     Jobs.ColoradoSpringsCO,
#     Payroll_F26,Payroll.ColoradoSpringsCO
#
# NewName is the name the NEW SHARED MAILBOX will take. It exists because of
# security groups: a mail-enabled security group is never deleted, only
# mail-disabled, so it keeps its own name and the mailbox has to be called
# something else. Blank NewName means the mailbox takes the list's own name,
# which is only possible when the group is going away.
#
#   1  EXPORT (on-premises)   Every attribute of every listed group -- owners,
#                             members, every proxy address, LegacyExchangeDN
#                             (the X500 that keeps replies to old mail
#                             working), Send As, restrictions, custom
#                             attributes -- into ONE report in the reports
#                             folder. Exports TWICE and compares; a directory
#                             that shifted mid-export is not safe to build on.
#                             Then HIDES from the address book every group that
#                             will be KEPT. Read-only apart from that hide.
#   2  DECOMMISSION (on-prem) Mail-disables every group. DELETES only a plain
#                             distribution group whose NewName is blank. A
#                             mail-enabled SECURITY group is never deleted --
#                             it may be securing shares or applications, and
#                             losing it would break them silently. Then WAIT
#                             about 30 minutes for a directory sync.
#   3  CHECK SYNC (cloud)     Per row, are the addresses step 4 needs actually
#                             free in Exchange Online? Asked by ADDRESS, not by
#                             object, because a kept security group's object
#                             never goes away -- it keeps syncing forever. Rows
#                             go Ready or Not ready; re-run until all Ready.
#   4  CREATE (cloud)         Per row: a shared mailbox (New-Mailbox -Shared),
#                             the old addresses as aliases, the X500,
#                             sent-items copy on for Send As and Send on
#                             Behalf; then a mail-enabled security group
#                             SMB-Access-<name> holding the old members, with
#                             Full Access and Send As over the mailbox and
#                             owners from Managers.csv. Idempotent -- anything
#                             already there is reported, not repeated.
#   5  VERIFY (cloud)         Read-only. For each row: does the mailbox exist
#                             and is it shared, how many aliases, does the
#                             access group exist, how many members, are the two
#                             permissions in place.
#
# THE REPORT IS THE STATE. Every step writes its result back into it, so the
# window can be closed during the sync wait and re-opened on the same batch
# with "Load report...". Steps 2 to 5 will not run against a row step 1 never
# exported.
#
# SELECT ROWS to work on a subset -- click one, shift-click a range, Ctrl+A
# for all, click empty space for none. With nothing selected a step runs the
# whole batch, and the label beside the step buttons says which it will be.
# This matters at step 4, where three rows out of fourteen are usually still
# waiting on a sync.
#
# A STEP BUTTON IS OUTLINED IN WHITE when it can be pressed and dimmed into the
# background when its input does not exist yet. Step 1 needs a batch; steps 2
# to 5 need at least one row that step 1 exported. Pressing a dimmed one is
# harmless -- it writes a line in the log saying what is missing.
#
# ADDRESSES ARE CARRIED AS THEY WERE wherever the tenant still accepts the
# domain; an address on a domain the tenant no longer accepts is carried on the
# primary mail domain instead, and the log says so. Step 4 of the old script
# rewrote EVERY address onto the primary domain, which silently dropped any
# address that was legitimately somewhere else.
# TOOLKIT-HELP-END

param(
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Common.ps1 lives at the TOOLKIT ROOT and there is exactly one of it. This
# walks up to find it rather than counting folders: these scripts sit three
# levels down, Set-Managers\ sits four, and the next project may sit somewhere
# else again. A Common.ps1 dropped into a project folder would be a fork
# nobody reviews.
$toolkitRoot = $PSScriptRoot
while ($toolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $toolkitRoot 'Common.ps1'))) {
    $toolkitRoot = Split-Path -Parent $toolkitRoot
}
if (-not $toolkitRoot) {
    throw "Common.ps1 was not found in any folder above '$PSScriptRoot'. This script must live under the toolkit folder."
}
. (Join-Path $toolkitRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

$script:ListSep    = ';'
$script:ReportPfx  = 'DLSharedMBX'
$script:InputName  = 'DL-Input.csv'
$script:MgrName    = 'Managers.csv'

#region Helpers ---------------------------------------------------------------
function onprem { param([string]$n, [hashtable]$p) Invoke-OnPremCommand -Name $n -Parameters $p }
function exo    { param([string]$n, [hashtable]$p) Invoke-EXOCommand    -Name $n -Parameters $p }

function ConvertTo-CsvList   { param([object]$Value) if ($null -eq $Value) { return '' }; return (@($Value | ForEach-Object { "$_" } | Where-Object { $_ }) -join $script:ListSep) }
function ConvertFrom-CsvList { param([string]$Text) if (-not "$Text".Trim()) { return @() }; return @("$Text" -split [regex]::Escape($script:ListSep) | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

function Test-OnPremReady { try { return [bool](Test-EmsAvailable).Available } catch { return $false } }
function Test-CloudReady  { try { return [bool](Test-ExoInProcess).Present }  catch { return $false } }

function Get-OnPremSmtp {
    # Recipient references (ManagedBy, AcceptMessagesOnlyFrom, ...) come back
    # as canonical names. The report stores addresses because only an address
    # survives a rename, a move, or being read on the other side of a sync.
    param([object]$Refs)
    $out = @()
    foreach ($r in @($Refs)) {
        $id = "$r"; if (-not $id) { continue }
        $x = $null
        try { $x = onprem 'Get-Recipient' @{ Identity = $id; ErrorAction = 'Stop' } } catch { }
        $out += $(if ($x -and "$($x.PrimarySmtpAddress)") { "$($x.PrimarySmtpAddress)" } else { $id })
    }
    return $out
}

function Test-SecurityGroup {
    <#
        Is this a mail-enabled SECURITY group?

        Read from the exported GroupType string, which for a mail-enabled
        security group contains 'SecurityEnabled'. This is the ONE place the
        question is answered, because a second definition -- matching on a name
        prefix, say -- could disagree with this one, and any group the two
        disagreed about would get the wrong treatment. The wrong treatment here
        is deletion.
    #>
    param($Row)
    return ("$($Row.GroupType)" -match 'SecurityEnabled')
}

function Get-AcceptedDomainList {
    # Cached for the life of the window. Used to decide whether an old address
    # can be carried as-is.
    if ($null -ne $script:AcceptedDomains) { return $script:AcceptedDomains }
    $script:AcceptedDomains = @()
    try {
        $script:AcceptedDomains = @(exo 'Get-AcceptedDomain' @{ ErrorAction = 'Stop' } | ForEach-Object { "$($_.DomainName)".ToLower() })
    } catch { }
    return $script:AcceptedDomains
}

function Resolve-CarriedAddress {
    <#
        The address to put on the new mailbox for an address the old list had.

        Carried AS IT WAS when the tenant still accepts its domain -- that is
        the whole point of carrying it, and rewriting it would break replies to
        anything sent before the migration. Only an address on a domain the
        tenant no longer accepts is moved to the primary mail domain, because
        Exchange would refuse it otherwise.

        Returns @{ Address = ...; Rewritten = $true|$false } so the caller can
        say in the log which happened.
    #>
    param([string]$Address)
    $a = "$Address".Trim()
    if ($a -notmatch '@') { return @{ Address = $a; Rewritten = $false } }
    $local, $domain = $a.Split('@')
    $accepted = Get-AcceptedDomainList
    if ($accepted.Count -eq 0 -or $accepted -contains "$domain".ToLower()) {
        return @{ Address = $a; Rewritten = $false }
    }
    return @{ Address = "$local@$(Get-MailDomain)"; Rewritten = $true }
}

function Get-AccessGroupName {
    <#
        SMB-Access-<mailbox name>, cleaned for use as an alias.

        An alias may not contain spaces or most punctuation, and Exchange caps
        it at 64. The mailbox name is operator-supplied through NewName, so
        neither can be assumed.
    #>
    param([string]$MailboxName)
    $s = ("SMB-Access-$MailboxName" -replace '[^A-Za-z0-9.\-_]', '')
    if ($s.Length -gt 64) { $s = $s.Substring(0, 64) }
    return $s.Trim('.-')
}

function Get-MailboxNameForRow {
    # NewName when there is one, otherwise the list's own name. A kept security
    # group still holds its name, which is why NewName exists at all.
    param($Row)
    $n = "$($Row.NewName)".Trim()
    if ($n) { return $n }
    return "$($Row.Name)".Trim()
}

function Test-RowBlocked {
    <#
        The one combination this tool refuses: a mail-enabled SECURITY group
        with no NewName.

        The group is kept, so it keeps its name and its alias. The mailbox
        would have to take the same name, and cannot. Rather than invent one
        and leave the operator to discover it later in the address book, the
        row is stopped here and reported. Returns the sentence, or ''.
    #>
    param($Row)
    if ((Test-SecurityGroup $Row) -and -not "$($Row.NewName)".Trim()) {
        return 'Security group with no NewName -- the group is kept and holds this name. Put a NewName in the input file.'
    }
    return ''
}

function Read-InputFile {
    <#
        DL-Input.csv, next to this script. Name,NewName -- the same file the
        four numbered scripts used, unchanged, so a half-prepared batch does
        not have to be retyped.
    #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Input file not found: $Path" }
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { throw "Input file is empty: $Path" }
    if (-not ($rows[0].PSObject.Properties.Name -contains 'Name')) {
        throw "Input file needs a 'Name' column (and a 'NewName' column, which may be blank)."
    }
    $out = @()
    foreach ($r in $rows) {
        $n = "$($r.Name)".Trim()
        if (-not $n) { continue }
        $out += [PSCustomObject]@{ Name = $n; NewName = "$($r.NewName)".Trim() }
    }
    if ($out.Count -eq 0) { throw "Input file has no usable rows: $Path" }
    return $out
}

function Read-ManagerList {
    <#
        Managers.csv, next to this script: an 'email' column. These become
        ManagedBy on every SMB-Access group the batch creates. Missing file is
        not fatal -- the groups are created ownerless and the log says so, which
        is better than refusing to build fourteen mailboxes over one CSV.
    #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $out = @()
    foreach ($m in @(Import-Csv -LiteralPath $Path)) {
        $e = "$($m.email)".Trim()
        if ($e) { $out += $e }
    }
    return @($out | Select-Object -Unique)
}

function New-ReportRow {
    <#
        One row of the report: everything about the group, then a status column
        per step. The attribute columns match AdminConsole\DLtoSharedMB.ps1's
        single-row report so the two are readable side by side and a row can be
        lifted from one to the other.
    #>
    param([Parameter(Mandatory)][string]$Name, [string]$NewName)
    $row = [ordered]@{
        Name        = $Name
        NewName     = "$NewName".Trim()
        MailboxName = ''
        DisplayName = ''
        Alias       = ''
        PrimarySmtpAddress = ''
        EmailAddresses     = ''
        LegacyExchangeDN   = ''
        GroupType          = ''
        SamAccountName     = ''
        Guid               = ''
        DistinguishedName  = ''
        OrganizationalUnit = ''
        ManagedBy          = ''
        Members            = ''
        MemberCount        = 0
        SendAs             = ''
        GrantSendOnBehalfTo = ''
        RequireSenderAuthenticationEnabled = ''
        AcceptMessagesOnlyFrom             = ''
        AcceptMessagesOnlyFromDLMembers    = ''
        RejectMessagesFrom                 = ''
        RejectMessagesFromDLMembers        = ''
        ModerationEnabled                  = ''
        ModeratedBy                        = ''
        HiddenFromAddressListsEnabled      = ''
        MailTip                            = ''
        Notes                              = ''
        WhenCreated                        = ''
    }
    for ($n = 1; $n -le 15; $n++) { $row["CustomAttribute$n"] = '' }
    for ($n = 1; $n -le 5; $n++)  { $row["ExtensionCustomAttribute$n"] = '' }
    $row['AccessGroup']  = ''
    $row['ExportedUtc']  = ''
    $row['Step1']        = ''
    $row['Step1Detail']  = ''
    $row['Hide']         = ''
    $row['Step2']        = ''
    $row['Step2Detail']  = ''
    $row['Step3']        = ''
    $row['Step3Detail']  = ''
    $row['Step4']        = ''
    $row['Step4Detail']  = ''
    $row['Step5']        = ''
    $row['Step5Detail']  = ''
    return [PSCustomObject]$row
}

function Read-GroupIntoRow {
    <#
        Fill one report row from the on-premises group. Throws when the group
        is not there, so the caller can mark the row NotFound and carry on with
        the rest of the batch -- one missing list must not stop thirteen good
        ones.
    #>
    param([Parameter(Mandatory)]$Row)
    $dg = onprem 'Get-DistributionGroup' @{ Identity = $Row.Name; ErrorAction = 'Stop' }
    $grp = $null; try { $grp = onprem 'Get-Group' @{ Identity = $dg.Identity; ErrorAction = 'Stop' } } catch { }
    $members = @(onprem 'Get-DistributionGroupMember' @{ Identity = $dg.Identity; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
                 ForEach-Object { if ("$($_.PrimarySmtpAddress)") { "$($_.PrimarySmtpAddress)" } else { "$($_.Name)" } })
    $sendAs = @()
    try {
        # On-premises Send As is an AD ACE, reported as DOMAIN\sam. SELF and the
        # well-known SIDs are Exchange's own and are not carried.
        foreach ($a in @(onprem 'Get-ADPermission' @{ Identity = $dg.DistinguishedName; ErrorAction = 'Stop' } |
                 Where-Object { -not $_.IsInherited -and $_.ExtendedRights -match 'Send-As' -and "$($_.User)" -notmatch 'NT AUTHORITY\\SELF|^S-1-5' })) {
            $sam = ("$($a.User)" -split '\\')[-1]
            $r = $null; try { $r = onprem 'Get-Recipient' @{ Identity = $sam; ErrorAction = 'Stop' } } catch { }
            $sendAs += $(if ($r) { "$($r.PrimarySmtpAddress)" } else { "$($a.User)" })
        }
    } catch { }

    $Row.DisplayName        = "$($dg.DisplayName)"
    $Row.Alias              = "$($dg.Alias)"
    $Row.PrimarySmtpAddress = "$($dg.PrimarySmtpAddress)"
    $Row.EmailAddresses     = ConvertTo-CsvList ($dg.EmailAddresses | ForEach-Object { $_.ProxyAddressString })
    $Row.LegacyExchangeDN   = "$($dg.LegacyExchangeDN)"
    $Row.GroupType          = "$($dg.GroupType)"
    $Row.SamAccountName     = "$($dg.SamAccountName)"
    $Row.Guid               = "$($dg.Guid)"
    $Row.DistinguishedName  = "$($dg.DistinguishedName)"
    $Row.OrganizationalUnit = "$($dg.OrganizationalUnit)"
    $Row.ManagedBy          = ConvertTo-CsvList (Get-OnPremSmtp $dg.ManagedBy)
    $Row.Members            = ConvertTo-CsvList $members
    $Row.MemberCount        = $members.Count
    $Row.SendAs             = ConvertTo-CsvList $sendAs
    $Row.GrantSendOnBehalfTo = ConvertTo-CsvList (Get-OnPremSmtp $dg.GrantSendOnBehalfTo)
    $Row.RequireSenderAuthenticationEnabled = "$($dg.RequireSenderAuthenticationEnabled)"
    $Row.AcceptMessagesOnlyFrom             = ConvertTo-CsvList (Get-OnPremSmtp $dg.AcceptMessagesOnlyFrom)
    $Row.AcceptMessagesOnlyFromDLMembers    = ConvertTo-CsvList (Get-OnPremSmtp $dg.AcceptMessagesOnlyFromDLMembers)
    $Row.RejectMessagesFrom                 = ConvertTo-CsvList (Get-OnPremSmtp $dg.RejectMessagesFrom)
    $Row.RejectMessagesFromDLMembers        = ConvertTo-CsvList (Get-OnPremSmtp $dg.RejectMessagesFromDLMembers)
    $Row.ModerationEnabled                  = "$($dg.ModerationEnabled)"
    $Row.ModeratedBy                        = ConvertTo-CsvList (Get-OnPremSmtp $dg.ModeratedBy)
    $Row.HiddenFromAddressListsEnabled      = "$($dg.HiddenFromAddressListsEnabled)"
    $Row.MailTip                            = "$($dg.MailTip)"
    $Row.Notes                              = $(if ($grp) { "$($grp.Notes)" } else { '' })
    $Row.WhenCreated                        = "$($dg.WhenCreated)"
    for ($n = 1; $n -le 15; $n++) { $Row."CustomAttribute$n" = "$($dg."CustomAttribute$n")" }
    for ($n = 1; $n -le 5; $n++)  { $Row."ExtensionCustomAttribute$n" = ConvertTo-CsvList $dg."ExtensionCustomAttribute$n" }
    $Row.MailboxName = Get-MailboxNameForRow $Row
    $Row.AccessGroup = Get-AccessGroupName -MailboxName $Row.MailboxName
    $Row.ExportedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
}

function Get-RowFingerprint {
    <#
        The part of a row that a second read must agree with.

        The two exports exist to prove the directory is not shifting underneath
        the run. Comparing whole CSV files (what the old step 1 did) also
        compares ExportedUtc, which differs by construction -- so the check
        either always failed or the timestamp had to be left out of the export.
        Comparing the attributes that matter is the same test without that
        problem.
    #>
    param($Row)
    return (@($Row.PrimarySmtpAddress, $Row.EmailAddresses, $Row.Members, $Row.ManagedBy,
              $Row.SendAs, $Row.GrantSendOnBehalfTo, $Row.LegacyExchangeDN, $Row.GroupType) -join '|')
}

function Save-Report {
    param([Parameter(Mandatory)][object[]]$Rows, [Parameter(Mandatory)][string]$Path)
    $Rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Find-NewestReport {
    $dir = Get-ReportPath
    return (Get-ChildItem -LiteralPath $dir -Filter "$($script:ReportPfx)_*.csv" -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1)
}
#endregion

#region Form ------------------------------------------------------------------
function Show-BatchForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'DL to Shared Mailbox (batch)'
    # WIDE ENOUGH FOR EVERY GRID COLUMN AT LAUNCH.
    #
    # $script:GridW is the sum of the column widths below plus room for the
    # vertical scrollbar, so the two cannot drift apart: add a column and this
    # number moves with it. A window that opens too narrow to show the status
    # columns is a window whose whole point is hidden until somebody drags it.
    #
    # Clamped to the working area at the end of this function, because a
    # 1280-wide form on a 1366-wide laptop is not an improvement.
    $script:GridW = 1221
    # ONE PLACE FOR THE THREE STATUS COLOURS, and they are bright ones.
    #
    # These sit on the theme's own background, which on the Console 2 themes is
    # nearly black. The muted 40,140,80 green that reads as "quietly fine" on a
    # white form is barely legible on a dark one, and a status colour that has
    # to be squinted at is not a status colour. Bright enough to carry on dark,
    # still clearly green/red/amber on light.
    $script:OkColor   = [System.Drawing.Color]::FromArgb(90, 220, 140)
    $script:BadColor  = [System.Drawing.Color]::FromArgb(245, 115, 115)
    $script:WarnColor = [System.Drawing.Color]::FromArgb(240, 175, 60)
    $script:ContentW = $script:GridW
    $form.ClientSize = New-Object System.Drawing.Size(($script:ContentW + 24), 760)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 600)

    $y = 12
    $lblConn = New-Object System.Windows.Forms.Label
    $lblConn.Location = New-Object System.Drawing.Point(12, $y); $lblConn.Size = New-Object System.Drawing.Size($script:ContentW, 18); $lblConn.Tag = 'NoTheme'
    $form.Controls.Add($lblConn)
    $y += 24

    # ---- input + report -----------------------------------------------------
    $grpIn = New-Object System.Windows.Forms.GroupBox
    $grpIn.Text = 'Batch'
    $grpIn.Location = New-Object System.Drawing.Point(12, $y); $grpIn.Size = New-Object System.Drawing.Size($script:ContentW, 92)
    $form.Controls.Add($grpIn)

    $btnEditIn = New-Object System.Windows.Forms.Button; $btnEditIn.Text = 'Edit input...'; $btnEditIn.Location = New-Object System.Drawing.Point(12, 22); $btnEditIn.Size = New-Object System.Drawing.Size(120, 28); $btnEditIn.Tag = 'Plain'; $grpIn.Controls.Add($btnEditIn)
    $btnLoadIn = New-Object System.Windows.Forms.Button; $btnLoadIn.Text = 'Start new batch'; $btnLoadIn.Location = New-Object System.Drawing.Point(140, 22); $btnLoadIn.Size = New-Object System.Drawing.Size(130, 28); $grpIn.Controls.Add($btnLoadIn)
    $btnLoadRep = New-Object System.Windows.Forms.Button; $btnLoadRep.Text = 'Load report...'; $btnLoadRep.Location = New-Object System.Drawing.Point(278, 22); $btnLoadRep.Size = New-Object System.Drawing.Size(130, 28); $btnLoadRep.Tag = 'Plain'; $grpIn.Controls.Add($btnLoadRep)
    $btnOpenRep = New-Object System.Windows.Forms.Button; $btnOpenRep.Text = 'Open report'; $btnOpenRep.Location = New-Object System.Drawing.Point(416, 22); $btnOpenRep.Size = New-Object System.Drawing.Size(110, 28); $btnOpenRep.Tag = 'Plain'; $btnOpenRep.Enabled = $false; $grpIn.Controls.Add($btnOpenRep)

    $lblRep = New-Object System.Windows.Forms.Label
    $lblRep.Location = New-Object System.Drawing.Point(534, 28); $lblRep.Size = New-Object System.Drawing.Size(($script:ContentW - 546), 18)
    $lblRep.ForeColor = [System.Drawing.Color]::DimGray; $lblRep.Text = '(no batch loaded)'; $grpIn.Controls.Add($lblRep)

    $lblCounts = New-Object System.Windows.Forms.Label
    $lblCounts.Location = New-Object System.Drawing.Point(12, 60); $lblCounts.Size = New-Object System.Drawing.Size(($script:ContentW - 24), 18)
    $lblCounts.ForeColor = [System.Drawing.Color]::DimGray; $grpIn.Controls.Add($lblCounts)
    $y += 100

    # ---- steps --------------------------------------------------------------
    $grpSteps = New-Object System.Windows.Forms.GroupBox
    $grpSteps.Text = 'Steps   (select rows to work on a subset; nothing selected means the whole batch)'
    $grpSteps.Location = New-Object System.Drawing.Point(12, $y); $grpSteps.Size = New-Object System.Drawing.Size($script:ContentW, 92)
    $form.Controls.Add($grpSteps)

    # NO 'Plain' TAGS ON THESE FIVE, and NONE OF THEM IS EVER .Enabled = $false.
    #
    # A disabled WinForms button draws its caption in the system's grey no
    # matter what ForeColor says -- and the system's grey is a near-black on a
    # dark theme, which is how "2 Decommission" ended up unreadable. There is
    # no property that fixes that; the colour is chosen inside the control.
    #
    # So availability is tracked in $form.StepOn instead of in .Enabled. The
    # button stays enabled, $styleSteps paints it (light grey caption, dimmed
    # face, no outline) and the handler turns the click into a sentence saying
    # what is missing. That is strictly better than a dead button: "why can't I
    # press this?" is answered instead of ignored.
    $btn1 = New-Object System.Windows.Forms.Button; $btn1.Text = "1  Export`r`n(on-prem)";      $btn1.Location = New-Object System.Drawing.Point(12, 20);  $btn1.Size = New-Object System.Drawing.Size(150, 40); $grpSteps.Controls.Add($btn1)
    $btn2 = New-Object System.Windows.Forms.Button; $btn2.Text = "2  Decommission`r`n(on-prem)"; $btn2.Location = New-Object System.Drawing.Point(170, 20); $btn2.Size = New-Object System.Drawing.Size(150, 40); $grpSteps.Controls.Add($btn2)
    $btn3 = New-Object System.Windows.Forms.Button; $btn3.Text = "3  Check sync`r`n(cloud)";     $btn3.Location = New-Object System.Drawing.Point(328, 20); $btn3.Size = New-Object System.Drawing.Size(150, 40); $grpSteps.Controls.Add($btn3)
    $btn4 = New-Object System.Windows.Forms.Button; $btn4.Text = "4  Create`r`n(cloud)";         $btn4.Location = New-Object System.Drawing.Point(486, 20); $btn4.Size = New-Object System.Drawing.Size(150, 40); $grpSteps.Controls.Add($btn4)
    $btn5 = New-Object System.Windows.Forms.Button; $btn5.Text = "5  Verify`r`n(cloud)";         $btn5.Location = New-Object System.Drawing.Point(644, 20); $btn5.Size = New-Object System.Drawing.Size(150, 40); $grpSteps.Controls.Add($btn5)

    $lblSel = New-Object System.Windows.Forms.Label
    $lblSel.Location = New-Object System.Drawing.Point(810, 32); $lblSel.Size = New-Object System.Drawing.Size(($script:ContentW - 822), 18)
    $lblSel.ForeColor = [System.Drawing.Color]::DimGray
    # Blank when nothing is selected. "Whole batch" was restating the caption
    # on the group box directly above it, and a label that is always there is a
    # label nobody reads when it finally says something.
    $lblSel.Text = ''
    $grpSteps.Controls.Add($lblSel)
    $y += 100

    # ---- grid ---------------------------------------------------------------
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, $y)
    $grid.Size = New-Object System.Drawing.Size($script:GridW, 280)
    $grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.ReadOnly = $true
    $grid.RowHeadersVisible = $false
    # SELECTION IS THE SUBSET, AND THERE IS NO TICK COLUMN.
    #
    # A checkbox column plus Tick all / Untick all was three controls doing
    # what the grid's own selection already does: click a row, shift-click a
    # range, Ctrl+A for all, click away for none. It also had a trap -- a
    # DataGridView checkbox does not commit its value until the cell loses
    # focus, so ticking a row and then pressing a step button read the tick as
    # it was BEFORE the click. Selection has no such delay.
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $true
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $form.Controls.Add($grid)

    # Widths total $script:GridW less ~20px for the vertical scrollbar. Change
    # one and change that constant, or the window opens too narrow again.
    foreach ($c in @(
        @{ N = 'Name';    H = 'Distribution list'; W = 210 },
        @{ N = 'Type';    H = 'Type';              W = 95  },
        @{ N = 'Mailbox'; H = 'New mailbox';       W = 200 },
        @{ N = 'Address'; H = 'Address';           W = 220 },
        @{ N = 'Members'; H = '# Members';         W = 80  },
        @{ N = 'S1';      H = '1 Export';          W = 78  },
        @{ N = 'S2';      H = '2 Decomm';          W = 88  },
        @{ N = 'S3';      H = '3 Sync';            W = 78  },
        @{ N = 'S4';      H = '4 Create';          W = 78  },
        @{ N = 'S5';      H = '5 Verify';          W = 78  })) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.N; $col.HeaderText = $c.H; $col.Width = $c.W; $col.ReadOnly = $true
        [void]$grid.Columns.Add($col)
    }
    $y += 288

    # ---- log ----------------------------------------------------------------
    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline = $true; $txtOut.ReadOnly = $true; $txtOut.ScrollBars = 'Vertical'
    $txtOut.Location = New-Object System.Drawing.Point(12, $y); $txtOut.Size = New-Object System.Drawing.Size($script:ContentW, 190)
    $txtOut.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor
                     [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)
    $y += 198

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(($script:ContentW - 88), $y); $btnClose.Size = New-Object System.Drawing.Size(100, 28)
    $btnClose.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnClose.Tag = 'Plain'
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    # STATE LIVES ON THE FORM, not in $script:. These handlers run while the
    # console (or another tool) may be the current scope, and a $script: read
    # from a scriptblock resolves against the nearest script scope AT RUN TIME
    # rather than the one it was written in -- see ARCHITECTURE on the status
    # strip for what that costs when it goes wrong.
    $form | Add-Member -NotePropertyName 'Busy'       -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'Rows'       -NotePropertyValue @()    -Force
    $form | Add-Member -NotePropertyName 'ReportPath' -NotePropertyValue ''     -Force

    $log = { param([string]$s) $txtOut.AppendText("$s`r`n") }

    # ---- how a step button says whether it can be pressed --------------------
    #
    # AVAILABLE: the theme's own button colour with a 2px WHITE outline.
    # NOT AVAILABLE YET: the same colour blended most of the way toward the
    # form's background, with no outline -- so it recedes into the window
    # instead of standing out from it.
    #
    # BLENDED, NOT A FIXED GREY, AND THAT IS THE POINT. A literal light grey
    # reads as "off" on a light theme and as the BRIGHTEST thing on the window
    # on a dark one, which is the opposite of what a disabled control should
    # do. Blending toward whatever the form is actually painted keeps "dimmer
    # than its neighbours" true on every theme, including ones added later.
    #
    # The unblended colour is captured ONCE per button the first time this
    # runs. Re-blending an already-blended colour would fade the buttons a
    # little further on every repaint until they vanished.
    $steps = @($btn1, $btn2, $btn3, $btn4, $btn5)
    $form | Add-Member -NotePropertyName 'StepBase' -NotePropertyValue @{} -Force
    $form | Add-Member -NotePropertyName 'StepOn'   -NotePropertyValue @{} -Force

    $blend = {
        param($From, $To, [double]$Amount)
        [System.Drawing.Color]::FromArgb(
            [int]($From.R + (($To.R - $From.R) * $Amount)),
            [int]($From.G + (($To.G - $From.G) * $Amount)),
            [int]($From.B + (($To.B - $From.B) * $Amount)))
    }

    $styleSteps = {
        foreach ($b in $steps) {
            $key = "$($b.Text)"
            if (-not $form.StepBase.ContainsKey($key)) {
                $form.StepBase[$key] = @{ Back = $b.BackColor; Fore = $b.ForeColor }
            }
            $base = $form.StepBase[$key]
            $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $b.UseVisualStyleBackColor = $false
            if ($form.StepOn[$key]) {
                $b.BackColor = $base.Back
                $b.ForeColor = $base.Fore
                $b.FlatAppearance.BorderSize = 2
                $b.FlatAppearance.BorderColor = [System.Drawing.Color]::White
            } else {
                # Face dimmed toward the window; caption a LIGHT grey that is
                # legible on either. This is the whole reason these buttons are
                # never actually disabled -- .NET would override the caption
                # with its own near-black grey and there is no way to stop it.
                $b.BackColor = & $blend $base.Back $form.BackColor 0.72
                $b.ForeColor = [System.Drawing.Color]::FromArgb(170, 175, 185)
                $b.FlatAppearance.BorderSize = 1
                $b.FlatAppearance.BorderColor = & $blend $base.Back $form.BackColor 0.45
            }
        }
    }

    # Says why a step cannot run yet, and returns $false so the handler stops.
    # Called from the click, not from the paint, because the answer is only
    # interesting to somebody who just tried.
    $stepReady = {
        param($Button, [string]$Why)
        if ($form.StepOn["$($Button.Text)"]) { return $true }
        & $log $Why
        return $false
    }

    $refreshConn = {
        $op = Test-OnPremReady; $cl = Test-CloudReady
        $lblConn.Text = "On-premises: $(if ($op) { 'connected' } else { 'NOT connected -- steps 1 and 2 need it' })     Cloud: $(if ($cl) { 'connected' } else { 'NOT connected -- steps 3, 4 and 5 need it' })"
        $lblConn.ForeColor = $(if ($op -and $cl) { $script:OkColor } else { $script:WarnColor })
    }

    $paint = {
        # Redraw the grid from the rows. The SELECTION is preserved by name,
        # because a step that ran on a subset must not throw that subset away
        # -- the operator's next move is usually another step on the same rows.
        $sel = @{}
        foreach ($gr in $grid.SelectedRows) { $sel["$($gr.Cells['Name'].Value)"] = $true }
        $grid.Rows.Clear()
        foreach ($r in $form.Rows) {
            $type = $(if (-not "$($r.GroupType)") { '' } elseif (Test-SecurityGroup $r) { 'Security' } else { 'Distribution' })
            $i = $grid.Rows.Add()
            $gr = $grid.Rows[$i]
            $gr.Cells['Name'].Value    = "$($r.Name)"
            $gr.Cells['Type'].Value    = $type
            $gr.Cells['Mailbox'].Value = "$($r.MailboxName)"
            $gr.Cells['Address'].Value = "$($r.PrimarySmtpAddress)"
            $gr.Cells['Members'].Value = "$($r.MemberCount)"
            foreach ($n in 1, 2, 3, 4, 5) {
                $v = "$($r."Step$n")"
                $gr.Cells["S$n"].Value = $v
                if ($v -match '^(Failed|Blocked|NotFound|Not ready)') { $gr.Cells["S$n"].Style.ForeColor = $script:BadColor }
                elseif ($v -match '^(Done|Ready|OK|Kept|Deleted|Created|Exists)') { $gr.Cells["S$n"].Style.ForeColor = $script:OkColor }
            }
        }
        $n = @($form.Rows).Count
        $sec = @($form.Rows | Where-Object { Test-SecurityGroup $_ }).Count
        $blocked = @($form.Rows | Where-Object { Test-RowBlocked $_ }).Count
        $lblCounts.Text = "$n list(s): $sec security (kept, mail-disabled only), $($n - $sec) distribution$(if ($blocked) { "   --   $blocked BLOCKED, see the log" })"
        # A STEP IS AVAILABLE ONLY WHEN ITS INPUT EXISTS, so the row of buttons
        # reads as the running order rather than five things you could press in
        # any order. Step 1 needs a batch; 2 to 5 need at least one row that
        # step 1 actually exported. Everything past that is checked in the
        # handler -- greying step 4 until every row is Ready would block the
        # normal case of creating the eleven that synced.
        $any = ($n -gt 0)
        $exported = @($form.Rows | Where-Object { $_.Step1 -eq 'Done' }).Count -gt 0
        $form.StepOn["$($btn1.Text)"] = $any
        foreach ($b in @($btn2, $btn3, $btn4, $btn5)) { $form.StepOn["$($b.Text)"] = $exported }
        $btnOpenRep.Enabled = [bool]"$($form.ReportPath)"
        & $styleSteps
        # Put the selection back -- or make sure there ISN'T one.
        #
        # A DataGridView selects its first row the moment rows are added. Left
        # alone, that would mean every repaint silently narrowed the next step
        # to row 1 while the label said so in small grey text beside the
        # buttons. Clearing CurrentCell first is what makes ClearSelection
        # stick; on its own the grid puts the selection straight back.
        $grid.CurrentCell = $null
        $grid.ClearSelection()
        if ($sel.Count) {
            foreach ($gr in $grid.Rows) { if ($sel["$($gr.Cells['Name'].Value)"]) { $gr.Selected = $true } }
        }
        & $showSel
    }

    $showSel = {
        $c = @($grid.SelectedRows).Count
        $lblSel.Text = $(if ($c -eq 0) { '' } else { "$c row(s) selected" })
    }
    $grid.Add_SelectionChanged({ & $showSel })

    $pick = {
        # The rows a step should act on: the selected ones, or all of them.
        $names = @()
        foreach ($gr in $grid.SelectedRows) { $names += "$($gr.Cells['Name'].Value)" }
        if ($names.Count -eq 0) { return @($form.Rows) }
        return @($form.Rows | Where-Object { $names -contains "$($_.Name)" })
    }

    $save = {
        if (-not "$($form.ReportPath)") { return }
        try { Save-Report -Rows @($form.Rows) -Path $form.ReportPath }
        catch { & $log "Report could not be written: $($_.Exception.Message)" }
    }

    # ---- edit input ----
    $btnEditIn.Add_Click({
        try {
            $p = Join-Path $PSScriptRoot $script:InputName
            if (-not (Test-Path -LiteralPath $p)) {
                'Name,NewName' | Set-Content -LiteralPath $p -Encoding UTF8
                & $log "Created an empty $($script:InputName)."
            }
            Invoke-Item -LiteralPath $p
            & $log "Opened $p -- save it, then press 'Start new batch'."
        } catch { Show-HandlerError -ErrorRecord $_ -Context 'Open the input file' }
    })

    # ---- start new batch ----
    $btnLoadIn.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $p = Join-Path $PSScriptRoot $script:InputName
            $in = Read-InputFile -Path $p
            $rows = @()
            foreach ($r in $in) { $rows += (New-ReportRow -Name $r.Name -NewName $r.NewName) }
            foreach ($r in $rows) { $r.MailboxName = Get-MailboxNameForRow $r }
            $form.Rows = $rows
            $form.ReportPath = Join-Path (Get-ReportPath) "$($script:ReportPfx)_$(Get-Date -Format 'yyyy-MM-dd_HHmm').csv"
            & $save
            $lblRep.Text = [System.IO.Path]::GetFileName($form.ReportPath)
            & $paint
            & $log "Batch of $($rows.Count) from $($script:InputName). Report: $($form.ReportPath)"
            & $log 'Nothing has been read from the directory yet -- run step 1.'
        } catch { & $log "Could not start the batch: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Read the input file' }
        finally { $form.Busy = $false }
    })

    # ---- load report ----
    $btnLoadRep.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.InitialDirectory = Get-ReportPath
            $dlg.Filter = "Batch reports ($($script:ReportPfx)_*.csv)|$($script:ReportPfx)_*.csv|All CSV|*.csv"
            $newest = Find-NewestReport
            if ($newest) { $dlg.FileName = $newest.Name }
            if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $rows = @(Import-Csv -LiteralPath $dlg.FileName)
            if ($rows.Count -eq 0) { & $log 'That report has no rows.'; return }
            $form.Rows = $rows
            $form.ReportPath = $dlg.FileName
            $lblRep.Text = [System.IO.Path]::GetFileName($form.ReportPath)
            & $paint
            & $log "Loaded $($rows.Count) row(s) from $($form.ReportPath)"
        } catch { & $log "Load failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Load report' }
        finally { $form.Busy = $false }
    })

    $btnOpenRep.Add_Click({
        try { if ("$($form.ReportPath)" -and (Test-Path -LiteralPath $form.ReportPath)) { Invoke-Item -LiteralPath $form.ReportPath } }
        catch { Show-HandlerError -ErrorRecord $_ -Context 'Open report' }
    })

    # ---- 1 export -----------------------------------------------------------
    $btn1.Add_Click({
        if ($form.Busy) { return }
        if (-not (& $stepReady $btn1 "No batch loaded. Press 'Start new batch' to read $($script:InputName), or 'Load report...' to carry on with one.")) { return }
        $form.Busy = $true
        try {
            if (-not (Test-OnPremReady)) { & $log 'On-premises Exchange is not connected in this window.'; return }
            $rows = & $pick
            if ($rows.Count -eq 0) { & $log 'No rows.'; return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            # FIRST PASS.
            $ok = 0; $bad = 0
            $first = @{}
            foreach ($r in $rows) {
                Write-ToolStep "Exporting $($r.Name) ..."
                try {
                    Read-GroupIntoRow -Row $r
                    $first["$($r.Name)"] = Get-RowFingerprint $r
                    $r.Step1 = 'Done'; $r.Step1Detail = "$($r.MemberCount) member(s)"
                    $ok++
                } catch {
                    $r.Step1 = 'NotFound'; $r.Step1Detail = "$($_.Exception.Message)"
                    & $log "  $($r.Name) -- NOT read: $($_.Exception.Message)"
                    $bad++
                }
            }
            & $log "Read $ok of $($rows.Count) list(s)$(if ($bad) { "; $bad could not be read" })."

            # SECOND PASS, AND THE COMPARISON.
            #
            # Two reads of the same directory a moment apart. If they disagree,
            # something else is writing to these groups right now and the export
            # is not a safe basis for step 2 -- which deletes. The old step 1
            # compared two whole CSV files; this compares the attributes that
            # matter, so the export timestamp cannot fail the check on its own.
            Write-ToolStep 'Second read for the consistency check ...'
            # The failures are collected as the ROW OBJECTS, not as sentences
            # about them. Parsing a name back out of a message would break on
            # the first list with a space in its name, and marking the wrong row
            # Failed here is worse than not checking at all.
            $moved = @()
            foreach ($r in @($rows | Where-Object { $_.Step1 -eq 'Done' })) {
                $probe = New-ReportRow -Name $r.Name -NewName $r.NewName
                try { Read-GroupIntoRow -Row $probe }
                catch { $moved += [PSCustomObject]@{ Row = $r; Why = 'vanished between reads' }; continue }
                if ((Get-RowFingerprint $probe) -ne $first["$($r.Name)"]) {
                    $moved += [PSCustomObject]@{ Row = $r; Why = 'changed between reads' }
                }
            }
            if ($moved.Count) {
                foreach ($m in $moved) { $m.Row.Step1 = 'Failed'; $m.Row.Step1Detail = $m.Why }
                & $log ''
                & $log "CONSISTENCY CHECK FAILED for $($moved.Count) list(s):"
                foreach ($m in $moved) { & $log "  $($m.Row.Name) -- $($m.Why)" }
                & $log 'Something else is writing to these groups. Do NOT run step 2 on them.'
            } else {
                & $log 'Consistency check passed: two reads agree.'
            }

            # THE HIDE, AND WHY IT IS HERE AND NOT IN STEP 2.
            #
            # Mail-disabling does NOT remove the cloud copy of a synced group.
            # The AD group keeps syncing and Exchange Online falls back to the
            # @...onmicrosoft.com address it appended at first sync, leaving the
            # list in the address book as a recipient that accepts mail and
            # silently drops it. Hiding first, and letting a sync carry the
            # hide, avoids that.
            #
            # ONLY GROUPS THAT WILL BE KEPT are hidden -- a group about to be
            # deleted has no such afterlife. That is the same test step 2 uses,
            # asked once, in Test-SecurityGroup.
            $keep = @($rows | Where-Object { $_.Step1 -eq 'Done' -and (Test-SecurityGroup $_) })
            if ($keep.Count -eq 0) {
                & $log 'No security groups in this batch, so nothing to hide.'
            } else {
                & $log ''
                & $log "Hiding $($keep.Count) kept group(s) from the address book ..."
                foreach ($r in $keep) {
                    Write-ToolStep "Hiding $($r.Name) ..."
                    if ("$($r.HiddenFromAddressListsEnabled)" -eq 'True') { $r.Hide = 'AlreadyHidden'; & $log "  $($r.Name) -- already hidden."; continue }
                    try {
                        # Learn (Set-DistributionGroup): -Identity,
                        # -HiddenFromAddressListsEnabled <$true|$false>.
                        #
                        # Bound on the Guid, and QUOTED. A name is a display
                        # value and can be ambiguous; a Guid cannot. The quotes
                        # are not cosmetic -- $dg.Guid is a [System.Guid] and
                        # DistributionGroupIdParameter refuses to convert from
                        # that type ("Cannot process argument transformation"),
                        # while it takes the string form happily.
                        onprem 'Set-DistributionGroup' @{ Identity = "$($r.Guid)"; HiddenFromAddressListsEnabled = $true; BypassSecurityGroupManagerCheck = $true; ErrorAction = 'Stop' } | Out-Null
                        $r.Hide = 'Hidden'; $r.HiddenFromAddressListsEnabled = 'True'
                        & $log "  $($r.Name) -- hidden."
                        Write-CompletedLog -Action 'DLBatchHide' -Target (Get-LogTarget $r.PrimarySmtpAddress $r.Name) -Result 'Success' -Detail 'HiddenFromAddressListsEnabled = True before decommission'
                    } catch {
                        $r.Hide = "Failed: $($_.Exception.Message)"
                        & $log "  $($r.Name) -- HIDE FAILED: $($_.Exception.Message)"
                        Write-FailedLog -Action 'DLBatchHide' -Target (Get-LogTarget $r.PrimarySmtpAddress $r.Name) -ErrorRecord $_
                    }
                }
            }

            # Blocked rows are named now rather than at step 2, so the input
            # file can be fixed while the sync wait has not started.
            foreach ($r in @($rows | Where-Object { $_.Step1 -eq 'Done' })) {
                $why = Test-RowBlocked $r
                if ($why) { & $log "  BLOCKED  $($r.Name) -- $why" }
            }

            & $save
            & $paint
            & $log ''
            & $log 'Step 1 done. Review the report, then run step 2.'
            Write-CompletedLog -Action 'DLBatchExport' -Target (Get-LogTarget '' 'batch') -Result $(if ($bad -or $moved.Count) { 'Partial' } else { 'Success' }) `
                -Detail "$ok exported, $bad not found, $($moved.Count) inconsistent; report $([System.IO.Path]::GetFileName($form.ReportPath))"
        } catch { & $log "Step 1 failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Export the batch (on-premises)' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    # ---- 2 decommission -----------------------------------------------------
    $btn2.Add_Click({
        if ($form.Busy) { return }
        if (-not (& $stepReady $btn2 'Step 2 needs step 1 first: nothing has been exported, so there is no record of what these groups contain.')) { return }
        $form.Busy = $true
        try {
            if (-not (Test-OnPremReady)) { & $log 'On-premises Exchange is not connected in this window.'; return }
            $rows = @((& $pick) | Where-Object { $_.Step1 -eq 'Done' })
            if ($rows.Count -eq 0) { & $log 'No rows with a completed step 1. Export first.'; return }

            # WHAT "DECOMMISSION" MEANS DEPENDS ENTIRELY ON THE GROUP'S TYPE.
            #
            # A plain distribution group exists only to carry mail. With the
            # address moving to a mailbox it has no reason to exist, so it is
            # deleted -- unless it has a NewName, which says keep it.
            #
            # A mail-enabled SECURITY group is NEVER deleted. It may also be
            # securing file shares, applications or other permissions, and
            # deleting it would break those silently and unrecoverably. It is
            # only mail-disabled: the address is freed, the AD group and
            # everything it secures stay exactly as they were.
            $del  = @($rows | Where-Object { -not (Test-SecurityGroup $_) -and -not "$($_.NewName)".Trim() })
            $keep = @($rows | Where-Object { (Test-SecurityGroup $_) -or "$($_.NewName)".Trim() })

            $a = [System.Windows.Forms.MessageBox]::Show(
                ("Decommission $($rows.Count) group(s) on-premises?" +
                 "`r`n`r`nMAIL-DISABLE AND KEEP: $($keep.Count)" +
                 "`r`n  every security group, plus any row with a NewName" +
                 "`r`n`r`nMAIL-DISABLE AND DELETE: $($del.Count)" +
                 "`r`n  $(if ($del.Count) { (@($del | ForEach-Object { $_.Name }) -join ', ') } else { '(none)' })" +
                 "`r`n`r`nDeletion cannot be undone. The report is the only record of what those groups contained:" +
                 "`r`n  $($form.ReportPath)" +
                 "`r`n`r`nAfterwards, wait about 30 minutes for a directory sync, then run step 3."),
                'Confirm decommission', 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { & $log 'Step 2 cancelled.'; return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            $done = 0; $fail = 0
            foreach ($r in $rows) {
                Write-ToolStep "Decommissioning $($r.Name) ..."
                $why = Test-RowBlocked $r
                if ($why) { $r.Step2 = 'Blocked'; $r.Step2Detail = $why; & $log "  BLOCKED  $($r.Name) -- $why"; $fail++; continue }
                $isSec  = Test-SecurityGroup $r
                $delete = (-not $isSec) -and (-not "$($r.NewName)".Trim())
                try {
                    # Learn (Disable-DistributionGroup): -Identity; removes the
                    # Exchange attributes and leaves the AD group in place.
                    # -BypassSecurityGroupManagerCheck because without it
                    # Exchange refuses anyone who is not in the group's
                    # ManagedBy, administrator or not.
                    onprem 'Disable-DistributionGroup' @{ Identity = "$($r.Guid)"; BypassSecurityGroupManagerCheck = $true; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                    if ($delete) {
                        # Learn (Remove-DistributionGroup): -Identity;
                        # -Confirm:$false; -BypassSecurityGroupManagerCheck.
                        onprem 'Remove-DistributionGroup' @{ Identity = "$($r.Guid)"; BypassSecurityGroupManagerCheck = $true; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        $r.Step2 = 'Deleted'; $r.Step2Detail = 'Mail-disabled, then removed'
                        & $log "  $($r.Name) -- mail-disabled and DELETED."
                        Write-CompletedLog -Action 'DLBatchDelete' -Target (Get-LogTarget $r.PrimarySmtpAddress $r.Name) -Result 'Success' -Detail "Remove-DistributionGroup; report $([System.IO.Path]::GetFileName($form.ReportPath))"
                    } else {
                        $r.Step2 = 'Kept'
                        $r.Step2Detail = $(if ($isSec) { 'Mail-disabled; security group kept' } else { 'Mail-disabled; kept (has a NewName)' })
                        & $log "  $($r.Name) -- mail-disabled, KEPT ($(if ($isSec) { 'security group' } else { 'has a NewName' }))."
                        Write-CompletedLog -Action 'DLBatchDisable' -Target (Get-LogTarget $r.PrimarySmtpAddress $r.Name) -Result 'Success' -Detail "Disable-DistributionGroup only; $($r.Step2Detail); report $([System.IO.Path]::GetFileName($form.ReportPath))"
                    }
                    $done++
                } catch {
                    $r.Step2 = 'Failed'; $r.Step2Detail = "$($_.Exception.Message)"
                    & $log "  $($r.Name) -- FAILED: $($_.Exception.Message)"
                    Write-FailedLog -Action 'DLBatchDecomm' -Target (Get-LogTarget $r.PrimarySmtpAddress $r.Name) -ErrorRecord $_
                    $fail++
                }
            }
            & $save
            & $paint
            & $log ''
            & $log "Step 2 done: $done decommissioned, $fail failed. WAIT about 30 minutes for a directory sync, then run step 3."
        } catch { & $log "Step 2 failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Decommission on-premises groups' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    # ---- 3 check sync -------------------------------------------------------
    $btn3.Add_Click({
        if ($form.Busy) { return }
        if (-not (& $stepReady $btn3 'Step 3 needs step 1 first: it checks the addresses the export recorded.')) { return }
        $form.Busy = $true
        try {
            if (-not (Test-CloudReady)) { & $log 'Exchange Online is not connected in this window.'; return }
            $rows = @((& $pick) | Where-Object { $_.Step1 -eq 'Done' })
            if ($rows.Count -eq 0) { & $log 'No rows with a completed step 1.'; return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            # ASKED BY ADDRESS, NOT BY OBJECT, AND THAT IS THE WHOLE POINT.
            #
            # "Has the object gone?" is the right question only for a group
            # that was deleted. A KEPT security group's object never goes away
            # -- it keeps syncing for ever -- so that question answered STILL
            # EXISTS on every run, permanently, on exactly the rows that carry
            # the listserv risk. A red light that must always be ignored is the
            # same as no light at all.
            #
            # What step 4 actually needs is the ADDRESSES. Mail-disabling frees
            # them while leaving the object in place, so they are usually free
            # well before the object would have vanished -- and every minute
            # between step 2 and step 4 is a minute those addresses bounce.
            $ready = 0; $notReady = 0
            foreach ($r in $rows) {
                Write-ToolStep "Checking $($r.Name) ..."
                $wanted = @()
                $mbName = Get-MailboxNameForRow $r
                if ("$($r.NewName)".Trim()) { $wanted += "$("$($r.NewName)".Trim())@$(Get-MailDomain)" }
                foreach ($e in (ConvertFrom-CsvList $r.EmailAddresses)) {
                    if ($e -match '^(?i)smtp:') { $wanted += (Resolve-CarriedAddress ($e.Substring(5))).Address }
                }
                $wanted = @($wanted | Where-Object { $_ } | Select-Object -Unique)
                if ($wanted.Count -eq 0) { $r.Step3 = 'Not ready'; $r.Step3Detail = 'No SMTP addresses in the report'; $notReady++; continue }

                $blocked = @()
                foreach ($addr in $wanted) {
                    $hit = $null
                    try {
                        # Cloud, named explicitly. Get-Recipient exists
                        # on-premises too and this check is meaningless against
                        # the wrong directory: it would report "sync complete"
                        # for objects the cloud never had.
                        $hit = exo 'Get-Recipient' @{ Filter = "EmailAddresses -eq 'smtp:$addr'"; ErrorAction = 'SilentlyContinue' }
                    } catch { }
                    if ($hit) { $blocked += "$addr (held by $(@($hit)[0].Name))" }
                }
                if ($blocked.Count -eq 0) {
                    $r.Step3 = 'Ready'; $r.Step3Detail = "$($wanted.Count) address(es) free"
                    & $log "  $($r.Name) -- READY ($($wanted.Count) address(es) free), mailbox will be '$mbName'."
                    $ready++
                } else {
                    $r.Step3 = 'Not ready'; $r.Step3Detail = ($blocked -join '; ')
                    & $log "  $($r.Name) -- NOT READY: $($blocked -join '; ')"
                    $notReady++
                }
            }
            & $save
            & $paint
            & $log ''
            & $log "Ready: $ready    Not ready: $notReady"
            if ($notReady -eq 0) { & $log 'All clear -- run step 4 now.' }
            else { & $log 'Re-run this in a few minutes. Step 4 can be run on the Ready rows in the meantime by selecting them first.' }
        } catch { & $log "Step 3 failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Check sync (cloud)' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    # ---- 4 create -----------------------------------------------------------
    $btn4.Add_Click({
        if ($form.Busy) { return }
        if (-not (& $stepReady $btn4 'Step 4 needs step 1 first: the mailbox is built from the exported attributes.')) { return }
        $form.Busy = $true
        try {
            if (-not (Test-CloudReady)) { & $log 'Exchange Online is not connected in this window.'; return }
            $rows = @((& $pick) | Where-Object { $_.Step1 -eq 'Done' })
            if ($rows.Count -eq 0) { & $log 'No rows with a completed step 1.'; return }

            $notReady = @($rows | Where-Object { $_.Step3 -ne 'Ready' })
            $a = [System.Windows.Forms.MessageBox]::Show(
                ("Create in Exchange Online for $($rows.Count) row(s)?" +
                 "`r`n`r`n  a shared mailbox with the old addresses as aliases and the X500" +
                 "`r`n  a mail-enabled security group SMB-Access-<name> holding the old members" +
                 "`r`n  Full Access and Send As for that group over the mailbox" +
                 "`r`n  owners from $($script:MgrName)" +
                 "`r`n`r`nAnything already there is left alone." +
                 $(if ($notReady.Count) { "`r`n`r`n$($notReady.Count) row(s) have NOT passed step 3 and will probably fail on a held address." } else { '' })),
                'Confirm create', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { & $log 'Step 4 cancelled.'; return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            $managers = Read-ManagerList -Path (Join-Path $PSScriptRoot $script:MgrName)
            $owners = @()
            foreach ($m in $managers) {
                $x = $null
                try { $x = exo 'Get-Recipient' @{ Identity = $m; ErrorAction = 'Stop' } } catch { }
                if ($x) { $owners += "$($x.PrimarySmtpAddress)" } else { & $log "  owner '$m' from $($script:MgrName) was not found in the cloud -- skipped." }
            }
            if ($owners.Count -eq 0) { & $log "No owners resolved from $($script:MgrName); the access groups will be created without ManagedBy." }

            $made = 0; $fail = 0
            foreach ($r in $rows) {
                $why = Test-RowBlocked $r
                if ($why) { $r.Step4 = 'Blocked'; $r.Step4Detail = $why; & $log "  BLOCKED  $($r.Name) -- $why"; $fail++; continue }
                $mbName = Get-MailboxNameForRow $r
                $grpName = Get-AccessGroupName -MailboxName $mbName
                $r.MailboxName = $mbName; $r.AccessGroup = $grpName
                & $log ''
                & $log "$($r.Name)  ->  $mbName"
                try {
                    # ---- addresses ----
                    $aliases = @(); $rewrote = @()
                    foreach ($e in (ConvertFrom-CsvList $r.EmailAddresses)) {
                        if ($e -notmatch '^(?i)smtp:') { continue }
                        $res = Resolve-CarriedAddress ($e.Substring(5))
                        $aliases += $res.Address
                        if ($res.Rewritten) { $rewrote += "$($e.Substring(5)) -> $($res.Address)" }
                    }
                    $aliases = @($aliases | Where-Object { $_ } | Select-Object -Unique)
                    foreach ($m in $rewrote) { & $log "  domain no longer accepted, carried as: $m" }

                    $primary = $(if ("$($r.NewName)".Trim()) { "$("$($r.NewName)".Trim())@$(Get-MailDomain)" }
                                 else { (Resolve-CarriedAddress $r.PrimarySmtpAddress).Address })
                    if (-not $primary) { throw 'No primary address could be worked out for this row.' }

                    # ---- mailbox ----
                    Write-ToolStep "Creating mailbox $mbName ..."
                    $mb = $null
                    try { $mb = exo 'Get-Mailbox' @{ Identity = $mbName; ErrorAction = 'Stop' } } catch { }
                    if ($mb) {
                        & $log "  mailbox exists already -- left alone."
                    } else {
                        # Learn (New-Mailbox): -Shared makes it a shared
                        # mailbox; -Name, -DisplayName, -Alias,
                        # -PrimarySmtpAddress. Alias is derived from the group's
                        # own alias when it had one, because that is what the
                        # rest of the tenant already knows this thing as.
                        $alias = ("$(if ("$($r.NewName)".Trim()) { "$($r.NewName)".Trim() } else { $r.Alias })" -replace '[^A-Za-z0-9.\-_]', '')
                        if ($alias.Length -gt 64) { $alias = $alias.Substring(0, 64) }
                        exo 'New-Mailbox' @{
                            Shared             = $true
                            Name               = $mbName
                            DisplayName        = $(if ("$($r.NewName)".Trim()) { $mbName } else { "$($r.DisplayName)" })
                            Alias              = $alias
                            PrimarySmtpAddress = $primary
                            ErrorAction        = 'Stop'
                        } | Out-Null
                        & $log "  mailbox created: $primary"
                        # Learn (Set-Mailbox): -MessageCopyForSentAsEnabled and
                        # -MessageCopyForSendOnBehalfEnabled, set AFTER
                        # creation because New-Mailbox does not take them. Both
                        # on, so replies do not vanish into an individual's Sent
                        # Items -- the single most common complaint after a
                        # shared mailbox goes live.
                        exo 'Set-Mailbox' @{ Identity = $mbName; MessageCopyForSentAsEnabled = $true; MessageCopyForSendOnBehalfEnabled = $true; ErrorAction = 'Stop' } | Out-Null
                    }

                    # ---- aliases and the X500 ----
                    # Learn (Set-Mailbox): -EmailAddresses @{Add="smtp:..."}.
                    # One Add at a time so a single bad address does not lose
                    # the rest, and each is reported.
                    foreach ($alias in $aliases) {
                        if ($alias -ieq $primary) { continue }
                        try { exo 'Set-Mailbox' @{ Identity = $mbName; EmailAddresses = @{ Add = "smtp:$alias" }; ErrorAction = 'Stop' } | Out-Null; & $log "  alias  $alias" }
                        catch { & $log "  alias  $alias -- NOT added: $($_.Exception.Message)" }
                    }
                    # THE X500 IS WHY REPLIES TO OLD MAIL STILL WORK. Outlook
                    # caches the old group's LegacyExchangeDN in every message
                    # it ever sent to it; without this proxy those replies fail
                    # with "could not be found in the address book". Costs one
                    # call and is invisible until it is missing.
                    if ("$($r.LegacyExchangeDN)".Trim()) {
                        try { exo 'Set-Mailbox' @{ Identity = $mbName; EmailAddresses = @{ Add = "X500:$($r.LegacyExchangeDN)" }; ErrorAction = 'Stop' } | Out-Null; & $log '  X500 added (replies to old mail keep working)' }
                        catch { & $log "  X500 -- NOT added: $($_.Exception.Message)" }
                    }

                    # ---- access group ----
                    Write-ToolStep "Access group for $mbName ..."
                    $g = $null
                    try { $g = exo 'Get-DistributionGroup' @{ Identity = $grpName; ErrorAction = 'Stop' } } catch { }
                    if ($g) {
                        & $log "  access group exists already -- left alone."
                    } else {
                        # Learn (New-DistributionGroup): -Type Security creates
                        # a MAIL-ENABLED SECURITY group, which is the only kind
                        # that can hold mailbox permissions. A plain
                        # distribution group cannot, and granting to one fails
                        # at the point of use rather than here.
                        exo 'New-DistributionGroup' @{
                            Name               = $grpName
                            Alias              = $grpName
                            PrimarySmtpAddress = "$grpName@$(Get-MailDomain)"
                            Type               = 'Security'
                            ErrorAction        = 'Stop'
                        } | Out-Null
                        & $log "  access group created: $grpName"
                    }

                    # ---- members ----
                    $members = @(ConvertFrom-CsvList $r.Members)
                    $have = @()
                    try { $have = @(exo 'Get-DistributionGroupMember' @{ Identity = $grpName; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } | ForEach-Object { "$($_.PrimarySmtpAddress)".ToLower() }) } catch { }
                    $addOk = 0; $addSkip = 0
                    foreach ($m in $members) {
                        $x = $null
                        try { $x = exo 'Get-Recipient' @{ Identity = $m; ErrorAction = 'Stop' } } catch { }
                        if (-not $x) { try { $x = exo 'Get-Recipient' @{ Filter = "EmailAddresses -eq 'smtp:$m'"; ErrorAction = 'SilentlyContinue' } } catch { } }
                        if (-not $x) { & $log "  member $m -- not found in the cloud, skipped."; $addSkip++; continue }
                        $addr = "$(@($x)[0].PrimarySmtpAddress)"
                        if ($have -contains $addr.ToLower()) { $addSkip++; continue }
                        try {
                            # Learn (Add-DistributionGroupMember): -Identity,
                            # -Member, -BypassSecurityGroupManagerCheck.
                            exo 'Add-DistributionGroupMember' @{ Identity = $grpName; Member = $addr; BypassSecurityGroupManagerCheck = $true; ErrorAction = 'Stop' } | Out-Null
                            $addOk++
                        } catch { & $log "  member $addr -- NOT added: $($_.Exception.Message)"; $addSkip++ }
                    }
                    & $log "  members: $addOk added, $addSkip already there or skipped, of $($members.Count)"

                    # ---- owners ----
                    if ($owners.Count) {
                        try {
                            # Learn (Set-DistributionGroup): -ManagedBy,
                            # -BypassSecurityGroupManagerCheck.
                            exo 'Set-DistributionGroup' @{ Identity = $grpName; ManagedBy = $owners; BypassSecurityGroupManagerCheck = $true; ErrorAction = 'Stop' } | Out-Null
                            & $log "  owners set: $($owners.Count)"
                        } catch { & $log "  owners -- NOT set: $($_.Exception.Message)" }
                    }

                    # ---- permissions ----
                    Write-ToolStep "Permissions on $mbName ..."
                    $curFA = @(); $curSA = @()
                    try { $curFA = @(exo 'Get-MailboxPermission'   @{ Identity = $mbName; ErrorAction = 'Stop' } | Where-Object { -not $_.IsInherited -and $_.AccessRights -contains 'FullAccess' } | ForEach-Object { "$($_.User)".ToLower() }) } catch { }
                    try { $curSA = @(exo 'Get-RecipientPermission' @{ Identity = $mbName; ErrorAction = 'Stop' } | Where-Object { $_.AccessRights -contains 'SendAs' } | ForEach-Object { "$($_.Trustee)".ToLower() }) } catch { }
                    $grpAddr = "$grpName@$(Get-MailDomain)"
                    if ($curFA -contains $grpAddr.ToLower() -or $curFA -contains $grpName.ToLower()) { & $log '  Full Access already granted to the access group' }
                    else {
                        # Learn (Add-MailboxPermission): -User, -AccessRights
                        # FullAccess, -InheritanceType All, -AutoMapping.
                        #
                        # AUTOMAPPING IS OFF, and it is not a preference.
                        # Auto-mapping only ever applies to a permission granted
                        # to an individual; granted to a GROUP it does nothing,
                        # so asking for it would be a claim the tool cannot
                        # keep. Members add the mailbox themselves.
                        exo 'Add-MailboxPermission' @{ Identity = $mbName; User = $grpName; AccessRights = 'FullAccess'; InheritanceType = 'All'; AutoMapping = $false; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        & $log '  Full Access granted to the access group'
                    }
                    if ($curSA -contains $grpAddr.ToLower() -or $curSA -contains $grpName.ToLower()) { & $log '  Send As already granted to the access group' }
                    else {
                        # Learn (Add-RecipientPermission): -Trustee,
                        # -AccessRights SendAs.
                        exo 'Add-RecipientPermission' @{ Identity = $mbName; Trustee = $grpName; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        & $log '  Send As granted to the access group'
                    }

                    # ---- send on behalf ----
                    $sob = @()
                    foreach ($s in (ConvertFrom-CsvList $r.GrantSendOnBehalfTo)) {
                        try { [void](exo 'Get-Recipient' @{ Identity = $s; ErrorAction = 'Stop' }); $sob += $s } catch { }
                    }
                    if ($sob.Count) {
                        # Learn (Set-Mailbox): -GrantSendOnBehalfTo @{Add=...}.
                        try { exo 'Set-Mailbox' @{ Identity = $mbName; GrantSendOnBehalfTo = @{ Add = $sob }; ErrorAction = 'Stop' } | Out-Null; & $log "  Send on Behalf carried over: $($sob.Count)" }
                        catch { & $log "  Send on Behalf -- NOT set: $($_.Exception.Message)" }
                    }

                    $r.Step4 = 'Created'; $r.Step4Detail = "$mbName <$primary>; group $grpName; $addOk member(s) added"
                    $made++
                    Write-CompletedLog -Action 'DLBatchCreate' -Target (Get-LogTarget $primary $mbName) -Result 'Success' `
                        -Detail "New-Mailbox -Shared; $($aliases.Count) alias(es); access group $grpName with $addOk member(s); FullAccess + SendAs"
                } catch {
                    $r.Step4 = 'Failed'; $r.Step4Detail = "$($_.Exception.Message)"
                    & $log "  FAILED: $($_.Exception.Message)"
                    Write-FailedLog -Action 'DLBatchCreate' -Target (Get-LogTarget $r.PrimarySmtpAddress $r.Name) -ErrorRecord $_
                    $fail++
                }
            }
            & $save
            & $paint
            & $log ''
            & $log "Step 4 done: $made created, $fail failed. Run step 5 to read it all back."
        } catch { & $log "Step 4 failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Create mailboxes and access groups (cloud)' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    # ---- 5 verify -----------------------------------------------------------
    $btn5.Add_Click({
        if ($form.Busy) { return }
        if (-not (& $stepReady $btn5 'Step 5 needs step 1 first: it checks the cloud against what the export recorded.')) { return }
        $form.Busy = $true
        try {
            if (-not (Test-CloudReady)) { & $log 'Exchange Online is not connected in this window.'; return }
            $rows = @((& $pick) | Where-Object { $_.Step1 -eq 'Done' })
            if ($rows.Count -eq 0) { & $log 'No rows with a completed step 1.'; return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            # READ-ONLY, AND WORTH THE MINUTE IT TAKES. Step 4 reports what it
            # did; this reports what is actually there. On a batch of fourteen
            # those are not reliably the same thing, and the difference is
            # cheaper to find now than in a ticket next week.
            $good = 0; $bad = 0
            foreach ($r in $rows) {
                Write-ToolStep "Verifying $($r.MailboxName) ..."
                $mbName = "$($r.MailboxName)"; if (-not $mbName) { $mbName = Get-MailboxNameForRow $r }
                $grpName = "$($r.AccessGroup)"; if (-not $grpName) { $grpName = Get-AccessGroupName -MailboxName $mbName }
                $problems = @(); $notes = @()
                $mb = $null
                try { $mb = exo 'Get-Mailbox' @{ Identity = $mbName; ErrorAction = 'Stop' } } catch { }
                if (-not $mb) { $problems += 'no mailbox' }
                else {
                    if ("$($mb.RecipientTypeDetails)" -ne 'SharedMailbox') { $problems += "mailbox is a $($mb.RecipientTypeDetails), not shared" }
                    $smtpCount = @(@($mb.EmailAddresses) | Where-Object { "$_" -match '^(?i)smtp:' }).Count
                    $x500 = @(@($mb.EmailAddresses) | Where-Object { "$_" -match '^(?i)x500:' }).Count
                    $notes += "$smtpCount address(es), $x500 X500"
                    if (-not $mb.MessageCopyForSentAsEnabled) { $problems += 'sent-items copy off for Send As' }
                    if (-not $mb.MessageCopyForSendOnBehalfEnabled) { $problems += 'sent-items copy off for Send on Behalf' }
                }
                $g = $null
                try { $g = exo 'Get-DistributionGroup' @{ Identity = $grpName; ErrorAction = 'Stop' } } catch { }
                if (-not $g) { $problems += 'no access group' }
                else {
                    $mc = 0
                    try { $mc = @(exo 'Get-DistributionGroupMember' @{ Identity = $grpName; ResultSize = 'Unlimited'; ErrorAction = 'Stop' }).Count } catch { }
                    $notes += "group has $mc member(s)"
                    $want = @(ConvertFrom-CsvList $r.Members).Count
                    if ($want -gt 0 -and $mc -lt $want) { $problems += "group has $mc of $want member(s)" }
                    if (-not @($g.ManagedBy).Count) { $notes += 'group has no owner' }
                }
                if ($mb -and $g) {
                    $fa = $false; $sa = $false
                    try { $fa = [bool]@(exo 'Get-MailboxPermission'   @{ Identity = $mbName; ErrorAction = 'Stop' } | Where-Object { -not $_.IsInherited -and $_.AccessRights -contains 'FullAccess' -and "$($_.User)" -match [regex]::Escape($grpName) }).Count } catch { }
                    try { $sa = [bool]@(exo 'Get-RecipientPermission' @{ Identity = $mbName; ErrorAction = 'Stop' } | Where-Object { $_.AccessRights -contains 'SendAs' -and "$($_.Trustee)" -match [regex]::Escape($grpName) }).Count } catch { }
                    if (-not $fa) { $problems += 'Full Access missing' }
                    if (-not $sa) { $problems += 'Send As missing' }
                }
                if ($problems.Count) {
                    $r.Step5 = 'Failed'; $r.Step5Detail = ($problems -join '; ')
                    & $log "  $mbName -- $($problems -join '; ')"
                    $bad++
                } else {
                    $r.Step5 = 'OK'; $r.Step5Detail = ($notes -join '; ')
                    & $log "  $mbName -- OK ($($notes -join '; '))"
                    $good++
                }
            }
            & $save
            & $paint
            & $log ''
            & $log "Step 5 done: $good good, $bad with problems."
            Write-CompletedLog -Action 'DLBatchVerify' -Target (Get-LogTarget '' 'batch') -Result $(if ($bad) { 'Partial' } else { 'Success' }) `
                -Detail "$good verified clean, $bad with problems; report $([System.IO.Path]::GetFileName($form.ReportPath))"
        } catch { & $log "Step 5 failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Verify (cloud)' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    $form.Add_Shown({
        & $refreshConn
        & $log "Input file: $(Join-Path $PSScriptRoot $script:InputName)"
        & $log "Reports:    $(Get-ReportPath)"
        $n = Find-NewestReport
        if ($n) { & $log "Newest batch report: $($n.Name)  --  'Load report...' to carry on with it." }
        & $log "'Start new batch' reads $($script:InputName) and begins a fresh report."
    })

    Set-ConsoleTheme -Form $form
    # AFTER the theme, because the enabled look IS the theme's button colour
    # and this is where the unblended value gets captured.
    & $styleSteps
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL to Shared Mailbox (batch)' | Out-Null
    # A 1260-wide window on a 1366-wide laptop is not an improvement. Give back
    # whatever does not fit and let the grid scroll horizontally instead.
    $work = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    if ($form.Width  -gt $work.Width)  { $form.Width  = $work.Width }
    if ($form.Height -gt $work.Height) { $form.Height = $work.Height }
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    [void]$form.ShowDialog()
}
#endregion

Show-BatchForm
