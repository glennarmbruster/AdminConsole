#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- every write here is on-premises
#                                 Set-DistributionGroup, called unqualified.
#                                 Exchange Online exports the same cmdlet name
#                                 and in one process the last connection wins.
# TOOLKIT-VARIANT: OnPremAD
<#
    Rename-OfficeGroups.ps1
    Finishes an office rename across a franchise office's mail-enabled groups.

    An office changes name. Somebody renames the AD group objects and stops
    there, because Name is the field you can see in the tool you renamed them
    in. DisplayName, Alias and the address list keep the old town, so the
    address book, the GAL and every reply-to still say Humboldt while the
    office has been Milan for a fortnight.

    Reasoning: ScriptNotes/Rename-OfficeGroups.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-20
#>
# TOOLKIT-HELP-BEGIN
# Finish an office rename on its distribution and mail-enabled security groups.
#
# Type the OLD CityST and the NEW one -- HumboldtTN and MilanTN -- and press
# Find.
#
# CITYST ONLY. Not 'AR.HumboldtTN', just 'HumboldtTN'. The
# prefixes are kept and rebuilt around the new name, so one pair of entries
# covers the whole office:
#
#     AR.HumboldtTN       ->  AR.MilanTN
#     HumboldtTN          ->  MilanTN
#     Jobs.HumboldtTN     ->  Jobs.MilanTN
#     Invoice.HumboldtTN  ->  Invoice.MilanTN
#
# Every group carrying either name is listed, whatever state it is in: not
# renamed at all, half renamed, or already done. Tick the ones you want, read
# the preview, then Apply.
#
# WHAT IT CHANGES, per group:
#   DisplayName        HumboldtTN            -> MilanTN
#   Alias              AR.HumboldtTN         -> AR.MilanTN
#   Email addresses    every smtp address containing the old CityST is rewritten,
#                      and the rewritten primary becomes the new primary
#   Name               only if you tick the box -- usually it is already done
#
# WHAT IT NEVER TOUCHES:
#   X500 addresses     these are what let old replies still resolve. Removing
#                      one breaks reply-to-old-mail, silently, months later.
#   Any address        without the old CityST in it. A group with a second,
#                      unrelated address keeps it.
#   Membership, owners, permissions, or anything that is not a name.
#
# THE OLD ADDRESS. By default it is REMOVED, because that is what the ticket
# usually means. Untick 'Remove the old addresses' to keep it as a secondary
# instead, and mail to the old address keeps arriving. Removing is the
# irreversible one -- an NDR is the first anyone hears about it.
#
# EMAIL ADDRESS POLICY. A group with the policy still enabled cannot have its
# addresses set by hand: Exchange refuses. Those groups are listed with a note
# and are skipped unless you tick the box to turn the policy off for them, which
# is a real change and is why it is not the default.
#
# ON-PREMISES ONLY, franchise directory. The change syncs up on the next cycle;
# nothing here writes to Exchange Online.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: On-premises Exchange, franchise directory. Writes DisplayName, Alias and email addresses on the groups you tick.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

#region Environment ------------------------------------------------------------
# THE GUARD IS ASYMMETRIC, and that is not an oversight -- see Group-Manager-Editor
# for the long version.
#
# This tool calls Set-DistributionGroup UNQUALIFIED, because there is no
# Invoke-EXOCommand equivalent for the management shell. An Exchange Online
# session in the same process would capture that name and the write would land
# in the cloud on a synced object, where it survives until the next sync
# overwrites it -- a change that appears to work and then quietly undoes itself.
$script:Blocked   = $false
$script:BlockText = ''

$emsState = Test-EmsAvailable
$exoState = Test-ExoInProcess

if ($exoState.Present) {
    $script:Blocked   = $true
    $script:BlockText = 'BLOCKED: Exchange Online is loaded in this process and would answer instead of on-premises Exchange. This must run in its own window.'
} elseif (-not $emsState.Available) {
    $script:Blocked   = $true
    $script:BlockText = 'NOT connected to on-premises Exchange -- nothing can run. Use Connect ExSE.'
}

# Pinned by Find, and every write reuses it. Splatting an empty hashtable adds
# nothing, so the call sites stay identical whether or not a controller was
# resolved.
$script:DCParam    = @{}
$script:DirectoryId = ''
#endregion

#region Actions ----------------------------------------------------------------
function Rename-CityState {
    <#
        Replaces the old CityST with the new one, case-insensitively,
        leaving the rest of the string alone.

        REGEX WITH AN ESCAPED PATTERN, not .Replace(). .Replace() on a string is
        case-SENSITIVE, so 'humboldttn' in an alias somebody typed in lower case
        would sail straight through and the group would come out half renamed --
        which is the exact condition this tool exists to clean up.

        [regex]::Escape on the CityST because it is operator-typed. A CityST with
        a dot in it -- and these all have one, 'AR.HumboldtTN' -- is a regex
        wildcard otherwise.
    #>
    param([string]$Text, [string]$OldCityST, [string]$NewCityST)
    if (-not "$Text") { return "$Text" }
    return [regex]::Replace("$Text", [regex]::Escape($OldCityST), $NewCityST, 'IgnoreCase')
}

function Get-OfficeGroup {
    <#
        Every mail-enabled group carrying either CityST, in the first directory
        that returns one.

        BOTH CITYSTATES, NOT JUST THE OLD ONE. The whole premise is that a rename
        was left half done, so there is no single field that reliably still says
        Humboldt. In the ticket that prompted this, Name already said MilanTN
        while DisplayName, Alias and the address said HumboldtTN -- search only
        for the old CityST and you find the group; search only for the new one
        and you also find it; search for both and you also find the groups
        somebody renamed differently, or not at all.

        Get-DistributionGroup covers mail-enabled UNIVERSAL SECURITY groups as
        well as distribution groups, which matters here: the office's main group
        is usually a security group and the rest are distribution groups.

        DIRECTORIES COME FROM Tenant.json, and the FIRST that returns a hit
        wins, with its controller pinned for every later call. A read and a
        write that land in different forests is the failure this prevents.
    #>
    param(
        [Parameter(Mandatory)][string]$OldCityST,
        [Parameter(Mandatory)][string]$NewCityST
    )

    $result = [ordered]@{
        Groups      = @()
        DirectoryId = ''
        Server      = ''
        Errors      = @()
    }

    $o = ConvertTo-OpathValue -Value $OldCityST
    $n = ConvertTo-OpathValue -Value $NewCityST

    # Name, DisplayName and Alias only. EmailAddresses is deliberately NOT in
    # the filter: it is multi-valued, its OPATH behaviour differs between
    # Exchange versions, and a group whose ADDRESS carries the CityST but whose
    # name, display name and alias do not is a case that does not occur here --
    # the alias and the address are generated from each other.
    $filter = "(Name -like '*$o*') -or (DisplayName -like '*$o*') -or (Alias -like '*$o*') -or " +
              "(Name -like '*$n*') -or (DisplayName -like '*$n*') -or (Alias -like '*$n*')"

    $dirs = @()
    try { $dirs = @(Get-TenantDirectory) } catch { }
    if ($dirs.Count -eq 0) { $dirs = @([PSCustomObject]@{ Id = '' }) }

    foreach ($d in $dirs) {
        $id = "$($d.Id)"

        # Resolve-DirectoryDomainController, NOT Get-TenantDirectoryServer.
        # ARCHITECTURE 14f: the latter returns a DOMAIN when a directory names
        # several controllers, and Exchange rejects a domain with "Active
        # Directory error 0x52 -- suitability of server". Empty is a valid
        # answer and means "pass nothing".
        $dc = ''
        try { if ($id) { $dc = Resolve-DirectoryDomainController -DirectoryId $id } } catch { }

        $p = @{ Filter = $filter; ResultSize = 'Unlimited'; ErrorAction = 'Stop' }
        if ("$dc".Trim()) { $p['DomainController'] = $dc }

        try {
            $found = @(Get-DistributionGroup @p)
        } catch {
            $result.Errors += "$(if ($id) { $id } else { 'default directory' }): $($_.Exception.Message)"
            continue
        }

        if ($found.Count -gt 0) {
            $result.Groups      = $found
            $result.DirectoryId = $id
            $result.Server      = $dc
            break
        }
    }

    return [PSCustomObject]$result
}

function Get-PlannedChange {
    <#
        What one group would become. Computes only -- writes nothing.

        Everything the operator is asked to approve is decided here, so the
        preview and the write cannot disagree: Invoke-OfficeRename applies this
        object and does no thinking of its own.
    #>
    param(
        [Parameter(Mandatory)]$Group,
        [Parameter(Mandatory)][string]$OldCityST,
        [Parameter(Mandatory)][string]$NewCityST,
        [bool]$RemoveOld,
        [bool]$RenameName
    )

    $curName    = "$($Group.Name)"
    $curDisplay = "$($Group.DisplayName)"
    $curAlias   = "$($Group.Alias)"
    $curPrimary = "$($Group.PrimarySmtpAddress)"

    $newName    = Rename-CityState -Text $curName    -OldCityST $OldCityST -NewCityST $NewCityST
    $newDisplay = Rename-CityState -Text $curDisplay -OldCityST $OldCityST -NewCityST $NewCityST
    $newAlias   = Rename-CityState -Text $curAlias   -OldCityST $OldCityST -NewCityST $NewCityST
    $newPrimary = Rename-CityState -Text $curPrimary -OldCityST $OldCityST -NewCityST $NewCityST

    # ---- the address list -------------------------------------------------
    #
    # REBUILT WHOLE, because Set-DistributionGroup -EmailAddresses REPLACES the
    # collection rather than merging into it -- the same trap as ManagedBy and
    # AcceptMessagesOnlyFromSendersOrMembers. Anything not carried across here
    # is deleted, so the rule is to copy everything and change only what
    # matches.
    $kept    = New-Object System.Collections.ArrayList
    $dropped = New-Object System.Collections.ArrayList
    $added   = New-Object System.Collections.ArrayList
    $seen    = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)

    $addEntry = {
        param([string]$Prefix, [string]$Value, $Bucket)
        $key = "$Prefix`:$Value"
        if ($seen.Add($Value)) { [void]$Bucket.Add($key) }
    }

    foreach ($raw in @($Group.EmailAddresses)) {
        $s = "$raw"
        if (-not $s) { continue }

        $idx = $s.IndexOf(':')
        if ($idx -lt 1) { [void]$kept.Add($s); continue }

        $prefix = $s.Substring(0, $idx)
        $value  = $s.Substring($idx + 1)

        # X500 AND EVERYTHING THAT IS NOT SMTP PASS THROUGH UNTOUCHED.
        #
        # X500 is the one worth naming: it carries the object's previous
        # legacyExchangeDN, and it is what lets a reply to a two-year-old mail
        # still resolve. Dropping one breaks that silently, and the report
        # arrives months later as "replies to old mail bounce".
        if ($prefix -notmatch '^(?i)smtp$') { [void]$kept.Add($s); continue }

        $renamed = Rename-CityState -Text $value -OldCityST $OldCityST -NewCityST $NewCityST

        if ($renamed -ieq $value) {
            # No CityST in it -- an unrelated address the office also has. Keep.
            & $addEntry 'smtp' $value $kept
            continue
        }

        & $addEntry 'smtp' $renamed $added
        if ($RemoveOld) { [void]$dropped.Add("smtp:$value") }
        else            { & $addEntry 'smtp' $value $kept }
    }

    # ONE PRIMARY, AND IT IS THE CASE THAT DECIDES IT. In an EmailAddresses
    # collection the single upper-case SMTP: entry is the primary; every other
    # is a secondary. Passing two upper-case entries is an error, passing none
    # leaves Exchange to choose.
    $final = New-Object System.Collections.ArrayList
    foreach ($e in @($added.ToArray()) + @($kept.ToArray())) {
        $i = $e.IndexOf(':')
        $v = $e.Substring($i + 1)
        $p = $e.Substring(0, $i)
        if ($p -match '^(?i)smtp$') {
            if ($v -ieq $newPrimary) { [void]$final.Add("SMTP:$v") } else { [void]$final.Add("smtp:$v") }
        } else {
            [void]$final.Add($e)
        }
    }

    # ---- notes: things worth knowing that this tool will not do ------------
    $notes = New-Object System.Collections.ArrayList

    if ($Group.EmailAddressPolicyEnabled) {
        [void]$notes.Add('Email address policy is ENABLED -- addresses cannot be set by hand')
    }
    if ("$($Group.SamAccountName)" -and (Rename-CityState -Text "$($Group.SamAccountName)" -OldCityST $OldCityST -NewCityST $NewCityST) -ine "$($Group.SamAccountName)") {
        [void]$notes.Add("SamAccountName still says $($Group.SamAccountName) -- not changed here")
    }
    if (-not ($final | Where-Object { $_ -cmatch '^SMTP:' })) {
        [void]$notes.Add('No primary could be determined -- skipped')
    }

    $changes = New-Object System.Collections.ArrayList
    if ($RenameName -and $newName -cne $curName)   { [void]$changes.Add('Name') }
    if ($newDisplay -cne $curDisplay)              { [void]$changes.Add('DisplayName') }
    if ($newAlias   -cne $curAlias)                { [void]$changes.Add('Alias') }
    if ($added.Count -gt 0 -or $dropped.Count -gt 0) { [void]$changes.Add('EmailAddresses') }

    return [PSCustomObject]@{
        Dn           = "$($Group.DistinguishedName)"
        Name         = $curName
        NewName      = $newName
        Display      = $curDisplay
        NewDisplay   = $newDisplay
        Alias        = $curAlias
        NewAlias     = $newAlias
        Primary      = $curPrimary
        NewPrimary   = $newPrimary
        Addresses    = @($final.ToArray())
        Added        = @($added.ToArray())
        Dropped      = @($dropped.ToArray())
        PolicyOn     = [bool]$Group.EmailAddressPolicyEnabled
        Notes        = @($notes.ToArray())
        Changes      = @($changes.ToArray())
        HasWork      = ($changes.Count -gt 0)
    }
}

function Format-PlannedChange {
    param([Parameter(Mandatory)]$Plan, [bool]$RenameName)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("=== $($Plan.Name) ===")

    if (-not $Plan.HasWork) {
        [void]$sb.AppendLine('  Nothing to change -- already carries the new CityST.')
        foreach ($n in $Plan.Notes) { [void]$sb.AppendLine("  NOTE: $n") }
        return $sb.ToString()
    }

    if ($RenameName -and $Plan.NewName -cne $Plan.Name) {
        [void]$sb.AppendLine("  Name         : $($Plan.Name)  ->  $($Plan.NewName)")
    }
    if ($Plan.NewDisplay -cne $Plan.Display) {
        [void]$sb.AppendLine("  DisplayName  : $($Plan.Display)  ->  $($Plan.NewDisplay)")
    }
    if ($Plan.NewAlias -cne $Plan.Alias) {
        [void]$sb.AppendLine("  Alias        : $($Plan.Alias)  ->  $($Plan.NewAlias)")
    }
    if ($Plan.NewPrimary -cne $Plan.Primary) {
        [void]$sb.AppendLine("  Primary SMTP : $($Plan.Primary)  ->  $($Plan.NewPrimary)")
    }
    foreach ($a in $Plan.Added)   { [void]$sb.AppendLine("  ADD          : $a") }
    foreach ($d in $Plan.Dropped) { [void]$sb.AppendLine("  REMOVE       : $d") }

    # The full resulting list is printed, not just the delta. EmailAddresses is
    # replaced wholesale, so what matters is what the collection will BE -- and
    # an X500 that quietly failed to make the list is only visible here.
    [void]$sb.AppendLine('  Resulting address list:')
    foreach ($a in $Plan.Addresses) { [void]$sb.AppendLine("      $a") }

    foreach ($n in $Plan.Notes) { [void]$sb.AppendLine("  NOTE: $n") }
    return $sb.ToString()
}

function Invoke-OfficeRename {
    <#
        Applies one planned change. Decides nothing.

        ONE Set-DistributionGroup CALL. DisplayName, Alias and the address list
        go together, so a group cannot end up with its alias changed and its
        addresses not -- which is the half-finished state this tool was written
        to repair, and it would be galling to create more of it.

        -EmailAddresses AND NOT -PrimarySmtpAddress. Microsoft Learn is explicit:
        "You can't use the EmailAddresses parameter and the PrimarySmtpAddress
        parameter in the same command." The primary is set by CASE instead --
        the single upper-case SMTP: entry -- and the docs are equally explicit
        that the value may appear only once.

        THE ONE THING THAT IS NOT CONFIRMED. The same page says PrimarySmtpAddress
        "updates the primary email address and WindowsEmailAddress property to
        the same value", and says nothing about whether EmailAddresses does the
        same to WindowsEmailAddress -- the 'mail' attribute. In practice Exchange
        keeps them together; I could not verify it, so the caller READS THE GROUP
        BACK and prints what actually landed rather than trusting either answer.

        -BypassSecurityGroupManagerCheck because Exchange checks ManagedBy on a
        mail-enabled SECURITY group regardless of RBAC, and the office's main
        group is usually one. Without it the write fails for a reason that has
        nothing to do with the operator's permissions.
    #>
    param(
        [Parameter(Mandatory)]$Plan,
        [bool]$RenameName,
        [bool]$DisablePolicy
    )

    $p = @{
        Identity                       = $Plan.Dn
        DisplayName                    = $Plan.NewDisplay
        Alias                          = $Plan.NewAlias
        EmailAddresses                 = $Plan.Addresses
        BypassSecurityGroupManagerCheck = $true
        ErrorAction                    = 'Stop'
    }
    if ($RenameName) { $p['Name'] = $Plan.NewName }

    # THE POLICY GOES OFF IN THE SAME CALL, not a separate one before it.
    # Two calls means a window where the policy is off and the addresses are
    # still the old ones, and if the second call fails the group is left in a
    # state nobody asked for.
    if ($DisablePolicy -and $Plan.PolicyOn) { $p['EmailAddressPolicyEnabled'] = $false }

    foreach ($k in $script:DCParam.Keys) { $p[$k] = $script:DCParam[$k] }

    # IDENTITY IS THE DISTINGUISHED NAME. ARCHITECTURE 14g: a resolved object's
    # Identity comes back as its AD common name, and handing that back is a
    # fresh forest-wide NAME LOOKUP rather than a bind -- which, on a tool whose
    # entire job is that two objects briefly share a name, could land on the
    # wrong one.
    Set-DistributionGroup @p | Out-Null

    # READ BACK, AND RETURN WHAT IS REALLY THERE.
    #
    # Not belt-and-braces: it settles the WindowsEmailAddress question above
    # with evidence instead of an assumption, on the first group of the first
    # run, in front of the person who can do something about it. A tool that
    # reports what it INTENDED is how "it said it worked" and "it did not work"
    # end up both being true.
    $rp = @{ Identity = $Plan.Dn; ErrorAction = 'Stop' }
    foreach ($k in $script:DCParam.Keys) { $rp[$k] = $script:DCParam[$k] }
    return (Get-DistributionGroup @rp)
}
#endregion

#region WinForms UI ------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Office Rename -- groups'
$form.ClientSize    = New-Object System.Drawing.Size(944, 652)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Tag       = 'NoTheme'
$lblWarn.Location  = New-Object System.Drawing.Point(12, 12)
$lblWarn.Size      = New-Object System.Drawing.Size(880, 34)
$lblWarn.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
if ($script:Blocked) {
    $lblWarn.Text      = $script:BlockText
    $lblWarn.ForeColor = [System.Drawing.Color]::Red
} else {
    $lblWarn.Text      = 'ON-PREMISES. Writes go to the franchise directory and sync up on the next cycle. Nothing here touches Exchange Online.'
    $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
}
$form.Controls.Add($lblWarn)

$lblOld           = New-Object System.Windows.Forms.Label
$lblOld.Text      = 'Old CityST:'
$lblOld.Location  = New-Object System.Drawing.Point(12, 56)
$lblOld.Size      = New-Object System.Drawing.Size(80, 18)
$form.Controls.Add($lblOld)

$txtOld           = New-Object System.Windows.Forms.TextBox
$txtOld.Location  = New-Object System.Drawing.Point(96, 53)
$txtOld.Size      = New-Object System.Drawing.Size(180, 24)
$txtOld.Font      = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOld)

$lblNew           = New-Object System.Windows.Forms.Label
$lblNew.Text      = 'New CityST:'
$lblNew.Location  = New-Object System.Drawing.Point(296, 56)
$lblNew.Size      = New-Object System.Drawing.Size(80, 18)
$form.Controls.Add($lblNew)

$txtNew           = New-Object System.Windows.Forms.TextBox
$txtNew.Location  = New-Object System.Drawing.Point(380, 53)
$txtNew.Size      = New-Object System.Drawing.Size(180, 24)
$txtNew.Font      = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtNew)

$btnFind          = New-Object System.Windows.Forms.Button
$btnFind.Text     = 'Find'
$btnFind.Location = New-Object System.Drawing.Point(576, 52)
$btnFind.Size     = New-Object System.Drawing.Size(100, 26)
$form.Controls.Add($btnFind)

$lblDir           = New-Object System.Windows.Forms.Label
$lblDir.Location  = New-Object System.Drawing.Point(688, 56)
$lblDir.Size      = New-Object System.Drawing.Size(244, 18)
$lblDir.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblDir)

# THE EXAMPLE IS ON THE FORM, not in the help behind the '?'.
#
# The first version labelled these boxes 'CityST', which is a word I made
# up while writing the search and which means nothing to the person holding the
# ticket. The ticket says HumboldtTN and MilanTN; the form should ask for
# exactly that, and show what happens to the prefixes so nobody has to guess
# whether to type 'AR.HumboldtTN' or just 'HumboldtTN'.
$lblHint           = New-Object System.Windows.Forms.Label
$lblHint.Text      = 'CityST only -- HumboldtTN, not AR.HumboldtTN. Prefixes are kept and rebuilt: AR.HumboldtTN becomes AR.MilanTN, Jobs.HumboldtTN becomes Jobs.MilanTN.'
$lblHint.Location  = New-Object System.Drawing.Point(12, 84)
$lblHint.Size      = New-Object System.Drawing.Size(920, 18)
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblHint)

# The grid gives up the 22px rather than the window growing -- it still shows
# about ten rows, and an office has four or five groups.
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 108)
$grid.Size     = New-Object System.Drawing.Size(920, 278)
$grid.AllowUserToAddRows    = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible     = $false
$grid.SelectionMode         = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.MultiSelect           = $false
$grid.BackgroundColor       = [System.Drawing.Color]::White
$grid.BorderStyle           = [System.Windows.Forms.BorderStyle]::Fixed3D
$grid.Font                  = New-Object System.Drawing.Font('Consolas', 9)
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue

$colSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSel.Name = 'Sel'; $colSel.HeaderText = ''; $colSel.Width = 30
[void]$grid.Columns.Add($colSel)
foreach ($c in @(
    @{ N = 'Name';    H = 'Name';            W = 150 },
    @{ N = 'Display'; H = 'DisplayName now'; W = 150 },
    @{ N = 'Alias';   H = 'Alias now';       W = 150 },
    @{ N = 'Primary'; H = 'Primary SMTP now'; W = 220 },
    @{ N = 'Notes';   H = 'Notes';           W = 210 })) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.N; $col.HeaderText = $c.H; $col.Width = $c.W; $col.ReadOnly = $true
    [void]$grid.Columns.Add($col)
}
$form.Controls.Add($grid)

$chkRemoveOld          = New-Object System.Windows.Forms.CheckBox
$chkRemoveOld.Text     = 'Remove the old addresses'
$chkRemoveOld.Location = New-Object System.Drawing.Point(12, 396)
$chkRemoveOld.Size     = New-Object System.Drawing.Size(200, 20)
$chkRemoveOld.Checked  = $true
$form.Controls.Add($chkRemoveOld)

$chkRenameName          = New-Object System.Windows.Forms.CheckBox
$chkRenameName.Text     = 'Also rename the Name attribute'
$chkRenameName.Location = New-Object System.Drawing.Point(224, 396)
$chkRenameName.Size     = New-Object System.Drawing.Size(230, 20)
$form.Controls.Add($chkRenameName)

$chkPolicy          = New-Object System.Windows.Forms.CheckBox
$chkPolicy.Text     = 'Turn off the email address policy where it blocks the change'
$chkPolicy.Location = New-Object System.Drawing.Point(464, 396)
$chkPolicy.Size     = New-Object System.Drawing.Size(360, 20)
$form.Controls.Add($chkPolicy)

$btnAll           = New-Object System.Windows.Forms.Button
$btnAll.Text      = 'All'
$btnAll.Location  = New-Object System.Drawing.Point(838, 393)
$btnAll.Size      = New-Object System.Drawing.Size(44, 24)
$form.Controls.Add($btnAll)

$btnNone          = New-Object System.Windows.Forms.Button
$btnNone.Text     = 'None'
$btnNone.Location = New-Object System.Drawing.Point(886, 393)
$btnNone.Size     = New-Object System.Drawing.Size(46, 24)
$form.Controls.Add($btnNone)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(12, 422)
$lblStatus.Size     = New-Object System.Drawing.Size(920, 20)
$form.Controls.Add($lblStatus)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 446)
$txtOut.Size       = New-Object System.Drawing.Size(920, 152)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$btnPreview          = New-Object System.Windows.Forms.Button
$btnPreview.Text     = 'Preview'
$btnPreview.Location = New-Object System.Drawing.Point(12, 610)
$btnPreview.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnPreview)

$btnApply          = New-Object System.Windows.Forms.Button
$btnApply.Text     = 'Apply changes'
$btnApply.Location = New-Object System.Drawing.Point(140, 610)
$btnApply.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnApply)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(822, 610)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnClose)

# Found groups, keyed by grid row. The grid holds display text; this holds the
# objects, because re-reading Exchange to answer "what did row 3 mean" is both
# slow and a chance for the answer to have changed underneath.
$script:Found = @()

function Get-CityStatePair {
    return [PSCustomObject]@{
        Old = $txtOld.Text.Trim()
        New = $txtNew.Text.Trim()
    }
}

function Test-CityStateReady {
    $t = Get-CityStatePair
    if (-not $t.Old -or -not $t.New) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'Enter both the old and the new CityST.'
        return $false
    }
    if ($t.Old -ieq $t.New) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'The two entries are the same -- there would be nothing to change.'
        return $false
    }
    return $true
}

function Get-SelectedPlan {
    <#
        The planned change for every ticked row, recomputed from the current
        checkbox settings.

        RECOMPUTED RATHER THAN CACHED, on purpose. 'Remove the old addresses'
        changes what the plan IS, so a plan built when the box was ticked and
        applied after it was unticked would do the opposite of what the screen
        says.
    #>
    $t = Get-CityStatePair
    $plans = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        if ("$($grid.Rows[$i].Cells['Sel'].Value)" -ne 'True') { continue }
        if ($i -ge $script:Found.Count) { continue }
        [void]$plans.Add((Get-PlannedChange -Group $script:Found[$i] -OldCityST $t.Old -NewCityST $t.New `
                            -RemoveOld $chkRemoveOld.Checked -RenameName $chkRenameName.Checked))
    }
    return @($plans.ToArray())
}

$btnFind.Add_Click({
    if ($script:Blocked) { $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick; $lblStatus.Text = 'Blocked -- see the message at the top.'; return }
    if (-not (Test-CityStateReady)) { return }

    $t = Get-CityStatePair
    $btnFind.Enabled = $false
    $form.Cursor     = [System.Windows.Forms.Cursors]::WaitCursor
    $grid.Rows.Clear()
    $script:Found = @()
    $txtOut.Text  = ''
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.Text = 'Searching...'
    $form.Refresh()

    try {
        $r = Get-OfficeGroup -OldCityST $t.Old -NewCityST $t.New

        $script:DirectoryId = $r.DirectoryId
        $script:DCParam     = @{}
        if ("$($r.Server)".Trim()) { $script:DCParam['DomainController'] = $r.Server }

        $lblDir.Text = if ($r.DirectoryId) { "Directory: $($r.DirectoryId)" } else { '' }

        $script:Found = @($r.Groups)
        foreach ($g in $script:Found) {
            $plan = Get-PlannedChange -Group $g -OldCityST $t.Old -NewCityST $t.New `
                        -RemoveOld $chkRemoveOld.Checked -RenameName $chkRenameName.Checked
            $note = if ($plan.Notes.Count) { ($plan.Notes -join '; ') }
                    elseif (-not $plan.HasWork) { 'already renamed' }
                    else { '' }
            # PRE-TICKED ONLY WHERE THERE IS WORK. A row that is already correct
            # starts unticked, so pressing Apply straight after Find does the
            # thing the operator would have chosen anyway.
            [void]$grid.Rows.Add($plan.HasWork, $plan.Name, $plan.Display, $plan.Alias, $plan.Primary, $note)
        }

        if ($r.Errors.Count -gt 0) { $txtOut.Text = ($r.Errors -join [Environment]::NewLine) }

        if ($script:Found.Count -eq 0) {
            $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
            $lblStatus.Text = "No mail-enabled group carries '$($t.Old)' or '$($t.New)'. Check the spelling, and that this console is connected to the right directory."
        } else {
            $work = @($script:Found | ForEach-Object {
                Get-PlannedChange -Group $_ -OldCityST $t.Old -NewCityST $t.New -RemoveOld $chkRemoveOld.Checked -RenameName $chkRenameName.Checked
            } | Where-Object { $_.HasWork }).Count
            $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
            $lblStatus.Text = "$($script:Found.Count) group(s) found, $work with something to change. Press Preview."
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroup'
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'Search failed -- see below.'
    } finally {
        $btnFind.Enabled = $true
        $form.Cursor     = [System.Windows.Forms.Cursors]::Default
    }
})

$btnAll.Add_Click({  for ($i = 0; $i -lt $grid.Rows.Count; $i++) { $grid.Rows[$i].Cells['Sel'].Value = $true  } })
$btnNone.Add_Click({ for ($i = 0; $i -lt $grid.Rows.Count; $i++) { $grid.Rows[$i].Cells['Sel'].Value = $false } })

$btnPreview.Add_Click({
    if (-not (Test-CityStateReady)) { return }
    $plans = Get-SelectedPlan
    if ($plans.Count -eq 0) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'Nothing ticked.'
        return
    }
    $sb = New-Object System.Text.StringBuilder
    foreach ($p in $plans) { [void]$sb.AppendLine((Format-PlannedChange -Plan $p -RenameName $chkRenameName.Checked)) }
    $txtOut.Text = $sb.ToString().TrimEnd()
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.Text = "$($plans.Count) group(s) previewed. Nothing has been written."
})

$btnApply.Add_Click({
    if ($script:Blocked) { $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick; $lblStatus.Text = 'Blocked -- see the message at the top.'; return }
    if (-not (Test-CityStateReady)) { return }

    $plans = @(Get-SelectedPlan | Where-Object { $_.HasWork })
    if ($plans.Count -eq 0) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'Nothing ticked that has anything to change.'
        return
    }

    # BLOCKED BY POLICY IS REPORTED BEFORE THE PROMPT, not discovered during the
    # run. A group whose address policy is on cannot have addresses set, and
    # finding that out one group into a five-group change leaves a half-renamed
    # office -- which is what this tool exists to fix.
    $blocked = @($plans | Where-Object { $_.PolicyOn -and -not $chkPolicy.Checked })
    if ($blocked.Count -gt 0) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = "$($blocked.Count) group(s) have the email address policy enabled. Tick the policy box, or untick those rows."
        $txtOut.Text = "These cannot be changed while their email address policy is enabled:" + [Environment]::NewLine +
                       (($blocked | ForEach-Object { "  $($_.Name)" }) -join [Environment]::NewLine)
        return
    }

    $lines = @($plans | ForEach-Object { "  $($_.Name)  ->  $($_.NewPrimary)" })
    $extra = @()
    if ($chkRemoveOld.Checked)  { $extra += 'The old addresses will be REMOVED -- mail to them will bounce.' }
    if ($chkRenameName.Checked) { $extra += 'The Name attribute will be changed as well.' }
    if ($chkPolicy.Checked)     { $extra += 'The email address policy will be turned OFF where it blocks the change.' }

    $confirm = Show-ToolDialog(
        ("Rename $($plans.Count) group(s)?" + [Environment]::NewLine + [Environment]::NewLine +
         ($lines -join [Environment]::NewLine) +
         $(if ($extra.Count) { [Environment]::NewLine + [Environment]::NewLine + ($extra -join [Environment]::NewLine) } else { '' })),
        'Confirm office rename',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
        $lblStatus.Text = 'Cancelled. Nothing was written.'
        return
    }

    $btnApply.Enabled = $false
    $form.Cursor      = [System.Windows.Forms.Cursors]::WaitCursor
    $sb = New-Object System.Text.StringBuilder
    $ok = 0; $bad = 0
    $t  = Get-CityStatePair

    try {
        foreach ($p in $plans) {
            # EACH GROUP IS INDEPENDENT. One that will not write is reported and
            # the rest continue -- abandoning the run on the second of five
            # would leave the office in exactly the mixed state being repaired.
            try {
                $after = Invoke-OfficeRename -Plan $p -RenameName $chkRenameName.Checked -DisablePolicy $chkPolicy.Checked

                # WHAT IS ON THE OBJECT NOW, not what was asked for.
                [void]$sb.AppendLine("OK      $($p.Name)")
                [void]$sb.AppendLine("          DisplayName        : $($after.DisplayName)")
                [void]$sb.AppendLine("          Alias              : $($after.Alias)")
                [void]$sb.AppendLine("          PrimarySmtpAddress : $($after.PrimarySmtpAddress)")
                [void]$sb.AppendLine("          WindowsEmailAddress: $($after.WindowsEmailAddress)")

                # THE ONE MISMATCH WORTH SHOUTING ABOUT. If 'mail' did not
                # follow the primary, every downstream consumer that reads the
                # AD attribute rather than Exchange -- which is most of them,
                # including this toolkit's own directory lookups -- still sees
                # the old town.
                if ("$($after.WindowsEmailAddress)" -and "$($after.WindowsEmailAddress)" -ine "$($after.PrimarySmtpAddress)") {
                    [void]$sb.AppendLine("          ^^ WindowsEmailAddress did NOT follow the primary. Fix it before closing the ticket.")
                }
                Write-CompletedLog -Action 'RenameOfficeGroup' -Target $p.NewPrimary -Result 'Success' `
                    -Environment $script:DirectoryId `
                    -Detail "$($t.Old) -> $($t.New); changed $($p.Changes -join ', '); was $($p.Primary)"
                $ok++
            } catch {
                [void]$sb.AppendLine("FAILED  $($p.Name)  -- $($_.Exception.Message)")
                # $p.Primary, NOT $p.Name: the rename FAILED, so the group still
                # answers to its old address. NewPrimary would name something
                # that does not exist.
                Write-FailedLog -Action 'RenameOfficeGroup' `
                    -Target (Get-LogTarget $p.Primary $p.Name) -ErrorRecord $_ `
                    -Environment $script:DirectoryId `
                    -Detail "$($t.Old) -> $($t.New)"
                $bad++
            }
        }

        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("$ok succeeded, $bad failed.")
        [void]$sb.AppendLine('Press Find again to re-read the groups and confirm.')
        $txtOut.Text = $sb.ToString()

        $lblStatus.ForeColor = if ($bad -eq 0) { [System.Drawing.Color]::ForestGreen } else { [System.Drawing.Color]::Firebrick }
        $lblStatus.Text = "$ok changed, $bad failed. On-premises only -- the cloud follows on the next sync."
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Office rename'
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = 'Failed -- see below.'
    } finally {
        $btnApply.Enabled = $true
        $form.Cursor      = [System.Windows.Forms.Cursors]::Default
    }
})

$btnClose.Add_Click({ $form.Close() })

if ($script:Blocked) {
    $btnFind.Enabled  = $false
    $btnApply.Enabled = $false
}

$form.AcceptButton = $btnFind

Set-ConsoleTheme -Form $form
Set-FormResizable -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Office Rename' | Out-Null

[void]$form.ShowDialog()
#endregion
