#Requires -Version 5.1
<#
    _Inventory-Shared.ps1
    Query and window code shared by the phase 1 group inventories.

    NOT A TOOL. It defines functions and does nothing else, so it carries no
    console guard and gets no button. The leading underscore is the signal --
    it sorts to the top of the folder and reads as "private" beside the two
    Inventory-*.ps1 files that dot-source it.

    WHY A THIRD FILE RATHER THAN TWO SELF-CONTAINED ONES. In-scope and
    out-of-scope differ by one list of recipient types. Everything else -- the
    directory scoping, the membership join, the nesting test, the owner
    resolution, the whole window -- is identical. Two copies of that would have
    to be corrected twice, and the first time only one of them got corrected
    the two reports would disagree about the same estate while both looking
    authoritative. A phase 1 inventory that contradicts itself is worse than no
    inventory.

    Author   : Glenn Armbruster
    Revised  : 2026-08-13
#>

#region Directory scoping ----------------------------------------------------

function Get-InventoryDirectory {
    <#
        The directory these reports run against.

        DELIBERATELY NOT THE CONSOLE'S -Environment. Every other tool takes the
        directory from whatever object the operator searched for, which is
        right when the tool acts ON that object. These two act on a whole
        directory, and taking the hint from the last search would mean that
        looking up a franchise user and then clicking Inventory silently
        produced a franchise report under a window captioned HQ.

        So it is the DEFAULT directory from Tenant.json, every time, and the
        window says which one on its face. Phase 1 is scoped to one directory
        by decision, not by accident of what was in the search box.
    #>
    $d = Get-DefaultTenantDirectory
    if (-not $d) {
        throw ('No directories are configured in Tenant.json, so there is nothing to inventory. ' +
               'Open Tenant Manager and add the on-premises directory first.')
    }
    return $d
}

function ConvertTo-LdapFilterValue {
    <#
        Escapes a value for use inside an LDAP filter, per RFC 4515.

        USES .Replace(), NOT -replace, AND THAT IS NOT A STYLE CHOICE. In
        -replace the REPLACEMENT half is substitution syntax where a backslash
        begins an escape, so '\5c' as a replacement does not mean what it
        reads as. .Replace() is a plain literal string method with no regex
        engine behind it on either side.

        Backslash is handled FIRST. Escaping it after the others would go back
        over the backslashes this function had just inserted and double them --
        '(' would come out as \5c28 rather than \28, and the filter would match
        nothing while looking correct.

        This matters because distinguished names really do contain these
        characters: a display name with a comma in it arrives as
        CN=Surname\, Firstname.
    #>
    param([string]$Value)
    $v = "$Value"
    $v = $v.Replace('\', '\5c')
    $v = $v.Replace('(', '\28')
    $v = $v.Replace(')', '\29')
    $v = $v.Replace('*', '\2a')
    return $v
}

#endregion

#region Data -----------------------------------------------------------------

function Get-GroupDnLookup {
    <#
        Of these distinguished names, which ones are groups?

        Returns a case-insensitive HashSet of the DNs that are.

        WHY NOT JUST ENUMERATE EVERY GROUP IN THE FOREST. That is the obvious
        implementation and it does not scale here: one directory in this estate
        holds hundreds of offices, and pulling every group object to answer a
        yes/no about a few thousand members is minutes of waiting for a column.

        WHY NOT ONE LOOKUP PER MEMBER. The other obvious implementation, and
        worse -- thousands of round trips.

        So: chunked LDAP. One filter per fifty names, each an OR of exact DN
        matches ANDed with objectClass=group, which returns only the members
        that are groups and nothing else. A few dozen queries instead of
        thousands, and exact rather than inferred.

        Chunks are routed to the directory the DN actually lives in, because a
        group in one directory can hold a member from another and querying the
        wrong server returns nothing -- which would read as "not a group"
        rather than as "asked the wrong machine".

        A DN that belongs to no configured directory is counted, reported, and
        treated as not-a-group. Foreign security principals and members from
        domains nobody listed land here. Silently counting them as ordinary
        members is the failure this branch exists to make visible.
    #>
    param(
        [string[]]$DistinguishedNames,
        [scriptblock]$Progress
    )

    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $result = [PSCustomObject]@{
        GroupDns   = $set
        Unroutable = 0
        Failed     = @()
    }

    $names = @($DistinguishedNames | Where-Object { "$_".Trim() } | Sort-Object -Unique)
    if ($names.Count -eq 0) { return $result }

    # Bucket by directory first. Doing it per-chunk instead would mix
    # directories inside one filter and send half of every chunk to a server
    # that cannot answer for it.
    $byServer = @{}
    $unroutable = 0
    foreach ($dn in $names) {
        $dir = Resolve-DistinguishedNameDirectory -DistinguishedName $dn
        if (-not $dir) { $unroutable++; continue }
        $srv = Get-TenantDirectoryServer -Id "$($dir.Id)"
        if (-not $srv) { $unroutable++; continue }
        if (-not $byServer.ContainsKey($srv)) { $byServer[$srv] = New-Object System.Collections.ArrayList }
        [void]$byServer[$srv].Add($dn)
    }
    $result.Unroutable = $unroutable

    $chunkSize = 50
    $done = 0
    $total = ($names.Count - $unroutable)
    if ($total -lt 1) { $total = 1 }

    foreach ($srv in @($byServer.Keys)) {
        $list = @($byServer[$srv])
        for ($i = 0; $i -lt $list.Count; $i += $chunkSize) {
            $end   = [Math]::Min($i + $chunkSize, $list.Count) - 1
            $chunk = @($list[$i..$end])

            $clauses = ($chunk | ForEach-Object {
                '(distinguishedName=' + (ConvertTo-LdapFilterValue -Value $_) + ')'
            }) -join ''
            $filter = '(&(objectClass=group)(|' + $clauses + '))'

            try {
                $found = @(Get-ADObject -LDAPFilter $filter -Server $srv -ResultSetSize $null -ErrorAction Stop)
                foreach ($f in $found) { [void]$set.Add("$($f.DistinguishedName)") }
            } catch {
                $result.Failed += "Nesting lookup against $srv ($($_.Exception.Message))"
            }

            $done += $chunk.Count
            if ($Progress) {
                & $Progress ("Checking members for nested groups -- $done of $total...") ([int](55 + 30.0 * $done / $total))
            }
        }
    }

    return $result
}

function Get-OwnerText {
    <#
        ManagedBy turned into something a human can act on.

        Exchange hands back an identity string, which is a canonical name or a
        DN depending on which cmdlet produced it and is useless in a planning
        spreadsheet -- the point of the column is "who do I email about this
        group", so it resolves to a name and an address.

        Cached by identity because a handful of people own a great many groups
        between them, and without the cache this is the slowest column in the
        report by a wide margin.

        An identity that will not resolve is returned AS-IS rather than blanked.
        A blank would read as "this group has no owner", which is a different
        and much more consequential fact -- it is the thing that makes a group
        hard to migrate. The raw string at least says someone is named.
    #>
    param(
        $ManagedBy,
        [string]$Server,
        [hashtable]$Cache
    )

    $out = @()
    foreach ($m in @($ManagedBy)) {
        $key = "$m".Trim()
        if (-not $key) { continue }

        if ($Cache.ContainsKey($key)) { $out += $Cache[$key]; continue }

        $text = $key
        try {
            $o = Get-ADObject -Identity $key -Server $Server -Properties DisplayName, mail, Name -ErrorAction Stop
            $nm = if ("$($o.DisplayName)".Trim()) { "$($o.DisplayName)" } else { "$($o.Name)" }
            $text = if ("$($o.mail)".Trim()) { "$nm <$($o.mail)>" } else { $nm }
        } catch {
            # Identity was not a DN, or the object is in another directory.
            # Leave the raw value; see the docstring on why not blank.
        }
        $Cache[$key] = $text
        $out += $text
    }

    return ($out -join '; ')
}

function Get-HqGroupInventory {
    <#
        The mail-enabled groups of one directory, restricted to the recipient
        types asked for, with membership and ownership filled in.

        TWO SOURCES, ON PURPOSE.

        Exchange answers what KIND of group this is. RecipientTypeDetails is an
        Exchange concept and the only reliable way to tell a room list from an
        ordinary distribution group -- both are universal distribution groups
        in Active Directory and differ by a flag Exchange owns. Deriving it
        from raw msExch* bit values was the alternative and it is exactly the
        sort of thing that is right until a version bump.

        Active Directory answers how BIG it is. The member attribute comes back
        in the same bulk read as everything else, where Exchange would charge a
        round trip per group for the same fact.

        The join is on distinguished name, which both sides carry and neither
        side rewrites.

        Every Exchange call goes through Invoke-ExchangeCommand -World OnPrem.
        Get-DistributionGroup exists in BOTH Exchanges and the last module
        imported wins the bare name, so an operator with a cloud session open
        would otherwise get a cloud inventory under an on-premises caption.
        Routing it explicitly turns that from a wrong answer into a refusal.
    #>
    param(
        [Parameter(Mandatory)][string[]]$RecipientTypeDetails,
        [scriptblock]$Progress
    )

    $dir      = Get-InventoryDirectory
    $dirId    = "$($dir.Id)"
    $adServer = Get-TenantDirectoryServer -Id $dirId
    if (-not $adServer) {
        throw "Directory '$dirId' has no domain or controller configured. Open Tenant Manager and set one."
    }

    $failed = @()
    $wanted = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($t in $RecipientTypeDetails) { [void]$wanted.Add("$t") }

    # SCOPE FIRST, QUERY SECOND.
    #
    # Exchange's default recipient scope is the DOMAIN hosting Exchange, not
    # the forest. In a multi-domain forest that means Get-DistributionGroup
    # returns a different set depending on what else has run in this process --
    # any tool that called Set-ADSessionScope has already widened it, and one
    # that has not leaves it narrow. A report whose row count depends on which
    # button you pressed first is a report nobody can reconcile.
    #
    # So it is set explicitly, every run, and the scoping is then done HERE by
    # distinguished name rather than left to Exchange.
    #
    # A failure is recorded rather than thrown: the report still runs and comes
    # back marked PARTIAL saying the scope could not be confirmed. Continuing
    # silently on an unknown scope is the one outcome that must not happen.
    & $Progress 'Setting the directory scope for this session...' 2
    try {
        $scope = Set-ADSessionScope -DirectoryId $dirId
        if ($scope.Fallback) { $failed += "Controller pin refused, continuing forest-wide. $($scope.Detail)" }
    } catch {
        $failed += ("Could not set the recipient scope, so this list may be missing groups from " +
                    "other domains in the forest ($($_.Exception.Message))")
    }

    & $Progress 'Querying on-premises Exchange for mail-enabled groups...' 5

    # -ResultSize Unlimited IS LOAD-BEARING. The default is 1000 and it
    # truncates SILENTLY -- no warning, no error, just a short list that looks
    # complete. On an estate this size that is not a hypothetical.
    $all = @(Invoke-ExchangeCommand -World OnPrem -Command 'Get-DistributionGroup' -Params @{
        ResultSize = 'Unlimited'; ErrorAction = 'Stop' })

    & $Progress "Scoping $($all.Count) group(s) to $dirId..." 20

    # Directory scope, then type. In that order because the type test is the
    # cheaper of the two to run on a smaller set.
    $scoped = @($all | Where-Object {
        (Test-DnInDirectory -DistinguishedName "$($_.DistinguishedName)" -DirectoryId $dirId) -and
        $wanted.Contains("$($_.RecipientTypeDetails)")
    })

    if ($scoped.Count -eq 0) {
        return [PSCustomObject]@{
            Rows = @(); Failed = $failed; Directory = $dirId; Server = $adServer
            Scanned = $all.Count; Unroutable = 0
        }
    }

    & $Progress "Reading membership from Active Directory for $($scoped.Count) group(s)..." 35

    # ONE bulk read rather than one per group. Indexed by DN so the join below
    # is a dictionary hit.
    $adByDn = @{}
    try {
        foreach ($g in @(Get-ADGroup -Filter 'mail -like "*"' -Server $adServer `
                            -Properties member, GroupCategory, GroupScope, DisplayName, mail `
                            -ErrorAction Stop)) {
            $adByDn["$($g.DistinguishedName)"] = $g
        }
    } catch {
        $failed += "Bulk Active Directory read against $adServer ($($_.Exception.Message))"
    }

    # Anything the bulk read missed -- a group whose mail attribute is empty
    # even though Exchange considers it mail-enabled -- is fetched singly. Rare
    # enough to be cheap, and leaving it out would show a real group with zero
    # members, which is the number that gets a group retired.
    $missing = @($scoped | Where-Object { -not $adByDn.ContainsKey("$($_.DistinguishedName)") })
    foreach ($m in $missing) {
        try {
            $g = Get-ADGroup -Identity "$($m.DistinguishedName)" -Server $adServer `
                    -Properties member, GroupCategory, GroupScope, DisplayName, mail -ErrorAction Stop
            $adByDn["$($g.DistinguishedName)"] = $g
        } catch {
            $failed += "Could not read '$($m.Name)' from Active Directory ($($_.Exception.Message))"
        }
    }

    & $Progress 'Collecting member lists...' 50

    $memberDns = New-Object System.Collections.ArrayList
    foreach ($s in $scoped) {
        $ad = $adByDn["$($s.DistinguishedName)"]
        if (-not $ad) { continue }
        foreach ($dn in @($ad.member)) { [void]$memberDns.Add("$dn") }
    }

    $nest = Get-GroupDnLookup -DistinguishedNames @($memberDns.ToArray()) -Progress $Progress
    foreach ($f in @($nest.Failed)) { $failed += $f }

    & $Progress 'Resolving owners...' 88

    $ownerCache = @{}
    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($s in $scoped) {
        $ad      = $adByDn["$($s.DistinguishedName)"]
        $members = @()
        if ($ad) { $members = @($ad.member) }

        $nested = 0
        foreach ($dn in $members) { if ($nest.GroupDns.Contains("$dn")) { $nested++ } }

        $groupType = ''
        if ($ad) {
            $cat = "$($ad.GroupCategory)"
            $scp = "$($ad.GroupScope)"
            $groupType = (@($cat, $scp) | Where-Object { $_ }) -join ', '
        }

        $name = "$($s.DisplayName)".Trim()
        if (-not $name) { $name = "$($s.Name)" }

        $rows.Add([PSCustomObject][ordered]@{
            GroupName            = $name
            PrimarySmtpAddress   = "$($s.PrimarySmtpAddress)"
            RecipientTypeDetails = "$($s.RecipientTypeDetails)"
            GroupType            = $groupType
            Members              = $members.Count
            HasNestedGroups      = $(if ($nested -gt 0) { 'Yes' } else { 'No' })
            NestedGroupCount     = $nested
            Owners               = (Get-OwnerText -ManagedBy $s.ManagedBy -Server $adServer -Cache $ownerCache)
            DistinguishedName    = "$($s.DistinguishedName)"
        })
    }

    return [PSCustomObject]@{
        Rows       = @($rows | Sort-Object GroupName)
        Failed     = $failed
        Directory  = $dirId
        Server     = $adServer
        Scanned    = $all.Count
        Unroutable = $nest.Unroutable
    }
}

function Get-HqDynamicGroupInventory {
    <#
        Dynamic distribution groups, name and address only.

        THE MISSING COLUMNS ARE THE POINT. A dynamic group stores a filter, not
        a membership, so there is no member attribute to count -- the only way
        to get a number is to run the group's own recipient filter against the
        directory, once per group. That is a real query each, and the answer is
        a snapshot that changes on its own the moment anyone's department
        changes.

        Phase 1 does not migrate these, so paying for a number that will be
        stale before the spreadsheet is read buys nothing. Name and address
        establish that they exist and what they are called, which is what the
        out-of-scope list is for.

        Note this cmdlet has NO -RecipientTypeDetails parameter, unlike
        Get-DistributionGroup -- so the type is stamped on rather than filtered.
    #>
    param([scriptblock]$Progress)

    $dir   = Get-InventoryDirectory
    $dirId = "$($dir.Id)"

    & $Progress 'Querying on-premises Exchange for dynamic distribution groups...' 92

    $rows   = New-Object System.Collections.Generic.List[object]
    $failed = @()

    try {
        $dyn = @(Invoke-ExchangeCommand -World OnPrem -Command 'Get-DynamicDistributionGroup' -Params @{
            ResultSize = 'Unlimited'; ErrorAction = 'Stop' })

        foreach ($d in $dyn) {
            if (-not (Test-DnInDirectory -DistinguishedName "$($d.DistinguishedName)" -DirectoryId $dirId)) { continue }
            $name = "$($d.DisplayName)".Trim()
            if (-not $name) { $name = "$($d.Name)" }
            $rows.Add([PSCustomObject][ordered]@{
                GroupName            = $name
                PrimarySmtpAddress   = "$($d.PrimarySmtpAddress)"
                RecipientTypeDetails = 'DynamicDistributionGroup'
                GroupType            = ''
                Members              = ''
                HasNestedGroups      = ''
                NestedGroupCount     = ''
                Owners               = ''
                DistinguishedName    = "$($d.DistinguishedName)"
            })
        }
    } catch {
        $failed += "Dynamic distribution groups ($($_.Exception.Message))"
    }

    return [PSCustomObject]@{ Rows = @($rows | Sort-Object GroupName); Failed = $failed }
}

function Add-ScopeReason {
    <#
        Rebuilds each row with the reason column in the right place.

        Rebuilt rather than Add-Member'd because Add-Member appends, and the
        CSV takes its column order from the property order. The grid puts the
        reason next to the recipient type, so the object has to as well or the
        two disagree -- which is exactly the sort of small inconsistency that
        makes someone distrust the whole file.
    #>
    param($Rows, [hashtable]$Reasons)

    $out = New-Object System.Collections.Generic.List[object]
    foreach ($r in @($Rows)) {
        $type   = "$($r.RecipientTypeDetails)"
        $reason = if ($Reasons.ContainsKey($type)) { $Reasons[$type] } else { "Out of scope for phase 1 ($type)" }

        $out.Add([PSCustomObject][ordered]@{
            GroupName            = $r.GroupName
            PrimarySmtpAddress   = $r.PrimarySmtpAddress
            RecipientTypeDetails = $type
            OutOfScopeReason     = $reason
            GroupType            = $r.GroupType
            Members              = $r.Members
            HasNestedGroups      = $r.HasNestedGroups
            NestedGroupCount     = $r.NestedGroupCount
            Owners               = $r.Owners
            DistinguishedName    = $r.DistinguishedName
        })
    }
    return @($out)
}

function Invoke-InventoryCollect {
    <#
        Runs whichever collection this window is for and returns one result
        object. Both reports go through here so the shape is guaranteed
        identical -- the window reads Rows, Failed, Directory, Server, Scanned
        and Unroutable and does not care which report it is showing.
    #>
    param(
        [Parameter(Mandatory)][string[]]$RecipientTypeDetails,
        [switch]$IncludeDynamic,
        [hashtable]$ReasonMap,
        [scriptblock]$Progress
    )

    $groups = Get-HqGroupInventory -RecipientTypeDetails $RecipientTypeDetails -Progress $Progress
    $rows   = @($groups.Rows)
    $failed = @($groups.Failed)

    if ($IncludeDynamic) {
        $dyn    = Get-HqDynamicGroupInventory -Progress $Progress
        $rows  += @($dyn.Rows)
        $failed += @($dyn.Failed)
    }

    if ($ReasonMap) {
        # Sorted by reason then name so the categories arrive as blocks. One
        # alphabetical run interleaves them and makes counting a category by
        # eye impossible.
        $rows = @(Add-ScopeReason -Rows $rows -Reasons $ReasonMap | Sort-Object OutOfScopeReason, GroupName)
    }

    return [PSCustomObject]@{
        Rows       = @($rows)
        Failed     = @($failed | Where-Object { $_ })
        Directory  = $groups.Directory
        Server     = $groups.Server
        Scanned    = $groups.Scanned
        Unroutable = $groups.Unroutable
    }
}

#endregion

#region Window ---------------------------------------------------------------

function Show-InventoryWindow {
    <#
        The window both inventories use.

        Grid on top, detail pane below, same shape as Log Viewer, because an
        operator who has learned one report should not have to learn the other.

        The caller supplies a title, a report prefix and the recipient types to
        collect. Everything else -- find box, previous runs, CSV write,
        progress, theme, corner placement, help button -- is identical between
        the two and lives here once.

        NO SCRIPTBLOCK PARAMETER, AND NO .GetNewClosure() ANYWHERE IN HERE.
        Both were in the first version and both were wrong for the same reason;
        see the block on $ctx below.
    #>
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)][string]$Blurb,
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string[]]$RecipientTypeDetails,
        [switch]$IncludeDynamic,
        [hashtable]$ReasonMap,
        # Hashtables of @{ Name = ...; Header = ...; Width = ... }.
        [object[]]$ExtraColumns = @()
    )

    $ReportDir = Get-ReportPath

    $form                 = New-Object System.Windows.Forms.Form
    $form.Text            = "Report -- $Title"
    $form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.ClientSize      = New-Object System.Drawing.Size(1204, 764)

    $lblInfo          = New-Object System.Windows.Forms.Label
    $lblInfo.Text     = $Blurb
    $lblInfo.Location = New-Object System.Drawing.Point(12, 12)
    $lblInfo.Size     = New-Object System.Drawing.Size(1180, 18)
    $form.Controls.Add($lblInfo)

    $lblSrc           = New-Object System.Windows.Forms.Label
    $lblSrc.Location  = New-Object System.Drawing.Point(12, 32)
    $lblSrc.Size      = New-Object System.Drawing.Size(1180, 18)
    $lblSrc.ForeColor = [System.Drawing.Color]::DimGray
    $lblSrc.Text      = 'Source: on-premises Exchange and Active Directory. Exchange Online is not queried.'
    $form.Controls.Add($lblSrc)

    $btnRun          = New-Object System.Windows.Forms.Button
    $btnRun.Text     = 'Run Report'
    $btnRun.Location = New-Object System.Drawing.Point(12, 58)
    $btnRun.Size     = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnRun)

    $btnOpen          = New-Object System.Windows.Forms.Button
    $btnOpen.Text     = 'Open Reports Folder'
    $btnOpen.Location = New-Object System.Drawing.Point(140, 58)
    $btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
    $form.Controls.Add($btnOpen)

    $lblFind          = New-Object System.Windows.Forms.Label
    $lblFind.Text     = 'Find:'
    $lblFind.Location = New-Object System.Drawing.Point(310, 64)
    $lblFind.Size     = New-Object System.Drawing.Size(34, 18)
    $form.Controls.Add($lblFind)

    $txtFind          = New-Object System.Windows.Forms.TextBox
    $txtFind.Location = New-Object System.Drawing.Point(346, 61)
    $txtFind.Size     = New-Object System.Drawing.Size(240, 22)
    $form.Controls.Add($txtFind)

    $btnClose          = New-Object System.Windows.Forms.Button
    $btnClose.Text     = 'Close'
    $btnClose.Location = New-Object System.Drawing.Point(1082, 58)
    $btnClose.Size     = New-Object System.Drawing.Size(110, 30)
    $btnClose.Anchor   = 'Top,Right'
    $form.Controls.Add($btnClose)

    $lblStatus          = New-Object System.Windows.Forms.Label
    $lblStatus.Text     = 'Idle -- click Run Report.'
    $lblStatus.Location = New-Object System.Drawing.Point(12, 132)
    $lblStatus.Size     = New-Object System.Drawing.Size(1180, 18)
    $lblStatus.Anchor   = 'Top,Left,Right'
    $form.Controls.Add($lblStatus)

    $pbar = Add-ReportProgressBar -Form $form -X 12 -Y 152 -Width 1180 -Height 12

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location                = New-Object System.Drawing.Point(12, 172)
    $grid.Size                    = New-Object System.Drawing.Size(1180, 400)
    $grid.Anchor                  = 'Top,Left,Right,Bottom'
    $grid.ReadOnly                = $true
    $grid.AllowUserToAddRows      = $false
    $grid.AllowUserToDeleteRows   = $false
    $grid.AllowUserToOrderColumns = $true
    $grid.RowHeadersVisible       = $false
    $grid.SelectionMode           = 'FullRowSelect'
    $grid.MultiSelect             = $true
    $grid.AutoSizeColumnsMode     = 'None'
    $form.Controls.Add($grid)

    $lblDetail          = New-Object System.Windows.Forms.Label
    $lblDetail.Text     = 'Selected group:'
    $lblDetail.Location = New-Object System.Drawing.Point(12, 582)
    $lblDetail.Size     = New-Object System.Drawing.Size(200, 18)
    $lblDetail.Anchor   = 'Left,Bottom'
    $form.Controls.Add($lblDetail)

    $txtDetail            = New-Object System.Windows.Forms.TextBox
    $txtDetail.Location   = New-Object System.Drawing.Point(12, 602)
    $txtDetail.Size       = New-Object System.Drawing.Size(1180, 122)
    $txtDetail.Multiline  = $true
    $txtDetail.ReadOnly   = $true
    $txtDetail.ScrollBars = 'Vertical'
    $txtDetail.Font       = New-Object System.Drawing.Font('Consolas', 9)
    $txtDetail.Anchor     = 'Left,Right,Bottom'
    $form.Controls.Add($txtDetail)

    $lblCount          = New-Object System.Windows.Forms.Label
    $lblCount.Location = New-Object System.Drawing.Point(12, 732)
    $lblCount.Size     = New-Object System.Drawing.Size(1000, 20)
    $lblCount.Anchor   = 'Left,Bottom'
    $form.Controls.Add($lblCount)

    $files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix $Prefix `
                    -X 12 -Y 94 -Width 1180 -StatusLabel $lblStatus

    # --- columns ------------------------------------------------------------
    # EXTRAS GO AFTER 'Recipient Type', NOT ON THE END. The only extra column
    # in use is the out-of-scope reason, and a reason belongs beside the type
    # it explains. Appending it after Owner(s) would have put the most-read
    # column in that report off the right-hand edge on a laptop.
    $spec = @(
        @{ N = 'GroupName';            H = 'Group Name';        W = 260 }
        @{ N = 'PrimarySmtpAddress';   H = 'Primary Email';      W = 250 }
        @{ N = 'RecipientTypeDetails'; H = 'Recipient Type';     W = 190 }
    )
    foreach ($x in $ExtraColumns) {
        $spec += @{ N = $x.Name; H = $x.Header; W = $x.Width }
    }
    $spec += @(
        @{ N = 'GroupType';        H = 'Group Type (AD)';   W = 150 }
        @{ N = 'Members';          H = '# Members';         W = 80  }
        @{ N = 'HasNestedGroups';  H = 'Has Nested Groups'; W = 120 }
        @{ N = 'NestedGroupCount'; H = '# Nested';          W = 75  }
        @{ N = 'Owners';           H = 'Owner(s)';          W = 300 }
    )
    # LAST, AND ALWAYS PRESENT. The selection handler matches on this column to
    # find the row behind the selected line, so it is not optional.
    $spec += @{ N = 'DistinguishedName'; H = 'Distinguished Name'; W = 420 }

    foreach ($s in $spec) {
        $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c.Name = $s.N; $c.HeaderText = $s.H; $c.Width = $s.W
        [void]$grid.Columns.Add($c)
    }

    # --- shared state, carried on .Tag --------------------------------------
    #
    # EVERY HANDLER READS $this.Tag. NOTHING HERE USES .GetNewClosure(), AND
    # THAT IS THE WHOLE POINT OF THIS BLOCK.
    #
    # This window is built inside a FUNCTION, so the handlers need some way to
    # reach $grid, $form and the collected rows. GetNewClosure() looks like the
    # answer -- it copies the local variables into the scriptblock -- and it
    # was the first thing tried here. It fails, and it fails LOUDLY but in a
    # place that hides the cause:
    #
    #     The term 'Complete-ProgressReporter' is not recognized...
    #
    # GetNewClosure binds the scriptblock to a NEW DYNAMIC MODULE. Variables
    # are copied in, so the variables work. But FUNCTION lookup inside a module
    # goes module scope, then GLOBAL -- and Common.ps1 is DOT-SOURCED into the
    # tool's SCRIPT scope, so none of its functions are global. Every
    # Common.ps1 call inside the handler therefore vanishes.
    #
    # Common.ps1 already records this twice, on Add-ReportFileControls and on
    # Add-ToolHelpButton. It is written here a third time because the wrong
    # answer is the one that looks right.
    #
    # A plain scriptblock keeps the session state it was created in, where the
    # functions are visible; the DATA rides on Tag instead.
    $ctx = [PSCustomObject]@{
        Form        = $form
        Grid        = $grid
        TxtFind     = $txtFind
        TxtDetail   = $txtDetail
        LblStatus   = $lblStatus
        LblCount    = $lblCount
        Bar         = $pbar
        BtnRun      = $btnRun
        Files       = $files
        Spec        = $spec
        Rows        = @()
        ReportDir   = $ReportDir
        Prefix      = $Prefix
        Title       = $Title
        Types       = $RecipientTypeDetails
        Dynamic     = [bool]$IncludeDynamic
        ReasonMap   = $ReasonMap
    }

    Add-Member -InputObject $ctx -MemberType ScriptMethod -Name Redraw -Value {
        $find = "$($this.TxtFind.Text)".Trim()
        $rows = @($this.Rows)
        if ($find) {
            $rows = @($rows | Where-Object {
                "$($_.GroupName)$($_.PrimarySmtpAddress)$($_.Owners)$($_.RecipientTypeDetails)" `
                    -like "*$find*"
            })
        }

        $this.Grid.SuspendLayout()
        $this.Grid.Rows.Clear()
        foreach ($r in $rows) {
            $vals = New-Object System.Collections.ArrayList
            foreach ($s in $this.Spec) { [void]$vals.Add("$($r.($s.N))") }
            [void]$this.Grid.Rows.Add($vals.ToArray())
        }
        $this.Grid.ResumeLayout()

        $shown = @($rows).Count
        $total = @($this.Rows).Count
        $this.LblCount.Text = if ($shown -eq $total) {
            "$total group(s)."
        } else {
            "$shown of $total group(s) shown -- filtered by '$find'."
        }
    }

    $btnRun.Tag   = $ctx
    $btnOpen.Tag  = $ctx
    $btnClose.Tag = $ctx
    $txtFind.Tag  = $ctx
    $grid.Tag     = $ctx

    $btnClose.Add_Click({ $this.Tag.Form.Close() })
    $btnOpen.Add_Click({ Open-ReportFolder -Path $this.Tag.ReportDir })
    $txtFind.Add_TextChanged({ $this.Tag.Redraw() })

    $grid.Add_SelectionChanged({
        $c = $this.Tag
        if (-not $c) { return }
        if ($this.SelectedRows.Count -eq 0) { $c.TxtDetail.Text = ''; return }
        $i = $this.SelectedRows[0].Index
        if ($i -lt 0) { return }
        # Matched on the DN, not the name. Two groups can share a display name
        # -- an old one and its replacement, mid-migration, is the normal case
        # -- and matching on the name would show the wrong one's detail with
        # nothing on screen to say so.
        $dn  = "$($this.Rows[$i].Cells['DistinguishedName'].Value)"
        $row = @($c.Rows | Where-Object { "$($_.DistinguishedName)" -eq $dn }) | Select-Object -First 1
        if (-not $row) { $c.TxtDetail.Text = ''; return }

        $sb = New-Object System.Text.StringBuilder
        foreach ($p in $row.PSObject.Properties) {
            [void]$sb.AppendLine(('{0,-22}{1}' -f $p.Name, "$($p.Value)"))
        }
        $c.TxtDetail.Text = $sb.ToString()
    })

    $btnRun.Add_Click({
        $c = $this.Tag
        $c.BtnRun.Enabled = $false
        $prev = $c.Form.Cursor
        $c.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $c.TxtDetail.Text = ''
        try {
            $progress = New-ProgressReporter -Label $c.LblStatus -Bar $c.Bar -Activity $c.Title

            $p = @{ RecipientTypeDetails = $c.Types; Progress = $progress }
            if ($c.Dynamic)   { $p['IncludeDynamic'] = $true }
            if ($c.ReasonMap) { $p['ReasonMap']      = $c.ReasonMap }
            $result = Invoke-InventoryCollect @p

            Complete-ProgressReporter -Activity $c.Title -Bar $c.Bar

            $c.Rows = @($result.Rows)
            $c.Redraw()

            $partial = (@($result.Failed).Count -gt 0)

            if (@($c.Rows).Count -eq 0) {
                $c.LblStatus.Text = 'No groups matched -- nothing written.'
                $c.TxtDetail.Text = 'No groups matched the phase 1 filter, so no CSV was written.'
                if ($partial) {
                    $c.TxtDetail.Text += [Environment]::NewLine + ($result.Failed -join [Environment]::NewLine)
                }
                return
            }

            if (-not (Test-Path -LiteralPath $c.ReportDir)) {
                New-Item -Path $c.ReportDir -ItemType Directory -Force | Out-Null
            }
            $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
            $csv   = Join-Path $c.ReportDir "$($c.Prefix)_$stamp.csv"
            $c.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

            $c.Files.SelectPath($csv)

            $suffix = if ($partial) { ' -- PARTIAL, see the detail pane' } else { '' }
            $c.LblStatus.Text = "Done. $(@($c.Rows).Count) group(s) written to $(Split-Path -Leaf $csv)$suffix"

            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine("Directory scanned    : $($result.Directory)")
            [void]$sb.AppendLine("Active Directory     : $($result.Server)")
            [void]$sb.AppendLine("Groups seen org-wide : $($result.Scanned)")
            [void]$sb.AppendLine("Matched this report  : $(@($c.Rows).Count)")
            if ($result.Unroutable -gt 0) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine("$($result.Unroutable) member(s) belong to no directory listed in Tenant.json.")
                [void]$sb.AppendLine('  They are counted as members but NOT tested for being nested groups,')
                [void]$sb.AppendLine('  so Has Nested Groups may read No where it should read Yes.')
            }
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("Saved to: $csv")
            if ($partial) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine('*** INCOMPLETE -- the following did not answer: ***')
                foreach ($f in $result.Failed) { [void]$sb.AppendLine("  $f") }
            }
            $c.TxtDetail.Text = $sb.ToString()

            $logResult = if ($partial) { 'Partial' } else { 'Success' }
            Write-CompletedLog -Action 'Report' -Target $c.Prefix -Result $logResult `
                -Detail "$($c.Title) -- $(@($c.Rows).Count) group(s) written to $(Split-Path -Leaf $csv)" `
                -Environment "$($result.Directory)"
        } catch {
            $c.LblStatus.Text = 'Failed.'
            $c.TxtDetail.Text = Format-ErrorDetail -ErrorRecord $_ -Context $c.Title
            Write-FailedLog -Action 'Report' -Target $c.Prefix -ErrorRecord $_
        } finally {
            Complete-ProgressReporter -Activity $c.Title
            $c.Form.Cursor    = $prev
            $c.BtnRun.Enabled = $true
        }
    })

    $form.StartPosition = 'Manual'
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    # AFTER the size is final -- Set-FormResizable takes the current size as the
    # minimum, so calling it earlier would pin the window to whatever it was
    # part-way through layout.
    Set-FormResizable -Form $form
    Set-ConsoleTheme -Form $form
    Add-ToolHelpButton -Form $form -ScriptPath $ScriptPath -Title $Title | Out-Null

    [void]$form.ShowDialog()
}

#endregion
