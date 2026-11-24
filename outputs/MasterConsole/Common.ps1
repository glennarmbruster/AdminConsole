#Requires -Version 5.1
<#
    Common.ps1
    Shared helper library for the Admin Console launcher and every
    standalone tool script (MailboxHealth.ps1, MessageTrace.ps1,
    InboxRules.ps1, etc.).

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

#region Constants (moved from the WinForms UI region, lines ~3672-3682) ------
$script:MailDomain   = 'expresspros.com'
$script:FranchiseDCs = @(
    'fran-vm-dc01.franchises.expresspersonnel.com'
    'fran-vm-dc02.franchises.expresspersonnel.com'
    'fran-vm-dc03.franchises.expresspersonnel.com'
    'fran-vm-dc04.franchises.expresspersonnel.com'
    'fran-vm-dc05.franchises.expresspersonnel.com'
    'fran-vm-dc06.franchises.expresspersonnel.com'
    'fran-vm-dc07.franchises.expresspersonnel.com'
    'fran-vm-dc08.franchises.expresspersonnel.com'
)
$script:HQDomain = 'expresspersonnel.com'
$script:ExchangeOnPremServer = 'exch-vm-hybrid1.expresspersonnel.com'
$script:FranchiseAutoDetectDC = 'fran-vm-dc01.franchises.expresspersonnel.com'

# Licence assignment is driven by membership of these on-prem groups, so
# membership IS the licence rather than a proxy for it. Shared here because
# more than one report needs them and the names must not drift apart.
# Names must match AD exactly -- a mismatch yields an empty result, not an
# error, so reports state which groups they could not find.
$script:LicenseGroups = @(
    [PSCustomObject]@{ Name = 'E1 License HQ';          Sku = 'E1'; Environment = 'HQ' }
    [PSCustomObject]@{ Name = 'E1 License Fran';        Sku = 'E1'; Environment = 'Franchise' }
    [PSCustomObject]@{ Name = 'E5 License HQ - Full';   Sku = 'E5'; Environment = 'HQ' }
    [PSCustomObject]@{ Name = 'E5 License Fran - Full'; Sku = 'E5'; Environment = 'Franchise' }
)

function ConvertTo-MatchableName {
    <#
        Normalises the characters that look identical in a group name but are
        not.

        'E5 License HQ - Full' typed in Word comes back with the hyphen turned
        into an EN DASH (U+2013), which renders almost the same as a real
        hyphen (U+002D). A name copied by eye then matches nothing, and the
        report says the group does not exist. Non-breaking spaces and smart
        quotes fail the same way and are just as invisible.

        This comment names the character rather than showing it. An earlier
        version pasted a real en dash in here to illustrate the point, which
        made a comment about invisible characters the only non-ASCII byte in
        the toolkit -- and the first thing an encoding round-trip would eat.
        See ARCHITECTURE 14i. The code below is the safe way to write one.
    #>
    param([string]$Text)
    if (-not $Text) { return '' }
    $t = "$Text"
    foreach ($d in @([char]0x2010, [char]0x2011, [char]0x2012, [char]0x2013,
                     [char]0x2014, [char]0x2015, [char]0x2212)) {
        $t = $t.Replace($d, '-')                       # dash family -> hyphen
    }
    $t = $t.Replace([char]0x00A0, ' ')                 # non-breaking space
    $t = $t.Replace([char]0x2018, "'").Replace([char]0x2019, "'")
    $t = $t.Replace([char]0x201C, '"').Replace([char]0x201D, '"')
    return ($t -replace '\s+', ' ').Trim().ToLower()
}

function Resolve-LicenseGroupObject {
    <#
        Finds a licence group by name in whichever forest holds it.

        Exact match first. If that fails, every group in the forest is compared
        on its NORMALISED name, so a Word-style dash or a stray non-breaking
        space cannot make a group invisible.

        A fallback match returns MatchedOn = 'normalised' and the real name, so
        callers can report the discrepancy rather than quietly papering over an
        inconsistency that is worth fixing at the source.
    #>
    param([string]$Name)

    $servers = @(
        [PSCustomObject]@{ Label = 'HQ';        Server = $script:HQDomain }
        [PSCustomObject]@{ Label = 'Franchise'; Server = $script:FranchiseAutoDetectDC }
    )

    foreach ($s in $servers) {
        if (-not $s.Server) { continue }
        try {
            $g = Get-ADGroup -Identity $Name -Server $s.Server -ErrorAction Stop
            if ($g) {
                return [PSCustomObject]@{
                    Group = $g; Server = $s.Server; Forest = $s.Label
                    MatchedOn = 'exact'; ActualName = "$($g.Name)"
                }
            }
        } catch { }
    }

    $target = ConvertTo-MatchableName $Name
    if (-not $target) { return $null }

    foreach ($s in $servers) {
        if (-not $s.Server) { continue }
        try {
            # Narrow the sweep with the first plain-text word so this is not a
            # full enumeration of every group in the forest.
            $stem = ($Name -split '[^\w]')[0]
            $filt = if ($stem) { "Name -like '$stem*'" } else { '*' }
            foreach ($g in (Get-ADGroup -Filter $filt -Server $s.Server -ErrorAction Stop)) {
                if ((ConvertTo-MatchableName $g.Name) -eq $target) {
                    return [PSCustomObject]@{
                        Group = $g; Server = $s.Server; Forest = $s.Label
                        MatchedOn = 'normalised'; ActualName = "$($g.Name)"
                    }
                }
            }
        } catch { }
    }
    return $null
}

function Get-LicensedUserMap {
    <#
        Returns @{ <lowercase DN> = @{ Groups=..; Skus=..; Environment=.. } }
        for every user in a licence group, plus the list of groups that could
        not be resolved.

        Built once and reused: several reports need "is this person licensed",
        and asking per user would be a call each.
    #>
    param([scriptblock]$Progress)

    $map      = @{}
    $notFound = @()
    $renamed  = @()
    $totals   = @{}

    $i = 0
    foreach ($lg in $script:LicenseGroups) {
        $i++
        if ($Progress) { & $Progress "Reading licence group '$($lg.Name)'..." }
        $found = Resolve-LicenseGroupObject -Name $lg.Name
        if (-not $found) { $notFound += $lg.Name; continue }

        # Surfaced here rather than left to each caller: the discrepancy is
        # discovered in this function, so throwing it away would mean every
        # caller either re-derives it or silently loses it.
        if ($found.MatchedOn -eq 'normalised') {
            $renamed += "$($lg.Name)  ->  actual AD name: $($found.ActualName)"
        }

        try {
            $members = @(Get-ADGroupMember -Identity $found.Group.DistinguishedName `
                            -Server $found.Server -Recursive -ErrorAction Stop |
                         Where-Object { $_.objectClass -eq 'user' })
        } catch { $notFound += "$($lg.Name) (unreadable)"; continue }

        $totals[$lg.Name] = $members.Count
        foreach ($m in $members) {
            $k = "$($m.DistinguishedName)".ToLower()
            if (-not $map.ContainsKey($k)) {
                $map[$k] = [PSCustomObject]@{
                    Groups      = $lg.Name
                    Skus        = $lg.Sku
                    Environment = $lg.Environment
                    Server      = $found.Server
                }
            } else {
                $e = $map[$k]
                $e.Groups = "$($e.Groups); $($lg.Name)"
                if ($e.Skus -notlike "*$($lg.Sku)*") { $e.Skus = "$($e.Skus); $($lg.Sku)" }
            }
        }
    }
    return [PSCustomObject]@{ Map = $map; NotFound = $notFound; Renamed = $renamed; Totals = $totals }
}
#endregion

#region Parsing (moved unchanged -- original lines 550-559) ------------------
function Resolve-ADAccountBySmtp {
    <#
        Finds the on-prem AD account behind a cloud SMTP address, across both
        forests, and reports whether it is enabled.

        Shared because two reports need the same answer and it must not drift:
        "is the person behind this cloud object still an employee" is decided by
        on-prem AD in this environment, the same way licence state is (see
        $script:LicenseGroups).

        Searched by mail, userPrincipalName AND proxyAddresses, because a cloud
        object's primary SMTP is frequently a proxy address on-prem rather than
        the mail attribute -- matching only on mail produces false "not found"
        results that read as "this person left" when they did not.

        Pass the same -Cache hashtable across a whole run. Group owners repeat
        heavily, and without caching a few hundred groups becomes thousands of
        identical directory lookups.

        Returns Found = $false with Searched populated when nothing matched, so
        the caller can distinguish "no such account" from "could not ask".
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [hashtable]$Cache
    )

    $key = $Smtp.ToLowerInvariant()
    if ($Cache -and $Cache.ContainsKey($key)) { return $Cache[$key] }

    # LDAP special characters, per RFC 4515. Rare in an address, fatal if present.
    $esc = $Smtp
    foreach ($p in @(@('\','\5c'), @('(','\28'), @(')','\29'), @('*','\2a'), @([char]0,'\00'))) {
        $esc = $esc.Replace($p[0], $p[1])
    }
    $filter = "(|(mail=$esc)(userPrincipalName=$esc)(proxyAddresses=smtp:$esc))"

    $servers = @($script:HQDomain) + @($script:FranchiseDCs)
    $result = [PSCustomObject]@{
        Smtp = $Smtp; Found = $false; Enabled = $null; SamAccountName = ''
        DisplayName = ''; Server = ''; Environment = ''; Error = ''
    }
    $errors = @()

    foreach ($srv in $servers) {
        if (-not $srv) { continue }
        try {
            $u = @(Get-ADUser -LDAPFilter $filter -Server $srv `
                        -Properties Enabled, mail, DisplayName -ErrorAction Stop)
        } catch { $errors += "$srv : $($_.Exception.Message)"; continue }

        if (@($u).Count -gt 0) {
            $hit = $u[0]
            $result.Found          = $true
            $result.Enabled        = [bool]$hit.Enabled
            $result.SamAccountName = "$($hit.SamAccountName)"
            $result.DisplayName    = "$($hit.DisplayName)"
            $result.Server         = $srv
            $result.Environment    = if ($srv -eq $script:HQDomain) { 'HQ' } else { 'Franchise' }
            break
        }
    }

    # Every directory refused: that is "could not ask", not "does not exist".
    if (-not $result.Found -and $errors.Count -eq @($servers).Count) {
        $result.Error = ($errors | Select-Object -First 1)
    }

    if ($Cache) { $Cache[$key] = $result }
    return $result
}

#region Configuration files ---------------------------------------------------
<#
    THE CODE IS MINE, THE CONFIGURATION IS YOURS.

    Every .json the console reads lives in Config\, and nothing in Config\ is
    ever shipped over. That makes the update instruction one rule with no
    exceptions to remember:

        replace the files, leave Config\ alone.

    A per-file exception list ("replace everything except these four, and don't
    miss Settings.json") is the kind of instruction that gets followed correctly
    nine times and then not. A folder boundary cannot be half-followed.

    Config\ holds:
        Tools.json      which tools appear, where, in what order
        Sections.json   the console's sections and row counts
        Themes.json     colour palettes
        Settings.json   mail domain, DC, folders, Edge profiles
        Images\         logos

    NOT in Config\: reports\, Logs\ and captures\. Those are OUTPUT, a different
    category again, and they are already relocatable through Settings.json.
    Folding them in here would blur the boundary this folder exists to make
    obvious.
#>
$script:ConfigFolderName = 'Config'
$script:ConfigFileNames  = @('Tools.json', 'Sections.json', 'Themes.json', 'Settings.json')

function Get-ConfigFolder {
    <#
        The Config folder, created on demand. ONE definition used by both reads
        and writes -- if the resolver could disagree with itself, ToolManager
        would write somewhere the console does not read and the two files would
        silently diverge.
    #>
    $dir = Join-Path $PSScriptRoot $script:ConfigFolderName
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
    return $dir
}

function Get-ConfigFilePath {
    <# Full path to a named config file inside Config\. #>
    param([Parameter(Mandatory)][string]$Name)
    return (Join-Path (Get-ConfigFolder) $Name)
}

function Get-ConfigImageFolder {
    <# Where logos live. Config\Images -- a logo is your content, not mine. #>
    $dir = Join-Path (Get-ConfigFolder) 'Images'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
    return $dir
}

function Invoke-ConfigMigration {
    <#
        Moves configuration left in the toolkit root into Config\, once.

        Runs on launch so an existing install fixes itself when the new files
        are copied over the top. Deliberately a MOVE, not a copy: leaving the old
        file behind would create exactly the two-locations-that-disagree problem
        this folder exists to prevent.

        A file already present in Config\ wins -- the root copy is then stale by
        definition, and is renamed rather than deleted so nothing is destroyed by
        a tool the user did not ask to reorganise anything.

        Returns what it did, so the console can say so rather than moving the
        user's files silently.
    #>
    $moved = @()
    $dir   = Get-ConfigFolder

    foreach ($n in $script:ConfigFileNames) {
        $old = Join-Path $PSScriptRoot $n
        $new = Join-Path $dir $n
        if (-not (Test-Path -LiteralPath $old)) { continue }
        try {
            if (Test-Path -LiteralPath $new) {
                $bak = "$old.superseded"
                Move-Item -LiteralPath $old -Destination $bak -Force -ErrorAction Stop
                $moved += "$n was already in Config\ -- the root copy is now $n.superseded"
            } else {
                Move-Item -LiteralPath $old -Destination $new -Force -ErrorAction Stop
                $moved += "$n moved into Config\"
            }
        } catch {
            $moved += "$n could not be moved: $($_.Exception.Message)"
        }
    }

    # Logos too, from the old Themes\Images.
    $oldImg = Join-Path $PSScriptRoot 'Themes\Images'
    if (Test-Path -LiteralPath $oldImg) {
        $newImg = Get-ConfigImageFolder
        foreach ($f in @(Get-ChildItem -LiteralPath $oldImg -File -ErrorAction SilentlyContinue)) {
            $dest = Join-Path $newImg $f.Name
            if (Test-Path -LiteralPath $dest) { continue }
            try {
                Move-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
                $moved += "$($f.Name) moved into Config\Images\"
            } catch { }
        }
    }
    return $moved
}
#endregion

#region Settings.json ---------------------------------------------------------
<#
    Locations and per-machine values, kept OUT of this file.

    Two reasons this is a separate file rather than constants here:

      * This file gets replaced when the toolkit is updated, and a setting
        stored in it disappears on every update -- silently, with a symptom that
        looks nothing like the cause.
      * Hardcoded absolute paths assume a drive letter exists. 'E:\Scripts\...'
        works on the machine it was written on and fails on a server built
        differently, at the moment it tries to write a log.

    The tenant constants above stay hardcoded on purpose -- this is the
    expresspros toolkit and that is the point of it. Settings.json only carries
    what varies by MACHINE.
#>
$script:SettingsFileName = 'Settings.json'

function Get-SettingsFilePath {
    return (Get-ConfigFilePath -Name $script:SettingsFileName)
}

function Get-ToolkitSettings {
    <#
        Reads Settings.json, or returns empty values if it is absent -- a
        missing file is the normal first-run state, not an error. Cached in a
        global rather than $script: so tools the console launches with '&' see
        the same answer instead of each re-reading it.
    #>
    param([switch]$Refresh)

    if ($global:ToolkitSettings -and -not $Refresh) { return $global:ToolkitSettings }

    $settings = [PSCustomObject]@{
        ReportsPath        = ''
        LogPath            = ''
        DisableRequestPath = ''
        OperatorEmail      = ''
        SpoAdminUrl        = ''
        ExchangeServerUri  = ''
        BrowserProfiles    = $null
    }
    $path = Get-SettingsFilePath

    if (Test-Path -LiteralPath $path) {
        try {
            $j = (Get-Content -LiteralPath $path -Raw -ErrorAction Stop) | ConvertFrom-Json -ErrorAction Stop
            foreach ($k in @('ReportsPath', 'LogPath', 'DisableRequestPath', 'OperatorEmail', 'SpoAdminUrl', 'ExchangeServerUri')) {
                if ($j.PSObject.Properties[$k]) { $settings.$k = "$($j.$k)".Trim() }
            }
            if ($j.PSObject.Properties['BrowserProfiles']) { $settings.BrowserProfiles = $j.BrowserProfiles }
        } catch {
            Write-Warning "Settings.json could not be read and is being ignored: $($_.Exception.Message)"
        }
    }
    $global:ToolkitSettings = $settings
    return $settings
}

function Save-ToolkitSettings {
    param(
        [string]$ReportsPath = '',
        [string]$LogPath = '',
        [string]$DisableRequestPath = '',
        [string]$OperatorEmail = '',
        [string]$SpoAdminUrl = '',
        [string]$ExchangeServerUri = '',
        [hashtable]$BrowserProfiles
    )
    $profiles = [ordered]@{}
    if ($BrowserProfiles) {
        foreach ($k in ($BrowserProfiles.Keys | Sort-Object)) { $profiles[$k] = "$($BrowserProfiles[$k])".Trim() }
    }
    $obj = [ordered]@{
        ReportsPath        = "$ReportsPath".Trim()
        LogPath            = "$LogPath".Trim()
        DisableRequestPath = "$DisableRequestPath".Trim()
        OperatorEmail      = "$OperatorEmail".Trim()
        SpoAdminUrl        = "$SpoAdminUrl".Trim()
        ExchangeServerUri  = "$ExchangeServerUri".Trim()
        BrowserProfiles    = $profiles
    }
    $json = ([PSCustomObject]$obj | ConvertTo-Json -Depth 3)
    $enc  = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText((Get-SettingsFilePath), ($json -replace "`r?`n", "`r`n"), $enc)
    $global:ToolkitSettings = $null
    return (Get-ToolkitSettings -Refresh)
}

function Resolve-ToolkitFolder {
    <#
        A configured folder, or one beside the toolkit. A relative configured
        value is relative to the TOOLKIT, not to whatever directory the shell
        happened to be sitting in when the console launched.

        An unreachable share falls back beside the toolkit with a warning rather
        than taking the tool down -- failing to write a log should not stop the
        work the log was recording.
    #>
    param([string]$Configured, [Parameter(Mandatory)][string]$DefaultLeaf)

    $path = if ($Configured) {
                if ([System.IO.Path]::IsPathRooted($Configured)) { $Configured }
                else { Join-Path $PSScriptRoot $Configured }
            } else { Join-Path $PSScriptRoot $DefaultLeaf }

    if (-not (Test-Path -LiteralPath $path)) {
        try { New-Item -Path $path -ItemType Directory -Force | Out-Null }
        catch {
            Write-Warning "Could not create '$path' ($($_.Exception.Message)). Falling back to the toolkit folder."
            $path = Join-Path $PSScriptRoot $DefaultLeaf
            if (-not (Test-Path -LiteralPath $path)) { New-Item -Path $path -ItemType Directory -Force | Out-Null }
        }
    }
    return $path
}

function Get-ReportPath {
    <#
        Where report CSVs are written. ONE definition -- this was previously
        written out in ten separate report scripts, the same
        one-value-many-copies shape that let a duplicated domain controller
        drift out of step.
    #>
    return (Resolve-ToolkitFolder -Configured (Get-ToolkitSettings).ReportsPath -DefaultLeaf 'reports')
}

function Get-LogPath {
    <# Where CompletedTasks.csv is written. #>
    return (Resolve-ToolkitFolder -Configured (Get-ToolkitSettings).LogPath -DefaultLeaf 'Logs')
}

function Get-DisableRequestPath {
    <#
        Folder holding the disable-request scripts and their queue CSV.

        Returned WITHOUT creating it, unlike the folders above: this one is
        expected to already exist and hold specific files. Creating an empty
        folder because the setting was wrong would turn "your path is wrong"
        into "the queue is empty", which reads as good news.
    #>
    $c = (Get-ToolkitSettings).DisableRequestPath
    if ($c) {
        if ([System.IO.Path]::IsPathRooted($c)) { return $c }
        return (Join-Path $PSScriptRoot $c)
    }
    return (Join-Path $PSScriptRoot 'DisableRequests')
}
function Get-OperatorEmail {
    <#
        Where a tool sends the operator's own copy of anything it mails out.

        This exists because the address was hardcoded to one person, so whoever
        ran the tool was never the one who received the copy -- which turned the
        copy from a check on your own work into mail about somebody else's.

        Settings.json first, then the running account's own mail attribute in
        AD. Returns '' when neither answers, and callers must treat that as
        "send no copy" rather than falling back to a default: a copy going
        quietly to the wrong person is worse than no copy at all.
    #>
    $configured = "$((Get-ToolkitSettings).OperatorEmail)".Trim()
    if ($configured) { return $configured }

    if ($global:ToolkitOperatorEmailProbed) { return "$($global:ToolkitOperatorEmail)" }
    $global:ToolkitOperatorEmailProbed = $true
    $global:ToolkitOperatorEmail = ''
    try {
        $sam = $env:USERNAME
        if ($sam) {
            $u = Get-ADUser -Identity $sam -Properties mail -ErrorAction Stop
            if ($u -and $u.mail) { $global:ToolkitOperatorEmail = "$($u.mail)" }
        }
    } catch { }
    return "$($global:ToolkitOperatorEmail)"
}

#endregion

function Get-TargetSmtp {
    param([string]$Text)
    if (-not $Text) { return $null }
    $t = $Text.Trim()
    if ($t -eq '') { return $null }
    if ($t -like '*@*') { return $t }
    return "$t@$($script:MailDomain)"
}
#endregion

#region Error formatting (moved unchanged -- original lines 561-587) --------
function Show-HandlerError {
    <#
        Surfaces an error from a WinForms event handler as something readable.

        An exception that escapes a Click handler is caught by .NET, not by
        PowerShell, and you get a CLR crash dialog: a hundred lines of
        System.Management.Automation internals, the loaded assembly list, and
        NOT the one thing you need -- which script and which line. This puts the
        PowerShell view back.

        Also writes a Failed row to the audit log, because an error the operator
        dismissed is otherwise the one kind of failure that leaves no trace.
    #>
    param(
        [Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [string]$Context = 'Action'
    )
    $inv  = $ErrorRecord.InvocationInfo
    $line = if ($inv) { "$($inv.ScriptName):$($inv.ScriptLineNumber)" } else { '(unknown location)' }
    $stmt = if ($inv -and $inv.Line) { "$($inv.Line)".Trim() } else { '' }

    $msg = "$Context failed." + [Environment]::NewLine + [Environment]::NewLine +
           "$($ErrorRecord.Exception.Message)" + [Environment]::NewLine + [Environment]::NewLine +
           "Where : $line"
    if ($stmt) { $msg += [Environment]::NewLine + "Line  : $stmt" }
    if ($ErrorRecord.FullyQualifiedErrorId) {
        $msg += [Environment]::NewLine + "Id    : $($ErrorRecord.FullyQualifiedErrorId)"
    }

    [void][System.Windows.Forms.MessageBox]::Show($msg, 'Error', 'OK', 'Error')
    try {
        Write-CompletedLog -Action 'Error' -Target $Context -Result 'Failed' `
            -Detail "$($ErrorRecord.Exception.Message) [$line]"
    } catch { }
}

function Format-ErrorDetail {
    param(
        [Parameter(Mandatory)] $ErrorRecord,
        [string]$Context = ''
    )
    $ex   = $ErrorRecord.Exception
    $cmd  = if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.MyCommand) {
                $ErrorRecord.InvocationInfo.MyCommand.Name
            } else { '(cmdlet unknown)' }
    $line = if ($ErrorRecord.InvocationInfo) { $ErrorRecord.InvocationInfo.ScriptLineNumber } else { '?' }
    $sb = New-Object System.Text.StringBuilder
    $prefix = if ($Context) { "ERROR [$Context]" } else { 'ERROR' }
    [void]$sb.AppendLine("$prefix -- $cmd (script line $line)")
    [void]$sb.AppendLine("    Message  : $($ex.Message)")
    [void]$sb.AppendLine("    Type     : $($ex.GetType().FullName)")
    if ($ErrorRecord.CategoryInfo) {
        [void]$sb.AppendLine("    Category : $($ErrorRecord.CategoryInfo.Category) ($($ErrorRecord.CategoryInfo.Reason))")
    }
    if ($ex.InnerException) {
        [void]$sb.AppendLine("    Inner    : $($ex.InnerException.Message)")
    }
    return $sb.ToString().TrimEnd()
}
#endregion

#region EXO cmdlet resolution (moved unchanged -- original lines 588-625) ---
function Get-EXOCommand {
    <#
        Resolve a cmdlet to the Exchange Online copy.

        Originally matched on module NAME -- tmpEXO_* then
        ExchangeOnlineManagement. That worked, but it hard-codes an
        implementation detail of implicit remoting: the generated module name.
        It is not contractual and has changed shape between module versions.

        Now delegates to Get-ExchangeModuleMap, which identifies the two worlds
        by what they EXPORT (on-premises exports Get-RemoteMailbox; the cloud
        does not). That is a property of Exchange rather than of the module
        loader.

        Name and signature kept so existing call sites are untouched. The
        name-based match survives as a fallback for an exotic load that the
        capability test does not recognise.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $map = Get-ExchangeModuleMap
    if (-not $map.Cloud) { $map = Get-ExchangeModuleMap -Refresh }
    if ($map.Cloud) {
        $cmd = Get-Command -Name "$($map.Cloud.Name)\$Name" -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd }
    }

    $all = @(Get-Command -Name $Name -All -ErrorAction SilentlyContinue)
    if ($all.Count -eq 0) {
        throw "No command named '$Name' exists in this session. Is Exchange Online connected (Connect-ExchangeOnline)?"
    }
    $cmd = $all | Where-Object { $_.ModuleName -like 'tmpEXO_*' -or $_.Source -like 'tmpEXO_*' } | Select-Object -First 1
    if (-not $cmd) {
        $cmd = $all | Where-Object { $_.ModuleName -eq 'ExchangeOnlineManagement' -or $_.Source -eq 'ExchangeOnlineManagement' } | Select-Object -First 1
    }
    if (-not $cmd) {
        $sources = ($all | ForEach-Object { "$($_.Name) [ModuleName=$($_.ModuleName); Source=$($_.Source)]" }) -join '  |  '
        throw "Found '$Name' but none of the loaded copies look like Exchange Online: $sources -- run 'Get-Command $Name -All | Select Name,ModuleName,Source' to confirm which one this should be; an on-prem Exchange module/snap-in may be shadowing it."
    }
    return $cmd
}
function Invoke-EXOCommand {
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    $cmd = Get-EXOCommand -Name $Name
    & $cmd @Parameters
}

function Get-OnPremCommand {
    <#
        The counterpart nobody wrote, which is why the hybrid scripts qualified
        their cloud calls and left their on-premises ones bare.

        An asymmetric guard is a guard with a hole in it: the cloud half was
        pinned down and the on-premises half went wherever import order sent
        it.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $map = Get-ExchangeModuleMap
    if (-not $map.OnPrem) { $map = Get-ExchangeModuleMap -Refresh }
    if (-not $map.OnPrem) {
        throw "'$Name' was requested against on-premises Exchange, which is not connected in this process. Refusing to run it against Exchange Online instead."
    }
    $cmd = Get-Command -Name "$($map.OnPrem.Name)\$Name" -ErrorAction SilentlyContinue
    if (-not $cmd) {
        throw "'$Name' is not exported by the on-premises Exchange module ($($map.OnPrem.Name))."
    }
    return $cmd
}
function Invoke-OnPremCommand {
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    $cmd = Get-OnPremCommand -Name $Name
    & $cmd @Parameters
}
#endregion

#region Object-type resolution (moved unchanged -- original lines 628-668) --
function Resolve-Target {
    param(
        [string]$Smtp,
        [string]$ADServer
    )
    $result = [ordered]@{
        Type              = 'NotFound'
        Smtp              = $Smtp
        ADGroupDN         = $null
        ResolutionErrors  = @()
    }
    try {
        Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'Mailbox'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (EXO) resolution attempt')
    }
    try {
        Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'UnifiedGroup'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroup (EXO) resolution attempt')
    }
    try {
        $grp = @(Get-ADGroup -Filter "mail -eq '$Smtp'" -Server $ADServer `
                    -Properties mail, ManagedBy, GroupCategory, GroupScope, msExchHideFromAddressLists `
                    -ErrorAction Stop)
        if ($grp.Count -eq 1) {
            $result.Type      = 'OnpremGroup'
            $result.ADGroupDN = $grp[0].DistinguishedName
            return [PSCustomObject]$result
        } elseif ($grp.Count -gt 1) {
            $result.ResolutionErrors += "ERROR [Get-ADGroup (on-prem) resolution attempt] -- $($grp.Count) AD groups matched mail '$Smtp'. Expected 0 or 1. Resolve the duplicate mail attribute before using this tool."
        }
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADGroup (on-prem) resolution attempt')
    }
    return [PSCustomObject]$result
}
#endregion

#region Inbox rule read (moved unchanged -- original lines 798-806) ---------
function Get-InboxRulesForMailbox {
    param([string]$Smtp)
    try {
        $rules = @(Invoke-EXOCommand -Name 'Get-InboxRule' -Parameters @{ Mailbox = $Smtp; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ Rules = $rules; Error = $null }
    } catch {
        return [PSCustomObject]@{ Rules = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-InboxRule') }
    }
}
#endregion

#region Completed-tasks logging ----------------------------------------------
#
# Schema v2. The v1 file was Task,DateTime -- one free-text sentence and a
# timestamp to the minute. That records that something happened; it does not let
# you ask what happened to a given mailbox, who did it, or whether it actually
# worked. Everything below exists to answer one of those three questions.
#
# Three deliberate choices:
#
#   Timestamp is ISO 8601 WITH OFFSET. The old format had no seconds, so two
#   actions in the same minute could not be ordered, and no zone, so logs from
#   tenants in different countries could not be merged without guessing. ISO
#   8601 also sorts correctly as plain text, which matters because this is a CSV
#   somebody will open in Notepad.
#
#   Result is its own column. Previously a partial success was recorded as
#   English inside the sentence ("-- 4 applied"), which is invisible to a filter.
#   Failures are logged too: a log that only records successes cannot be used to
#   reconstruct a bad afternoon, which is the one job an audit trail has.
#
#   RunId groups the rows of a single operation. ResourceDuplicator writes one
#   row for the create and another for the settings, and without a shared id
#   there is nothing tying them together -- or tying either to the failure row
#   when the third step dies.
#
$script:LogColumns = @(
    'Timestamp', 'Operator', 'Machine', 'Script',
    'Action', 'Target', 'Result', 'Detail', 'RunId', 'Environment'
)
$script:LogHeaderLine = ($script:LogColumns -join ',')

# Controlled vocabulary. Free text is kept in Detail; Action exists so that
# "how many mailboxes did we disable last month" is a filter rather than a grep.
$script:LogActionMap = @{
    'add' = 'Add'; 'added' = 'Add'; 'bulk-added' = 'Add'
    'remove' = 'Remove'; 'removed' = 'Remove'; 'bulk-removed' = 'Remove'
    'delete' = 'Delete'; 'deleted' = 'Delete'
    'create' = 'Create'; 'created' = 'Create'
    'enable' = 'Enable'; 'enabled' = 'Enable'
    'disable' = 'Disable'; 'disabled' = 'Disable'
    'grant' = 'Grant'; 'granted' = 'Grant'
    'revoke' = 'Revoke'; 'revoked' = 'Revoke'
    'set' = 'Configure'; 'configured' = 'Configure'; 'edited' = 'Configure'
    'restore' = 'Restore'; 'restored' = 'Restore'
    'released' = 'Release'
    'ran' = 'Run'
}

function Get-LogOperator {
    <#
        DOMAIN\user -- the identity that actually held the token, not a display
        name. On a jump box where several people share a console this is the only
        thing that attributes the change to a person.
    #>
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        if ($id) { return $id }
    } catch { }
    if ($env:USERDOMAIN) { return "$($env:USERDOMAIN)\$($env:USERNAME)" }
    return "$($env:USERNAME)"
}

function Get-LogScriptName {
    <#
        The tool that made the call, not this file. Common.ps1 is dot-sourced, so
        it appears in the call stack for every logged action and has to be walked
        past.
    #>
    try {
        foreach ($frame in (Get-PSCallStack)) {
            if (-not $frame.ScriptName) { continue }
            $leaf = Split-Path -Leaf $frame.ScriptName
            if ($leaf -eq 'Common.ps1') { continue }
            return $leaf
        }
    } catch { }
    return 'unknown'
}

function Get-ToolRunId {
    <#
        One id per tool, for the life of the process.

        Held in a GLOBAL keyed by script name rather than a $script: variable,
        because tools are launched with & inside the console's own process: they
        share globals but each gets its own script scope. A $script: variable
        here would be $null in every tool, and a single global without the key
        would give every tool the same id -- which is worse than none, because it
        would falsely group unrelated actions.
    #>
    param([string]$ScriptName)
    if (-not $global:ToolkitRunIds) { $global:ToolkitRunIds = @{} }
    $key = $ScriptName
    if (-not $key) { $key = 'unknown' }
    if (-not $global:ToolkitRunIds.ContainsKey($key)) {
        $global:ToolkitRunIds[$key] = [guid]::NewGuid().ToString('N').Substring(0, 12)
    }
    return $global:ToolkitRunIds[$key]
}

function ConvertTo-CsvField {
    <#
        RFC 4180 quoting. The v1 writer used a bare -f format string, so a task
        containing a double quote produced a row that no CSV reader could parse
        -- and the corruption was invisible until someone opened the file months
        later. Newlines are flattened for the same reason.
    #>
    param([string]$Value)
    if ($null -eq $Value) { $Value = '' }
    $v = $Value -replace '\r?\n', ' '
    $v = $v -replace '"', '""'
    return '"' + $v + '"'
}

function Split-LegacyTask {
    <#
        Best-effort Action and Target from a v1-style sentence.

        Deliberately conservative: it returns blank rather than a guess. The full
        sentence always survives in Detail, so a blank column costs nothing,
        whereas a wrong one is a filter that quietly returns the wrong answer --
        which is worse than no filter at all.
    #>
    param([string]$Task)
    $out = @{ Action = ''; Target = ''; Environment = '' }
    if (-not $Task) { return $out }

    $first = ($Task.Trim() -split '\s+')[0]
    if ($first) {
        $k = $first.ToLower()
        if ($script:LogActionMap.ContainsKey($k)) { $out.Action = $script:LogActionMap[$k] }
    }

    # An SMTP address is the only token that can be identified as the target
    # without understanding the sentence.
    $m = [regex]::Match($Task, '[\w.+''-]+@[\w-]+\.[\w.-]+')
    if ($m.Success) { $out.Target = $m.Value }

    # The tools already write "(Franchise)", "(HQ)" or "(EXO)" into the sentence.
    # That is an existing convention at a known position, so reading it back is
    # recovery rather than guesswork.
    $e = [regex]::Match($Task, '\((HQ|Franchise|EXO|On-Prem[^);]*)')
    if ($e.Success) { $out.Environment = $e.Groups[1].Value.Trim() }

    return $out
}

function Get-CompletedLogPath {
    <#
        Returns the file to append to, or $null when the caller must be warned.

        A v1 file is NOT rewritten or appended to. Export-Csv -Append throws on a
        header mismatch in 5.1, and silently widening someone's existing audit
        file is not a decision a logging function should make on its own -- so v2
        rows go to a clearly named sibling until the old file is moved aside.
    #>
    $folder = Get-LogPath
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }
    $primary = Join-Path $folder 'CompletedTasks.csv'

    if (-not (Test-Path -LiteralPath $primary)) {
        $script:LogHeaderLine | Out-File -FilePath $primary -Encoding utf8
        return $primary
    }

    try { $head = (Get-Content -LiteralPath $primary -TotalCount 1 -ErrorAction Stop) } catch { $head = '' }
    if ("$head".Trim() -eq $script:LogHeaderLine) { return $primary }

    $fallback = Join-Path $folder 'CompletedTasks-v2.csv'
    if (-not (Test-Path -LiteralPath $fallback)) {
        $script:LogHeaderLine | Out-File -FilePath $fallback -Encoding utf8
        Write-Warning ("CompletedTasks.csv uses the old Task,DateTime format. " +
            "Writing to CompletedTasks-v2.csv instead. Rename or archive the old " +
            "file and the original name will be used again.")
    }
    return $fallback
}

function Write-CompletedLog {
    <#
        .SYNOPSIS
        Appends one row to the completed-tasks audit log.

        .DESCRIPTION
        -Task is the v1 parameter and still works: it lands in Detail, and Action
        and Target are derived from it where that can be done safely. New callers
        should pass -Action, -Target and -Result explicitly, which is the only way
        Result can be anything other than Success.
    #>
    param(
        [string]$Task,
        [string]$Action,
        [string]$Target,
        [ValidateSet('Success', 'Partial', 'Failed')][string]$Result = 'Success',
        [string]$Detail,
        [string]$Environment,
        [string]$RunId
    )

    if (-not $Detail) { $Detail = $Task }
    if (-not $Detail -or $Detail.Trim() -eq '') {
        if (-not $Action -and -not $Target) { return }
    }

    if (-not $Action -or -not $Target -or -not $Environment) {
        $parsed = Split-LegacyTask -Task $Detail
        if (-not $Action)      { $Action      = $parsed.Action }
        if (-not $Target)      { $Target      = $parsed.Target }
        if (-not $Environment) { $Environment = $parsed.Environment }
    }

    $scriptName = Get-LogScriptName
    if (-not $RunId) { $RunId = Get-ToolRunId -ScriptName $scriptName }
    if (-not $Environment) { $Environment = Get-LogEnvironmentLabel }

    # Round-tripping through DateTimeOffset keeps the machine's real UTC offset,
    # including whichever side of a DST change today falls on.
    $ts = [DateTimeOffset]::Now.ToString('yyyy-MM-ddTHH:mm:sszzz')

    $values = @(
        $ts, (Get-LogOperator), $env:COMPUTERNAME, $scriptName,
        $Action, $Target, $Result, "$Detail".Trim(), $RunId, $Environment
    )
    $line = (($values | ForEach-Object { ConvertTo-CsvField -Value $_ }) -join ',')

    $file = Get-CompletedLogPath
    if (-not $file) { return }
    try { Add-Content -Path $file -Value $line -ErrorAction Stop }
    catch { Write-Warning "Could not write audit log: $($_.Exception.Message)" }
}

function Write-FailedLog {
    <#
        Convenience wrapper so an error path is one line. Without this the catch
        blocks stay empty, which is exactly how a log ends up recording only
        successes.
    #>
    param(
        [string]$Action,
        [string]$Target,
        [string]$Detail,
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    $d = $Detail
    if ($ErrorRecord) {
        $msg = $ErrorRecord.Exception.Message
        if ($d) { $d = "$d -- $msg" } else { $d = $msg }
    }
    Write-CompletedLog -Action $Action -Target $Target -Result 'Failed' -Detail $d
}
function Get-LogEnvironmentLabel {
    <#
        A fallback only. Environment cannot be derived reliably at log time --
        the operator's own forest is not necessarily the target's, and this
        console reaches both. Callers that know should pass -Environment; the
        legacy parser recovers it from the "(HQ)" / "(Franchise)" convention the
        tools already follow. This is what is left when neither applies.
    #>
    if ($global:ToolkitEnvironment) { return "$($global:ToolkitEnvironment)" }
    if ($env:USERDNSDOMAIN) { return $env:USERDNSDOMAIN }
    return ''
}
#endregion

#region NEW -- parameter-based AD server resolution --------------------------
function Get-ADServerString {
    param(
        [Parameter(Mandatory)][ValidateSet('Franchise', 'HQ')][string]$Environment,
        [string]$DomainController
    )
    if ($Environment -eq 'HQ') { return $script:HQDomain }
    return $DomainController
}
#endregion

#region NEW -- auto-detect HQ vs Franchise -----------------------------------
function Find-TargetEnvironment {
    param([string]$Smtp)
    $errors = @()
    $hqServer = Get-ADServerString -Environment 'HQ' -DomainController $null
    try {
        $hqHit = @(Get-ADObject -Filter "mail -eq '$Smtp'" -Server $hqServer -ErrorAction Stop)
        if ($hqHit.Count -gt 0) {
            return [PSCustomObject]@{ Environment = 'HQ'; ADServer = $hqServer; Errors = $errors }
        }
    } catch {
        $errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'HQ AD lookup (environment auto-detect)')
    }
    $franServer = Get-ADServerString -Environment 'Franchise' -DomainController $script:FranchiseAutoDetectDC
    try {
        $franHit = @(Get-ADObject -Filter "mail -eq '$Smtp'" -Server $franServer -ErrorAction Stop)
        if ($franHit.Count -gt 0) {
            return [PSCustomObject]@{ Environment = 'Franchise'; ADServer = $franServer; Errors = $errors }
        }
    } catch {
        $errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Franchise AD lookup (environment auto-detect)')
    }
    return [PSCustomObject]@{ Environment = $null; ADServer = $null; Errors = $errors }
}
#endregion

#region NEW -- auto-detect HQ vs Franchise BY sAMAccountName ------------------
function Find-UserEnvironmentBySam {
    param([Parameter(Mandatory)][string]$SamAccountName)

    $errors = @()
    $sam = $SamAccountName.Trim()

    $hqServer = Get-ADServerString -Environment 'HQ' -DomainController $null
    try {
        $hqHit = Get-ADUser -Identity $sam -Server $hqServer -ErrorAction Stop
        if ($hqHit) {
            return [PSCustomObject]@{
                Environment = 'HQ'; ADServer = $hqServer; DCParam = ''
                User = $hqHit; Errors = $errors
            }
        }
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
    } catch {
        $errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'HQ AD lookup (sAMAccountName auto-detect)')
    }

    $franServer = Get-ADServerString -Environment 'Franchise' -DomainController $script:FranchiseAutoDetectDC
    try {
        $franHit = Get-ADUser -Identity $sam -Server $franServer -ErrorAction Stop
        if ($franHit) {
            return [PSCustomObject]@{
                Environment = 'Franchise'; ADServer = $franServer
                DCParam = $script:FranchiseAutoDetectDC
                User = $franHit; Errors = $errors
            }
        }
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
    } catch {
        $errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Franchise AD lookup (sAMAccountName auto-detect)')
    }

    return [PSCustomObject]@{
        Environment = $null; ADServer = $null; DCParam = $null
        User = $null; Errors = $errors
    }
}
#endregion

#region NEW -- consistent window placement -----------------------------------
function Set-FormScreenCorner {
    param(
        [Parameter(Mandatory)]$Form,
        [ValidateSet('TopLeft', 'BottomLeft')][string]$Corner = 'TopLeft'
    )
    try {
        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual

        $x = $wa.Left
        if ($Corner -eq 'TopLeft') {
            $y = $wa.Top
        } else {
            $y = $wa.Bottom - $Form.Height
            if ($y -lt $wa.Top) { $y = $wa.Top }
        }
        $Form.Location = New-Object System.Drawing.Point($x, $y)
    } catch {
        $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    }
}
#endregion

#region Theming --------------------------------------------------------------
# Deliberately NOT $script:-scoped state for the path.
#
# Set-ConsoleTheme is called from tool scripts that AdminConsole launches with
# the call operator. Those run in their OWN script scope, so a $script:
# variable assigned when Common.ps1 was dot-sourced into the console is $null
# by the time the function runs inside the child -- which fed $null to
# Test-Path -LiteralPath and threw. $PSScriptRoot inside a function always
# resolves to the file that DEFINED the function, so it is scope-proof.
$script:ThemeCache = $null

function Get-ThemeFilePath {
    return (Get-ConfigFilePath -Name 'Themes.json')
}

function ConvertTo-ThemeColor {
    param([string]$Hex)
    if ([string]::IsNullOrWhiteSpace($Hex)) { return $null }
    try { return [System.Drawing.ColorTranslator]::FromHtml($Hex) } catch { return $null }
}

function Get-ConsoleTheme {
    <#
        The resolved theme is cached in a GLOBAL, not a $script: variable.

        Tool scripts are launched by the console with the call operator, so
        they run in their own script scope and see none of the console's
        $script: state -- including the tenant-domain cache this resolution
        depends on. When the tenant lookup came up empty in a child, $Name fell
        back to ActiveTheme, which is normally the disabled 'Default', and the
        tool silently opened unthemed. A global is inherited by children, so
        the console's answer travels with them.

        It also means Get-AcceptedDomain is called once per session rather than
        once per tool launch.
    #>
    param([string]$Name, [switch]$Refresh)

    if ($Refresh) {
        $script:ThemeCache = $null
        $global:AdminConsoleTheme = $null
    }
    if (-not $Name -and $global:AdminConsoleTheme) { return $global:AdminConsoleTheme }
    if (-not $script:ThemeCache) {
        $themeFile = Get-ThemeFilePath
        if (-not $themeFile) { return $null }
        if (-not (Test-Path -LiteralPath $themeFile)) { return $null }
        try {
            $script:ThemeCache = Get-Content -LiteralPath $themeFile -Raw |
                                 ConvertFrom-Json
        } catch { return $null }
    }
    $cfg = $script:ThemeCache

    if (-not $Name) {
        $Name = $cfg.ActiveTheme
        if ($cfg.FollowTenant) {
            $d = Get-ConnectedTenantDomain
            if ($d -and $cfg.TenantThemes.PSObject.Properties.Name -contains $d) {
                $Name = $cfg.TenantThemes.$d
            }
        }
    }
    if (-not $Name) { return $null }
    if ($cfg.Themes.PSObject.Properties.Name -notcontains $Name) { return $null }

    $t = $cfg.Themes.$Name
    if (-not $t.Enabled) { return $null }
    Add-Member -InputObject $t -NotePropertyName 'Name' -NotePropertyValue $Name -Force
    $global:AdminConsoleTheme = $t
    return $t
}

function Get-ShadedColor {
    <#
        Nudges a colour lighter or darker by a fixed amount, staying in range.

        Direction is chosen from the base colour's own brightness: a dark theme
        gets a LIGHTER stripe, a light theme gets a darker one. A fixed
        direction would produce an invisible stripe on half the themes, which is
        the same as having no stripe at all.
    #>
    param(
        [Parameter(Mandatory)][System.Drawing.Color]$Base,
        [int]$Amount = 12
    )
    $lift = (($Base.R + $Base.G + $Base.B) / 3) -lt 128
    $d = if ($lift) { $Amount } else { -$Amount }
    $clamp = { param($v) [Math]::Max(0, [Math]::Min(255, [int]$v)) }
    return [System.Drawing.Color]::FromArgb(
        (& $clamp ($Base.R + $d)), (& $clamp ($Base.G + $d)), (& $clamp ($Base.B + $d)))
}

function Get-BlendedColor {
    <#
        Mixes two colours. Amount is how far to move from -From toward -To,
        0.0 to 1.0. Used to dim read-only text toward its own background, which
        reads as deliberately inactive at any brightness -- unlike a hardcoded
        grey, which disappears on a dark theme and shouts on a light one.
    #>
    param(
        [Parameter(Mandatory)][System.Drawing.Color]$From,
        [Parameter(Mandatory)][System.Drawing.Color]$To,
        [double]$Amount = 0.5
    )
    if ($Amount -lt 0) { $Amount = 0 } elseif ($Amount -gt 1) { $Amount = 1 }
    $mix = { param($a, $b) [int][Math]::Round(($a * (1 - $Amount)) + ($b * $Amount)) }
    return [System.Drawing.Color]::FromArgb(
        (& $mix $From.R $To.R), (& $mix $From.G $To.G), (& $mix $From.B $To.B))
}

function Set-FormDoubleBuffered {
    param([Parameter(Mandatory)]$Form)
    try {
        [System.Windows.Forms.Control].GetProperty('DoubleBuffered',
            'Instance,NonPublic').SetValue($Form, $true, $null)
    } catch { }
}

function Add-FormGradient {
    param([Parameter(Mandatory)]$Form, [string]$FromHex, [string]$ToHex)
    $c1 = ConvertTo-ThemeColor $FromHex
    $c2 = ConvertTo-ThemeColor $ToHex
    if (-not $c1 -or -not $c2) { return }
    Set-FormDoubleBuffered -Form $Form
    $Form.Add_Paint({
        param($src, $e)
        try {
            $r = $src.ClientRectangle
            if ($r.Width -le 0 -or $r.Height -le 0) { return }
            $b = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                    $r, $c1, $c2,
                    [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
            $e.Graphics.FillRectangle($b, $r)
            $b.Dispose()
        } catch { }
    }.GetNewClosure())
}

function Set-FlatStyleSafely {
    <#
        Sets FlatStyle only where the control actually has it.

        Not every control does -- TextBox, NumericUpDown and ProgressBar among
        them -- and assigning a property that does not exist is a terminating
        error that takes the whole theme pass down. The symptom is a tool that
        refuses to open with a message about a property, which reads like
        corruption rather than a styling problem.
    #>
    param([Parameter(Mandatory)]$Control)
    if ($Control.PSObject.Properties['FlatStyle']) {
        $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    }
}

function Set-ConsoleTheme {
    param(
        [Parameter(Mandatory)]$Form,
        $Theme,
        [switch]$NoGradient
    )
    if (-not $Theme) { $Theme = Get-ConsoleTheme }
    if (-not $Theme) { return }

    $formBack   = ConvertTo-ThemeColor $Theme.FormBack
    $formText   = ConvertTo-ThemeColor $Theme.FormText
    $panelBack  = ConvertTo-ThemeColor $Theme.PanelBack
    $groupText  = ConvertTo-ThemeColor $Theme.GroupText
    $btnBack    = ConvertTo-ThemeColor $Theme.ButtonBack
    $btnText    = ConvertTo-ThemeColor $Theme.ButtonText
    $btnBorder  = ConvertTo-ThemeColor $Theme.ButtonBorder
    $btnHover   = ConvertTo-ThemeColor $Theme.ButtonHover
    $btnDown    = ConvertTo-ThemeColor $Theme.ButtonDown
    $inputBack  = ConvertTo-ThemeColor $Theme.InputBack
    $inputText  = ConvertTo-ThemeColor $Theme.InputText

    # Selected radio/checkbox text. Falls back to the button border colour,
    # which is already the theme's accent, so an older Themes.json without this
    # role still gets a visible selection rather than none.
    $selText = ConvertTo-ThemeColor $Theme.SelectedText
    if (-not $selText) { $selText = ConvertTo-ThemeColor $Theme.ButtonBorder }

    # ServiceNow green. Defaults are used when the theme does not define them,
    # so the colour is consistent across every theme rather than disappearing
    # on the ones that predate this.
    $snBack   = ConvertTo-ThemeColor $Theme.ServiceNowBack
    $snText   = ConvertTo-ThemeColor $Theme.ServiceNowText
    $snHover  = ConvertTo-ThemeColor $Theme.ServiceNowHover
    $snBorder = ConvertTo-ThemeColor $Theme.ServiceNowBorder
    if (-not $snBack)   { $snBack   = [System.Drawing.Color]::FromArgb(10, 92, 66) }
    if (-not $snText)   { $snText   = [System.Drawing.Color]::FromArgb(255, 255, 255) }
    if (-not $snHover)  { $snHover  = [System.Drawing.Color]::FromArgb(14, 120, 86) }
    if (-not $snBorder) { $snBorder = [System.Drawing.Color]::FromArgb(22, 160, 114) }

    if ($formBack) { $Form.BackColor = $formBack }
    if ($formText) { $Form.ForeColor = $formText }

    if (-not $NoGradient -and $Theme.GradientTo -and $Theme.FormBack) {
        Add-FormGradient -Form $Form -FromHex $Theme.FormBack -ToHex $Theme.GradientTo
    }

    $apply = {
        param($parent)
        foreach ($c in $parent.Controls) {
            # Tag-based opt-out. Controls whose colour carries MEANING -- the
            # ExO/SPO connection state, the ServiceNow highlight -- set
            # Tag = 'NoTheme' and are skipped, so a theme can never disguise a
            # red "not connected" as decoration.
            if ("$($c.Tag)" -like '*NoTheme*') { continue }

            switch -Regex ($c.GetType().Name) {
                # Anchored. 'Button' unanchored also matches RadioButton, and
                # switch -Regex falls through every matching branch without a
                # break -- so a RadioButton ran the button styling AND its own,
                # and anything reaching a FlatStyle line that lacks the property
                # threw. Set-FlatStyleSafely below is the second guard.
                '^Button$' {
                    if ($btnBack) {
                        Set-FlatStyleSafely -Control $c
                        $c.BackColor = $btnBack
                        if ($btnText)   { $c.ForeColor = $btnText }
                        if ($btnBorder) { $c.FlatAppearance.BorderColor = $btnBorder }
                        if ($btnHover)  { $c.FlatAppearance.MouseOverBackColor = $btnHover }
                        if ($btnDown)   { $c.FlatAppearance.MouseDownBackColor = $btnDown }
                    }
                    # ServiceNow buttons get the ServiceNow green so the most
                    # used destination is findable without reading labels.
                    # Matched on the button TEXT, not Tag -- tool buttons carry
                    # their tool object in Tag, so that slot is taken.
                    if ($snBack -and "$($c.Text)" -match '(?i)service\s*now') {
                        Set-FlatStyleSafely -Control $c
                        $c.BackColor = $snBack
                        if ($snText)   { $c.ForeColor = $snText }
                        if ($snHover)  { $c.FlatAppearance.MouseOverBackColor = $snHover }
                        if ($snBorder) { $c.FlatAppearance.BorderColor = $snBorder }
                    }
                }
                'GroupBox' {
                    $c.BackColor = [System.Drawing.Color]::Transparent
                    if ($groupText) { $c.ForeColor = $groupText }
                }
                '^(RadioButton|CheckBox)$' {
                    # WinForms draws the radio dot and tick with system colours,
                    # which vanish against a dark BackColor -- the control is
                    # readable but you cannot tell which option is selected.
                    # Colour and weight carry the state instead of the glyph.
                    $c.BackColor = [System.Drawing.Color]::Transparent
                    Set-FlatStyleSafely -Control $c

                    # Fonts built once per control and stashed on it. Creating a
                    # Font on every state change would leak a GDI handle each
                    # time the user clicks.
                    Add-Member -InputObject $c -NotePropertyName '_FontRegular' `
                        -NotePropertyValue $c.Font -Force
                    Add-Member -InputObject $c -NotePropertyName '_FontBold' `
                        -NotePropertyValue (New-Object System.Drawing.Font($c.Font, [System.Drawing.FontStyle]::Bold)) -Force
                    Add-Member -InputObject $c -NotePropertyName '_ClrOn'  -NotePropertyValue $selText  -Force
                    Add-Member -InputObject $c -NotePropertyName '_ClrOff' -NotePropertyValue $formText -Force

                    # $this inside the handler is the control that raised it, so
                    # one shared scriptblock serves every radio and checkbox.
                    $c.Add_CheckedChanged({
                        if ($this.Checked) {
                            if ($this._ClrOn)   { $this.ForeColor = $this._ClrOn }
                            if ($this._FontBold) { $this.Font = $this._FontBold }
                        } else {
                            if ($this._ClrOff)      { $this.ForeColor = $this._ClrOff }
                            if ($this._FontRegular) { $this.Font = $this._FontRegular }
                        }
                    })

                    # Apply the current state now, so whatever is already
                    # selected is visible before the user touches anything.
                    if ($c.Checked) {
                        if ($selText) { $c.ForeColor = $selText }
                        $c.Font = $c._FontBold
                    } elseif ($formText) {
                        $c.ForeColor = $formText
                    }
                }
                '^(Label|LinkLabel)$' {
                    $c.BackColor = [System.Drawing.Color]::Transparent
                    if ($formText) { $c.ForeColor = $formText }
                }
                '^(TextBox|RichTextBox|ComboBox|ListBox|MaskedTextBox|NumericUpDown)$' {
                    if ($inputBack) { $c.BackColor = $inputBack }
                    if ($inputText) { $c.ForeColor = $inputText }
                }
                '^(ListView|TreeView)$' {
                    if ($inputBack) { $c.BackColor = $inputBack }
                    if ($inputText) { $c.ForeColor = $inputText }
                }
                '^DataGridView$' {
                    <#
                        A DataGridView does NOT paint its cells from BackColor
                        and ForeColor. Setting only those themed the strip of
                        empty space below the last row and left every cell on
                        its own defaults -- which is why ToolManager's 32 rows
                        were dark text on a dark background and needed
                        squinting. Cells come from DefaultCellStyle, headers
                        from ColumnHeadersDefaultCellStyle, and neither inherits
                        from the control.

                        EnableHeadersVisualStyles = $false is the non-obvious
                        one: while it is $true, Windows draws the headers with
                        the system visual style and silently ignores every
                        header colour set here.
                    #>
                    if ($inputBack) {
                        $c.BackgroundColor = $inputBack
                        $c.DefaultCellStyle.BackColor = $inputBack
                    }
                    if ($inputText) { $c.DefaultCellStyle.ForeColor = $inputText }

                    # Zebra striping. With three dozen rows this does more for
                    # readability than the text colour does -- it gives the eye
                    # a row boundary to track along.
                    if ($inputBack) {
                        $c.AlternatingRowsDefaultCellStyle.BackColor =
                            (Get-ShadedColor -Base $inputBack -Amount 14)
                        if ($inputText) { $c.AlternatingRowsDefaultCellStyle.ForeColor = $inputText }
                    }

                    if ($btnBack) { $c.DefaultCellStyle.SelectionBackColor = $btnBack }
                    $selFore = if ($selText) { $selText } elseif ($btnText) { $btnText } else { $null }
                    if ($selFore) { $c.DefaultCellStyle.SelectionForeColor = $selFore }

                    $c.EnableHeadersVisualStyles = $false
                    if ($panelBack) {
                        $c.ColumnHeadersDefaultCellStyle.BackColor = $panelBack
                        $c.RowHeadersDefaultCellStyle.BackColor    = $panelBack
                        $c.ColumnHeadersDefaultCellStyle.SelectionBackColor = $panelBack
                        $c.RowHeadersDefaultCellStyle.SelectionBackColor    = $panelBack
                    }
                    $hdrFore = if ($groupText) { $groupText } elseif ($formText) { $formText } else { $null }
                    if ($hdrFore) {
                        $c.ColumnHeadersDefaultCellStyle.ForeColor = $hdrFore
                        $c.RowHeadersDefaultCellStyle.ForeColor    = $hdrFore
                        $c.ColumnHeadersDefaultCellStyle.SelectionForeColor = $hdrFore
                        $c.RowHeadersDefaultCellStyle.SelectionForeColor    = $hdrFore
                    }
                    if ($btnBorder) { $c.GridColor = $btnBorder }

                    # Read-only columns get dimmed text so "you cannot edit
                    # this" reads as deliberate rather than as a rendering
                    # fault -- the same complaint that started this fix.
                    if ($inputText -and $inputBack) {
                        $dim = Get-BlendedColor -From $inputText -To $inputBack -Amount 0.45
                        foreach ($col in @($c.Columns)) {
                            if ($col.ReadOnly) { $col.DefaultCellStyle.ForeColor = $dim }
                        }
                    }
                }
                '^(Panel|FlowLayoutPanel|TableLayoutPanel|TabPage|SplitterPanel)$' {
                    if ($panelBack) { $c.BackColor = $panelBack }
                    if ($formText)  { $c.ForeColor = $formText }
                }
            }
            if ($c.Controls -and $c.Controls.Count -gt 0) { & $apply $c }
        }
    }
    & $apply $Form
}

function Get-CmdletParameterNames {
    <#
        Returns the parameter names a cmdlet actually declares.

        Exists because guessing is unreliable: -PageSize and
        -RecipientTypeDetails were both assumed onto Get-EXOCasMailbox and both
        rejected. The ExchangeOnlineManagement surface varies by version and
        tenant, so the command object in front of us is the only source of
        truth -- not the documentation.
    #>
    param([string]$Name)
    try {
        $cmd = Get-EXOCommand -Name $Name
        if ($cmd.Parameters) { return @($cmd.Parameters.Keys) }
    } catch { }
    try {
        $cmd = Get-Command -Name $Name -ErrorAction Stop
        if ($cmd.Parameters) { return @($cmd.Parameters.Keys) }
    } catch { }
    return @()
}

function Invoke-EXOBulkQuery {
    <#
        A bulk Exchange Online query that survives the two things that reliably
        break one: an unsupported property, and a dropped connection.

          * Only parameters the installed cmdlet declares are passed.
          * "InvalidProperties = X, Y" is parsed, the offenders dropped, retried.
          * Transport faults retry with exponential backoff.

        Note this does NOT fix an oversized response -- retrying an enormous
        query reproduces an enormous query. Callers dealing with thousands of
        objects should slice the request (by recipient type, by alias letter)
        and call this per slice.
    #>
    param(
        [Parameter(Mandatory)][string]$Cmdlet,
        [string[]]$Properties,
        [hashtable]$Extra = @{},
        [scriptblock]$Progress,
        [int]$MaxAttempts = 5
    )
    $supported = Get-CmdletParameterNames -Name $Cmdlet
    $remaining = @($Properties)

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $params = @{ ErrorAction = 'Stop' }
        if ($supported -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }
        if ($remaining -and $supported -contains 'Properties') { $params['Properties'] = $remaining }
        foreach ($k in $Extra.Keys) {
            if ($supported -contains $k) { $params[$k] = $Extra[$k] }
        }
        try {
            return @(Invoke-EXOCommand -Name $Cmdlet -Parameters $params)
        } catch {
            $msg = "$($_.Exception.Message)"

            $m = [regex]::Match($msg, 'InvalidProperties\s*=\s*(.+?)\s*$')
            if ($m.Success) {
                $bad = $m.Groups[1].Value -split '\s*,\s*' | Where-Object { $_ }
                $remaining = @($remaining | Where-Object { $bad -notcontains $_ })
                if (-not $remaining) { throw }
                if ($Progress) { & $Progress "Dropped unsupported: $($bad -join ', '). Retrying..." }
                continue
            }

            $transient = 'underlying connection was closed|operation has timed out|' +
                         'unexpected error occurred on a receive|task was canceled|' +
                         'Unable to read data from the transport|forcibly closed'
            if (($msg -match $transient) -and $attempt -lt $MaxAttempts) {
                $wait = [math]::Pow(2, $attempt)
                if ($Progress) { & $Progress "Connection dropped. Retrying in $wait s..." }
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
    throw "Could not query $Cmdlet after $MaxAttempts attempts."
}

function Get-AcceptedDomainList {
    <#
        Accepted domains for the connected tenant, lowercased.

        "External" is only meaningful relative to this list -- a forward to
        another accepted domain is internal mail, not exfiltration.
    #>
    param([switch]$Refresh)
    if (-not $Refresh -and $global:AdminConsoleAcceptedDomains) {
        return $global:AdminConsoleAcceptedDomains
    }
    try {
        # ResultSize is mandatory here, not tidiness. Get-AcceptedDomain
        # defaults to 1000 and a tenant with more than that gets a SILENTLY
        # TRUNCATED list plus a warning on the host stream. Truncation is the
        # dangerous half: this list is what "external" is measured against, so
        # a domain past the cut-off would make internal mail read as
        # exfiltration in the External Forwarding report.
        $d = @(Invoke-EXOCommand -Name 'Get-AcceptedDomain' -Parameters @{
                   ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
               ForEach-Object { "$($_.DomainName)".ToLower() } | Where-Object { $_ })
        if ($d.Count -gt 0) {
            $global:AdminConsoleAcceptedDomains = $d
            return $d
        }
    } catch { }
    return @()
}

function Get-SmtpDomain {
    param([string]$Address)
    if (-not $Address) { return '' }
    $a = "$Address"
    # Recipient values arrive in several shapes: plain SMTP, "Display Name
    # [SMTP:user@dom]", or an X500/legacy DN. Pull the last @-part of whatever
    # is there rather than assuming a clean address.
    $m = [regex]::Match($a, '([A-Za-z0-9._%+\-]+)@([A-Za-z0-9.\-]+\.[A-Za-z]{2,})')
    if ($m.Success) { return $m.Groups[2].Value.ToLower() }
    return ''
}

function New-ProgressReporter {
    <#
        Returns a scriptblock that reports progress to every place at once:
        a WinForms label, an optional progress bar, and Write-Progress in the
        PowerShell window.

        Long reports look identical to hung reports. A run that queries
        thousands of mailboxes can sit silent for minutes, and the only honest
        difference between "working" and "stalled" is whether something is
        still changing on screen.

        Usage:
            $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Protocols'
            & $p 'Querying UserMailbox...'
            & $p 'Half way' 50          # optional percent

        DoEvents is called deliberately: these tools are single-threaded, so
        without it the form never repaints and Windows greys it out as
        "Not Responding" -- which looks exactly like the crash we are trying
        to disprove.
    #>
    param(
        $Label,
        $Bar,
        [string]$Activity = 'Working',
        [switch]$NoConsole
    )
    $state = [pscustomobject]@{
        Label     = $Label
        Bar       = $Bar
        Activity  = $Activity
        NoConsole = [bool]$NoConsole
        Started   = Get-Date
    }
    return {
        param([string]$Message, [int]$Percent = -1)
        $elapsed = (Get-Date) - $state.Started
        $stamp   = '{0:mm\:ss}' -f $elapsed

        if ($state.Label) {
            $state.Label.Text = "[$stamp] $Message"
        }
        if ($state.Bar) {
            if ($Percent -ge 0) {
                $state.Bar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                $state.Bar.Value = [Math]::Max(0, [Math]::Min(100, $Percent))
            } else {
                # Unknown duration -- marquee says "alive" without lying about
                # how far along it is.
                $state.Bar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
            }
        }
        if (-not $state.NoConsole) {
            if ($Percent -ge 0) {
                Write-Progress -Activity $state.Activity -Status $Message -PercentComplete $Percent
            } else {
                Write-Progress -Activity $state.Activity -Status $Message
            }
        }
        [System.Windows.Forms.Application]::DoEvents()
    }.GetNewClosure()
}

function Complete-ProgressReporter {
    param([string]$Activity = 'Working', $Bar)
    try { Write-Progress -Activity $Activity -Completed } catch { }
    if ($Bar) {
        $Bar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $Bar.Value = 100
    }
}

function Add-ReportProgressBar {
    <#
        Standard progress bar for report windows. Placed by the caller; this
        just keeps the styling consistent across every report.
    #>
    param(
        [Parameter(Mandatory)]$Form,
        [int]$X = 12, [int]$Y = 0, [int]$Width = 720, [int]$Height = 14
    )
    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Location = New-Object System.Drawing.Point($X, $Y)
    $pb.Size     = New-Object System.Drawing.Size($Width, $Height)
    $pb.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $pb.Value    = 0
    $pb.Tag      = 'NoTheme'
    $Form.Controls.Add($pb)
    return $pb
}

#region Report files ---------------------------------------------------------
<#
    Reading a report must not require running one.

    Every report window used to keep "Open Folder" and "Open CSV" disabled until
    a run completed, so looking at yesterday's output meant regenerating it --
    which on the slower reports is minutes of Exchange calls to read a file that
    was already on disk. The reports directory is the archive; these controls
    expose it.
#>

function Open-ReportFolder {
    <#
        Opens the reports directory, creating it if this is a first run.

        The directory is created rather than reported missing because an empty
        folder is the honest answer to "show me my reports" before any have been
        written -- and explorer.exe on a non-existent path fails silently, which
        would look like a broken button.
    #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
    Start-Process explorer.exe $Path
}

function Get-ReportFileList {
    <#
        Previous report runs, newest first.

        Sorting is by the timestamp in the FILENAME, not LastWriteTime, because
        the filename records when the report was *run*. A file that has been
        copied, restored from backup or touched by a sync client carries a
        mistaken mtime, and sorting by it would silently reorder history. Files
        that don't parse fall back to LastWriteTime rather than being dropped.

        Row counts are deliberately NOT computed here. Counting rows means
        opening every file on every refresh, which turns a folder of large CSVs
        into a visible stall. The caller reads the count for the ONE file the
        user selects.
    #>
    param(
        [Parameter(Mandatory)][string]$Directory,
        [string]$Prefix
    )
    if (-not (Test-Path -LiteralPath $Directory)) { return @() }

    $pattern = if ($Prefix) { "$Prefix`_*.csv" } else { '*.csv' }
    $files = @(Get-ChildItem -LiteralPath $Directory -Filter $pattern -File -ErrorAction SilentlyContinue)

    $out = foreach ($f in $files) {
        # Prefix_yyyy-MM-dd_HHmm.csv
        $stamp = $null
        if ($f.BaseName -match '_(\d{4}-\d{2}-\d{2})_(\d{2})(\d{2})$') {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParseExact(
                    "$($Matches[1]) $($Matches[2]):$($Matches[3])",
                    'yyyy-MM-dd HH:mm',
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [System.Globalization.DateTimeStyles]::None,
                    [ref]$parsed)) {
                $stamp = $parsed
            }
        }
        $fromName = ($null -ne $stamp)
        if (-not $fromName) { $stamp = $f.LastWriteTime }

        # Everything before the _yyyy-MM-dd_HHmm tail is the report name.
        $report = $f.BaseName -replace '_\d{4}-\d{2}-\d{2}_\d{4}$', ''

        [PSCustomObject]@{
            Path          = $f.FullName
            Name          = $f.Name
            Report        = $report
            RunAt         = $stamp
            StampFromName = $fromName
            SizeKB        = [math]::Round($f.Length / 1KB, 1)
        }
    }
    return @($out | Sort-Object RunAt -Descending)
}

function Get-ReportRowCount {
    <#
        Data rows in a CSV -- lines minus the header. Returns $null rather than
        0 if the file can't be read, because "0 rows" and "couldn't tell" must
        not display the same way.
    #>
    param([Parameter(Mandatory)][string]$Path)
    try {
        $n = 0
        $reader = [System.IO.File]::OpenText($Path)
        try { while ($null -ne $reader.ReadLine()) { $n++ } } finally { $reader.Dispose() }
        if ($n -le 1) { return 0 }
        return $n - 1
    } catch { return $null }
}

function Add-ReportFileControls {
    <#
        A "Previous runs" row for a report window: label, dropdown of prior
        runs of THIS report, and an Open button.

        Returns an object carrying Refresh() and SelectPath() so the caller can
        repopulate after writing a new file. State lives on the returned object,
        not in a $script: variable, because these controls are used by scripts
        the console launches with '&' -- where Common.ps1's script scope is not
        the caller's script scope and any $script: read would resolve to $null.

        The list refreshes on Form.Activated as well, so a file written by
        another report (or dropped in by hand) appears on returning to the
        window. That removes the need for a Refresh button.
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Form]$Form,
        [Parameter(Mandatory)][string]$ReportDir,
        [Parameter(Mandatory)][string]$Prefix,
        [int]$X = 12,
        [int]$Y = 0,
        [int]$Width = 720,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = 'Previous runs:'
    $lbl.Location = New-Object System.Drawing.Point($X, ($Y + 5))
    $lbl.Size     = New-Object System.Drawing.Size(90, 18)
    $Form.Controls.Add($lbl)

    $comboX = $X + 94
    $btnW   = 102
    $btnX   = $X + $Width - $btnW
    $comboW = $btnX - $comboX - 4

    $combo               = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = 'DropDownList'
    $combo.Location      = New-Object System.Drawing.Point($comboX, ($Y + 2))
    $combo.Size          = New-Object System.Drawing.Size($comboW, 22)
    $Form.Controls.Add($combo)

    $btn          = New-Object System.Windows.Forms.Button
    $btn.Text     = 'Open'
    $btn.Location = New-Object System.Drawing.Point($btnX, $Y)
    $btn.Size     = New-Object System.Drawing.Size($btnW, 28)
    $Form.Controls.Add($btn)

    $ctl = [PSCustomObject]@{
        Label       = $lbl
        Combo       = $combo
        OpenButton  = $btn
        Directory   = $ReportDir
        Prefix      = $Prefix
        StatusLabel = $StatusLabel
        Files       = @()
    }

    Add-Member -InputObject $ctl -MemberType ScriptMethod -Name Refresh -Value {
        $keep = $null
        if ($this.Combo.SelectedIndex -ge 0 -and $this.Combo.SelectedIndex -lt @($this.Files).Count) {
            $keep = $this.Files[$this.Combo.SelectedIndex].Path
        }

        $this.Files = @(Get-ReportFileList -Directory $this.Directory -Prefix $this.Prefix)
        $this.Combo.Items.Clear()

        if (@($this.Files).Count -eq 0) {
            [void]$this.Combo.Items.Add('(no previous runs in the reports folder)')
            $this.Combo.SelectedIndex = 0
            $this.Combo.Enabled       = $false
            $this.OpenButton.Enabled  = $false
            return
        }

        foreach ($f in $this.Files) {
            # A '~' marks a timestamp taken from the file's mtime because the
            # name didn't carry one -- so a copied or renamed file is visibly
            # less trustworthy in the list rather than silently equal.
            $mark = if ($f.StampFromName) { ' ' } else { '~' }
            [void]$this.Combo.Items.Add(
                ('{0}{1}    {2,8:N1} KB    {3}' -f $mark,
                    $f.RunAt.ToString('yyyy-MM-dd  HH:mm'), $f.SizeKB, $f.Name))
        }
        $this.Combo.Enabled      = $true
        $this.OpenButton.Enabled = $true

        $idx = 0
        if ($keep) {
            for ($i = 0; $i -lt @($this.Files).Count; $i++) {
                if ($this.Files[$i].Path -eq $keep) { $idx = $i; break }
            }
        }
        $this.Combo.SelectedIndex = $idx
    }

    Add-Member -InputObject $ctl -MemberType ScriptMethod -Name SelectPath -Value {
        param([string]$Path)
        $this.Refresh()
        for ($i = 0; $i -lt @($this.Files).Count; $i++) {
            if ($this.Files[$i].Path -eq $Path) { $this.Combo.SelectedIndex = $i; return }
        }
    }

    Add-Member -InputObject $ctl -MemberType ScriptMethod -Name SelectedPath -Value {
        if (@($this.Files).Count -eq 0) { return $null }
        $i = $this.Combo.SelectedIndex
        if ($i -lt 0 -or $i -ge @($this.Files).Count) { return $null }
        return $this.Files[$i].Path
    }

    # $ctl travels on .Tag, and these handlers deliberately do NOT use
    # .GetNewClosure().
    #
    # GetNewClosure binds a scriptblock to a NEW DYNAMIC MODULE. Local variables
    # are copied in -- which is why $ctl resolved -- but FUNCTION lookup inside
    # a module goes module scope then global, and never sees the functions this
    # file put in the caller's SCRIPT scope. So Get-ReportRowCount came back
    # "not recognized" from inside the handler while working fine everywhere
    # else. Same family as the $script: trap: the scope a scriptblock runs in is
    # not the scope it was written in.
    $btn.Tag   = $ctl
    $combo.Tag = $ctl

    $btn.Add_Click({
        $ctl = $this.Tag
        $p = $ctl.SelectedPath()
        if (-not $p) { return }
        if (-not (Test-Path -LiteralPath $p)) {
            # Deleted or moved since the list was built.
            if ($ctl.StatusLabel) { $ctl.StatusLabel.Text = 'That file is no longer there -- list refreshed.' }
            $ctl.Refresh()
            return
        }
        Start-Process $p
    })

    $combo.Add_SelectedIndexChanged({
        $ctl = $this.Tag
        if (-not $ctl) { return }
        if (-not $ctl.StatusLabel) { return }
        if (@($ctl.Files).Count -eq 0) { return }
        $i = $ctl.Combo.SelectedIndex
        if ($i -lt 0 -or $i -ge @($ctl.Files).Count) { return }
        $f = $ctl.Files[$i]

        # Clearing the combo resets SelectedIndex to -1, so this fires on every
        # Refresh -- including the one on Form.Activated. Reading a very large
        # CSV on each window focus would be a stall with no visible cause, so
        # above 20 MB the size is reported instead of the row count.
        if ($f.SizeKB -gt 20480) {
            $ctl.StatusLabel.Text = "Selected: $($f.Name) -- $([math]::Round($f.SizeKB/1024,1)) MB (too large to count rows quickly)."
            return
        }
        $rows = Get-ReportRowCount -Path $f.Path
        $ctl.StatusLabel.Text = if ($null -eq $rows) {
            "Selected: $($f.Name) -- row count unavailable."
        } else {
            "Selected: $($f.Name) -- $rows data row(s)."
        }
    })

    # The form's Tag holds a LIST: more than one report window could in
    # principle carry these controls, and overwriting Tag would silently
    # disconnect the first.
    if ($Form.Tag -is [System.Collections.IList]) { [void]$Form.Tag.Add($ctl) }
    else { $Form.Tag = [System.Collections.ArrayList]@($ctl) }

    $Form.Add_Activated({
        foreach ($c in @($this.Tag)) { if ($c -and $c.PSObject.Methods['Refresh']) { $c.Refresh() } }
    })

    $ctl.Refresh()
    return $ctl
}
#endregion

#region Tool capability detection --------------------------------------------
<#
    Grey out a tool the current window cannot actually run.

    This works only because AdminConsole launches tools with the call operator
    (& $path), so a tool runs in the SAME PROCESS as the console. Exchange
    Online and SharePoint sessions are per-process (see the architecture notes,
    section 4), so the console's connection state genuinely IS the tool's state.
    If tools were ever launched with Start-Process, every judgement here becomes
    a guess and this whole mechanism must be reconsidered.

    Requirements are DETECTED by reading each script rather than declared in
    Tools.json. A declared field would have to be maintained by hand for every
    tool and would silently go stale the moment a script's dependencies changed
    -- the same drift shape as the duplicated $DefaultDC. Reading the script
    keeps one source of truth: the code.

    KNOWN LIMITS, both deliberate:

    * Dot-sourced files are NOT followed. Every tool dot-sources Common.ps1,
      and Common.ps1 itself calls Get-ADUser, Get-SPOTenant and Exchange
      cmdlets -- following it would tag every tool as needing everything, which
      would grey out the entire console. A tool is judged on what IT calls.

    * Ordinary strings are NOT stripped, only comments and here-strings. A
      cmdlet named in a string is usually a real dependency (for example
      Get-CmdletParameterNames -Name 'Get-SPOSite'), so removing them would
      lose signal. A cmdlet named only in a comment WOULD be a false positive,
      which is why comments go.
#>

$script:CapabilityDefs = @(
    [PSCustomObject]@{
        Id       = 'ExO'
        Label    = 'Exchange Online'
        Patterns = @(
            '\b(?:Get|Set)-EXO[A-Za-z]+'
            '\bConnect-ExchangeOnline\b'
            '\bGet-ConnectionInformation\b'
            '\b(?:Get|Set)-Mailbox\b'
            '\bGet-CASMailbox\b'
            '\b(?:Get|Set|New|Remove)-(?:InboxRule|DistributionGroup|UnifiedGroup|MailContact)\b'
            '\bGet-UnifiedGroupLinks\b'
            '\b(?:Get|Add|Remove)-(?:MailboxPermission|RecipientPermission|MailboxFolderPermission)\b'
            '\bGet-(?:MailboxStatistics|MailboxAutoReplyConfiguration|MailboxRegionalConfiguration)\b'
            '\bGet-(?:Recipient|AcceptedDomain|OrganizationConfig)\b'
            # Open-ended on purpose: Microsoft versions these cmdlet names in
            # place. Get-MessageTrace was retired in favour of
            # Get-MessageTraceV2, and a trailing \b here matched neither the V2
            # form nor Get-MessageTraceDetailV2 -- so MessageTrace.ps1 scanned
            # as needing nothing at all. Assume any future V3 does the same.
            '\bGet-MessageTrace[A-Za-z0-9]*'
            '\b(?:Get|Release|Delete|Preview|Export)-QuarantineMessage[A-Za-z0-9]*'
            '\bStart-HistoricalSearch\b'
        )
    }
    [PSCustomObject]@{
        Id       = 'SPO'
        Label    = 'SharePoint Online'
        Patterns = @(
            '\b(?:Get|Set|Add|New|Remove|Connect|Disconnect|Test|Request)-SPO[A-Za-z]*'
            '\bConnect-PnPOnline\b'
            '\bGet-PnP[A-Za-z]+'
        )
    }
    [PSCustomObject]@{
        Id       = 'EMS'
        Label    = 'Exchange on-premises'
        Patterns = @(
            # Deliberately only the on-prem-EXCLUSIVE families. Get-Mailbox and
            # friends exist in both worlds and resolve by module load order, so
            # they prove nothing about which one is present.
            '\b(?:Get|Set|New|Remove|Enable|Disable)-RemoteMailbox\b'
            '\bGet-ExchangeServer\b'
            '\bUpdate-Recipient\b'
            # Set-ADServerSettings pins the Exchange shell to a domain
            # controller and exists ONLY on-premises. It is the marker for
            # tools that drive on-prem group membership: Add- and
            # Remove-DistributionGroupMember are exported by both worlds, so
            # they identify nothing, and a tool relying on them alone scans as
            # needing no connection at all and never greys out.
            '\bSet-ADServerSettings\b'
        )
    }
    [PSCustomObject]@{
        Id       = 'AD'
        Label    = 'On-premises Active Directory'
        Patterns = @(
            '\b(?:Get|Set|New|Remove|Add|Enable|Disable|Unlock|Search|Move|Rename)-AD[A-Za-z]+'
        )
    }
)

# Compiled once and reused. A [regex] built with the Compiled option costs
# more to construct and far less to run, which is the right trade for something
# invoked against every script in Tools.json.
$script:LineCommentRegex = New-Object System.Text.RegularExpressions.Regex(
    '^(?:[^''"#]|''(?:[^'']*''|[^'']*$)|"(?:[^"]*"|[^"]*$))*',
    [System.Text.RegularExpressions.RegexOptions]::Compiled)

function Remove-PowerShellComment {
    <#
        Strips block comments, line comments and here-strings from script text so
        a cmdlet mentioned in prose is not mistaken for a dependency.

        Here-strings go because they carry HTML and JavaScript in the web tools,
        and a '#' inside a CSS colour is not a comment. Line comments are removed
        only when the '#' is not inside a quoted string, so a URL fragment or a
        hex colour in ordinary code survives.
    #>
    param([Parameter(Mandatory)][string]$Text)

    # Here-strings first: their contents must not be parsed as anything.
    $t = [regex]::Replace($Text, "@`"[\s\S]*?`"@", '', 'Singleline')
    $t = [regex]::Replace($t,    "@'[\s\S]*?'@",   '', 'Singleline')
    # Block comments.
    $t = [regex]::Replace($t, '<#[\s\S]*?#>', '', 'Singleline')

    # One regex per line instead of a character-by-character loop.
    #
    # The loop it replaces ran in interpreted PowerShell over every character
    # of every registered script -- 438,000 of them across 31 tools -- on the
    # UI thread, before the console's form painted. That was most of the
    # cold-launch delay.
    #
    # Reads as: from the start of the line, consume anything that is not a
    # quote or a hash, OR a complete quoted string, OR an unterminated quote
    # running to end of line. The match therefore stops at the first hash that
    # is NOT inside a string, which is exactly where the loop cut.
    #
    # The unterminated-quote alternatives are not decoration. A double-quoted
    # string may legally span lines in PowerShell, so a line can open a quote
    # and never close it; without those branches the match would stop at the
    # opening quote and silently truncate real code.
    #
    # Verified line-for-line identical to the old loop across all 47 scripts
    # in the toolkit before replacing it.
    $out = New-Object System.Text.StringBuilder
    foreach ($line in ($t -split "`n")) {
        [void]$out.AppendLine($script:LineCommentRegex.Match($line).Value)
    }
    return $out.ToString()
}

function Get-CapabilityCachePath {
    <#
        Where the capability scan results are remembered between windows.

        LOCALAPPDATA, deliberately NOT the toolkit folder. The toolkit is meant
        to run from a share, and a cache file beside it would be written by
        several operators at once -- contention over a file that exists purely
        to save a second. Per-user and per-machine also matches what the cache
        describes: this machine's copy of these scripts.
    #>
    $dir = Join-Path $env:LOCALAPPDATA 'ExpressAdminToolkit'
    if (-not (Test-Path -LiteralPath $dir)) {
        try { New-Item -Path $dir -ItemType Directory -Force | Out-Null } catch { return $null }
    }
    return (Join-Path $dir 'capability-cache.json')
}

function Import-ScriptCapabilityCache {
    <#
        Loads the on-disk cache into the in-process one, at most once.

        Every failure path here is a silent fall-through to an empty cache. A
        corrupt or unreadable cache must cost a rescan, never a broken console
        -- this is an optimisation, and an optimisation that can fail the
        program is not one.
    #>
    if ($global:ToolCapabilityCacheLoaded) { return }
    $global:ToolCapabilityCacheLoaded = $true
    if (-not $global:ToolCapabilityCache) { $global:ToolCapabilityCache = @{} }

    $path = Get-CapabilityCachePath
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { return }
    try {
        $j = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        foreach ($p in $j.PSObject.Properties) {
            $global:ToolCapabilityCache[$p.Name] = [PSCustomObject]@{
                Path     = "$($p.Value.Path)"
                Found    = [bool]$p.Value.Found
                # @() matters: ConvertFrom-Json hands back a bare string for a
                # one-element array, and every caller does .Count on this.
                Requires = @($p.Value.Requires)
            }
        }
    } catch { }
}

function Save-ScriptCapabilityCache {
    <#
        Writes the cache back, only if something was actually scanned.

        Called once by the console after it has built its buttons rather than
        on every entry, so a cold start is one write instead of thirty-one.
        A tool that never calls this loses nothing but the saving.

        Keys carry the file's last-write ticks, so an edited script simply
        misses and rescans. Entries whose key no longer matches a current file
        are dropped on save, which is what stops the file growing forever as
        scripts are revised.
    #>
    if (-not $global:ToolCapabilityCacheDirty) { return }
    $path = Get-CapabilityCachePath
    if (-not $path) { return }

    try {
        $keep = @{}
        foreach ($k in $global:ToolCapabilityCache.Keys) {
            $parts = "$k" -split '\|'
            if ($parts.Count -ne 2) { continue }
            try {
                $fi = Get-Item -LiteralPath $parts[0] -ErrorAction Stop
                if ("$($fi.LastWriteTimeUtc.Ticks)" -eq $parts[1]) {
                    $keep[$k] = $global:ToolCapabilityCache[$k]
                }
            } catch { }
        }
        $json = [PSCustomObject]$keep | ConvertTo-Json -Depth 4
        $enc  = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($path, ($json -replace "`r?`n", "`r`n"), $enc)
        $global:ToolCapabilityCacheDirty = $false
    } catch { }
}

# Cmdlets that CREATE a connection rather than consume one. A script calling
# any of these provides that capability, so it is never greyed for lacking it.
#
# ANCHORED AT STATEMENT POSITION -- (?m)^\s* -- and that is the whole trick.
# Remove-PowerShellComment strips comments but NOT string literals, and several
# report tools carry an error message reading "...run Connect-SPOService
# first." An unanchored pattern matched that sentence and marked those tools as
# providing SharePoint, which would have left them enabled with no connection
# behind them. A mention is not a call; only a call starts a line.
$script:CapabilityProvides = @(
    # The optional '$x =' allows for the assignment form -- Connect-ExSE does
    # '$session = New-PSSession @params', which a bare start-of-line anchor
    # would miss, leaving the on-premises connect button greyed for wanting the
    # connection it creates.
    @{ Id = 'ExO'; Pattern = '(?m)^\s*(\$\w+\s*=\s*)?Connect-ExchangeOnline\b' }
    @{ Id = 'SPO'; Pattern = '(?m)^\s*(\$\w+\s*=\s*)?Connect-SPOService\b' }
    @{ Id = 'EMS'; Pattern = '(?m)^\s*(\$\w+\s*=\s*)?(Connect-ExchangeServer|New-PSSession)\b' }
)

function Get-ScriptCapabilityRequirement {

    <#
        Which capabilities a script needs, inferred from the cmdlets it calls.

        Cached by path AND last-write time, so editing a script re-scans it but
        a relaunch does not re-read two dozen files. The cache is global rather
        than $script: because the console and the tools it launches share a
        process but not a script scope (section 10), and it is BACKED BY DISK so
        a brand-new PowerShell window does not pay the scan either -- which is
        what it used to do, every time, on the UI thread before the form
        painted.
    #>
    param([Parameter(Mandatory)][string]$Path)

    Import-ScriptCapabilityCache

    try {
        $fi = Get-Item -LiteralPath $Path -ErrorAction Stop
    } catch {
        # A missing script is not an unmet capability -- it is a broken
        # Tools.json entry, which the console reports separately.
        return [PSCustomObject]@{ Path = $Path; Found = $false; Requires = @() }
    }

    $key = '{0}|{1}' -f $fi.FullName, $fi.LastWriteTimeUtc.Ticks
    if ($global:ToolCapabilityCache.ContainsKey($key)) {
        return $global:ToolCapabilityCache[$key]
    }

    $needs = @()
    try {
        $code = Remove-PowerShellComment -Text (Get-Content -LiteralPath $fi.FullName -Raw -ErrorAction Stop)
        foreach ($cap in $script:CapabilityDefs) {
            foreach ($p in $cap.Patterns) {
                if ([regex]::IsMatch($code, $p, 'IgnoreCase')) { $needs += $cap.Id; break }
            }
        }

        # A script that ESTABLISHES a connection cannot be gated on already
        # having it. Connect-ExO.ps1 calls Connect-ExchangeOnline, the scanner
        # read that as "needs Exchange Online", and the console greyed out the
        # one button that could have fixed it -- you could not connect because
        # you were not connected.
        #
        # Providing a capability outranks appearing to need it. Written as a
        # pattern rather than a filename so any future connect script inherits
        # this without anyone remembering to add it.
        foreach ($prov in $script:CapabilityProvides) {
            if ([regex]::IsMatch($code, $prov.Pattern, 'IgnoreCase')) {
                $needs = @($needs | Where-Object { $_ -ne $prov.Id })
            }
        }
    } catch { }

    $result = [PSCustomObject]@{
        Path     = $fi.FullName
        Found    = $true
        Requires = @($needs)
    }
    $global:ToolCapabilityCache[$key] = $result
    $global:ToolCapabilityCacheDirty = $true
    return $result
}

#region Tool help ---------------------------------------------------------------
<#
    A one-screen answer to "what does this tool do".

    Not documentation. The ScriptNotes files are documentation and are written
    for whoever maintains the toolkit. This is for the person who just clicked
    a button called "Returning User" and would like to know what happens next,
    before it happens.

    Two rules shape the design:

    1. THE TEXT LIVES IN THE SCRIPT. A Help.json would drift the first time
       somebody edited a script without editing the config, and would not
       survive the script being copied into the SRG or ANZ toolkit. A marker
       block in the header travels with the code:

           # TOOLKIT-HELP-BEGIN
           # Restores a user who was disabled in error.
           # TOOLKIT-HELP-END

    2. THE "RUNS AGAINST" LINE IS DERIVED, NOT TYPED. Get-ScriptCapabilityRequirement
       already classifies every script by which cmdlets it calls -- that is what
       greys the console buttons out. Reusing it means the help and the grey-out
       cannot disagree, and the line cannot go stale when the script changes.
       Typing it by hand would create a second source of truth for one fact,
       which is how the DC-selector and the write-scope bugs both happened.
#>

#region PIM activation timer ----------------------------------------------------
<#
    A local record of when PIM was last activated, and therefore how long is
    left before it has to be done again.

    THIS IS A NOTE, NOT A FACT. Entra is not consulted -- the tenant will not
    consent to the Graph permissions that would make it authoritative, and the
    portal does not put the Active-assignments tab in a URL either. What it
    knows is when the Activate button was pressed, plus a fixed duration.

    That is honest and useful because activation always goes through the
    button. If it ever does not, the countdown is wrong -- so the display
    always shows WHEN it was recorded, never a bare number. A stale timestamp
    is visibly stale; a bare "2h left" would not be.

    Stored per user under LOCALAPPDATA, NOT in Config\. Config ships to other
    engineers, and this is personal, machine-local state -- shipping it would
    hand someone else a countdown for an activation they never made.
#>

function Get-PimStatePath {
    $dir = Join-Path $env:LOCALAPPDATA 'ExpressAdminToolkit'
    if (-not (Test-Path -LiteralPath $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null } catch { }
    }
    return (Join-Path $dir 'pim-activation.json')
}

function Get-PimDurationHours {
    # 4 hours is this tenant's PIM policy. In Settings so a tenant with a
    # different maximum does not need a code change.
    $s = $null
    try { $s = Get-ToolkitSettings } catch { }
    if ($s -and $s.PSObject.Properties['PimDurationHours']) {
        $h = 0
        if ([int]::TryParse("$($s.PimDurationHours)", [ref]$h) -and $h -gt 0) { return $h }
    }
    return 4
}

function Get-PimActivation {
    <#
        Returns @{ Groups = <DateTimeOffset|null>; Roles = <DateTimeOffset|null> }
        -- the moment each was last recorded as activated.
    #>
    $out = @{ Groups = $null; Roles = $null }
    $p = Get-PimStatePath
    if (-not (Test-Path -LiteralPath $p)) { return $out }
    try {
        $j = Get-Content -LiteralPath $p -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        foreach ($k in 'Groups', 'Roles') {
            if ($j.PSObject.Properties[$k] -and "$($j.$k)".Trim()) {
                $dt = [DateTimeOffset]::MinValue
                # Round-trip format, parsed as round-trip. Parsing a stored
                # timestamp with local-culture rules is how a file written in
                # one region becomes unreadable in another.
                if ([DateTimeOffset]::TryParse("$($j.$k)", [System.Globalization.CultureInfo]::InvariantCulture,
                        [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$dt)) {
                    $out[$k] = $dt
                }
            }
        }
    } catch { }
    return $out
}

function Set-PimActivation {
    param([Parameter(Mandatory)][ValidateSet('Groups', 'Roles')][string]$Which)
    $state = Get-PimActivation
    $state[$Which] = [DateTimeOffset]::Now
    $obj = [PSCustomObject]@{
        Groups = if ($state.Groups) { $state.Groups.ToString('o') } else { $null }
        Roles  = if ($state.Roles)  { $state.Roles.ToString('o') }  else { $null }
    }
    try {
        $obj | ConvertTo-Json | Set-Content -LiteralPath (Get-PimStatePath) -Encoding UTF8 -ErrorAction Stop
    } catch { }
    return $state[$Which]
}

function Clear-PimActivation {
    param([Parameter(Mandatory)][ValidateSet('Groups', 'Roles')][string]$Which)
    $state = Get-PimActivation
    $state[$Which] = $null
    $obj = [PSCustomObject]@{
        Groups = if ($state.Groups) { $state.Groups.ToString('o') } else { $null }
        Roles  = if ($state.Roles)  { $state.Roles.ToString('o') }  else { $null }
    }
    try { $obj | ConvertTo-Json | Set-Content -LiteralPath (Get-PimStatePath) -Encoding UTF8 -ErrorAction Stop } catch { }
}

function Get-PimRemaining {
    <#
        Time left, or $null if never recorded. NEGATIVE when expired -- the
        caller decides how to say so, and "expired 20 minutes ago" is more
        useful than "expired".
    #>
    param([Parameter(Mandatory)][ValidateSet('Groups', 'Roles')][string]$Which)
    $at = (Get-PimActivation)[$Which]
    if (-not $at) { return $null }
    return ($at.AddHours((Get-PimDurationHours)) - [DateTimeOffset]::Now)
}

function Format-PimSpan {
    <#
        [Math]::Floor, NOT [int].

        The PowerShell [int] cast ROUNDS -- and rounds half to even -- it does
        not truncate. [int](239 / 60) is 4, so 3h59m displayed as "4h 59m" and
        looked like a five-hour timer on a four-hour policy. Reported
        2026-08-07 as "it's counting down from 5 hours".

        Anywhere a total is split into units, the quotient must be floored.
    #>
    param([TimeSpan]$Span)
    $m = [int][Math]::Floor([Math]::Abs($Span.TotalMinutes))
    $h = [int][Math]::Floor($m / 60); $m = $m % 60
    if ($h -gt 0) { return "{0}h {1:00}m" -f $h, $m }
    return "{0}m" -f $m
}

function Get-PimStatusLine {
    <#
        One short line for the console header. Empty string when nothing has
        ever been recorded, so the label can simply hide rather than showing
        a meaningless zero state.
    #>
    $parts = New-Object System.Collections.ArrayList
    foreach ($k in 'Groups', 'Roles') {
        $r = Get-PimRemaining -Which $k
        if ($null -eq $r) { continue }
        if ($r.TotalSeconds -le 0) { [void]$parts.Add("$k EXPIRED") }
        else { [void]$parts.Add("$k $(Format-PimSpan -Span $r)") }
    }
    if ($parts.Count -eq 0) { return '' }
    return ('PIM: ' + ($parts -join '  |  '))
}

function Get-PimStatusColour {
    <#
        Green with time to spare, amber under 30 minutes, red once expired.
        Thirty minutes because that is roughly the point at which starting
        something long becomes a bad idea.
    #>
    $worst = $null
    foreach ($k in 'Groups', 'Roles') {
        $r = Get-PimRemaining -Which $k
        if ($null -eq $r) { continue }
        if ($null -eq $worst -or $r -lt $worst) { $worst = $r }
    }
    if ($null -eq $worst)                { return [System.Drawing.Color]::DimGray }
    if ($worst.TotalSeconds -le 0)       { return [System.Drawing.Color]::FromArgb(255, 105, 97) }
    if ($worst.TotalMinutes -le 30)      { return [System.Drawing.Color]::FromArgb(230, 170, 60) }
    return [System.Drawing.Color]::FromArgb(80, 220, 130)
}
#endregion

#region Version and integrity ---------------------------------------------------
<#
    Two different questions, and they need two different answers.

    "WHAT VERSION ARE YOU ON?" is answered by a number -- Version.json at the
    toolkit root.

    "DOES YOUR COPY MATCH MINE?" is NOT answered by that number. A version
    string is a CLAIM: edit a script, forget to bump it, and the toolkit now
    reports a version it is not. The only honest answer is a per-file hash,
    generated from the master and compared on the copy. A hash cannot be wrong
    by omission -- forgetting to update it makes the check FAIL, which is the
    right direction to fail in.

    So Version.json carries both: one version, and a SHA-256 per code file.

    WHAT IS HASHED: .ps1 files, at the toolkit root and under Projects.
    WHAT IS NOT:    Config\ (Tools.json, Settings.json and friends are meant to
                    differ per engineer -- hashing them would report drift on
                    every machine and train everyone to ignore the result),
                    reports, logs, exports, and Version.json itself.

    Regenerate on the master with Publish-Toolkit.ps1 after any code
    change. That is a discipline, and the check compensates for slips: it
    reports the manifest as STALE when any .ps1 is newer than the manifest.
#>

function Get-ToolkitVersionPath {
    return (Join-Path $PSScriptRoot 'Version.json')
}

function Get-ToolkitVersion {
    <#
        Returns the version object, or a clearly-marked unknown. Never throws:
        a missing manifest is the normal state of a working copy that has not
        been stamped yet, not an error worth stopping the console for.
    #>
    $out = [PSCustomObject]@{
        Version  = 'unversioned'
        Released = $null
        Files    = @{}
        Found    = $false
    }
    $p = Get-ToolkitVersionPath
    if (-not (Test-Path -LiteralPath $p)) { return $out }
    try {
        $j = Get-Content -LiteralPath $p -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $files = @{}
        if ($j.PSObject.Properties['Files'] -and $j.Files) {
            foreach ($n in $j.Files.PSObject.Properties.Name) { $files[$n] = "$($j.Files.$n)" }
        }
        return [PSCustomObject]@{
            Version  = if ($j.PSObject.Properties['Version']) { "$($j.Version)" } else { 'unversioned' }
            Released = if ($j.PSObject.Properties['Released']) { "$($j.Released)" } else { $null }
            Files    = $files
            Found    = $true
        }
    } catch { return $out }
}

function Get-ToolkitCodeFile {
    <#
        The files the manifest covers. One definition, used by both the
        generator and the checker -- if these two ever disagreed, every copy
        would report drift that is not there.
    #>
    $root = $PSScriptRoot
    $list = New-Object System.Collections.ArrayList
    foreach ($f in (Get-ChildItem -LiteralPath $root -Filter '*.ps1' -File -ErrorAction SilentlyContinue)) {
        [void]$list.Add($f.FullName)
    }
    $proj = Join-Path $root 'Projects'
    if (Test-Path -LiteralPath $proj) {
        foreach ($f in (Get-ChildItem -LiteralPath $proj -Filter '*.ps1' -File -Recurse -ErrorAction SilentlyContinue)) {
            [void]$list.Add($f.FullName)
        }
    }
    return @($list.ToArray() | Sort-Object)
}

function Get-ToolkitRelativePath {
    param([Parameter(Mandatory)][string]$FullPath)
    $root = $PSScriptRoot.TrimEnd('\')
    if ($FullPath.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        return $FullPath.Substring($root.Length).TrimStart('\')
    }
    return (Split-Path -Leaf $FullPath)
}

function Get-FileHashHex {
    param([Parameter(Mandatory)][string]$Path)
    try { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant() }
    catch { return $null }
}

function Test-ToolkitIntegrity {
    <#
        Compares this copy against the manifest.

        Returns counts plus the actual lists, because "3 files differ" is a
        headline and "which three" is the answer.
    #>
    $v = Get-ToolkitVersion
    $modified = New-Object System.Collections.ArrayList
    $missing  = New-Object System.Collections.ArrayList
    $extra    = New-Object System.Collections.ArrayList
    $matched  = 0
    $stale    = $false

    if (-not $v.Found) {
        return [PSCustomObject]@{
            Version = $v.Version; Released = $v.Released; HasManifest = $false
            Matched = 0; Modified = @(); Missing = @(); Extra = @(); Stale = $false
            Summary = 'No Version.json -- this copy is not stamped.'
        }
    }

    $onDisk = @{}
    foreach ($f in (Get-ToolkitCodeFile)) { $onDisk[(Get-ToolkitRelativePath -FullPath $f)] = $f }

    foreach ($rel in $v.Files.Keys) {
        if (-not $onDisk.ContainsKey($rel)) { [void]$missing.Add($rel); continue }
        $h = Get-FileHashHex -Path $onDisk[$rel]
        if ($h -and $h -eq $v.Files[$rel]) { $matched++ } else { [void]$modified.Add($rel) }
    }
    foreach ($rel in $onDisk.Keys) {
        if (-not $v.Files.ContainsKey($rel)) { [void]$extra.Add($rel) }
    }

    # Timestamps, kept as an EXPLANATION and never as a verdict.
    #
    # A file newer than the manifest but hashing the same is byte-for-byte
    # identical -- touched, not changed. Treating that as a problem would turn
    # a perfect copy amber on any copy method that resets timestamps, which is
    # exactly the false alarm this region's header warns against. So Stale
    # never fails a clean check on its own; it only explains a failing one
    # ("that is why -- the master was edited and not re-stamped").
    try {
        $mTime = (Get-Item -LiteralPath (Get-ToolkitVersionPath)).LastWriteTimeUtc
        foreach ($f in $onDisk.Values) {
            if ((Get-Item -LiteralPath $f).LastWriteTimeUtc -gt $mTime) { $stale = $true; break }
        }
    } catch { }

    $summary =
        if ($modified.Count -eq 0 -and $missing.Count -eq 0 -and $extra.Count -eq 0) {
            "v$($v.Version) -- all $matched files match"
        } else {
            "v$($v.Version) -- $($modified.Count) modified, $($missing.Count) missing, $($extra.Count) extra"
        }
    if ($stale -and ($modified.Count -or $missing.Count -or $extra.Count)) {
        $summary += '  (manifest is older than the code -- re-stamp the master)'
    }

    return [PSCustomObject]@{
        Version = $v.Version; Released = $v.Released; HasManifest = $true
        Matched = $matched
        Modified = @($modified.ToArray()); Missing = @($missing.ToArray()); Extra = @($extra.ToArray())
        Stale = $stale
        Summary = $summary
    }
}
#endregion

function Get-ToolHelpText {
    <#
        Reads the TOOLKIT-HELP block out of a script.

        Returns $null when there is none, so callers can decide whether to
        show a button at all rather than showing an empty dialog.
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
    } catch { return $null }

    $out = New-Object System.Collections.ArrayList
    $inBlock = $false
    foreach ($l in $lines) {
        if ($l -match 'TOOLKIT-HELP-BEGIN') { $inBlock = $true; continue }
        if ($l -match 'TOOLKIT-HELP-END')   { break }
        if ($inBlock) {
            # Strip the leading '#' and ONE following space. Stripping all
            # leading whitespace would flatten deliberate indentation, and
            # these blocks use it for bullets.
            $text = $l -replace '^\s*#', ''
            if ($text.StartsWith(' ')) { $text = $text.Substring(1) }
            [void]$out.Add($text)
        }
    }
    if ($out.Count -eq 0) { return $null }
    return (($out -join "`r`n").Trim())
}

<#
    A SECOND, stricter classification -- used only by the help line.

    $script:CapabilityDefs answers "might this script need X?", and is
    deliberately generous: over-claiming makes the console grey a button that
    would have worked, which is a minor annoyance. Under-claiming lets someone
    click a tool that then fails, which is worse. Generous is correct there.

    Stating a fact to a person is the opposite problem, and reusing the
    generous list produced two wrong answers on the first test:

      * SetDLMembership.ps1 claimed Exchange Online, because Get-DistributionGroup
        is in the ExO pattern list. That cmdlet exists in BOTH worlds -- it is
        the exact ambiguity that cost most of 2026-08-06 -- and the tool is
        on-premises only.
      * HeaderAnalyzer.ps1 claimed Active Directory, because it contained a
        function called Get-AddressDomain and 'Get-AD' + 'dressDomain' matches
        the AD cmdlet pattern. The tool connects to nothing at all.

    So this list contains only markers that are EXCLUSIVE to one world, plus a
    whitelist of real AD nouns rather than an open-ended -AD* pattern. Anything
    ambiguous is reported as ambiguous, which is the honest answer and is also
    useful: "Exchange, whichever you are connected to" is a true statement about
    a tool that calls Get-Mailbox and nothing more specific.
#>
$script:HelpCapabilityDefs = @(
    [PSCustomObject]@{
        Id       = 'ExO'
        Label    = 'Exchange Online'
        Patterns = @(
            '\b(?:Get|Set)-EXO[A-Za-z]+'
            '\bConnect-ExchangeOnline\b'
            '\bGet-ConnectionInformation\b'
            '\bInvoke-EXOCommand\b'
            '\bGet-UnifiedGroup[A-Za-z]*'
            '\bGet-MessageTrace[A-Za-z0-9]*'
            '\b(?:Get|Release|Delete|Preview|Export)-QuarantineMessage[A-Za-z0-9]*'
            '\bStart-HistoricalSearch\b'
        )
    }
    [PSCustomObject]@{
        Id       = 'SPO'
        Label    = 'SharePoint Online'
        Patterns = @(
            '\b(?:Get|Set|Add|New|Remove|Connect|Disconnect|Test|Request)-SPO[A-Za-z]*'
            '\bConnect-PnPOnline\b'
            '\bGet-PnP[A-Za-z]+'
        )
    }
    [PSCustomObject]@{
        Id       = 'EMS'
        Label    = 'Exchange on-premises'
        Patterns = @(
            '\b(?:Get|Set|New|Remove|Enable|Disable)-RemoteMailbox\b'
            '\bGet-ExchangeServer\b'
            '\bGet-ExchangeCertificate\b'
            '\bSet-ADServerSettings\b'
            '\bUpdate-Recipient\b'
            '\bInvoke-OnPremCommand\b'
            # Exchange cmdlets despite the AD prefix, and on-premises only.
            '\b(?:Get|Add|Remove)-ADPermission\b'
        )
    }
    [PSCustomObject]@{
        Id       = 'AD'
        Label    = 'On-premises Active Directory'
        Patterns = @(
            # Whitelisted nouns, not -AD[A-Za-z]+. The open form matches any
            # local function whose name happens to begin Get-AD, which is how
            # an offline tool ended up claiming a directory connection.
            '\b(?:Get|Set|New|Remove|Add|Enable|Disable|Unlock|Search|Move|Rename)-AD(?:User|Group|GroupMember|PrincipalGroupMembership|Object|Computer|OrganizationalUnit|Domain|DomainController|RootDSE|Account|AccountPassword)\b'
        )
    }
    [PSCustomObject]@{
        Id       = 'Graph'
        Label    = 'Microsoft Entra (Microsoft Graph)'
        Patterns = @(
            # Added 2026-08-06 after PimActivation.ps1 reported "connects to no
            # directory and sends nothing anywhere" -- while talking to Graph.
            #
            # An absent capability does not read as "unknown", it reads as
            # "nothing", and "nothing" is the most reassuring sentence this
            # tool can print. A missing entry is therefore not a gap, it is a
            # false statement of safety. Anything added to the toolkit that
            # reaches a new service needs a row here on the same day.
            '\bConnect-MgGraph\b'
            '\bInvoke-MgGraphRequest\b'
            '\bGet-Mg[A-Z][A-Za-z]*'
            'graph\\.microsoft\\.com'
        )
    }
)

# Cmdlets that exist identically in BOTH Exchanges. Their presence proves the
# script talks to Exchange; it proves nothing about WHICH.
$script:AmbiguousExchangePattern =
    '\b(?:Get|Set|New|Remove|Add|Enable|Disable)-(?:Mailbox|MailboxPermission|MailboxFolderPermission|MailboxStatistics|CASMailbox|Recipient|RecipientPermission|DistributionGroup|DistributionGroupMember|MailContact|MailUser|InboxRule|AcceptedDomain|OrganizationConfig|MailboxAutoReplyConfiguration)\b'

function Get-ToolRunsAgainstText {
    <#
        The derived first line: which directories this script actually talks to.

        Derived rather than typed so it cannot go stale when the script changes,
        and so it cannot disagree with itself. See $script:HelpCapabilityDefs
        for why it does NOT reuse the grey-out classification.
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return 'Runs against: unknown -- the script could not be read.'
    }
    try {
        $raw  = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $code = Remove-PowerShellComment -Text $raw
    } catch {
        return 'Runs against: unknown -- the script could not be read.'
    }

    # An explicit override, for the case derivation genuinely cannot answer:
    #
    #     # TOOLKIT-HELP-RUNS: <text>
    #
    # This is NOT a licence to type what could be derived -- the whole point of
    # deriving is that it cannot go stale. It exists because AdminConsole.ps1
    # is a launcher: it runs against whatever you connect it to, and scanning
    # its own code describes the launcher's plumbing rather than the answer to
    # the question. Use it where the truth is not in the file's cmdlets, and
    # nowhere else.
    if ($raw -match '(?im)^\s*#\s*TOOLKIT-HELP-RUNS:\s*(.+?)\s*$') {
        return "Runs against: $($Matches[1])"
    }

    $labels = New-Object System.Collections.ArrayList
    $ids    = New-Object System.Collections.ArrayList
    foreach ($cap in $script:HelpCapabilityDefs) {
        foreach ($p in $cap.Patterns) {
            if ([regex]::IsMatch($code, $p, 'IgnoreCase')) {
                [void]$labels.Add("$($cap.Label)"); [void]$ids.Add("$($cap.Id)"); break
            }
        }
    }

    # A script that declares its own process is on-premises by construction --
    # that declaration exists precisely because it must not share a process
    # with the cloud. It settles the ambiguous case without guessing.
    $declaresOnPrem = [regex]::IsMatch($code, 'TOOLKIT-REQUIRES:\s*OwnProcess', 'IgnoreCase')
    if (-not $declaresOnPrem) {
        try {
            $head = (Get-Content -LiteralPath $Path -TotalCount 40 -ErrorAction Stop) -join "`n"
            $declaresOnPrem = [regex]::IsMatch($head, 'TOOLKIT-REQUIRES:\s*OwnProcess', 'IgnoreCase')
        } catch { }
    }

    $ambiguous = [regex]::IsMatch($code, $script:AmbiguousExchangePattern, 'IgnoreCase')
    if ($ambiguous -and -not ($ids -contains 'ExO') -and -not ($ids -contains 'EMS')) {
        if ($declaresOnPrem) {
            [void]$labels.Add('Exchange on-premises')
        } else {
            # Honest, and genuinely informative: these cmdlets go wherever the
            # session points, which is a real property of the tool rather than
            # a gap in the analysis.
            [void]$labels.Add('Exchange (on-premises or Online -- whichever this session is connected to)')
        }
    }

    if ($labels.Count -eq 0) {
        return 'Runs against: nothing -- this tool connects to no directory and sends nothing anywhere.'
    }
    return ('Runs against: ' + (($labels | Select-Object -Unique) -join ', '))
}

function Show-ToolHelp {
    <#
        The dialog. Derived line, a rule, then the script's own words.

        Read-only multiline TextBox rather than a MessageBox: help that runs
        past a few lines gets a scrollbar instead of a dialog taller than the
        screen, and the text stays selectable so it can be pasted into a
        ticket.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Title = ''
    )

    $help = Get-ToolHelpText -Path $Path
    $runs = Get-ToolRunsAgainstText -Path $Path
    if (-not $Title) { $Title = [System.IO.Path]::GetFileNameWithoutExtension($Path) }

    if (-not $help) {
        $help = "No overview has been written for this tool yet.`r`n`r`nAdd one between TOOLKIT-HELP-BEGIN and TOOLKIT-HELP-END in the script header and it will appear here."
    }

    $body = "$runs`r`n" + ('-' * 68) + "`r`n`r`n$help"

    $f               = New-Object System.Windows.Forms.Form
    $f.Text          = "About: $Title"
    $f.Size          = New-Object System.Drawing.Size(620, 460)
    $f.StartPosition = 'CenterParent'
    $f.MinimizeBox   = $false
    $f.MaximizeBox   = $false
    $f.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $tb            = New-Object System.Windows.Forms.TextBox
    $tb.Multiline  = $true
    $tb.ReadOnly   = $true
    $tb.ScrollBars = 'Vertical'
    $tb.WordWrap   = $true
    $tb.Location   = New-Object System.Drawing.Point(12, 12)
    $tb.Size       = New-Object System.Drawing.Size(580, 360)
    $tb.Font       = New-Object System.Drawing.Font('Consolas', 9)
    $tb.BackColor  = [System.Drawing.Color]::WhiteSmoke
    $tb.ForeColor  = [System.Drawing.Color]::Black
    $tb.Tag        = 'NoTheme'
    $tb.Text       = $body
    # Caret at the top. A ReadOnly TextBox set after creation otherwise opens
    # scrolled to the end on longer text.
    $tb.Select(0, 0)
    $f.Controls.Add($tb)

    $ok          = New-Object System.Windows.Forms.Button
    $ok.Text     = 'Close'
    $ok.Location = New-Object System.Drawing.Point(492, 382)
    $ok.Size     = New-Object System.Drawing.Size(100, 30)
    # Resolved from the sender, not from a captured $f. Same reasoning as
    # above -- no variable capture, nothing to rebind, nothing to lose.
    $ok.Add_Click({ $this.FindForm().Close() })
    $f.Controls.Add($ok)
    $f.AcceptButton = $ok
    $f.CancelButton = $ok

    Set-ConsoleTheme -Form $f
    [void]$f.ShowDialog()
    $f.Dispose()
}

function Add-ToolHelpButton {
    <#
        Puts a '?' in the top-right corner of a form, in the same place on
        every tool.

        Consistency is the entire point -- a help button you have to look for
        is one nobody presses. It anchors Top+Right so it stays in the corner
        if the form is resized, and it is added LAST so it draws above any
        control it overlaps (WinForms z-order is creation order: earlier is
        on top).

        Tag carries 'NoTheme' plus the script path and title, pipe-separated.
        The theme's opt-out is a substring test, so the button still keeps its
        own colours while Tag doubles as the handler's data channel.
    #>
    param(
        [Parameter(Mandatory)]$Form,
        [Parameter(Mandatory)][string]$ScriptPath,
        [string]$Title = ''
    )

    $btn          = New-Object System.Windows.Forms.Button
    $btn.Text     = '?'
    $btn.Name     = 'btnToolHelp'
    $btn.Size     = New-Object System.Drawing.Size(28, 26)

    <#
        Finds a clear spot rather than assuming the corner is free.

        A static sweep on 2026-08-06 found the corner position sitting on real
        content on 15 of 40 forms -- a text box in HeaderAnalyzer, a grid in
        QuickLaunchTool, a button in Test-Toolkit, and the right-hand end of a
        status label on a dozen report tools. Fixing fifteen layouts by hand
        would fix today and not tomorrow.

        Done at run time on purpose. Only here can the ACTUAL rendered text
        width be measured; a source scan has to estimate it, and estimating it
        is what made the first two passes of that sweep wrong in both
        directions -- once by ignoring that a Label clips at its own Width, and
        once by counting an overlap with a label's empty right-hand end as a
        collision.

        A Label is an obstacle only as far as its text actually reaches.
        Everything else is an obstacle for its whole width, because anything
        can be drawn anywhere inside it.
    #>
    $y = 8
    $x = $Form.ClientSize.Width - 40
    foreach ($c in $Form.Controls) {
        if ("$($c.Tag)" -like '*ToolkitLogo*') { $x = $c.Left - 8 - $btn.Width; break }
    }

    $obstacles = New-Object System.Collections.ArrayList
    foreach ($c in $Form.Controls) {
        if ("$($c.Name)" -eq 'btnToolHelp') { continue }
        if ("$($c.Tag)" -like '*ToolkitLogo*') { continue }
        if (-not $c.Visible -and $c.GetType().Name -ne 'Label') { continue }

        $bounds = $c.Bounds
        if ($c -is [System.Windows.Forms.Label] -and -not $c.AutoSize) {
            try {
                $tw = [System.Windows.Forms.TextRenderer]::MeasureText($c.Text, $c.Font).Width
                if ($tw -lt $c.Width) {
                    $bounds = New-Object System.Drawing.Rectangle($c.Left, $c.Top, ([Math]::Max($tw, 1)), $c.Height)
                }
            } catch { }
        }
        [void]$obstacles.Add($bounds)
    }

    # Try the corner, then a short walk left.
    #
    # The walk rarely helps and that is worth knowing: on most of these forms
    # the top row is ONE label spanning the full width, so there is no
    # horizontal gap to find. It is kept for the few forms where a small nudge
    # does clear a button or a short label.
    $limit = $Form.ClientSize.Width - 200
    $start = $x
    $found = $false
    while ($x -gt $limit) {
        $rect = New-Object System.Drawing.Rectangle($x, $y, $btn.Width, $btn.Height)
        $clash = $false
        foreach ($o in $obstacles) { if ($rect.IntersectsWith($o)) { $clash = $true; break } }
        if (-not $clash) { $found = $true; break }
        $x -= 8
    }

    if (-not $found) {
        # No gap anywhere along the top row, so make one: widen the form and put
        # the button in the new margin.
        #
        # The alternatives were both worse. Sitting on top of a control hides
        # something the operator needs. Shrinking the offending label clips real
        # text -- and the labels involved are the ones already running to the
        # window edge, so they are exactly the ones with no slack to give.
        #
        # Widening only affects controls anchored Right, which follow the edge
        # as intended; absolutely positioned controls stay put and the new space
        # becomes an empty gutter, which is precisely what is wanted.
        $gutter = $btn.Width + 16
        $Form.ClientSize = New-Object System.Drawing.Size(
            ($Form.ClientSize.Width + $gutter), $Form.ClientSize.Height)
        $x = $Form.ClientSize.Width - $btn.Width - 12

        # A logo already placed sits at the old right edge; move it along.
        foreach ($c in $Form.Controls) {
            if ("$($c.Tag)" -like '*ToolkitLogo*') {
                $c.Left = $c.Left + $gutter
                $x = $c.Left - 8 - $btn.Width
            }
        }
    }

    $btn.Location = New-Object System.Drawing.Point($x, $y)
    $btn.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btn.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.BackColor = [System.Drawing.Color]::FromArgb(52, 88, 148)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(120, 170, 235)
    $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(72, 118, 190)

    # NO .GetNewClosure(). It looks like the right way to capture $ScriptPath
    # and it is the reason the first version threw
    # "Show-ToolHelp is not recognized".
    #
    # GetNewClosure rebinds the scriptblock to a NEW dynamic module scope whose
    # parent is the global scope. Common.ps1 is dot-sourced into each tool's
    # SCRIPT scope, so its functions are not global -- and the closure
    # therefore cannot see them. A plain scriptblock keeps the session state it
    # was created in, where Show-ToolHelp is perfectly visible.
    #
    # The data rides on Tag instead. Set-ConsoleTheme's opt-out test is
    # "$($c.Tag) -like '*NoTheme*'", a substring match, so a composite string
    # still opts out -- which is what makes this slot usable rather than
    # 'taken'.
    $btn.Tag = "NoTheme|$ScriptPath|$Title"
    $btn.Add_Click({
        $parts = "$($this.Tag)" -split '\|', 3
        if ($parts.Count -ge 2) {
            Show-ToolHelp -Path $parts[1] -Title $(if ($parts.Count -ge 3) { $parts[2] } else { '' })
        }
    })

    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.SetToolTip($btn, 'What does this tool do?')

    $Form.Controls.Add($btn)
    $btn.BringToFront()
    return $btn
}
#endregion

function Get-ExchangeModuleMap {
    <#
        Which loaded module owns the on-premises verbs, and which owns the
        cloud ones.

        Some scripts genuinely need BOTH Exchanges in one process -- the hybrid
        disable pipeline reads and writes on-premises AD and RemoteMailbox
        objects and then acts on the cloud mailbox. Isolating them into their
        own process (TOOLKIT-REQUIRES: OwnProcess) would break them, so they
        need the other answer: say which Exchange you mean, every time.

        Both worlds export Get-Mailbox, Set-Mailbox, Get-DistributionGroup and
        dozens more, and the last import wins the bare name. A hybrid script
        that calls the bare name is relying on connection ORDER -- which works
        right up until someone connects in the other order, and then writes
        land in the wrong directory with no error.

        Telling the modules apart, version-independently:

          on-premises -- exports Get-RemoteMailbox. The *-RemoteMailbox family
                         is on-prem-exclusive, which is the same property
                         Test-EmsAvailable relies on.
          cloud       -- exports Get-Mailbox but NOT Get-RemoteMailbox.

        Deliberately not matched on module NAME. Implicit remoting generates
        names like tmpEXO_a1b2c3 and they are not stable across versions or
        sessions.
    #>
    param([switch]$Refresh)

    if (-not $Refresh -and $script:ExchangeModuleMap) { return $script:ExchangeModuleMap }

    $onPrem = $null; $cloud = $null
    foreach ($m in @(Get-Module)) {
        $ec = $m.ExportedCommands
        if (-not $ec -or -not $ec.ContainsKey('Get-Mailbox')) { continue }
        if ($ec.ContainsKey('Get-RemoteMailbox')) {
            if (-not $onPrem) { $onPrem = $m }
        } else {
            if (-not $cloud) { $cloud = $m }
        }
    }

    $script:ExchangeModuleMap = [PSCustomObject]@{
        OnPrem = $onPrem
        Cloud  = $cloud
        Both   = [bool]($onPrem -and $cloud)
    }
    return $script:ExchangeModuleMap
}

function Invoke-ExchangeCommand {
    <#
        Run an Exchange cmdlet against a NAMED directory, not against whichever
        module happens to have won the name.

            Invoke-ExchangeCommand -World OnPrem -Command 'Set-Mailbox' -Params @{
                Identity = $smtp; HiddenFromAddressListsEnabled = $true
            }

        When only one Exchange is loaded this resolves to the same command the
        bare name would have. The value is entirely in the case where both are
        loaded -- and in the case where a reader has to work out, months later,
        which directory a line was aimed at. The answer is now in the line.

        Throws rather than falling back if the requested world is absent. A
        hybrid script that silently ran its on-premises half against the cloud
        is exactly the failure being prevented here, and a fallback would
        reintroduce it.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('OnPrem', 'Cloud')][string]$World,
        [Parameter(Mandatory)][string]$Command,
        [hashtable]$Params = @{}
    )

    $map = Get-ExchangeModuleMap
    $mod = if ($World -eq 'OnPrem') { $map.OnPrem } else { $map.Cloud }

    if (-not $mod) {
        # Re-read once before giving up: a connect may have happened after the
        # map was first built, and the cache would still say absent.
        $map = Get-ExchangeModuleMap -Refresh
        $mod = if ($World -eq 'OnPrem') { $map.OnPrem } else { $map.Cloud }
    }
    if (-not $mod) {
        $what = if ($World -eq 'OnPrem') { 'on-premises Exchange' } else { 'Exchange Online' }
        throw "$Command was requested against $what, which is not connected in this process. Refusing to run it against the other directory."
    }

    $qualified = Get-Command -Name "$($mod.Name)\$Command" -ErrorAction SilentlyContinue
    if (-not $qualified) {
        throw "$Command is not exported by the $World Exchange module ($($mod.Name))."
    }
    & $qualified @Params
}

function Test-ExoInProcess {
    <#
        Is Exchange Online ALSO live in this process?

        This is not the same question as Test-EmsAvailable, and conflating the
        two cost three rounds of misdiagnosis on 2026-08-06.

        Test-EmsAvailable answers "are the on-premises cmdlets loaded". It can
        be TRUE while Exchange Online is loaded as well -- and the two share
        names. Get-DistributionGroup, Remove-DistributionGroupMember,
        Set-Mailbox and dozens more exist in both worlds. Whichever module was
        imported last owns the name, so a tool can report "connected to
        on-premises Exchange", genuinely have EMS available, and still have
        every call routed to the cloud.

        What it looks like when that happens: an on-premises object is found
        and identified correctly, and the write fails with

            The operation on Identity "X" failed because it's out of the
            current user's write scope

        which is Exchange Online refusing to modify a directory-synchronised
        object. The sentence continues "...because the object is being
        synchronized from your on-premises organization", and that tail is what
        gives it away -- so never truncate an error message in a log window.

        A tool that writes on-premises must therefore check BOTH: that EMS is
        available, and that EXO is not competing for the same verbs.
    #>
    # v3+ answers directly. This is the reliable path.
    $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
    if ($gci) {
        try {
            $conns = @(Get-ConnectionInformation -ErrorAction Stop)
            $live  = @($conns | Where-Object { $_.State -eq 'Connected' -or -not $_.State })
            if ($live.Count -gt 0) {
                return [PSCustomObject]@{
                    Present = $true
                    Detail  = "Exchange Online is connected in this process ($(($live | ForEach-Object { $_.Organization }) -join ', ')). It shares cmdlet names with on-premises Exchange."
                }
            }
        } catch { }
    }

    # v2 and remote-session fallback: an implicit remoting session pointed at
    # the service. Matching on the URI rather than the name because the
    # generated session names are not stable.
    try {
        $ps = @(Get-PSSession -ErrorAction Stop | Where-Object {
            $_.ConfigurationName -eq 'Microsoft.Exchange' -and
            "$($_.Runspace.ConnectionInfo.ConnectionUri)" -match 'outlook\.office365\.com|outlook\.office\.com|ps\.compliance\.protection'
        })
        if ($ps.Count -gt 0) {
            return [PSCustomObject]@{
                Present = $true
                Detail  = 'An Exchange Online remoting session is open in this process. It shares cmdlet names with on-premises Exchange.'
            }
        }
    } catch { }

    return [PSCustomObject]@{ Present = $false; Detail = 'No Exchange Online connection in this process.' }
}

function Test-EmsAvailable {
    <#
        On-premises Exchange, tested via *-RemoteMailbox.

        That family exists ONLY on-prem, so its presence is unambiguous -- which
        is the same property that makes NewUser.ps1 safe to run beside an open
        Exchange Online session (section 2). Testing with Get-Mailbox would
        prove nothing, because both worlds export it.
    #>
    $c = Get-Command -Name 'New-RemoteMailbox' -ErrorAction SilentlyContinue
    if (-not $c) {
        return [PSCustomObject]@{
            Available = $false
            Detail    = 'Not available -- the *-RemoteMailbox cmdlets are not loaded in this process. Open an Exchange Management Shell window, or import the on-prem session.'
        }
    }

    # Loaded is not the same as working. Test-AdAvailable already learned this
    # and answers it with Get-ADRootDSE; without the equivalent here, an EMS
    # session whose server has gone away reports as connected and then fails on
    # the first Set-RemoteMailbox -- which in a disable batch means every user
    # fails at step one and the retry queue fills with the same misleading error.
    #
    # Get-RemoteMailbox is on-prem-exclusive, so this proves WHICH Exchange
    # answered as well as THAT one did. -ResultSize 1 keeps it cheap on an org
    # with thousands of mailboxes.
    #
    # -WarningAction SilentlyContinue is REQUIRED, not tidiness. Asking for one
    # row out of thousands makes Exchange say
    #
    #     WARNING: There are more results available than are currently
    #     displayed. To view them, increase the value for the ResultSize
    #     parameter.
    #
    # which is true, expected, and exactly what was asked for -- but it prints
    # on the host stream at every launch and every connection repaint, so it
    # reads as a fault in the console rather than as Exchange answering the
    # question correctly. The warning cannot be avoided by raising ResultSize;
    # it is caused by LIMITING the query at all, so the only fix is to silence
    # this one call. See ARCHITECTURE section 14b -- this is the opposite trap
    # to the one Get-AcceptedDomain has, and confusing the two cost a day.
    try {
        $null = Get-RemoteMailbox -ResultSize 1 -ErrorAction Stop -WarningAction SilentlyContinue
        return [PSCustomObject]@{ Available = $true; Detail = 'Exchange on-premises is loaded and answering.' }
    } catch {
        return [PSCustomObject]@{
            Available = $false
            Detail    = "Not available -- the on-prem cmdlets are loaded but the Exchange server did not answer: $($_.Exception.Message)"
        }
    }
}

function Test-AdAvailable {
    <#
        The ActiveDirectory module plus a domain controller that answers.

        Get-ADRootDSE is the cheapest possible round trip -- it proves the
        module is loaded AND a DC is reachable, which "is the module present"
        alone does not.
    #>
    if (-not (Get-Command -Name 'Get-ADUser' -ErrorAction SilentlyContinue)) {
        return [PSCustomObject]@{
            Available = $false
            Detail    = 'Not available -- the ActiveDirectory module is not loaded. Run Import-Module ActiveDirectory.'
        }
    }
    try {
        $null = Get-ADRootDSE -Server $script:HQDomain -ErrorAction Stop
        return [PSCustomObject]@{ Available = $true; Detail = "Active Directory reachable via $($script:HQDomain)." }
    } catch {
        return [PSCustomObject]@{
            Available = $false
            Detail    = "Not available -- the module is loaded but $($script:HQDomain) did not answer: $($_.Exception.Message)"
        }
    }
}

function Get-CapabilityStateMap {
    <#
        Probes all four capabilities once and caches the answer for the life of
        the window. Pass -Refresh after connecting to something.

        NOTE the limit this cannot cross: a stale Exchange Online session
        reports as connected and then fails on the next real call with a
        transport error (section 11). So an enabled button means "nothing is
        obviously missing", not "this will work". Greying catches not-connected;
        it cannot catch connected-but-dead.
    #>
    param([switch]$Refresh, [switch]$SkipSpo)

    if ($global:CapabilityStateMap -and -not $Refresh) { return $global:CapabilityStateMap }

    $map = @{}

    $exo = Test-ExoConnectionState
    $map['ExO'] = [PSCustomObject]@{
        Id = 'ExO'; Label = 'Exchange Online'
        Available = [bool]$exo.Connected; Unknown = $false; Detail = "$($exo.Detail)"
    }

    # SharePoint is the expensive probe -- Get-SPOTenant is a network round trip
    # -- so the console skips it on the startup pass. Unknown is NOT the same as
    # unavailable: an unchecked capability must not grey anything out, or every
    # SharePoint tool would be dead until something happened to check.
    if ($SkipSpo) {
        # Carry a real previous answer forward rather than downgrading it to
        # Unknown. Without this, any cheap refresh would un-grey the SharePoint
        # tools it had correctly greyed a moment earlier -- the buttons would
        # flicker between blocked and available depending on which refresh ran
        # last, which is worse than either answer.
        $prior = $null
        if ($global:CapabilityStateMap) { $prior = $global:CapabilityStateMap['SPO'] }
        if ($prior -and -not $prior.Unknown) {
            $map['SPO'] = $prior
        } else {
            $map['SPO'] = [PSCustomObject]@{
                Id = 'SPO'; Label = 'SharePoint Online'
                Available = $true; Unknown = $true
                Detail    = 'Not checked yet. Tools that need SharePoint are left enabled rather than greyed on a guess.'
            }
        }
    } else {
        $spo = Test-SpoConnectionState
        $map['SPO'] = [PSCustomObject]@{
            Id = 'SPO'; Label = 'SharePoint Online'
            Available = [bool]$spo.Connected; Unknown = $false; Detail = "$($spo.Detail)"
        }
    }

    $ems = Test-EmsAvailable
    $map['EMS'] = [PSCustomObject]@{
        Id = 'EMS'; Label = 'Exchange on-premises'
        Available = $ems.Available; Unknown = $false; Detail = $ems.Detail
    }

    $ad = Test-AdAvailable
    $map['AD'] = [PSCustomObject]@{
        Id = 'AD'; Label = 'On-premises Active Directory'
        Available = $ad.Available; Unknown = $false; Detail = $ad.Detail
    }

    $global:CapabilityStateMap = $map
    return $map
}

function Get-ToolAvailability {
    <#
        Whether a tool can run right now, and a tooltip explaining any 'no'.

        FAILS OPEN. A script that cannot be read, or whose scan found nothing,
        is reported as available. Every existing Tools.json entry predates this
        feature, so a mechanism that blocked anything it could not classify
        would grey out the whole console the first time it ran -- turning a
        convenience into an outage.
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [hashtable]$StateMap
    )

    if (-not $StateMap) { $StateMap = Get-CapabilityStateMap }

    $req = Get-ScriptCapabilityRequirement -Path $ScriptPath
    $missing = @()
    $unknown = @()
    foreach ($id in @($req.Requires)) {
        if (-not $StateMap.ContainsKey($id)) { continue }
        if ($StateMap[$id].Unknown)          { $unknown += $StateMap[$id]; continue }
        if (-not $StateMap[$id].Available)   { $missing += $StateMap[$id] }
    }

    if (@($missing).Count -eq 0) {
        $needText = if (@($req.Requires).Count -gt 0) {
            'Needs: ' + ((@($req.Requires) | ForEach-Object { $StateMap[$_].Label }) -join ', ')
        } else {
            'No connection required.'
        }
        foreach ($u in $unknown) { $needText += "`r`n$($u.Label) -- $($u.Detail)" }
        return [PSCustomObject]@{
            Enabled = $true; Requires = @($req.Requires); Missing = @(); Reason = $needText
        }
    }

    $lines = @('Unavailable in this window:')
    foreach ($m in $missing) { $lines += "  $($m.Label) -- $($m.Detail)" }
    return [PSCustomObject]@{
        Enabled  = $false
        Requires = @($req.Requires)
        Missing  = @($missing | ForEach-Object { $_.Id })
        Reason   = ($lines -join "`r`n")
    }
}
#endregion

#region Browser profiles -----------------------------------------------------
<#
    Different destinations need different identities. ServiceNow wants the
    everyday account; the M365 portals want the admin account. Opening both
    with Start-Process lands them in whichever profile Edge last used, which
    means a sign-in prompt, or worse, the wrong tenant.

    Map purpose -> Edge profile DISPLAY name (the name shown in Edge's profile
    switcher). The display name is deliberately what goes here: Edge's command
    line wants the *directory* name ('Profile 1'), which is an implementation
    detail nobody remembers and which changes if profiles are recreated.
    Get-EdgeProfileDirectory does that translation.
#>
$script:BrowserProfiles = @{
    ServiceNow = 'ServiceNow'   # ticket queue
    Admin      = 'Admin'        # M365 / Entra / Defender portals -- elevated
    GWA        = 'GWA'          # everyday account, for future links
}

function Get-ExchangeServerUri {
    <#
        The remote PowerShell endpoint on the on-premises Exchange server.

        Settings.json first, then built from $script:ExchangeOnPremServer.
        Accepts either a bare host name or a full URI in the setting, because
        both are things a person reasonably types, and rejecting one of them
        would be a support call rather than a feature.

        http, not https, and that is not an oversight. Microsoft: "The
        ConnectionUri value uses http, not https because the session is
        encrypted by the Kerberos token in the payload. Kerberos encryption
        removes the need to encrypt the session again using https."
    #>
    param()

    $configured = ''
    try { $configured = "$((Get-ToolkitSettings).ExchangeServerUri)".Trim() } catch { }
    if (-not $configured) { $configured = "$script:ExchangeOnPremServer".Trim() }
    if (-not $configured) { return $null }

    if ($configured -match '^https?://') { return $configured.TrimEnd('/') + '/' }
    return "http://$configured/PowerShell/"
}

function Get-SpoAdminUrl {
    <#
        The SharePoint admin centre URL for this tenant.

        Settings.json first, then derived. Deriving it is safe because the admin
        centre host is not a free choice -- it is always
        https://<tenant>-admin.sharepoint.com -- so the only unknown is the
        tenant label, and there are two places that already know it.

        Session organisation is preferred over the mail domain because the
        SharePoint tenant label follows the *.onmicrosoft.com prefix, not the
        vanity mail domain. They match here; they do not always, and a tenant
        that renamed its mail domain would otherwise get a URL that has never
        existed.

        The setting exists because derivation cannot cover a vanity SharePoint
        host or a tenant whose SPO label genuinely differs. Set it and nothing
        is guessed.
    #>
    param()

    $configured = ''
    try { $configured = "$((Get-ToolkitSettings).SpoAdminUrl)".Trim() } catch { }
    if ($configured) {
        if ($configured -notmatch '^https?://') { $configured = "https://$configured" }
        return $configured.TrimEnd('/')
    }

    $label = ''
    $org = Get-SessionOrganization
    if ($org) { $label = ("$org" -split '\.')[0] }
    if (-not $label) { $label = ("$script:MailDomain" -split '\.')[0] }
    if (-not $label) { return $null }

    return "https://$label-admin.sharepoint.com"
}

function Get-BrowserProfileMap {
    <#
        Purpose -> Edge profile display name, with Settings.json overriding per
        purpose. Merged rather than replaced, so naming only one purpose in
        Settings.json leaves the others on their defaults instead of deleting
        them -- a partial config should narrow behaviour, not remove it.

        These are per-MACHINE: the profile names are whatever Edge was set up
        with, and a name that does not exist falls back to the last-used profile
        -- which usually works, so the mistake shows up as the wrong account
        rather than as an error.
    #>
    $map = @{}
    foreach ($k in $script:BrowserProfiles.Keys) { $map[$k] = $script:BrowserProfiles[$k] }
    $s = Get-ToolkitSettings
    if ($s.BrowserProfiles) {
        foreach ($p in $s.BrowserProfiles.PSObject.Properties) {
            if ("$($p.Value)".Trim()) { $map[$p.Name] = "$($p.Value)".Trim() }
        }
    }
    return $map
}

function Get-EdgeExecutable {
    foreach ($p in @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles          'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:LOCALAPPDATA          'Microsoft\Edge\Application\msedge.exe')
    )) {
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    return $null
}

function Get-EdgeProfiles {
    <#
        Lists Edge profiles as DisplayName -> Directory.

        Edge stores this in its 'Local State' JSON under
        profile.info_cache, keyed by directory name with the display name
        inside. Run this once to see what your profiles are actually called:

            . .\Common.ps1 ; Get-EdgeProfiles
    #>
    $state = Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data\Local State'
    if (-not (Test-Path -LiteralPath $state)) { return @() }
    try {
        $j = Get-Content -LiteralPath $state -Raw | ConvertFrom-Json
        $cache = $j.profile.info_cache
        if (-not $cache) { return @() }
        return @(
            foreach ($dir in $cache.PSObject.Properties.Name) {
                [PSCustomObject]@{
                    DisplayName = "$($cache.$dir.name)"
                    Directory   = $dir
                }
            }
        )
    } catch { return @() }
}

function Get-EdgeProfileDirectory {
    <#
        Resolves an Edge profile display name to its directory name.
        Accepts a directory name directly too, so 'Profile 1' also works.
    #>
    param([string]$Name)
    if (-not $Name) { return $null }
    $all = Get-EdgeProfiles
    if ($all.Count -eq 0) { return $null }

    $hit = $all | Where-Object { $_.DisplayName -ieq $Name } | Select-Object -First 1
    if ($hit) { return $hit.Directory }
    $hit = $all | Where-Object { $_.Directory -ieq $Name } | Select-Object -First 1
    if ($hit) { return $hit.Directory }
    # Partial match last: 'GWA' should find 'GWA (work)'.
    $hit = $all | Where-Object { $_.DisplayName -like "*$Name*" } | Select-Object -First 1
    if ($hit) { return $hit.Directory }
    return $null
}

function Start-UrlInBrowserProfile {
    <#
        Opens a URL in a named Edge profile, falling back cleanly.

        Fallback order, and it matters that this never hard-fails: a browser
        that will not open is a worse outcome than one that opens in the wrong
        profile, because the second is obvious and recoverable.
          1. Edge with the resolved profile directory
          2. Edge with no profile switch
          3. Start-Process (default browser)

        Returns a status string so the caller can say which happened rather
        than leaving the user guessing why they got a sign-in prompt.
    #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [string]$ProfileName,
        [string]$Purpose
    )
    if (-not $ProfileName -and $Purpose) {
        $profileMap = Get-BrowserProfileMap
        if ($profileMap.ContainsKey($Purpose)) {
            $ProfileName = $profileMap[$Purpose]
        } else {
            # A typo'd purpose would otherwise open in the last-used profile
            # and look like it worked -- which is the exact failure this
            # function exists to prevent.
            Write-Warning ("Unknown browser purpose '{0}'. Known: {1}" -f
                           $Purpose, (($profileMap.Keys | Sort-Object) -join ', '))
        }
    }

    $edge = Get-EdgeExecutable
    if ($edge -and $ProfileName) {
        $dir = Get-EdgeProfileDirectory -Name $ProfileName
        if ($dir) {
            try {
                # The quotes around $dir are load-bearing.
                #
                # Edge profile directories are named 'Profile 2', 'Profile 5'
                # and so on -- they contain a SPACE. Start-Process joins
                # -ArgumentList entries with spaces and does not quote them, so
                # an unquoted value arrives at Edge as two arguments:
                #   --profile-directory=Profile   and   2
                # Edge then creates a new profile directory literally called
                # 'Profile', and treats the leftover '2' as a URL -- which it
                # resolves as the IP 0.0.0.2. Every purpose ends up in that same
                # new profile, which is exactly the wrong-identity failure this
                # function exists to prevent.
                Start-Process -FilePath $edge `
                    -ArgumentList @("--profile-directory=`"$dir`"", "`"$Url`"")
                return "Opened in Edge profile '$ProfileName' ($dir)."
            } catch { }
        } else {
            $names = (Get-EdgeProfiles | ForEach-Object { $_.DisplayName }) -join ', '
            try {
                Start-Process -FilePath $edge -ArgumentList @("`"$Url`"")
                return "Edge profile '$ProfileName' not found. Opened in the last-used profile. Available: $names"
            } catch { }
        }
    }
    if ($edge) {
        try {
            Start-Process -FilePath $edge -ArgumentList @("`"$Url`"")
            return 'Opened in Edge (no profile specified).'
        } catch { }
    }
    Start-Process $Url
    return 'Opened in the default browser.'
}
#endregion

function Get-TenantLogoPath {
    <#
        Convention over configuration: a logo is used if one is present, and
        the tenant-specific name wins over the generic one. Nothing is
        hardcoded to a tenant, so the same folder serves any of them.

            Config\Images\logo.<tenant-domain>.png    preferred
            Config\Images\logo.png                    fallback
    #>
    param([string]$Domain)

    $dir = Get-ConfigImageFolder
    if (-not (Test-Path -LiteralPath $dir)) { return $null }
    if (-not $Domain) { $Domain = Get-ConnectedTenantDomain }

    if ($Domain) {
        foreach ($ext in @('png', 'jpg', 'gif', 'bmp')) {
            $p = Join-Path $dir "logo.$Domain.$ext"
            if (Test-Path -LiteralPath $p) { return $p }
        }
    }
    foreach ($ext in @('png', 'jpg', 'gif', 'bmp')) {
        $p = Join-Path $dir "logo.$ext"
        if (Test-Path -LiteralPath $p) { return $p }
    }
    return $null
}

function Add-FormLogo {
    param(
        [Parameter(Mandatory)]$Form,
        [int]$MaxHeight = 40,
        [ValidateSet('TopRight', 'BottomRight')][string]$Corner = 'BottomRight',
        [int]$Margin = 10
    )
    $path = Get-TenantLogoPath
    if (-not $path) { return $null }
    try {
        $img = [System.Drawing.Image]::FromFile($path)

        # MaxHeight is a REQUEST, not a size.
        #
        # Doubling the logo on 2026-08-06 turned out to be impossible to do
        # blindly: at 2x it landed on the tab strip in Disable-User, on a button
        # in LogViewer, and on the input box in AdminConsole. Every form has a
        # different amount of room in that corner, and a constant cannot know
        # which.
        #
        # So it shrinks until it fits. The form gets the biggest logo its own
        # layout allows, up to what was asked for, and nothing is ever covered
        # by a decoration. Making room is then a deliberate layout change rather
        # than something discovered by looking at a screenshot.
        $h = $MaxHeight
        $w = [int]($img.Width * ($h / $img.Height))
        while ($h -gt 16) {
            $w  = [int]($img.Width * ($h / $img.Height))
            $tx = $Form.ClientSize.Width - $w - $Margin
            $ty = if ($Corner -eq 'TopRight') { $Margin }
                  else { $Form.ClientSize.Height - $h - $Margin }
            $rect = New-Object System.Drawing.Rectangle($tx, $ty, $w, $h)

            $clash = $false
            foreach ($c in $Form.Controls) {
                # The help button is allowed to be here -- it gets moved out of
                # the way below rather than being treated as an obstacle.
                if ("$($c.Name)" -eq 'btnToolHelp') { continue }
                if ($rect.IntersectsWith($c.Bounds)) { $clash = $true; break }
            }
            if (-not $clash) { break }
            $h -= 4
        }
        $w = [int]($img.Width * ($h / $img.Height))

        $pb = New-Object System.Windows.Forms.PictureBox
        $pb.Image    = $img
        $pb.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $pb.Size     = New-Object System.Drawing.Size($w, $h)
        $pb.BackColor = [System.Drawing.Color]::Transparent
        # Tagged so Add-ToolHelpButton can find it and sit to its left.
        $pb.Tag      = 'NoTheme|ToolkitLogo'
        $pb.Name     = 'pbToolkitLogo'

        $x = $Form.ClientSize.Width - $w - $Margin
        $y = if ($Corner -eq 'TopRight') { $Margin }
             else { $Form.ClientSize.Height - $h - $Margin }
        $pb.Location = New-Object System.Drawing.Point($x, $y)
        $pb.Anchor   = if ($Corner -eq 'TopRight') {
                           [System.Windows.Forms.AnchorStyles]::Top -bor
                           [System.Windows.Forms.AnchorStyles]::Right
                       } else {
                           [System.Windows.Forms.AnchorStyles]::Bottom -bor
                           [System.Windows.Forms.AnchorStyles]::Right
                       }
        $Form.Controls.Add($pb)
        $pb.BringToFront()

        # If the help button was added first, move it clear. Handling both
        # orders means neither function has to be called before the other,
        # which is one less rule to remember and get wrong.
        if ($Corner -eq 'TopRight') {
            foreach ($c in $Form.Controls) {
                if ("$($c.Name)" -eq 'btnToolHelp') {
                    $c.Location = New-Object System.Drawing.Point(($x - $c.Width - 8), $c.Top)
                    $c.BringToFront()
                }
            }
        }
        return $pb
    } catch { return $null }
}
#endregion

#region NEW -- connection state probes (for AdminConsole's Connect buttons) --
function Test-ExoConnectionState {
    param()
    $r = [ordered]@{ Connected = $false; Detail = ''; Organization = $null }

    $probe = @(Get-Command -Name 'Get-Mailbox' -All -ErrorAction SilentlyContinue |
        Where-Object { $_.ModuleName -like 'tmpEXO_*' -or $_.Source -like 'tmpEXO_*' -or
                       $_.ModuleName -eq 'ExchangeOnlineManagement' -or $_.Source -eq 'ExchangeOnlineManagement' })
    if ($probe.Count -eq 0) {
        $r.Detail = 'Not connected -- no Exchange Online cmdlets in this PowerShell process. Connect-ExchangeOnline is per-process, so connecting in another window does not count.'
        return [PSCustomObject]$r
    }

    $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
    if (-not $gci) {
        $r.Connected = $true
        $r.Detail = 'Exchange Online cmdlets are loaded. This module version has no Get-ConnectionInformation, so live session state could not be verified.'
        return [PSCustomObject]$r
    }
    try {
        $conns = @(& $gci -ErrorAction Stop)
        if ($conns.Count -eq 0) {
            $r.Detail = 'Not connected -- cmdlets are loaded but there is no active session (it may have expired). Run Connect-ExchangeOnline again.'
            return [PSCustomObject]$r
        }
        $healthy = $conns | Where-Object {
            (-not $_.PSObject.Properties['State']       -or "$($_.State)"       -eq 'Connected') -and
            (-not $_.PSObject.Properties['TokenStatus'] -or "$($_.TokenStatus)" -ne 'Expired')
        }
        $chosen = $conns | Select-Object -First 1
        if ($chosen.PSObject.Properties['Organization']) { $r.Organization = "$($chosen.Organization)" }
        if (@($healthy).Count -gt 0) {
            $r.Connected = $true
            $r.Detail = "Connected to Exchange Online$(if ($r.Organization) { " ($($r.Organization))" })."
        } else {
            $r.Detail = 'Not connected -- the existing session reports itself expired or disconnected. Run Connect-ExchangeOnline again.'
        }
    } catch {
        $r.Connected = $true
        $r.Detail = "Exchange Online cmdlets are loaded; Get-ConnectionInformation could not be read ($($_.Exception.Message))."
    }
    return [PSCustomObject]$r
}

function Get-SessionOrganization {
    <#
        The connected tenant's organisation name, straight from the Exchange
        Online module. Local call, no network -- the module already holds this
        from Connect-ExchangeOnline.

        Returns the *.onmicrosoft.com routing domain, which is nobody's mail
        address but is an unambiguous tenant identifier. Good enough to answer
        "am I where I think I am"; not what you want on screen if a vanity
        domain is available.
    #>
    $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
    if (-not $gci) { return $null }
    try {
        $c = @(& $gci -ErrorAction Stop) | Select-Object -First 1
        if ($c -and $c.PSObject.Properties['Organization']) {
            $o = "$($c.Organization)".Trim()
            if ($o) { return $o }
        }
    } catch { }
    return $null
}

function Test-TenantDomainMismatch {
    <#
        Replaces a check that used to fall out of enumerating every accepted
        domain: is the configured mail domain actually this tenant's?

        The old form confirmed the configured domain appeared in the accepted
        list. This asks the session directly instead -- no network, and it
        catches the case that actually hurts, which is being connected to the
        wrong tenant entirely rather than having a typo in a setting.

        Deliberately loose. A tenant's vanity domain rarely resembles its
        onmicrosoft prefix, so only an obvious disagreement is reported, and
        it is reported as a question rather than an error. Returns $null when
        there is nothing to say.
    #>
    param()

    $org = Get-SessionOrganization
    if (-not $org) { return $null }

    # The parentheses in (Get-ToolkitSettings).MailDomain are load-bearing.
    # Without them PowerShell parses the whole thing as a COMMAND NAME rather
    # than a property access, throws CommandNotFound straight into the catch,
    # and yields nothing. The symptom would be Settings.json being ignored
    # while the code constant still worked -- close to invisible.
    $configured = ''
    try { $configured = "$((Get-ToolkitSettings).MailDomain)".Trim() } catch { }
    if (-not $configured) { $configured = "$script:MailDomain".Trim() }
    if (-not $configured) { return $null }

    # An onmicrosoft prefix usually contains the organisation's short name, so
    # a shared leading label is treated as agreement.
    $orgLabel = ("$org" -split '\.')[0]
    $cfgLabel = ("$configured" -split '\.')[0]
    if (-not $orgLabel -or -not $cfgLabel) { return $null }
    if ($orgLabel -ieq $cfgLabel) { return $null }
    if ($orgLabel -like "*$cfgLabel*" -or $cfgLabel -like "*$orgLabel*") { return $null }
    if ("$org" -ieq "$configured") { return $null }

    return "Mail domain is set to '$configured' but this session is connected to '$org'. If that is not the tenant you meant, reconnect before making changes."
}

function Get-ConnectedTenantDomain {
    <#
        Which tenant is this console pointed at.

        THIS RUNS AT LAUNCH, and again on every connection-button repaint, so
        its cost is paid over and over. It must not enumerate anything.

        Three sources, cheapest first:

          1. $script:MailDomain / Settings.json -- already known, zero calls.
          2. Get-ConnectionInformation -- the module's own record of the
             session. LOCAL, no network round trip, and already read once by
             Test-ExoConnectionState during the same startup.
          3. Get-AcceptedDomain -- the network. Only reached when the first two
             are blank.

        The earlier version did (3) unconditionally and then filtered the
        result by $script:MailDomain -- fetching a thousand-plus domains to
        confirm a value it already had, on every repaint. That is where the
        console's startup lag came from.

        Losing the list also loses the implicit check that the configured
        domain really belongs to the connected tenant. Get-ConnectionInformation
        answers that better and for free -- see Test-TenantDomainMismatch.
    #>
    param()
    # Global for the same reason as the theme cache: child script scopes do
    # not inherit $script: state from the console that launched them.
    if ($global:AdminConsoleTenantDomain) { return $global:AdminConsoleTenantDomain }
    if ($script:ConnectedTenantDomainCache) { return $script:ConnectedTenantDomainCache }

    # (1) Configured. Settings.json wins over the code-level constant, matching
    # every other consumer of the mail domain.
    # The parentheses in (Get-ToolkitSettings).MailDomain are load-bearing.
    # Without them PowerShell parses the whole thing as a COMMAND NAME rather
    # than a property access, throws CommandNotFound straight into the catch,
    # and yields nothing. The symptom would be Settings.json being ignored
    # while the code constant still worked -- close to invisible.
    $configured = ''
    try { $configured = "$((Get-ToolkitSettings).MailDomain)".Trim() } catch { }
    if (-not $configured) { $configured = "$script:MailDomain".Trim() }
    if ($configured) {
        $script:ConnectedTenantDomainCache = $configured
        $global:AdminConsoleTenantDomain   = $configured
        return $configured
    }

    # (2) The session's own record. No network.
    $org = Get-SessionOrganization
    if ($org) {
        $script:ConnectedTenantDomainCache = $org
        $global:AdminConsoleTenantDomain   = $org
        return $org
    }

    # (3) Network, last resort.
    try {
        # ResultSize: see Get-AcceptedDomainList. Without it the cmdlet stops
        # at 1000 domains, and the console's whole identity -- tenant label,
        # theme binding, HQ/Franchise detection -- is decided from whatever
        # arbitrary slice came back first.
        $accepted = @(Invoke-EXOCommand -Name 'Get-AcceptedDomain' -Parameters @{
            ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
        if ($accepted.Count -eq 0) { return $null }

        $names = @($accepted | ForEach-Object { "$($_.DomainName)" })

        # Order matters, and 'first non-onmicrosoft' is deliberately NOT first.
        # In a tenant with hundreds of accepted domains that is whichever one
        # Exchange happened to return first -- a franchise office, not the
        # organisation. The tenant's own Default answers the question properly;
        # it is only skipped when it is the .onmicrosoft.com fallback, which
        # identifies the tenant correctly but reads as noise on screen.
        # No $script:MailDomain test here -- this branch is only reached when
        # the configured domain was blank, so that comparison could only ever
        # match nothing.
        $pick = $null
        if (-not $pick) {
            $def = $accepted | Where-Object { $_.Default -eq $true } | Select-Object -First 1
            if ($def -and "$($def.DomainName)" -notlike '*.onmicrosoft.com') {
                $pick = "$($def.DomainName)"
            }
        }
        if (-not $pick) {
            $pick = $names | Where-Object { $_ -notlike '*.onmicrosoft.com' } | Select-Object -First 1
        }
        if (-not $pick) {
            $def = $accepted | Where-Object { $_.Default -eq $true } | Select-Object -First 1
            if ($def) { $pick = "$($def.DomainName)" }
        }
        if (-not $pick) { $pick = $names[0] }

        if ($pick) {
            $script:ConnectedTenantDomainCache = $pick
            $global:AdminConsoleTenantDomain  = $pick
            return $pick
        }
    } catch {
    }
    return $null
}

function Clear-ConnectedTenantDomainCache {
    $script:ConnectedTenantDomainCache = $null
    $global:AdminConsoleTenantDomain  = $null
}

function Test-SpoConnectionState {
    param()
    $r = [ordered]@{ Connected = $false; Detail = '' }
    if (-not (Get-Command -Name 'Get-SPOTenant' -ErrorAction SilentlyContinue)) {
        $r.Detail = 'Not connected -- SharePoint Online cmdlets are not loaded in this process. Import Microsoft.Online.SharePoint.PowerShell and run Connect-SPOService.'
        return [PSCustomObject]$r
    }
    try {
        $null = Get-SPOTenant -ErrorAction Stop
        $r.Connected = $true
        $r.Detail = 'Connected to SharePoint Online.'
    } catch {
        $r.Detail = "Not connected -- SPO cmdlets are loaded but the session is missing or expired: $($_.Exception.Message)"
    }
    return [PSCustomObject]$r
}
#endregion

#region NEW -- shared "enter a target" prompt for standalone tool launches --
function Show-TargetPromptDialog {
    param(
        [string]$InitialSmtp,
        [string]$Title = 'Enter Target'
    )
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(480, 168)
    $dlg.StartPosition = 'CenterScreen'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = 'Mailbox or Group (email address):'
    $lblUser.Location = New-Object System.Drawing.Point(12, 12)
    $lblUser.AutoSize = $true
    $dlg.Controls.Add($lblUser)
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(12, 30)
    $txtUser.Size = New-Object System.Drawing.Size(440, 22)
    $txtUser.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtUser.Text = $InitialSmtp
    $dlg.Controls.Add($txtUser)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = 'HQ and Franchise are checked automatically -- no environment selection needed.'
    $lblHint.Location = New-Object System.Drawing.Point(12, 58)
    $lblHint.Size = New-Object System.Drawing.Size(440, 18)
    $lblHint.ForeColor = [System.Drawing.Color]::DimGray
    $lblHint.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Italic)
    $dlg.Controls.Add($lblHint)

    $lblErr = New-Object System.Windows.Forms.Label
    $lblErr.Location = New-Object System.Drawing.Point(12, 78)
    $lblErr.Size = New-Object System.Drawing.Size(440, 20)
    $lblErr.ForeColor = [System.Drawing.Color]::Red
    $dlg.Controls.Add($lblErr)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(280, 100)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Location = New-Object System.Drawing.Point(368, 100)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk

    $btnOk.Add_Click({
        $smtp = Get-TargetSmtp -Text $txtUser.Text
        if (-not $smtp) { $lblErr.Text = 'Enter a mailbox or group email address.'; return }
        $dlg.Tag = $smtp
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    })

    $result = $dlg.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $dlg.Tag
}

function Resolve-ToolTarget {
    param(
        [string]$SmtpAddress,
        [string]$Environment,
        [string]$DomainController,
        [string]$ToolTitle = 'Enter Target'
    )
    $smtp = Get-TargetSmtp -Text $SmtpAddress
    $envOk = $Environment -in @('Franchise', 'HQ')
    if ($smtp -and $envOk -and ($Environment -ne 'Franchise' -or $DomainController)) {
        return [PSCustomObject]@{
            Smtp        = $smtp
            Environment = $Environment
            ADServer    = Get-ADServerString -Environment $Environment -DomainController $DomainController
        }
    }
    if (-not $smtp) {
        $smtp = Show-TargetPromptDialog -InitialSmtp $SmtpAddress -Title $ToolTitle
        if (-not $smtp) { return $null }
    }
    $envResult = Find-TargetEnvironment -Smtp $smtp
    if (-not $envResult.Environment) {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [void][System.Windows.Forms.MessageBox]::Show(
            "'$smtp' was not found in HQ or Franchise AD.`n`nDouble-check the address and try again.",
            $ToolTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $null
    }
    return [PSCustomObject]@{
        Smtp        = $smtp
        Environment = $envResult.Environment
        ADServer    = $envResult.ADServer
    }
}
#endregion
