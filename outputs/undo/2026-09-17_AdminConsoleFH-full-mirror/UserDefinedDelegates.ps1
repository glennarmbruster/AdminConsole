#Requires -Version 5.1
<#
    UserDefinedDelegates.ps1
    Folder permissions on the five personal folders, plus Send On Behalf, in one
    place -- readable and editable.

    WHY ALL FIVE TOGETHER. A delegate who can see the calendar but not the inbox
    is the commonest "it half works" report, and answering it today means five
    separate Get-MailboxFolderPermission calls typed by hand. Seeing them side
    by side is the point of the tool.

    -Environment IS ACCEPTED AND USED, and passed to every Write-*Log call.
#>
# TOOLKIT-HELP-RUNS: READS and WRITES mailbox folder permissions and Send On Behalf in Exchange Online.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Who has access to somebody's Inbox, Calendar, Tasks, Contacts and Notes --
# and Send On Behalf -- with the ability to change it.
#
# Type an alias, the first part of an address, or a whole address. Press Load.
#
# THIS WRITES. Add, Change and Remove act immediately on the mailbox. Every one
# asks first and names what it is about to do.
#
# DEFAULT AND ANONYMOUS ARE NOT PEOPLE AND CANNOT BE REMOVED. They are the
# baseline for everybody who has no explicit entry, so Exchange refuses to
# delete them. Set them to None instead -- the tool does that automatically if
# you press Remove on one.
#
# SEND ON BEHALF IS NOT SEND AS. On Behalf shows the recipient "Bob on behalf of
# Alice"; Send As is indistinguishable from Alice sending it herself. They are
# different permissions in different places, and this tool only does the first.
# Send As lives in Mailbox Permissions.
#
# FOLDER NAMES ARE LOOKED UP, NOT ASSUMED. A mailbox created in another language
# has a Calendar folder called something else, and "\Calendar" simply fails on
# it. The real name is read from the mailbox each time.
#
# THE ACCESS LEVELS THAT MATTER MOST:
#   Reviewer          can read. The right answer for most calendar requests.
#   Editor            can read, create, change and delete. Full working access.
#   AvailabilityOnly  free/busy only, no detail. Calendar only.
#   Owner             everything including changing permissions. Rarely correct.
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

# THE DOMAIN IS READ FROM Tenant.json, NOT WRITTEN HERE.
#
# It was a literal, which is the one thing a .ps1 in this toolkit may not
# contain -- Test-Toolkit fails the build on a domain in code, and the reason
# is that the same script ships to more than one place. Get-MailDomain returns
# the tenant's primary domain and '' if it cannot, which is the same answer an
# empty box gives: the operator types the full address instead.
$script:DefaultDomain = ''
try { $script:DefaultDomain = Get-MailDomain } catch { }
$script:Mbx      = $null
$script:Folders  = @{}     # scope -> resolved '\Name' path
$script:Rows     = New-Object System.Collections.ArrayList
$script:OnBehalf = New-Object System.Collections.ArrayList

# THE SCOPES, IN THE ORDER PEOPLE ASK ABOUT THEM. Calendar first would be more
# honest about frequency, but Inbox first matches how the mailbox reads.
$script:Scopes = @('Inbox', 'Calendar', 'Tasks', 'Contacts', 'Notes')

# The levels Exchange accepts. Listed rather than free text because a typo here
# fails at the write, after the confirmation.
$script:Levels = @(
    'Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author',
    'NonEditingAuthor', 'Reviewer', 'Contributor', 'AvailabilityOnly',
    'LimitedDetails', 'None'
)

#region Lookup ------------------------------------------------------------------
# THIS FILE'S OWN COPY, DELIBERATELY NAMED DIFFERENTLY FROM Common.ps1's.
#
# It used to be called Resolve-Target, the same as the function in
# Common.ps1 -- which this file dot-sources. The local definition wins, so
# the calls below reached THIS one and always had; but the shadowing was a
# trap with two edges. A fix made in Common never arrived here, and anyone
# tidying up an apparent duplicate would have deleted this and left the
# calls binding silently to Common's version, which does not take the same
# parameters. That is a run-time failure at the click, on a Saturday.
#
# The two are genuinely different functions that happened to share a name,
# so they now have different names. Renamed 2026-09-12; behaviour unchanged.
function Resolve-DelegateTarget {
    <#
        Alias, first part, or full address -- to exactly one mailbox.

        EXACTLY ONE OR IT IS AN ERROR. This tool grants access to somebody's
        mail. Resolving 'smith' to the first of four is not a recoverable
        mistake.
    #>
    param([Parameter(Mandatory)][string]$Identity)

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Mbx = $null }
    $id = "$Identity".Trim()
    if (-not $id) { $r.Reason = 'Nothing entered.'; return $r }

    # A bare name gets the default domain. A pasted full address is left alone,
    # because appending to it produces bob@x.com@contoso.com and a
    # spectacularly unhelpful "not found".
    $try = @($id)
    if ($id -notlike '*@*') { $try += "$id@$($script:DefaultDomain)" }

    foreach ($cand in $try) {
        try {
            $m = @(Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $cand; ErrorAction = 'Stop' })
            if (@($m).Count -eq 1) { $r.Mbx = $m[0]; $r.Ok = $true; return $r }
            if (@($m).Count -gt 1) { $r.Reason = "'$cand' matched $(@($m).Count) mailboxes. Be exact."; return $r }
        } catch { }
    }

    $r.Reason = "'$id' did not resolve to a mailbox. Try the alias or the full address."
    return $r
}

function Resolve-FolderPaths {
    <#
        The real folder names for this mailbox.

        LOOKED UP, NOT ASSUMED, and this is not pedantry. A mailbox created in
        German has a Calendar folder named 'Kalender', and '\Calendar' throws
        "folder not found" -- which reads as a permissions problem and is not.

        A SCOPE THAT RETURNS NOTHING IS RECORDED AS ABSENT rather than skipped
        silently. A mailbox with no Notes folder is a fact worth showing.
    #>
    param([Parameter(Mandatory)][string]$Upn)

    $map = @{}
    foreach ($scope in $script:Scopes) {
        try {
            $f = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{
                Identity = $Upn; FolderScope = $scope; ErrorAction = 'Stop' } |
                Where-Object { $_.FolderPath -and ($_.FolderPath -split '/').Count -le 2 } |
                Select-Object -First 1)
            if (@($f).Count -gt 0) { $map[$scope] = "$($f[0].Name)" }
        } catch { }
    }
    return $map
}

function Get-AllFolderPermissions {
    param([Parameter(Mandatory)][string]$Upn)

    $script:Rows.Clear()
    $problems = New-Object System.Collections.ArrayList

    foreach ($scope in $script:Scopes) {
        if (-not $script:Folders.ContainsKey($scope)) {
            [void]$problems.Add("$scope -- no such folder on this mailbox.")
            continue
        }
        $path = "$Upn`:\$($script:Folders[$scope])"
        try {
            foreach ($p in @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{
                    Identity = $path; ErrorAction = 'Stop' })) {
                [void]$script:Rows.Add([PSCustomObject]@{
                    Folder = $scope
                    FolderName = "$($script:Folders[$scope])"
                    User   = "$($p.User)"
                    Access = (@($p.AccessRights) -join ', ')
                    Path   = $path
                })
            }
        } catch {
            [void]$problems.Add("$scope -- $($_.Exception.Message)")
        }
    }
    return @($problems.ToArray())
}

function Get-OnBehalfList {
    param($Mbx)
    $script:OnBehalf.Clear()
    foreach ($g in @($Mbx.GrantSendOnBehalfTo)) {
        if ("$g") { [void]$script:OnBehalf.Add("$g") }
    }
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'User Defined Delegates'
$form.ClientSize      = New-Object System.Drawing.Size(1180, 820)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$grpWho          = New-Object System.Windows.Forms.GroupBox
$grpWho.Text     = 'Mailbox'
$grpWho.Location = New-Object System.Drawing.Point(14, 46)
$grpWho.Size     = New-Object System.Drawing.Size(1152, 76)
$grpWho.Anchor   = 'Top,Left,Right'
[void]$form.Controls.Add($grpWho)

$lblWho          = New-Object System.Windows.Forms.Label
$lblWho.Text     = 'Alias or address'
$lblWho.Location = New-Object System.Drawing.Point(14, 29)
$lblWho.Size     = New-Object System.Drawing.Size(110, 20)
[void]$grpWho.Controls.Add($lblWho)

$txtWho          = New-Object System.Windows.Forms.TextBox
$txtWho.Location = New-Object System.Drawing.Point(130, 26)
$txtWho.Size     = New-Object System.Drawing.Size(300, 22)
[void]$grpWho.Controls.Add($txtWho)

$btnLoad           = New-Object System.Windows.Forms.Button
$btnLoad.Text      = 'Load'
$btnLoad.Location  = New-Object System.Drawing.Point(442, 24)
$btnLoad.Size      = New-Object System.Drawing.Size(110, 26)
$btnLoad.FlatStyle = 'Flat'
[void]$grpWho.Controls.Add($btnLoad)

$lblWhoHint           = New-Object System.Windows.Forms.Label
$lblWhoHint.Text      = "A bare name gets @$($script:DefaultDomain) added. A full address is used as typed."
$lblWhoHint.Location  = New-Object System.Drawing.Point(568, 29)
$lblWhoHint.Size      = New-Object System.Drawing.Size(560, 20)
$lblWhoHint.ForeColor = [System.Drawing.Color]::Gray
$lblWhoHint.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grpWho.Controls.Add($lblWhoHint)

# ---- folder permissions -------------------------------------------------------
$grpPerm          = New-Object System.Windows.Forms.GroupBox
$grpPerm.Text     = 'Folder permissions'
$grpPerm.Location = New-Object System.Drawing.Point(14, 130)
$grpPerm.Size     = New-Object System.Drawing.Size(1152, 340)
$grpPerm.Anchor   = 'Top,Left,Right'
[void]$form.Controls.Add($grpPerm)

$gridPerm                     = New-Object System.Windows.Forms.DataGridView
$gridPerm.Location            = New-Object System.Drawing.Point(12, 22)
$gridPerm.Size                = New-Object System.Drawing.Size(1128, 234)
$gridPerm.ReadOnly            = $true
$gridPerm.AllowUserToAddRows  = $false
$gridPerm.AutoSizeColumnsMode = 'Fill'
$gridPerm.SelectionMode       = 'FullRowSelect'
$gridPerm.MultiSelect         = $false
$gridPerm.Anchor              = 'Top,Left,Right'
[void]$grpPerm.Controls.Add($gridPerm)

$lblFolder          = New-Object System.Windows.Forms.Label
$lblFolder.Text     = 'Folder'
$lblFolder.Location = New-Object System.Drawing.Point(12, 268)
$lblFolder.Size     = New-Object System.Drawing.Size(50, 20)
[void]$grpPerm.Controls.Add($lblFolder)

$cboFolder               = New-Object System.Windows.Forms.ComboBox
$cboFolder.DropDownStyle = 'DropDownList'
$cboFolder.Location      = New-Object System.Drawing.Point(66, 265)
$cboFolder.Size          = New-Object System.Drawing.Size(120, 22)
foreach ($s in $script:Scopes) { [void]$cboFolder.Items.Add($s) }
$cboFolder.SelectedIndex = 0
[void]$grpPerm.Controls.Add($cboFolder)

$lblUser          = New-Object System.Windows.Forms.Label
$lblUser.Text     = 'User'
$lblUser.Location = New-Object System.Drawing.Point(198, 268)
$lblUser.Size     = New-Object System.Drawing.Size(40, 20)
[void]$grpPerm.Controls.Add($lblUser)

$txtUser          = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(240, 265)
$txtUser.Size     = New-Object System.Drawing.Size(250, 22)
[void]$grpPerm.Controls.Add($txtUser)

$lblLevel          = New-Object System.Windows.Forms.Label
$lblLevel.Text     = 'Level'
$lblLevel.Location = New-Object System.Drawing.Point(502, 268)
$lblLevel.Size     = New-Object System.Drawing.Size(44, 20)
[void]$grpPerm.Controls.Add($lblLevel)

$cboLevel               = New-Object System.Windows.Forms.ComboBox
$cboLevel.DropDownStyle = 'DropDownList'
$cboLevel.Location      = New-Object System.Drawing.Point(550, 265)
$cboLevel.Size          = New-Object System.Drawing.Size(170, 22)
foreach ($lv in $script:Levels) { [void]$cboLevel.Items.Add($lv) }
$cboLevel.SelectedItem = 'Reviewer'
[void]$grpPerm.Controls.Add($cboLevel)

$btnAdd           = New-Object System.Windows.Forms.Button
$btnAdd.Text      = 'Add'
$btnAdd.Location  = New-Object System.Drawing.Point(736, 263)
$btnAdd.Size      = New-Object System.Drawing.Size(96, 26)
$btnAdd.FlatStyle = 'Flat'
[void]$grpPerm.Controls.Add($btnAdd)

$btnChange           = New-Object System.Windows.Forms.Button
$btnChange.Text      = 'Change level'
$btnChange.Location  = New-Object System.Drawing.Point(840, 263)
$btnChange.Size      = New-Object System.Drawing.Size(110, 26)
$btnChange.FlatStyle = 'Flat'
[void]$grpPerm.Controls.Add($btnChange)

$btnRemove           = New-Object System.Windows.Forms.Button
$btnRemove.Text      = 'Remove'
$btnRemove.Location  = New-Object System.Drawing.Point(958, 263)
$btnRemove.Size      = New-Object System.Drawing.Size(96, 26)
$btnRemove.FlatStyle = 'Flat'
$btnRemove.ForeColor = [System.Drawing.Color]::Firebrick
[void]$grpPerm.Controls.Add($btnRemove)

$lblPermHint           = New-Object System.Windows.Forms.Label
$lblPermHint.Text      = 'Change and Remove act on the SELECTED ROW. Add uses the Folder, User and Level boxes. Default and Anonymous cannot be removed -- Remove sets them to None instead.'
$lblPermHint.Location  = New-Object System.Drawing.Point(12, 296)
$lblPermHint.Size      = New-Object System.Drawing.Size(1128, 34)
$lblPermHint.ForeColor = [System.Drawing.Color]::Gray
$lblPermHint.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grpPerm.Controls.Add($lblPermHint)

# ---- send on behalf -----------------------------------------------------------
$grpSob          = New-Object System.Windows.Forms.GroupBox
$grpSob.Text     = 'Send On Behalf'
$grpSob.Location = New-Object System.Drawing.Point(14, 478)
$grpSob.Size     = New-Object System.Drawing.Size(1152, 150)
$grpSob.Anchor   = 'Top,Left,Right'
[void]$form.Controls.Add($grpSob)

$lstSob          = New-Object System.Windows.Forms.ListBox
$lstSob.Location = New-Object System.Drawing.Point(12, 22)
$lstSob.Size     = New-Object System.Drawing.Size(420, 84)
[void]$grpSob.Controls.Add($lstSob)

$txtSob          = New-Object System.Windows.Forms.TextBox
$txtSob.Location = New-Object System.Drawing.Point(444, 22)
$txtSob.Size     = New-Object System.Drawing.Size(250, 22)
[void]$grpSob.Controls.Add($txtSob)

$btnSobAdd           = New-Object System.Windows.Forms.Button
$btnSobAdd.Text      = 'Add'
$btnSobAdd.Location  = New-Object System.Drawing.Point(704, 20)
$btnSobAdd.Size      = New-Object System.Drawing.Size(96, 26)
$btnSobAdd.FlatStyle = 'Flat'
[void]$grpSob.Controls.Add($btnSobAdd)

$btnSobRemove           = New-Object System.Windows.Forms.Button
$btnSobRemove.Text      = 'Remove selected'
$btnSobRemove.Location  = New-Object System.Drawing.Point(808, 20)
$btnSobRemove.Size      = New-Object System.Drawing.Size(130, 26)
$btnSobRemove.FlatStyle = 'Flat'
$btnSobRemove.ForeColor = [System.Drawing.Color]::Firebrick
[void]$grpSob.Controls.Add($btnSobRemove)

$lblSobHint           = New-Object System.Windows.Forms.Label
$lblSobHint.Text      = 'On Behalf shows the recipient "Bob on behalf of Alice". Send As -- which looks like Alice sent it -- is a different permission and lives in Mailbox Permissions.'
$lblSobHint.Location  = New-Object System.Drawing.Point(444, 52)
$lblSobHint.Size      = New-Object System.Drawing.Size(690, 54)
$lblSobHint.ForeColor = [System.Drawing.Color]::Gray
$lblSobHint.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$grpSob.Controls.Add($lblSobHint)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(14, 638)
$txtOut.Size       = New-Object System.Drawing.Size(1152, 132)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(1046, 780)
$btnClose.Size      = New-Object System.Drawing.Size(120, 28)
$btnClose.FlatStyle = 'Flat'
$btnClose.Anchor    = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
[void]$form.Controls.Add($btnClose)
#endregion

#region Display -----------------------------------------------------------------
function Update-PermGrid {
    $t = New-Object System.Data.DataTable
    foreach ($c in @('Folder', 'Folder name', 'User', 'Access')) { [void]$t.Columns.Add($c) }
    foreach ($r in @($script:Rows | Sort-Object Folder, User)) {
        [void]$t.Rows.Add($r.Folder, $r.FolderName, $r.User, $r.Access)
    }
    $gridPerm.DataSource = $t

    foreach ($gr in $gridPerm.Rows) {
        $u = "$($gr.Cells['User'].Value)"
        $a = "$($gr.Cells['Access'].Value)"
        # DEFAULT AND ANONYMOUS ARE GREYED because they are not delegates and
        # confusing them for one is how somebody "removes a user" and changes
        # the baseline for the whole organisation instead.
        if ($u -eq 'Default' -or $u -eq 'Anonymous') { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray }
        elseif ($a -match '(?i)owner') { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
    }
}

function Get-SelectedPermRow {
    if ($gridPerm.SelectedRows.Count -eq 0) { return $null }
    $i = $gridPerm.SelectedRows[0].Index
    $sorted = @($script:Rows | Sort-Object Folder, User)
    if ($i -lt 0 -or $i -ge @($sorted).Count) { return $null }
    return $sorted[$i]
}

function Invoke-Load {
    $res = Resolve-DelegateTarget -Identity "$($txtWho.Text)"
    if (-not $res.Ok) { $txtOut.Text = $res.Reason; return }

    $script:Mbx = $res.Mbx
    $upn = "$($script:Mbx.PrimarySmtpAddress)"
    $btnLoad.Enabled = $false
    $txtOut.Text = "Reading $upn ..."
    $form.Refresh()

    try {
        $script:Folders = Resolve-FolderPaths -Upn $upn
        $problems = Get-AllFolderPermissions -Upn $upn
        Get-OnBehalfList -Mbx $script:Mbx

        Update-PermGrid
        $lstSob.Items.Clear()
        foreach ($s in $script:OnBehalf) { [void]$lstSob.Items.Add($s) }

        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("$($script:Mbx.DisplayName)  <$upn>   [$($script:Mbx.RecipientTypeDetails)]")
        [void]$L.Add('')
        [void]$L.Add("$(@($script:Rows).Count) folder permission entry(ies) across $(@($script:Folders.Keys).Count) folder(s).")
        foreach ($s in $script:Scopes) {
            $n = @($script:Rows | Where-Object { $_.Folder -eq $s }).Count
            $nm = $(if ($script:Folders.ContainsKey($s)) { $script:Folders[$s] } else { '(absent)' })
            [void]$L.Add(("   {0,-10} {1,-22} {2} entry(ies)" -f $s, $nm, $n))
        }
        [void]$L.Add('')
        [void]$L.Add("Send On Behalf: $(if (@($script:OnBehalf).Count) { @($script:OnBehalf) -join ', ' } else { 'nobody' })")

        # NON-DEFAULT ENTRIES ARE THE ANSWER TO "who has access". Default and
        # Anonymous are always present and are not what anybody is asking about.
        $real = @($script:Rows | Where-Object { $_.User -ne 'Default' -and $_.User -ne 'Anonymous' -and $_.Access -notmatch '(?i)^none$' })
        [void]$L.Add('')
        if (@($real).Count -eq 0) {
            [void]$L.Add('NO NAMED DELEGATE ON ANY FOLDER. Only Default and Anonymous, which are the')
            [void]$L.Add('baseline for everybody rather than a grant to anyone in particular.')
        } else {
            [void]$L.Add("NAMED DELEGATES ($(@($real).Count)):")
            foreach ($r in $real) { [void]$L.Add(("   {0,-10} {1,-34} {2}" -f $r.Folder, $r.User, $r.Access)) }
        }

        if (@($problems).Count -gt 0) {
            [void]$L.Add('')
            [void]$L.Add('NOT EVERYTHING COULD BE READ -- anything below is missing from the grid:')
            foreach ($p in $problems) { [void]$L.Add("   $p") }
        }
        $txtOut.Text = ($L -join [Environment]::NewLine)

        Write-CompletedLog -Action 'Report' -Target (Get-LogTarget $upn 'User Defined Delegates') -Result 'Success' `
            -Detail "Read folder permissions and Send On Behalf. $(@($script:Rows).Count) entry(ies)." -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Load'
    } finally {
        $btnLoad.Enabled = $true
    }
}

function Confirm-Change {
    param([string]$Text, [string]$Title)
    return (Show-ToolDialog($Text, $Title,
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2) -eq [System.Windows.Forms.DialogResult]::OK)
}

function Invoke-AddPerm {
    if (-not $script:Mbx) { $txtOut.Text = 'Load a mailbox first.'; return }
    $scope = "$($cboFolder.SelectedItem)"
    $who   = "$($txtUser.Text)".Trim()
    $lvl   = "$($cboLevel.SelectedItem)"
    if (-not $who) { $txtOut.Text = 'Enter a user to add.'; return }
    if (-not $script:Folders.ContainsKey($scope)) { $txtOut.Text = "This mailbox has no $scope folder."; return }

    $upn  = "$($script:Mbx.PrimarySmtpAddress)"
    $path = "$upn`:\$($script:Folders[$scope])"
    if (-not (Confirm-Change -Title 'Add permission' -Text "Give $who $lvl on the $scope folder of $upn?")) { return }

    try {
        [void](Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' -Parameters @{
            Identity = $path; User = $who; AccessRights = $lvl; ErrorAction = 'Stop' })
        $txtOut.Text = "Added $who as $lvl on $scope. Reloading ..."
        Write-CompletedLog -Action 'Modify' -Target (Get-LogTarget $upn 'User Defined Delegates') -Result 'Success' `
            -Detail "Added $who as $lvl on $scope." -Environment $Environment
        Invoke-Load
    } catch {
        # ALREADY PRESENT IS THE COMMONEST FAILURE and Change is the fix, so say
        # that rather than leaving the operator with a raw exception.
        $m = "$($_.Exception.Message)"
        if ($m -match '(?i)already') {
            $txtOut.Text = "$who already has a permission on $scope. Select that row and use Change level instead." + [Environment]::NewLine + $m
        } else {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add-MailboxFolderPermission'
        }
        Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $upn 'User Defined Delegates') `
            -Detail "Add $who as $lvl on $scope failed. $m" -Environment $Environment
    }
}

function Invoke-ChangePerm {
    $row = Get-SelectedPermRow
    if (-not $row) { $txtOut.Text = 'Select a row to change.'; return }
    $lvl = "$($cboLevel.SelectedItem)"
    $upn = "$($script:Mbx.PrimarySmtpAddress)"
    if (-not (Confirm-Change -Title 'Change level' -Text "Change $($row.User) on $($row.Folder) from $($row.Access) to $lvl?")) { return }

    try {
        [void](Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{
            Identity = $row.Path; User = $row.User; AccessRights = $lvl; ErrorAction = 'Stop' })
        $txtOut.Text = "Changed $($row.User) on $($row.Folder) to $lvl. Reloading ..."
        Write-CompletedLog -Action 'Modify' -Target (Get-LogTarget $upn 'User Defined Delegates') -Result 'Success' `
            -Detail "Set $($row.User) to $lvl on $($row.Folder)." -Environment $Environment
        Invoke-Load
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-MailboxFolderPermission'
        Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $upn 'User Defined Delegates') `
            -Detail "Set $($row.User) to $lvl on $($row.Folder) failed. $($_.Exception.Message)" -Environment $Environment
    }
}

function Invoke-RemovePerm {
    $row = Get-SelectedPermRow
    if (-not $row) { $txtOut.Text = 'Select a row to remove.'; return }
    $upn = "$($script:Mbx.PrimarySmtpAddress)"

    # DEFAULT AND ANONYMOUS CANNOT BE DELETED. Exchange refuses, and the error
    # is unhelpful. Setting them to None is the equivalent action and is what
    # the operator meant.
    $isBuiltIn = ($row.User -eq 'Default' -or $row.User -eq 'Anonymous')

    $msg = if ($isBuiltIn) {
        "$($row.User) is a built-in entry and cannot be deleted -- it is the baseline for everybody with no explicit permission." + [Environment]::NewLine + [Environment]::NewLine +
        "Set it to None on $($row.Folder) instead?"
    } else {
        "Remove $($row.User) from the $($row.Folder) folder of $upn?" + [Environment]::NewLine + [Environment]::NewLine +
        "They currently have: $($row.Access)"
    }
    if (-not (Confirm-Change -Title 'Remove permission' -Text $msg)) { return }

    try {
        if ($isBuiltIn) {
            [void](Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{
                Identity = $row.Path; User = $row.User; AccessRights = 'None'; ErrorAction = 'Stop' })
            $txtOut.Text = "Set $($row.User) to None on $($row.Folder). Reloading ..."
            Write-CompletedLog -Action 'Modify' -Target (Get-LogTarget $upn 'User Defined Delegates') -Result 'Success' `
                -Detail "Set built-in $($row.User) to None on $($row.Folder)." -Environment $Environment
        } else {
            [void](Invoke-EXOCommand -Name 'Remove-MailboxFolderPermission' -Parameters @{
                Identity = $row.Path; User = $row.User; Confirm = $false; ErrorAction = 'Stop' })
            $txtOut.Text = "Removed $($row.User) from $($row.Folder). Reloading ..."
            Write-CompletedLog -Action 'Modify' -Target (Get-LogTarget $upn 'User Defined Delegates') -Result 'Success' `
                -Detail "Removed $($row.User) from $($row.Folder)." -Environment $Environment
        }
        Invoke-Load
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-MailboxFolderPermission'
        Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $upn 'User Defined Delegates') `
            -Detail "Remove $($row.User) from $($row.Folder) failed. $($_.Exception.Message)" -Environment $Environment
    }
}

function Invoke-SobAdd {
    if (-not $script:Mbx) { $txtOut.Text = 'Load a mailbox first.'; return }
    $who = "$($txtSob.Text)".Trim()
    if (-not $who) { $txtOut.Text = 'Enter a user to add.'; return }
    $upn = "$($script:Mbx.PrimarySmtpAddress)"
    if (-not (Confirm-Change -Title 'Add Send On Behalf' -Text "Let $who send on behalf of $upn?")) { return }

    try {
        # @{Add=} RATHER THAN A WHOLE-LIST ASSIGNMENT. Passing an array replaces
        # everything, so anyone already on the list would be silently dropped --
        # and nobody would notice until they could not send.
        [void](Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $upn; GrantSendOnBehalfTo = @{ Add = $who }; ErrorAction = 'Stop' })
        $txtOut.Text = "Added $who to Send On Behalf. Reloading ..."
        Write-CompletedLog -Action 'Modify' -Target (Get-LogTarget $upn 'User Defined Delegates') -Result 'Success' `
            -Detail "Added $who to GrantSendOnBehalfTo." -Environment $Environment
        Invoke-Load
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-Mailbox GrantSendOnBehalfTo Add'
        Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $upn 'User Defined Delegates') `
            -Detail "Add $who to GrantSendOnBehalfTo failed. $($_.Exception.Message)" -Environment $Environment
    }
}

function Invoke-SobRemove {
    if (-not $script:Mbx) { $txtOut.Text = 'Load a mailbox first.'; return }
    if (-not $lstSob.SelectedItem) { $txtOut.Text = 'Select somebody in the list to remove.'; return }
    $who = "$($lstSob.SelectedItem)"
    $upn = "$($script:Mbx.PrimarySmtpAddress)"
    if (-not (Confirm-Change -Title 'Remove Send On Behalf' -Text "Stop $who sending on behalf of $upn?")) { return }

    try {
        [void](Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $upn; GrantSendOnBehalfTo = @{ Remove = $who }; ErrorAction = 'Stop' })
        $txtOut.Text = "Removed $who from Send On Behalf. Reloading ..."
        Write-CompletedLog -Action 'Modify' -Target (Get-LogTarget $upn 'User Defined Delegates') -Result 'Success' `
            -Detail "Removed $who from GrantSendOnBehalfTo." -Environment $Environment
        Invoke-Load
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-Mailbox GrantSendOnBehalfTo Remove'
        Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $upn 'User Defined Delegates') `
            -Detail "Remove $who from GrantSendOnBehalfTo failed. $($_.Exception.Message)" -Environment $Environment
    }
}
#endregion

$btnLoad.Add_Click({ try { Invoke-Load } catch { Show-HandlerError -ErrorRecord $_ -Context 'Load' } })
$btnAdd.Add_Click({ try { Invoke-AddPerm } catch { Show-HandlerError -ErrorRecord $_ -Context 'Add' } })
$btnChange.Add_Click({ try { Invoke-ChangePerm } catch { Show-HandlerError -ErrorRecord $_ -Context 'Change level' } })
$btnRemove.Add_Click({ try { Invoke-RemovePerm } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove' } })
$btnSobAdd.Add_Click({ try { Invoke-SobAdd } catch { Show-HandlerError -ErrorRecord $_ -Context 'Add on behalf' } })
$btnSobRemove.Add_Click({ try { Invoke-SobRemove } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove on behalf' } })

# Selecting a row puts its folder and level in the boxes, so Change level is a
# two-click operation rather than a re-typing exercise.
$gridPerm.Add_SelectionChanged({
    try {
        $row = Get-SelectedPermRow
        if (-not $row) { return }
        $cboFolder.SelectedItem = $row.Folder
        $first = (@($row.Access -split ',')[0]).Trim()
        if ($script:Levels -contains $first) { $cboLevel.SelectedItem = $first }
    } catch { }
})

$txtWho.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        try { Invoke-Load } catch { Show-HandlerError -ErrorRecord $_ -Context 'Load' }
    }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'User Defined Delegates' | Out-Null

$form.Add_Shown({
    try {
        if ($SmtpAddress) { $txtWho.Text = "$SmtpAddress" }
        $txtOut.Text = 'Enter an alias or address and press Load.' + [Environment]::NewLine +
            [Environment]::NewLine +
            "Environment: $(if ($Environment) { $Environment } else { '(none passed)' })" + [Environment]::NewLine +
            [Environment]::NewLine +
            'Shows Inbox, Calendar, Tasks, Contacts and Notes permissions plus Send On' + [Environment]::NewLine +
            'Behalf. Add, Change and Remove write to the mailbox immediately and each' + [Environment]::NewLine +
            'one asks first.'
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening User Defined Delegates' }
})

[void]$form.ShowDialog()
