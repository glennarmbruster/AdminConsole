#Requires -Version 5.1
<#
    ParkInPlace.ps1
    Hide tools from THIS TENANT's console without moving any files.

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    2026-09-16  The tenant-park note is no longer overwritten by 'Unsaved
                changes.' one line later, and the section checklist ticks from
                the list the chosen scope saves to -- reading the tenant's
                sections while writing the personal one silently un-parked a
                person's own sections.
    2026-09-16  The Parked tick reflects the CHOSEN scope, not the union of both
                lists: the save wrote every ticked row into whichever store the
                radio named, so 'everyone' pushed personal hides onto the tenant.
                The section list follows this operator's own sections.
#>
# TOOLKIT-HELP-RUNS: Reads the script folder and writes this tenant's parked list. Moves no files and touches no mailbox.
# TOOLKIT-VARIANT: Any
# TOOLKIT-HELP-BEGIN
# Park a tool IN PLACE: the script stays in the folder, its button disappears
# from this tenant's console, and Add-ConsoleAdditions stops offering it.
#
# WHY THIS IS NOT DisabledTools\. That MOVES the file, for every tenant at once,
# and is the right answer when a tool is retired. This is per TENANT and moves
# nothing. The on-premises Exchange tools are real tools worth keeping -- they
# are simply meaningless in a cloud-only tenant, and one folder serving three
# tenants cannot say that by moving a file, because the file is shared and the
# answer is not.
#
# TWO WAYS TO PARK:
#
#   By capability   Park 'On-premises Exchange' and every tool that declares
#                   itself on-premises goes with it. Parking the Connect button
#                   on its own leaves twenty tools on the grid that cannot work
#                   without it, each greying out with its own explanation at the
#                   moment somebody presses it.
#
#   By script       Tick the ones you want. Sort any column, filter by name,
#                   tick several, park them together. Project tools under
#                   Projects\ are listed by their folder path.
#
#   By section      Tick a section name and the whole section -- heading and
#                   every button, including ones added later -- is left off
#                   this tenant's console. Independent of the script ticks:
#                   park the section, or some of its scripts, or both.
#
# THERE IS NO SELECT-ALL, deliberately. Parking every tool produces an empty
# console and is never a thing somebody meant to do in one click. Ticking
# thirty boxes is slow on purpose.
#
# UN-PARKING IS THE SAME SCREEN: untick and save. Nothing was moved, so nothing
# has to be found and put back.
#
# TWO SCOPES, PICKED AT THE TOP OF THE WINDOW:
#
#   Everyone   this tenant does not HAVE these tools. Stored in the tenant
#              file, so it travels with the tenant and applies to whoever
#              opens the console.
#   Just me    I do not want to look at these. Stored in your own profile
#              under Config\Users, and nobody else is affected.
#
# A tool the TENANT parks stays hidden whatever an individual picks -- the
# grid's 'Hidden by' column says which list is doing it.
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

# CAPABILITY GROUPS -- the variants a script can declare, in words.
#
# Read from each script's TOOLKIT-VARIANT marker rather than from a list of
# file names here, so a tool added next month is grouped correctly without this
# file knowing it exists.
$script:Groups = @(
    [pscustomobject]@{ Label = 'On-premises Exchange'; Variants = @('OnPrem')
                       Note  = 'Tools that run Exchange cmdlets against an on-premises server.' }
    [pscustomobject]@{ Label = 'On-premises AD';       Variants = @('OnPremAD')
                       Note  = 'Tools that read or write Active Directory directly.' }
    [pscustomobject]@{ Label = 'Everything on-premises'; Variants = @('OnPrem', 'OnPremAD', 'Dual')
                       Note  = 'Both of the above, plus tools that do each in one window.' }
    [pscustomobject]@{ Label = 'Cloud only';           Variants = @('Cloud')
                       Note  = 'Exchange Online, SharePoint Online and Graph tools.' }
)

#region Model -------------------------------------------------------------------
# NEVER OFFERED FOR PARKING, whatever the grid shows.
#
# Park Common.ps1 or the console itself and the toolkit does not start; park
# this tool and there is no way back in. A parking screen that can lock you out
# of parking is not a safety feature. These are hidden from the list entirely
# rather than shown-and-refused, because a row you cannot use is a row somebody
# spends time trying to use.
#
# ONLY THOSE THREE. Everything else -- the publisher, the tester, the revert
# script, the managers -- is listed and parkable. Parking moves nothing and
# the Dashboard launches them by path, not by button, so parking one costs
# nothing and stops Add-ConsoleAdditions offering it. An earlier version kept
# eleven names off the list and the first question was why REVERT-MultiTenant
# could not be parked; there was no good answer.
$script:NeverPark = @(
    'AdminConsole.ps1', 'Common.ps1', 'ParkInPlace.ps1'
)

function Get-ParkCandidate {
    <#
        Every script that MAY be parked, with what is known about it.

        MineOnly picks WHICH list the Parked tick reflects: this operator's
        personal hides, or the tenant's. It is not a display filter -- the save
        writes back exactly what is ticked, so the tick and the destination
        have to name the same list.

        OnButton says whether Tools.json currently gives it a button. A script
        with no button is already invisible, so parking it changes nothing
        today -- but it still belongs on the list, because parking is also what
        stops Add-ConsoleAdditions offering it on the next publish.
    #>
    param([switch]$MineOnly)

    # BOTH LISTS, KEPT APART. The grid has to say WHY a script is hidden --
    # ticking a box that is already ticked because the tenant parked it, and
    # then finding you cannot untick it, is the confusing half of this.
    $parked = Get-ParkedScriptIndex
    $userParked = (Get-UserParked).Scripts

    $onButton = @{}
    try {
        $tp = Get-ConfigFilePath -Name 'Tools.json'
        if (Test-Path -LiteralPath $tp) {
            foreach ($t in @(Get-Content -LiteralPath $tp -Raw -Encoding UTF8 | ConvertFrom-Json)) {
                $k = Get-ParkedKey -Script "$($t.Script)"
                if ($k) { $onButton[$k] = "$($t.Name)" }
            }
        }
    } catch { }

    # ROOT AND Projects\, NOTHING ELSE. Project tools live under
    # Projects\<name>\ and Tools.json points at them there, so they are as
    # parkable as anything in the root. Config\, Logs\, DisabledTools\ and
    # the rest hold no tools and are not walked.
    $files = @(Get-ChildItem -LiteralPath $root -Filter '*.ps1' -File -ErrorAction SilentlyContinue)
    $projDir = Join-Path $root 'Projects'
    if (Test-Path -LiteralPath $projDir) {
        $files += @(Get-ChildItem -LiteralPath $projDir -Filter '*.ps1' -File -Recurse -ErrorAction SilentlyContinue)
    }

    $rows = @()
    foreach ($f in @($files | Sort-Object FullName)) {
        if ($script:NeverPark -contains $f.Name) { continue }
        if ($f.Name -like '*Query.ps1') { continue }

        # Script is shown and saved as the path RELATIVE TO THE ROOT --
        # 'Foo.ps1' or 'Projects\X\Foo.ps1' -- because project folders reuse
        # file names and the bare leaf cannot tell them apart.
        $rel = $f.FullName.Substring($root.TrimEnd('\').Length + 1)
        $key = Get-ParkedKey -Script $rel
        $byTenant = $parked.ContainsKey($key)
        $byMe     = $userParked.ContainsKey($key)
        # THE TICK MEANS THIS SCOPE'S LIST, NOT THE UNION OF BOTH. It used to
        # be ($byTenant -or $byMe), which read correctly on screen and saved
        # catastrophically: the save writes every ticked row into whichever
        # store the radio names, so in 'everyone' scope one click pushed this
        # operator's PRIVATE hides onto the whole tenant, and in 'just me'
        # scope it copied the tenant list into the profile, where a later
        # tenant un-park could never reach it. HiddenBy still shows both, so
        # nothing is lost from the display.
        $rows += [pscustomobject]@{
            Parked   = [bool]$(if ($MineOnly) { $byMe } else { $byTenant })
            Script   = $rel
            Button   = $(if ($onButton.ContainsKey($key)) { $onButton[$key] } else { '' })
            Variant  = (Get-ScriptVariant -Path $f.FullName)
            HiddenBy = $(if ($byTenant -and $byMe) { 'tenant + me' } elseif ($byTenant) { 'tenant' } elseif ($byMe) { 'me' } else { '' })
            ByTenant = [bool]$byTenant
            Reason   = $(if ($byTenant) { "$($parked[$key].Reason)" } elseif ($byMe) { "$($userParked[$key].Reason)" } else { '' })
            Modified = $f.LastWriteTime
        }
    }
    return @($rows)
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = $(if (Get-ActiveTenantLabel) { "Park In Place -- $(Get-ActiveTenantLabel)" } else { 'Park In Place' })
$form.ClientSize      = New-Object System.Drawing.Size(1080, 854)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'CenterScreen'
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

# THE TENANT IS NAMED AT THE TOP AND IT IS NOT DECORATION. Everything this
# window does affects one tenant, and the commonest way to get it wrong is not
# noticing which one is open.
$lblWho          = New-Object System.Windows.Forms.Label
$lblWho.Location = New-Object System.Drawing.Point(12, 10)
$lblWho.Size     = New-Object System.Drawing.Size(1056, 20)
$lblWho.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWho.Text     = $(if (Get-ActiveTenantLabel) {
                        "Parking affects '$(Get-ActiveTenantLabel)' only. Other tenants keep every tool."
                    } else {
                        'Single-tenant install -- parking affects this console.'
                    })
$form.Controls.Add($lblWho)

# ---- who is this parking for ---------------------------------------------------
# TWO DIFFERENT STATEMENTS, AND THEY ARE NOT INTERCHANGEABLE.
#
#   Tenant   this tenant does not HAVE these tools -- cloud-only, so the
#            on-premises ones are meaningless here. A fact about the tenant,
#            true for whoever opens the console.
#   Me       I do not want to LOOK at these. A preference, mine alone.
#
# Tenant parking wins: a person can hide more, nobody can un-hide what the
# tenant parked. Which is why the two lists are stored separately rather than
# merged -- see Get-EffectiveParked.
$grpWho          = New-Object System.Windows.Forms.GroupBox
$grpWho.Text     = 'Who is this for'
$grpWho.Location = New-Object System.Drawing.Point(12, 34)
$grpWho.Size     = New-Object System.Drawing.Size(1056, 48)
$form.Controls.Add($grpWho)

$rbMe = New-Object System.Windows.Forms.RadioButton
$rbMe.Text     = "Just me ($(Get-UserProfileKey))"
$rbMe.Location = New-Object System.Drawing.Point(12, 18)
$rbMe.Size     = New-Object System.Drawing.Size(240, 22)
$grpWho.Controls.Add($rbMe)

$rbTenant = New-Object System.Windows.Forms.RadioButton
$rbTenant.Text     = $(if (Get-ActiveTenantLabel) { "Everyone in '$(Get-ActiveTenantLabel)'" } else { 'Everyone using this console' })
$rbTenant.Location = New-Object System.Drawing.Point(262, 18)
$rbTenant.Size     = New-Object System.Drawing.Size(300, 22)
$grpWho.Controls.Add($rbTenant)

# DEFAULTS TO WHICHEVER LIST IS ALREADY IN USE, so the common case -- open it,
# tick one more, save -- cannot silently move a tenant's list into one person's
# profile or the other way round.
if ((Get-UserParked).Scripts.Count -gt 0 -or @((Get-UserParked).Sections).Count -gt 0) {
    $rbMe.Checked = $true
} else {
    $rbTenant.Checked = $true
}

$lblWhoNote           = New-Object System.Windows.Forms.Label
$lblWhoNote.Location  = New-Object System.Drawing.Point(576, 20)
$lblWhoNote.Size      = New-Object System.Drawing.Size(468, 20)
$lblWhoNote.ForeColor = [System.Drawing.Color]::DimGray
$lblWhoNote.Text      = 'A tool the tenant parks stays hidden for everybody, whatever a person picks.'
$grpWho.Controls.Add($lblWhoNote)

# ---- capability parking --------------------------------------------------------
$grpCap          = New-Object System.Windows.Forms.GroupBox
$grpCap.Text     = 'Park a whole capability'
$grpCap.Location = New-Object System.Drawing.Point(12, 88)
$grpCap.Size     = New-Object System.Drawing.Size(1056, 74)
$form.Controls.Add($grpCap)

$cboCap               = New-Object System.Windows.Forms.ComboBox
$cboCap.DropDownStyle = 'DropDownList'
$cboCap.Location      = New-Object System.Drawing.Point(12, 24)
$cboCap.Size          = New-Object System.Drawing.Size(240, 22)
foreach ($g in $script:Groups) { [void]$cboCap.Items.Add($g.Label) }
$cboCap.SelectedIndex = 0
$grpCap.Controls.Add($cboCap)

$lblCapNote          = New-Object System.Windows.Forms.Label
$lblCapNote.Location = New-Object System.Drawing.Point(262, 27)
$lblCapNote.Size     = New-Object System.Drawing.Size(560, 34)
$lblCapNote.ForeColor = [System.Drawing.Color]::DimGray
$grpCap.Controls.Add($lblCapNote)

$btnCapPark          = New-Object System.Windows.Forms.Button
$btnCapPark.Text     = 'Tick these'
$btnCapPark.Location = New-Object System.Drawing.Point(830, 22)
$btnCapPark.Size     = New-Object System.Drawing.Size(100, 26)
$grpCap.Controls.Add($btnCapPark)

$btnCapClear          = New-Object System.Windows.Forms.Button
$btnCapClear.Text     = 'Untick these'
$btnCapClear.Location = New-Object System.Drawing.Point(938, 22)
$btnCapClear.Size     = New-Object System.Drawing.Size(100, 26)
$grpCap.Controls.Add($btnCapClear)

# ---- section parking -----------------------------------------------------------
# A SECTION, INDEPENDENTLY OF ITS SCRIPTS. Ticking the eight Projects scripts
# hides Projects today; the ninth one added next month comes straight back.
# Parking the section says what was meant -- this tenant does not do these --
# and holds however many tools the section gains later. The two lists are
# separate on purpose: park a section and keep its scripts unticked, park
# some of its scripts and leave the section up, or both. Tools.json stays one
# file for every tenant.
$grpSec          = New-Object System.Windows.Forms.GroupBox
$grpSec.Text     = 'Park whole sections'
$grpSec.Location = New-Object System.Drawing.Point(12, 168)
$grpSec.Size     = New-Object System.Drawing.Size(1056, 74)
$form.Controls.Add($grpSec)

$clbSec              = New-Object System.Windows.Forms.CheckedListBox
$clbSec.Location     = New-Object System.Drawing.Point(12, 20)
$clbSec.Size         = New-Object System.Drawing.Size(560, 46)
$clbSec.CheckOnClick = $true
$clbSec.MultiColumn  = $true
$clbSec.ColumnWidth  = 180
$clbSec.BorderStyle  = 'FixedSingle'
$grpSec.Controls.Add($clbSec)

$lblSecNote           = New-Object System.Windows.Forms.Label
$lblSecNote.Location  = New-Object System.Drawing.Point(584, 22)
$lblSecNote.Size      = New-Object System.Drawing.Size(460, 44)
$lblSecNote.ForeColor = [System.Drawing.Color]::DimGray
$lblSecNote.Text      = 'A ticked section and every tool in it stay off this tenant''s console, ' +
                        'including tools added to it later. Independent of the script ticks below.'
$grpSec.Controls.Add($lblSecNote)

# Section names come from the SAME list the console builds from: this
# operator's personal sections when they have them, Sections.json otherwise.
# Reading only the shared file meant somebody with a personal layout could not
# park the sections their console actually shows, and saw names it does not.
function Update-SectionChecklist {
    <#
        Fill the section checklist for the scope currently chosen.

        THE TICK AND THE DESTINATION MUST NAME THE SAME LIST -- the lesson the
        script rows learned above, applied here. Test-SectionParked answers for
        the TENANT only, so ticking from it while saving to the personal list
        meant a person's own parked sections came up unticked and were silently
        un-parked by the next save, and the tenant's came along into their
        profile. Read whichever list this scope writes.

        REBUILT ON EVERY SCOPE FLIP, for the same reason the grid is.
    #>
    $clbSec.Items.Clear()
    $mine = $rbMe.Checked
    $on = @{}
    try {
        $src = $(if ($mine) { @((Get-UserParked).Sections) } else { @(Get-ParkedSection) })
        foreach ($x in $src) { if ("$x".Trim()) { $on["$x".Trim().ToLowerInvariant()] = $true } }
    } catch { }
    foreach ($n in @($script:SectionNames)) {
        if (-not $n) { continue }
        [void]$clbSec.Items.Add($n, $on.ContainsKey($n.ToLowerInvariant()))
    }
}

# THE NAMES THIS OPERATOR'S CONSOLE ACTUALLY SHOWS: their own sections when
# they have them, Sections.json otherwise. Reading only the shared file meant
# somebody with a personal layout could not park the sections they see, and
# saw names their console does not use.
$script:SectionNames = @()
try {
    $uSec = $null
    try { $uSec = Get-UserLayoutProperty -Name 'Sections' } catch { }
    if ($uSec -and @($uSec).Count -gt 0) {
        $script:SectionNames = @(@($uSec) | ForEach-Object { "$($_.Name)".Trim() } | Where-Object { $_ })
    } else {
        $secCfg = Get-Content -LiteralPath (Get-ConfigFilePath -Name 'Sections.json') -Raw -ErrorAction Stop |
                  ConvertFrom-Json -ErrorAction Stop
        $script:SectionNames = @(@($secCfg.Sections) | ForEach-Object { "$($_.Name)".Trim() } | Where-Object { $_ })
    }
    Update-SectionChecklist
} catch {
    $lblSecNote.Text = "The section list could not be read: $($_.Exception.Message)"
}

# ---- filter --------------------------------------------------------------------
$lblFind          = New-Object System.Windows.Forms.Label
$lblFind.Text     = 'Filter:'
$lblFind.Location = New-Object System.Drawing.Point(12, 254)
$lblFind.Size     = New-Object System.Drawing.Size(44, 20)
$form.Controls.Add($lblFind)

$txtFind          = New-Object System.Windows.Forms.TextBox
$txtFind.Location = New-Object System.Drawing.Point(58, 251)
$txtFind.Size     = New-Object System.Drawing.Size(260, 22)
$form.Controls.Add($txtFind)

$chkOnlyParked          = New-Object System.Windows.Forms.CheckBox
$chkOnlyParked.Text     = 'Show only parked'
$chkOnlyParked.Location = New-Object System.Drawing.Point(332, 253)
$chkOnlyParked.Size     = New-Object System.Drawing.Size(140, 20)
$form.Controls.Add($chkOnlyParked)

$lblCount          = New-Object System.Windows.Forms.Label
$lblCount.Location = New-Object System.Drawing.Point(480, 254)
$lblCount.Size     = New-Object System.Drawing.Size(588, 20)
$lblCount.TextAlign = 'MiddleRight'
$form.Controls.Add($lblCount)

# ---- grid ----------------------------------------------------------------------
$grid                        = New-Object System.Windows.Forms.DataGridView
$grid.Location               = New-Object System.Drawing.Point(12, 280)
$grid.Size                   = New-Object System.Drawing.Size(1056, 500)
$grid.Anchor                 = 'Top,Left,Right,Bottom'
$grid.AllowUserToAddRows     = $false
$grid.AllowUserToDeleteRows  = $false
$grid.SelectionMode          = 'FullRowSelect'
$grid.MultiSelect            = $true
$grid.AutoSizeColumnsMode    = 'Fill'
$grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$form.Controls.Add($grid)

$btnSave          = New-Object System.Windows.Forms.Button
$btnSave.Text     = 'Save'
$btnSave.Location = New-Object System.Drawing.Point(12, 792)
$btnSave.Size     = New-Object System.Drawing.Size(120, 30)
$btnSave.Anchor   = 'Left,Bottom'
$form.Controls.Add($btnSave)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(144, 798)
$lblStatus.Size     = New-Object System.Drawing.Size(790, 20)
$lblStatus.Anchor   = 'Left,Right,Bottom'
$form.Controls.Add($lblStatus)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(948, 792)
$btnClose.Size     = New-Object System.Drawing.Size(120, 30)
$btnClose.Anchor   = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)
#endregion

#region Grid --------------------------------------------------------------------
# THE MODEL IS THE TRUTH, THE GRID IS A VIEW OF IT.
#
# Ticks have to survive filtering and sorting. If the checkbox state lived only
# in the DataGridView, typing in the filter box would rebuild the rows and throw
# away every tick made so far -- silently, which is the worst way to lose work.
# So a tick writes to $script:Model and the grid is redrawn from it.
$script:Model = @(Get-ParkCandidate -MineOnly:$rbMe.Checked)

function Get-ModelRow {
    param([string]$Script)
    $k = "$Script".ToLowerInvariant()
    foreach ($m in $script:Model) { if ($m.Script.ToLowerInvariant() -eq $k) { return $m } }
    return $null
}

function Update-Grid {
    <#
        Redraw from the model, honouring the filter and the current sort.

        THE SORT IS PRESERVED ACROSS REDRAWS. Retyping in the filter box and
        having the list jump back to alphabetical is the kind of small
        rudeness that makes a screen tiring to use.
    #>
    $q = "$($txtFind.Text)".Trim()
    $rows = @($script:Model)
    if ($q) {
        $rows = @($rows | Where-Object {
            $_.Script -like "*$q*" -or $_.Button -like "*$q*" -or $_.Variant -like "*$q*"
        })
    }
    if ($chkOnlyParked.Checked) { $rows = @($rows | Where-Object { $_.Parked }) }

    $sortCol = ''; $sortAsc = $true
    if ($grid.SortedColumn) { $sortCol = $grid.SortedColumn.Name; $sortAsc = ($grid.SortOrder -ne 'Descending') }

    $t = New-Object System.Data.DataTable
    [void]$t.Columns.Add('Parked',  [bool])
    [void]$t.Columns.Add('Script',  [string])
    [void]$t.Columns.Add('Button',  [string])
    [void]$t.Columns.Add('Variant', [string])
    [void]$t.Columns.Add('Reason',  [string])
    [void]$t.Columns.Add('Modified',[datetime])
    foreach ($r in $rows) {
        [void]$t.Rows.Add($r.Parked, $r.Script, $r.Button, $r.Variant, $r.Reason, $r.Modified)
    }

    $grid.DataSource = $t
    if ($grid.Columns.Contains('Parked')) {
        $grid.Columns['Parked'].Width = 60
        $grid.Columns['Parked'].AutoSizeMode = 'None'
        # NO SORT ON THE TICK COLUMN, and no select-all in its header either.
        # WinForms does not put a select-all box on a bound boolean column, so
        # there is nothing to remove -- but sorting BY the tick invites 'sort
        # then drag down the column', which is select-all with extra steps.
        $grid.Columns['Parked'].SortMode = 'NotSortable'
    }
    foreach ($c in @('Script','Button','Variant','HiddenBy','Reason','Modified','ByTenant')) {
        if ($grid.Columns.Contains($c)) { $grid.Columns[$c].ReadOnly = $true }
    }
    # ByTenant rides along for the guard in the tick handler; it is not a
    # column anyone needs to read -- 'Hidden by' already says it in words.
    if ($grid.Columns.Contains('ByTenant')) { $grid.Columns['ByTenant'].Visible = $false }
    if ($grid.Columns.Contains('Modified')) { $grid.Columns['Modified'].DefaultCellStyle.Format = 'yyyy-MM-dd HH:mm' }
    if ($grid.Columns.Contains('Script'))   { $grid.Columns['Script'].FillWeight = 90 }
    if ($grid.Columns.Contains('Reason'))   { $grid.Columns['Reason'].FillWeight = 120 }

    if ($sortCol -and $grid.Columns.Contains($sortCol)) {
        $dir = if ($sortAsc) { [System.ComponentModel.ListSortDirection]::Ascending }
               else          { [System.ComponentModel.ListSortDirection]::Descending }
        try { $grid.Sort($grid.Columns[$sortCol], $dir) } catch { }
    }

    $p = @($script:Model | Where-Object { $_.Parked }).Count
    $lblCount.Text = "$($rows.Count) shown of $($script:Model.Count)   --   $p parked"
}

# A tick writes straight through to the model. CommitEdit forces the cell to
# finish editing first: without it the value is read one click behind, which
# looks like the first tick not registering.
$grid.Add_CurrentCellDirtyStateChanged({
    if ($grid.IsCurrentCellDirty) {
        $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})
$grid.Add_CellValueChanged({
    param($s, $e)
    if ($e.RowIndex -lt 0) { return }
    if ($grid.Columns[$e.ColumnIndex].Name -ne 'Parked') { return }
    try {
        $name = "$($grid.Rows[$e.RowIndex].Cells['Script'].Value)"
        $m = Get-ModelRow -Script $name
        if ($m) {
            $now = [bool]$grid.Rows[$e.RowIndex].Cells['Parked'].Value
            # UN-TICKING HERE CANNOT UN-HIDE A TENANT PARK, and saying so is
            # better than refusing the click. In this scope the tick is the
            # PERSONAL list only, so a tenant-parked script shows unticked with
            # 'tenant' in HiddenBy; removing it from a personal list it was
            # never on changes nothing, and the button still will not appear.
            $note = ''
            if ($rbMe.Checked -and $m.ByTenant) {
                $note = "'$name' is parked for the whole tenant -- your personal list does not control it. Switch to '$($rbTenant.Text)' to un-park it."
            }
            $m.Parked = $now
            if (-not $m.Parked) { $m.Reason = '' }
        }
        $p = @($script:Model | Where-Object { $_.Parked }).Count
        $lblCount.Text = ($lblCount.Text -replace '--\s+\d+ parked', "--   $p parked")
        # THE NOTE OUTRANKS 'Unsaved changes.' -- it used to be written first
        # and then overwritten one line later, which is the same as not
        # writing it: the operator unticked a tenant park, was told nothing,
        # saved, and the button was still gone.
        $lblStatus.Text = $(if ($note) { $note } else { 'Unsaved changes.' })
    } catch { }
})

# SWITCHING SCOPE REBUILDS THE GRID. The ticks mean a different list either
# side of that radio button, and leaving the old ones on screen is how somebody
# saves one list's contents into the other.
$rbMe.Add_CheckedChanged({
    try {
        $script:Model = @(Get-ParkCandidate -MineOnly:$rbMe.Checked)
        Update-Grid
        Update-SectionChecklist
        $lblStatus.Text = $(if ($rbMe.Checked) {
            "Showing what $(Get-UserProfileKey) has hidden. Tenant parks are listed under HiddenBy but are not ticked -- they are not yours to change here."
        } else {
            'Showing what is parked for the whole tenant.'
        })
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Switching parking scope' }
})

$txtFind.Add_TextChanged({ try { Update-Grid } catch { } })
$chkOnlyParked.Add_CheckedChanged({ try { Update-Grid } catch { } })
#endregion

#region Handlers ----------------------------------------------------------------
function Set-CapabilityTicks {
    <#
        Tick or untick every script declaring the selected capability.

        This is a TICK, NOT A SAVE. The grid then shows exactly what is about to
        happen and Save is still a separate press -- 'park twenty tools' is not
        something to do on one click with no look at the list first.
    #>
    param([bool]$On)

    $g = $script:Groups[$cboCap.SelectedIndex]
    if (-not $g) { return }
    $names = @(Get-ScriptsForVariant -Variant $g.Variants)
    $hit = 0
    foreach ($n in $names) {
        $m = Get-ModelRow -Script $n
        if (-not $m) { continue }          # NeverPark, or not in the root
        if ($m.Parked -ne $On) { $hit++ }
        $m.Parked = $On
        if ($On -and -not $m.Reason) { $m.Reason = "$($g.Label) not used by this tenant" }
        if (-not $On) { $m.Reason = '' }
    }
    Update-Grid
    $verb = if ($On) { 'Ticked' } else { 'Unticked' }
    $lblStatus.Text = "$verb $hit of $($names.Count) '$($g.Label)' script(s). Nothing is saved until you press Save."
}

$cboCap.Add_SelectedIndexChanged({
    try {
        $g = $script:Groups[$cboCap.SelectedIndex]
        if (-not $g) { return }
        $n = @(Get-ScriptsForVariant -Variant $g.Variants).Count
        $lblCapNote.Text = "$($g.Note)`r`n$n script(s) declare it."
    } catch { }
})
$btnCapPark.Add_Click({  try { Set-CapabilityTicks -On $true }  catch { Show-HandlerError -ErrorRecord $_ -Context 'Ticking a capability' } })
$btnCapClear.Add_Click({ try { Set-CapabilityTicks -On $false } catch { Show-HandlerError -ErrorRecord $_ -Context 'Unticking a capability' } })

$btnSave.Add_Click({
    try {
        $want = @($script:Model | Where-Object { $_.Parked })
        $secs = @()
        foreach ($i in $clbSec.CheckedIndices) { $secs += "$($clbSec.Items[$i])" }

        # PARKING EVERYTHING IS REFUSED. There is no select-all, but forty
        # individual ticks reach the same place, and an empty console is never
        # what somebody meant. The number is a sanity check, not a policy: it
        # exists to catch a mistake, so it names one.
        if (($want.Count -ge $script:Model.Count) -or
            ($clbSec.Items.Count -gt 0 -and $secs.Count -ge $clbSec.Items.Count)) {
            $null = Show-ToolDialog(
                ('That would park every tool and leave an empty console.' + [Environment]::NewLine + [Environment]::NewLine +
                 'Untick at least one, or use DisabledTools\ if the whole toolkit is being retired.'),
                'Not saved', 'OK', 'Warning')
            return
        }

        $before = $(if ($rbMe.Checked) { (Get-UserParked).Scripts.Count } else { @(Get-ParkedScript).Count })
        $entries = @()
        foreach ($m in $want) {
            $entries += [pscustomobject]@{ Script = ".\$($m.Script)"; Reason = "$($m.Reason)"; When = '' }
        }
        # ROUTED BY THE SCOPE CHOSEN AT THE TOP. Both stores hold the same
        # shape -- entries plus section names -- so only the destination
        # differs; and both are written whole every time, because sending
        # only the scripts would silently un-park every section.
        if ($rbMe.Checked) {
            # Carried across untouched: a parking change must not reset the
            # arrangement or the theme this person already has.
            $layout = Get-UserLayoutOwn
            if (-not $layout) { $layout = [PSCustomObject]@{} }
            Add-Member -InputObject $layout -NotePropertyName 'Parked' -NotePropertyValue ([PSCustomObject][ordered]@{
                Scripts  = @($entries)
                Sections = @(@($secs) | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Sort-Object -Unique)
            }) -Force
            $path = Save-UserLayout -Layout $layout
        } else {
            $path = Save-ParkedScript -Entries $entries -Sections $secs
        }

        $who = $(if ($rbMe.Checked) { "for $(Get-UserProfileKey) only" } else { 'for this tenant' })
        $lblStatus.Text = "Saved $($entries.Count) script(s), $($secs.Count) section(s) parked $who (was $before script(s)). Reopen the console to see the change."
        Write-CompletedLog -Action 'ParkInPlace' -Target $path -Result 'Success' `
            -Detail "$($entries.Count) script(s), $($secs.Count) section(s) parked $who." -Environment $Environment

        $null = Show-ToolDialog(
            ("$($entries.Count) script(s) and $($secs.Count) section(s) are parked $who." + [Environment]::NewLine + [Environment]::NewLine +
             "Written to the Parked block in:" + [Environment]::NewLine + "  $path" + [Environment]::NewLine + [Environment]::NewLine +
             'The console reads this at startup, so close and reopen it to see the buttons change.'),
            'Parked', 'OK', 'Information')
    } catch {
        $lblStatus.Text = 'Save failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Saving the parked list'
        Write-FailedLog -Action 'ParkInPlace' -Target 'Parked list' -ErrorRecord $_ -Environment $Environment
    }
})
#endregion

#region Show --------------------------------------------------------------------
Update-Grid
try {
    $g = $script:Groups[0]
    $lblCapNote.Text = "$($g.Note)`r`n$(@(Get-ScriptsForVariant -Variant $g.Variants).Count) script(s) declare it."
} catch { }
$lblStatus.Text = "Parked block in $(Get-ParkedListPath)"

Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Park In Place' | Out-Null
[void]$form.ShowDialog()
#endregion
