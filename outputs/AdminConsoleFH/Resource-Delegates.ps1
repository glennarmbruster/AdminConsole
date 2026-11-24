#Requires -Version 5.1
<#
    Resource-Delegates.ps1
    Who approves bookings for a room, adding or removing them in bulk, and the
    switch that decides whether any of them are asked.

    DELEGATES AND AutomateProcessing ARE WRITTEN IN ONE CALL. They only mean
    anything together -- a delegate list is inert unless the mode is AutoUpdate,
    and AutoUpdate with no delegates makes the resource unbookable. Two separate
    writes would allow exactly that state to be reached by a failure rather than
    by a choice.

    SET-CALENDARPROCESSING REPLACES THE WHOLE LIST. There is no -Add or
    -Remove for ResourceDelegates: whatever is passed becomes the entire set.
    So every change here is read, modify, write -- and the read is repeated
    immediately before the write, because a list that somebody else edited in
    between would otherwise be silently overwritten by this window's stale copy.

    REMOVING THE LAST DELEGATE FROM AN AutoUpdate RESOURCE MAKES IT PERMANENTLY
    UNAVAILABLE. Requests are marked tentative and wait for an approval that
    nobody can now give. No error is raised anywhere and the room simply stops
    working, so that case is confirmed separately and in different words.

    Nothing is written until Save.
#>
# TOOLKIT-HELP-RUNS: Reads and WRITES ResourceDelegates on a room or equipment mailbox. Changes are immediate on Save.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# The people who approve booking requests for a room, and how to change them.
#
# Type a resource in the box at the top. AN ADDRESS OR JUST THE FIRST PART
# WORKS -- "conf-room-3" is retried as "conf-room-3@<the tenant's primary
# domain>" before it gives up, and the same applies to the delegates you add.
#
# SEVERAL DELEGATES AT ONCE. The add box takes a list: one per line, or
# separated by semicolons or commas. Paste a column out of a spreadsheet or a
# recipient line out of Outlook. Spaces are NOT separators, because a display
# name is a valid thing to type. Every entry is looked up before it goes on the
# list, and the result box says which were added, which were already there and
# which did not resolve -- the ones that failed stay in the box so you can
# correct them without re-typing the rest. Ctrl+Enter adds.
#
# DELEGATES ONLY DO ANYTHING WHEN AutomateProcessing IS AutoUpdate, and that
# switch is on this form. All three values it can take:
#
#   AutoUpdate  -- every request is held TENTATIVE and the delegates below
#                  approve or decline it. If the list is empty, NOTHING IS EVER
#                  APPROVED and the room looks permanently busy to everyone.
#   AutoAccept  -- the room books itself by its booking policy. Nobody vets
#                  anything; the delegates are stored and are not asked.
#   None        -- nothing is automatic. Requests sit in the resource mailbox
#                  until a person opens it and answers by hand.
#
# THE TWO ARE SAVED TOGETHER, in one write, because either alone is a trap: a
# list with the wrong mode does nothing, and AutoUpdate with an empty list stops
# the resource being bookable at all. That combination is confirmed separately
# and in different words, whether you reach it by removing the last delegate or
# by switching the mode.
#
# Calendar Processing remains the tool for everything ELSE about how a room
# decides -- the reply text, recurring meetings, conflict rules. This form
# carries only the one switch the delegate list depends on.
#
# THE WHOLE LIST IS REWRITTEN ON SAVE, because that is the only thing Exchange
# offers. If somebody else changes the delegates while this window is open, Save
# refuses and asks you to reload rather than quietly reverting their work.
#
# A DELEGATE THAT NO LONGER RESOLVES is shown as UNRESOLVED with the raw value
# the service returned. It is usually a deleted account, and it is left exactly
# as it is unless you remove it -- rewriting a value nobody can identify is how
# a list gets quietly corrupted.
#
# For delegates on a PERSON's mailbox, this is the wrong tool: that is folder
# permissions and Send On Behalf, and User Defined Delegates does it.
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

$script:Target  = Get-TargetSmtp -Text $SmtpAddress
$script:Cp      = $null
$script:Mode    = ''
# The list as it stands in this window, one object per delegate.
$script:Working = New-Object System.Collections.ArrayList
# The list exactly as the service last gave it to us. This is what the stale
# check tests against, and it is the reason a second operator's change cannot be
# silently reverted from here.
$script:Loaded  = @()
# The same list as resolved objects, kept so that a delegate the operator has
# already removed from the working list can still be named in the confirmation
# dialog rather than appearing as a raw directory path.
$script:LoadedEntries = @()

#region Read / resolve ----------------------------------------------------------
function Get-CandidateIdentities {
    <#
        What the operator typed, and the same thing with the tenant's primary
        domain on the end.

        BARE ALIASES ARE WHAT PEOPLE HAVE IN THEIR HEADS. Making somebody find
        a full address to add one name to a list is the friction that ends in
        the list not being maintained at all.

        THE DOMAIN COMES FROM Tenant.json, never a literal. TEN002 fails the
        build on a hardcoded domain and it would be wrong at the next tenant
        regardless.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $out = New-Object System.Collections.ArrayList
    $v = "$Text".Trim()
    if (-not $v) { return @() }
    [void]$out.Add($v)

    if ($v -notmatch '@') {
        $dom = ''
        try { $dom = "$(Get-Prop -Object (Get-TenantConfig) -Name 'PrimaryDomain' -Default '')".Trim() } catch { }
        if ($dom) { [void]$out.Add("$v@$dom") }
    }
    return @($out.ToArray())
}

function Resolve-DelegatePrincipal {
    <#
        Turn what was typed into something Set-CalendarProcessing will accept.

        RESOLVED BEFORE IT GOES ON THE LIST, not at save time. An unresolvable
        name accepted into the list would fail the whole write -- one bad entry
        rejects the entire call, because the list is written as a unit -- and it
        would fail at the moment the operator was least expecting it.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Address = ''; Display = ''; Kind = '' }
    $candidates = @(Get-CandidateIdentities -Text $Text)
    if (@($candidates).Count -eq 0) { $r.Reason = 'Type an address or the first part of one.'; return $r }

    foreach ($c in $candidates) {
        try {
            $rec = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $c; ErrorAction = 'Stop' }
            if ($rec) {
                $r.Ok      = $true
                $r.Address = "$($rec.PrimarySmtpAddress)"
                $r.Display = "$($rec.DisplayName)"
                $r.Kind    = "$($rec.RecipientTypeDetails)"
                return $r
            }
        } catch { }
    }

    $r.Reason = "'$("$Text".Trim())' did not resolve. Tried: $(@($candidates) -join ', ')."
    if (@($candidates).Count -eq 1 -and "$Text" -notmatch '@') {
        $r.Reason = "$($r.Reason) No primary domain is set in Tenant.json, so the bare name could not be retried against one -- Tenant Manager sets it."
    }
    return $r
}

function Get-ResourceState {
    <#
        The resource, its type, and its calendar processing.

        A PERSON'S MAILBOX IS REFUSED, and the refusal names the right tool.
        ResourceDelegates exists only on a resource; on a user mailbox the
        thing called a delegate is folder permissions plus Send On Behalf,
        which is a different mechanism with a different tool. Silently reading
        an empty list off a user mailbox would look like an answer.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Cp = $null; Display = ''; Kind = ''; Address = '' }
    $candidates = @(Get-CandidateIdentities -Text $Text)
    if (@($candidates).Count -eq 0) { $r.Reason = 'Enter a resource above and press Look Up.'; return $r }

    $mbx = $null
    $last = ''
    foreach ($c in $candidates) {
        try {
            $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $c; ErrorAction = 'Stop' }
            if ($mbx) { break }
        } catch { $last = $_.Exception.Message }
    }
    if (-not $mbx) {
        $r.Reason = "No mailbox found. Tried: $(@($candidates) -join ', ')."
        if ($last) { $r.Reason = "$($r.Reason)$([Environment]::NewLine)$last" }
        return $r
    }

    $r.Display = "$($mbx.DisplayName)"
    $r.Kind    = "$($mbx.RecipientTypeDetails)"
    $r.Address = "$($mbx.PrimarySmtpAddress)"

    if ($r.Kind -ne 'RoomMailbox' -and $r.Kind -ne 'EquipmentMailbox') {
        $r.Reason = "'$($r.Display)' is a $($r.Kind), not a room or equipment mailbox." + [Environment]::NewLine + [Environment]::NewLine +
            'ResourceDelegates only exists on a resource. Delegates on a person''s mailbox are folder permissions and Send On Behalf, which is a different mechanism -- use User Defined Delegates for that.'
        return $r
    }

    try {
        $r.Cp = Invoke-EXOCommand -Name 'Get-CalendarProcessing' -Parameters @{ Identity = $r.Address; ErrorAction = 'Stop' }
        $r.Ok = $true
    } catch {
        $r.Reason = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-CalendarProcessing'
    }
    return $r
}

function Get-DelegateRawList {
    <#
        ResourceDelegates off a calendar-processing object, as plain strings.

        WHAT COMES BACK IS NOT AN ADDRESS. Exchange returns these as directory
        identities -- typically a canonical name like "contoso.com/Users/Jane
        Doe" -- so the raw value is kept for writing back and resolved
        separately for display. Assuming SMTP here is what would turn a
        perfectly good list into one nobody recognises.
    #>
    param($Cp)

    $out = New-Object System.Collections.ArrayList
    if (-not $Cp) { return @() }
    foreach ($d in @($Cp.ResourceDelegates)) {
        $s = "$d".Trim()
        if ($s) { [void]$out.Add($s) }
    }
    return @($out.ToArray())
}

function Get-DelegateEntry {
    <#
        One raw delegate value, resolved for display where possible.

        AN UNRESOLVED ENTRY IS KEPT, NOT DROPPED. It is usually a deleted
        account, and it is still on the resource whether this window can name
        it or not. Dropping it from the list would delete it on the next Save
        without anybody choosing to -- so it is shown, marked, and written back
        untouched using the exact value the service gave us.
    #>
    param([Parameter(Mandatory)][string]$Raw)

    $e = [PSCustomObject]@{
        Raw = $Raw; Address = ''; Display = ''; Kind = ''; Resolved = $false
    }
    try {
        $rec = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $Raw; ErrorAction = 'Stop' }
        if ($rec) {
            $e.Address  = "$($rec.PrimarySmtpAddress)"
            $e.Display  = "$($rec.DisplayName)"
            $e.Kind     = "$($rec.RecipientTypeDetails)"
            $e.Resolved = $true
        }
    } catch { }
    return $e
}

function Format-DelegateEntry {
    param([Parameter(Mandatory)]$Entry)
    if ($Entry.Resolved) {
        return "$($Entry.Display)  <$($Entry.Address)>   [$($Entry.Kind)]"
    }
    return "UNRESOLVED -- $($Entry.Raw)   (kept exactly as stored; remove it if the account is gone)"
}

function Get-WriteValue {
    <#
        What goes to the service for one entry: THE VALUE THE SERVICE ITSELF
        GAVE US, for anything read from the resource, and the resolved address
        for anything typed in this window.

        BOTH ARE $Entry.Raw, and that is the point. Writing the SMTP address
        back in place of the canonical name Exchange returned would be a
        needless translation, and it would break every comparison in this file:
        the "nothing has changed" test, the added/removed diff and the check for
        somebody else's edit all compare against the list as it was read, so the
        two have to be the same strings. An earlier draft returned Address here
        and would have reported every untouched delegate as both added and
        removed on the first save.
    #>
    param([Parameter(Mandatory)]$Entry)
    return $Entry.Raw
}

function Get-DelegateKey {
    <#
        Who an entry IS, for telling one person from another.

        SEPARATE FROM Get-WriteValue ON PURPOSE. A delegate already on the
        resource arrives as a canonical name; the same person typed into the box
        arrives as an SMTP address. Comparing the strings that get WRITTEN would
        call those two different people and put both on the list.
    #>
    param([Parameter(Mandatory)]$Entry)
    if ($Entry.Resolved -and $Entry.Address) { return "$($Entry.Address)".ToLowerInvariant() }
    return "$($Entry.Raw)".ToLowerInvariant()
}

function Get-ComparableList {
    <# A list flattened for comparison: lower case, sorted, joined. #>
    param($Values)
    return (@($Values | ForEach-Object { "$_".Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object) -join '|')
}
#endregion

#region Form --------------------------------------------------------------------
# NO Anchor ON ANYTHING. The anchors were left off for a scaler that
# surrendered them, and it was deleted on 2026-09-11 (AGENTS.md rule 8). This
# window is fixed-layout in a sizable frame, which is how the whole toolkit
# behaved before scaling existed -- nothing to put back.
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Resource Delegates'
$form.ClientSize      = New-Object System.Drawing.Size(940, 784)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$lblHead          = New-Object System.Windows.Forms.Label
$lblHead.Location = New-Object System.Drawing.Point(14, 12)
$lblHead.Size     = New-Object System.Drawing.Size(840, 22)
$lblHead.Font     = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($lblHead)

$lblModeCap          = New-Object System.Windows.Forms.Label
$lblModeCap.Text     = 'Automate processing'
$lblModeCap.Location = New-Object System.Drawing.Point(14, 42)
$lblModeCap.Size     = New-Object System.Drawing.Size(130, 20)
$lblModeCap.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($lblModeCap)

# DROPDOWNLIST, NOT AN EDITABLE COMBO. AutomateProcessing takes exactly three
# values and a typo in a free-text box would be rejected by the service at save
# time, which is the worst moment to discover it.
#
# THE LOADED VALUE IS ADDED IF IT IS NOT ONE OF THE THREE. A resource carrying
# something this build has not heard of must not have it silently rewritten to
# whatever happened to be first in the list.
$cboMode               = New-Object System.Windows.Forms.ComboBox
$cboMode.Location      = New-Object System.Drawing.Point(150, 39)
$cboMode.Size          = New-Object System.Drawing.Size(180, 24)
$cboMode.DropDownStyle = 'DropDownList'
[void]$cboMode.Items.AddRange(@('AutoAccept', 'AutoUpdate', 'None'))
[void]$form.Controls.Add($cboMode)

$lblModeState          = New-Object System.Windows.Forms.Label
$lblModeState.Location = New-Object System.Drawing.Point(340, 42)
$lblModeState.Size     = New-Object System.Drawing.Size(570, 20)
[void]$form.Controls.Add($lblModeState)

# WHAT THE SELECTED MODE MEANS FOR THIS RESOURCE, on its own row above the list
# rather than buried in the output box. It describes the PENDING selection, not
# what is stored, because the pending one is what Save will make true.
$lblMode          = New-Object System.Windows.Forms.Label
$lblMode.Location = New-Object System.Drawing.Point(14, 68)
$lblMode.Size     = New-Object System.Drawing.Size(900, 34)
$lblMode.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$form.Controls.Add($lblMode)

$grpList          = New-Object System.Windows.Forms.GroupBox
$grpList.Text     = 'Delegates'
$grpList.Location = New-Object System.Drawing.Point(14, 106)
$grpList.Size     = New-Object System.Drawing.Size(912, 316)
[void]$form.Controls.Add($grpList)

$lstDelegates               = New-Object System.Windows.Forms.ListBox
$lstDelegates.Location      = New-Object System.Drawing.Point(12, 22)
$lstDelegates.Size          = New-Object System.Drawing.Size(888, 160)
$lstDelegates.SelectionMode = 'MultiExtended'
[void]$grpList.Controls.Add($lstDelegates)

$lblAddCap          = New-Object System.Windows.Forms.Label
$lblAddCap.Text     = 'Add delegates -- one per line, or separated by ; or ,'
$lblAddCap.Location = New-Object System.Drawing.Point(12, 190)
$lblAddCap.Size     = New-Object System.Drawing.Size(500, 18)
$lblAddCap.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
[void]$grpList.Controls.Add($lblAddCap)

# MULTILINE, so a column pasted out of a spreadsheet or a semicolon-separated
# list pasted out of Outlook both work without being reformatted by hand. That
# is the point of the change: adding eight delegates was eight round trips
# through a one-line box.
#
# ENTER IS NOT HIJACKED HERE. In a multiline box Enter means "next line", which
# is exactly what somebody typing a list wants; Ctrl+Enter adds. A single-line
# box could take Enter and this cannot.
$txtAdd            = New-Object System.Windows.Forms.TextBox
$txtAdd.Location   = New-Object System.Drawing.Point(12, 210)
$txtAdd.Size       = New-Object System.Drawing.Size(500, 58)
$txtAdd.Multiline  = $true
$txtAdd.ScrollBars = 'Vertical'
$txtAdd.WordWrap   = $false
[void]$grpList.Controls.Add($txtAdd)

$lblAddHint          = New-Object System.Windows.Forms.Label
$lblAddHint.Text     = ('Addresses, or just the first part -- a bare name is retried against the tenant domain.' + [Environment]::NewLine +
                        'Display names work too, so spaces are NOT treated as separators.' + [Environment]::NewLine +
                        'Every entry is looked up before it goes on the list. Ctrl+Enter adds.')
$lblAddHint.Location = New-Object System.Drawing.Point(520, 210)
$lblAddHint.Size     = New-Object System.Drawing.Size(380, 58)
[void]$grpList.Controls.Add($lblAddHint)

$btnAdd           = New-Object System.Windows.Forms.Button
$btnAdd.Text      = 'Add delegates'
$btnAdd.Location  = New-Object System.Drawing.Point(12, 276)
$btnAdd.Size      = New-Object System.Drawing.Size(200, 28)
$btnAdd.FlatStyle = 'Flat'
[void]$grpList.Controls.Add($btnAdd)

$btnRemove           = New-Object System.Windows.Forms.Button
$btnRemove.Text      = 'Remove selected'
$btnRemove.Location  = New-Object System.Drawing.Point(220, 276)
$btnRemove.Size      = New-Object System.Drawing.Size(170, 28)
$btnRemove.FlatStyle = 'Flat'
[void]$grpList.Controls.Add($btnRemove)

$btnSave           = New-Object System.Windows.Forms.Button
$btnSave.Text      = 'Save delegates and mode'
$btnSave.Location  = New-Object System.Drawing.Point(14, 432)
$btnSave.Size      = New-Object System.Drawing.Size(240, 28)
$btnSave.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnSave)

$btnReload           = New-Object System.Windows.Forms.Button
$btnReload.Text      = 'Reload'
$btnReload.Location  = New-Object System.Drawing.Point(264, 432)
$btnReload.Size      = New-Object System.Drawing.Size(120, 28)
$btnReload.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnReload)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(14, 470)
$txtOut.Size       = New-Object System.Drawing.Size(912, 300)
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
# B6.
$txtOut.WordWrap   = $true
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
[void]$form.Controls.Add($txtOut)

function Get-PendingMode {
    <# The mode the operator has selected, which is what Save will write. #>
    if (-not $cboMode.SelectedItem) { return '' }
    return "$($cboMode.SelectedItem)"
}

function Update-ModeLabel {
    <#
        What the SELECTED mode would mean for the list in front of the
        operator, phrased as this resource's situation rather than as general
        advice.

        THE PENDING VALUE, NOT THE STORED ONE. The dropdown is the thing being
        edited, so a description of what is still in Exchange would contradict
        the control immediately above it. The stored value is reported beside
        the dropdown instead, and only when the two differ.
    #>
    $n    = @($script:Working).Count
    $mode = Get-PendingMode

    if (-not $script:Cp) {
        $lblMode.Text = ''
        $lblModeState.Text = ''
        return
    }

    if ($mode -and $script:Mode -and $mode -ne $script:Mode) {
        $lblModeState.Text = "currently $script:Mode -- this changes on Save"
    } elseif ($script:Mode) {
        $lblModeState.Text = "unchanged ($script:Mode)"
    } else {
        $lblModeState.Text = ''
    }

    switch ($mode) {
        'AutoUpdate' {
            if ($n -eq 0) {
                $lblMode.Text = 'AutoUpdate WITH NO DELEGATES: every request waits for an approval nobody can give, so the resource appears permanently unavailable. No error is raised anywhere.'
            } else {
                $lblMode.Text = "AutoUpdate -- the $n delegate(s) below approve or decline every booking request for this resource."
            }
        }
        'AutoAccept' {
            $lblMode.Text = "AutoAccept -- the resource books itself by policy. The $n delegate(s) below are stored and are NOT approving anything; the booking policy decides instead."
        }
        'None' {
            $lblMode.Text = "None -- nothing is automatic on this resource. Requests sit in the mailbox unanswered unless somebody opens it, and the $n delegate(s) below are inert."
        }
        default {
            $lblMode.Text = "'$mode' is not one of the three values this build knows. It is left exactly as it is unless you change the dropdown."
        }
    }
}

function Update-DelegateList {
    $lstDelegates.Items.Clear()
    foreach ($e in $script:Working) { [void]$lstDelegates.Items.Add((Format-DelegateEntry -Entry $e)) }
    Update-ModeLabel
}

function Update-DelegateView {
    $script:Cp = $null
    $script:Mode = ''
    $script:Loaded = @()
    $script:LoadedEntries = @()
    $script:Working.Clear()
    $lstDelegates.Items.Clear()
    $lblHead.Text = ''
    Update-ModeLabel

    if (-not "$script:Target".Trim()) {
        $txtOut.Text = 'Enter a room or equipment mailbox above and press Look Up.' + [Environment]::NewLine +
            'An address or just the first part both work.'
        return
    }

    $res = Get-ResourceState -Text $script:Target
    if (-not $res.Ok) { $txtOut.Text = $res.Reason; return }

    # The resolved address, so everything after this point works against what
    # the service actually calls this resource rather than what was typed.
    $script:Target = $res.Address
    $script:Cp     = $res.Cp
    $script:Mode   = "$($res.Cp.AutomateProcessing)"

    $form.Text    = "Resource Delegates: $script:Target"
    $lblHead.Text = "$($res.Display)  --  $($res.Kind)"

    # THE DROPDOWN IS SET FROM THE RESOURCE, and a value this build does not
    # know is added rather than ignored -- selecting nothing would leave the
    # control blank and Save would then write a blank mode over a real one.
    if ($script:Mode -and -not $cboMode.Items.Contains($script:Mode)) {
        [void]$cboMode.Items.Add($script:Mode)
    }
    $cboMode.SelectedItem = $script:Mode

    $raw = @(Get-DelegateRawList -Cp $res.Cp)
    $script:Loaded = $raw
    foreach ($v in $raw) { [void]$script:Working.Add((Get-DelegateEntry -Raw $v)) }
    # A snapshot, not the live list: the working list is about to be edited.
    $script:LoadedEntries = @($script:Working.ToArray())
    Update-DelegateList

    $L = New-Object System.Collections.ArrayList
    [void]$L.Add("$($res.Display)  <$script:Target>")
    [void]$L.Add("  AutomateProcessing : $script:Mode")
    [void]$L.Add("  ResourceDelegates  : $(@($raw).Count) entr(y/ies)")
    [void]$L.Add('')

    $unres = @($script:Working | Where-Object { -not $_.Resolved })
    if (@($unres).Count -gt 0) {
        [void]$L.Add("$(@($unres).Count) delegate(s) DID NOT RESOLVE. That is normally a deleted account. They are shown as stored and are written back untouched -- remove them if they should be gone:")
        foreach ($u in $unres) { [void]$L.Add("    $($u.Raw)") }
        [void]$L.Add('')
    }

    if ($script:Mode -eq 'AutoUpdate' -and @($raw).Count -eq 0) {
        [void]$L.Add('THIS RESOURCE IS EFFECTIVELY BROKEN RIGHT NOW. It is set to AutoUpdate, which holds every request as tentative until a delegate approves it, and it has no delegates. Nobody can book it and nothing reports a fault. Add at least one delegate, or have Calendar Processing set to AutoAccept.')
        [void]$L.Add('')
    }
    if ($script:Mode -ne 'AutoUpdate' -and @($raw).Count -gt 0) {
        [void]$L.Add("The delegates on this resource are not approving anything: AutomateProcessing is '$script:Mode', not AutoUpdate. The list is stored and honoured only in AutoUpdate. This is legitimate -- a resource can carry a list ready for a mode change -- but nobody on it is currently being asked anything.")
        [void]$L.Add('')
    }

    [void]$L.Add('Nothing is written until you press Save. Save rewrites the whole list, because that is the only operation Exchange offers.')
    $txtOut.Text = ($L -join [Environment]::NewLine)

    Write-CompletedLog -Action 'Read' -Target (Get-LogTarget $script:Target 'Resource Delegates') -Result 'Success' `
        -Detail "Read ResourceDelegates on $script:Target. $(@($raw).Count) entr(y/ies), AutomateProcessing=$script:Mode." -Environment $Environment
}

function Invoke-DelegateSave {
    if (-not $script:Cp) { $txtOut.Text = 'Look a resource up first.'; return }

    $want    = @($script:Working | ForEach-Object { Get-WriteValue -Entry $_ })
    $wantCmp = Get-ComparableList -Values $want
    $haveCmp = Get-ComparableList -Values $script:Loaded

    $wantMode    = Get-PendingMode
    if (-not $wantMode) { $txtOut.Text = 'No automate processing value is selected.'; return }
    $modeChanged = ($wantMode -ne $script:Mode)
    $listChanged = ($wantCmp -ne $haveCmp)

    # Nothing to do. Said plainly rather than writing values identical to the
    # stored ones, which would log a change that did not happen.
    if (-not $listChanged -and -not $modeChanged) { $txtOut.Text = 'Nothing has changed.'; return }

    # RE-READ BEFORE WRITING. The whole list is replaced, so a change made by
    # somebody else since this window loaded would be reverted without either
    # of us knowing. Refused rather than merged: guessing which of two people
    # meant what is not this tool's decision to make.
    #
    # THE MODE IS CHECKED THE SAME WAY. It is one field rather than a list, but
    # reverting somebody's switch from AutoUpdate back to AutoAccept is a bigger
    # change than reverting a name, not a smaller one.
    $fresh = Get-ResourceState -Text $script:Target
    if (-not $fresh.Ok) {
        $txtOut.Text = "Could not re-read the resource before writing, so nothing was changed." + [Environment]::NewLine + $fresh.Reason
        return
    }
    $freshRaw  = @(Get-DelegateRawList -Cp $fresh.Cp)
    $freshMode = "$($fresh.Cp.AutomateProcessing)"

    if ((Get-ComparableList -Values $freshRaw) -ne $haveCmp -or $freshMode -ne $script:Mode) {
        $txtOut.Text = "SOMEBODY ELSE HAS CHANGED THIS RESOURCE since this window read it. Nothing was written." + [Environment]::NewLine + [Environment]::NewLine +
            "  This window loaded  : $(@($script:Loaded).Count) delegate(s), AutomateProcessing $script:Mode" + [Environment]::NewLine +
            "  On the resource now : $(@($freshRaw).Count) delegate(s), AutomateProcessing $freshMode" + [Environment]::NewLine +
            "    $(@($freshRaw) -join [Environment]::NewLine + '    ')" + [Environment]::NewLine + [Environment]::NewLine +
            'Saving would have replaced their work with this window''s older copy -- Set-CalendarProcessing has no add or remove, only replace. Press Reload and make the change again.'
        Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $script:Target 'Resource Delegates') `
            -Detail "Save refused: resource changed underneath (loaded $(@($script:Loaded).Count) delegates/$script:Mode, now $(@($freshRaw).Count)/$freshMode)." -Environment $Environment
        return
    }

    # Compared on the values that were READ and will be WRITTEN, which are the
    # same strings for anything that came off the resource. See Get-WriteValue.
    $loadedLower = @($script:Loaded | ForEach-Object { "$_".ToLowerInvariant() })
    $wantLower   = @($want | ForEach-Object { "$_".ToLowerInvariant() })
    $added   = @($script:Working | Where-Object { $loadedLower -notcontains (Get-WriteValue -Entry $_).ToLowerInvariant() })
    $removed = @($script:Loaded | Where-Object { $wantLower -notcontains "$_".ToLowerInvariant() })

    # THE DANGEROUS CASE GETS ITS OWN DIALOG AND ITS OWN WORDS: an AutoUpdate
    # resource with nobody to approve anything is silently unbookable.
    #
    # TESTED ON THE PENDING MODE, which is the change that made this reachable
    # two ways. It used to be only "you removed the last delegate"; now the
    # dropdown means "you switched a resource with no delegates to AutoUpdate"
    # lands in exactly the same state, and that one is the easier accident --
    # the list is already empty and looks like nothing to do with the switch.
    if (@($want).Count -eq 0 -and $wantMode -eq 'AutoUpdate') {
        $why = 'You are about to remove the last delegate.'
        if ($modeChanged) { $why = "You are about to set it to AutoUpdate while it has no delegates." }

        $msg = 'THIS WILL STOP ANYONE BOOKING THIS RESOURCE.' + [Environment]::NewLine + [Environment]::NewLine +
            "AutoUpdate holds every request as tentative until a delegate approves it. $why" + [Environment]::NewLine + [Environment]::NewLine +
            'Every request will then wait for an approval that nobody can give. The resource will appear permanently unavailable and nothing will report a fault.' + [Environment]::NewLine + [Environment]::NewLine +
            'If the resource should book itself instead, choose AutoAccept. Continue?'
        $r = Show-ToolDialog($msg, 'Leave this resource unbookable',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Stop,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    } else {
        # Named the way the operator recognises them, not by the string being
        # written -- a canonical directory path in a confirmation dialog is not
        # something anybody can check.
        $what = New-Object System.Collections.ArrayList
        foreach ($a in $added) {
            if ($a.Resolved) { [void]$what.Add("  ADD     $($a.Display)  <$($a.Address)>") }
            else             { [void]$what.Add("  ADD     $($a.Raw)") }
        }
        foreach ($d in $removed) {
            # A removed entry is no longer in the working list, so its friendly
            # name comes from the snapshot taken when the resource was read.
            $lbl = "$d"
            foreach ($e in $script:LoadedEntries) {
                if ("$($e.Raw)" -eq "$d" -and $e.Resolved) { $lbl = "$($e.Display)  <$($e.Address)>" }
            }
            [void]$what.Add("  REMOVE  $lbl")
        }

        # THE MODE CHANGE LEADS, because it is the bigger change of the two. A
        # switch that alters how every future request is handled should not be
        # read as a footnote under a list of names.
        if ($modeChanged) {
            [void]$what.Insert(0, '')
            [void]$what.Insert(0, "  MODE    $script:Mode -> $wantMode")
        }

        $msg = "Write the delegates and automate processing on $script:Target." + [Environment]::NewLine + [Environment]::NewLine +
            ($what -join [Environment]::NewLine) + [Environment]::NewLine + [Environment]::NewLine +
            "The list will hold $(@($want).Count) delegate(s) afterwards, with AutomateProcessing $wantMode." + [Environment]::NewLine + [Environment]::NewLine +
            'THE WHOLE LIST IS REPLACED. Set-CalendarProcessing offers no add or remove, so this writes the complete set shown in the window.'

        # What the chosen mode will actually mean, stated at the point of no
        # return rather than only on the form behind the dialog.
        if ($modeChanged -and $wantMode -eq 'AutoAccept') {
            $msg = "$msg$([Environment]::NewLine)$([Environment]::NewLine)AutoAccept MEANS NOBODY VETS BOOKINGS ANY MORE. The resource will accept and decline requests itself, by its booking policy, and the delegates below stop being asked. If this resource was being approved by a person, that stops."
        } elseif ($modeChanged -and $wantMode -eq 'None') {
            $msg = "$msg$([Environment]::NewLine)$([Environment]::NewLine)None MEANS NOTHING IS AUTOMATIC. Requests will sit in the resource mailbox unanswered unless somebody opens it and replies by hand. A resource left like this looks broken to whoever is trying to book it."
        } elseif ($modeChanged -and $wantMode -eq 'AutoUpdate') {
            $msg = "$msg$([Environment]::NewLine)$([Environment]::NewLine)AutoUpdate MEANS EVERY REQUEST NOW WAITS FOR A HUMAN. The $(@($want).Count) delegate(s) above become responsible for approving or declining bookings, and anything they do not answer stays tentative."
        } elseif ($wantMode -ne 'AutoUpdate') {
            $msg = "$msg$([Environment]::NewLine)$([Environment]::NewLine)AutomateProcessing is '$wantMode', not AutoUpdate, so this list is stored and nobody on it will be asked to approve anything until that changes."
        }

        $r = Show-ToolDialog($msg, 'Save delegates and mode',
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($r -ne [System.Windows.Forms.DialogResult]::OK) { return }
    }

    $detail = "Set-CalendarProcessing on $script:Target -- AutomateProcessing $script:Mode -> $wantMode; ResourceDelegates $(@($script:Loaded).Count) -> $(@($want).Count). Added: $(@($added | ForEach-Object { Get-WriteValue -Entry $_ }) -join '; '). Removed: $(@($removed) -join '; ')."

    try {
        # ONE CALL, BOTH SETTINGS, and that is not tidiness. The delegates and
        # the switch only mean anything together: two calls would leave a moment
        # where the resource is on AutoUpdate with the old list, and if the
        # second failed it would stay there -- which is the unbookable state
        # this tool warns about, reached by accident instead of by choice.
        #
        # An empty list must be passed as an empty array rather than omitted:
        # omitting the parameter changes nothing at all, which would read as a
        # successful removal.
        [void](Invoke-EXOCommand -Name 'Set-CalendarProcessing' -Parameters @{
            Identity           = $script:Target
            ResourceDelegates  = $want
            AutomateProcessing = $wantMode
            ErrorAction        = 'Stop'
        })

        # READ BACK, BOTH OF THEM. These are settings people configure once and
        # then rely on for months, and a write that reports success while
        # changing nothing is the failure that costs most to find later.
        $after = Get-ResourceState -Text $script:Target
        $afterRaw  = @()
        $afterMode = ''
        if ($after.Ok) {
            $afterRaw  = @(Get-DelegateRawList -Cp $after.Cp)
            $afterMode = "$($after.Cp.AutomateProcessing)"
        }

        Update-DelegateView

        $countOk = (@($afterRaw).Count -eq @($want).Count)
        $modeOk  = ($afterMode -eq $wantMode)

        if ($after.Ok -and $countOk -and $modeOk) {
            $txtOut.Text = "SAVED. $script:Target now has $(@($afterRaw).Count) delegate(s) and AutomateProcessing $afterMode." + [Environment]::NewLine +
                [Environment]::NewLine + $txtOut.Text
            Write-CompletedLog -Action 'Modify' -Target (Get-LogTarget $script:Target 'Resource Delegates') -Result 'Success' `
                -Detail $detail -Environment $Environment
        } else {
            $bad = New-Object System.Collections.ArrayList
            if (-not $countOk) { [void]$bad.Add("$(@($afterRaw).Count) delegate(s) where $(@($want).Count) were expected") }
            if (-not $modeOk)  { [void]$bad.Add("AutomateProcessing reads '$afterMode' where '$wantMode' was expected") }
            $txtOut.Text = "THE CHANGE DID NOT FULLY TAKE. The write raised no error, but reading the resource back shows $(@($bad) -join ', ')." + [Environment]::NewLine +
                [Environment]::NewLine + $txtOut.Text
            Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $script:Target 'Resource Delegates') `
                -Detail "$detail READ-BACK MISMATCH ($(@($bad) -join '; ')) -- NO ERROR RAISED." -Environment $Environment
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-CalendarProcessing (resource delegates)'
        Write-FailedLog -Action 'Modify' -Target (Get-LogTarget $script:Target 'Resource Delegates') `
            -Detail $detail -ErrorRecord $_ -Environment $Environment
    }
}
#endregion

#region Handlers ----------------------------------------------------------------
function Split-DelegateInput {
    <#
        The box, split into candidate entries.

        NEWLINE, SEMICOLON, COMMA AND TAB, AND DELIBERATELY NOT SPACE. Those
        four cover a column pasted from a spreadsheet, an Outlook recipient
        line, and a CSV cell. Space is excluded because Get-Recipient accepts a
        DISPLAY NAME -- splitting on space would turn "Jane Doe" into two
        lookups that both fail, and the operator would have no idea why.

        DE-DUPLICATED CASE-INSENSITIVELY, IN ORDER. A pasted list with the same
        address twice is a paste, not a request to look somebody up twice.
    #>
    param([string]$Text)

    $seen = @{}
    $out  = New-Object System.Collections.ArrayList
    foreach ($p in @("$Text" -split '[\r\n;,\t]+')) {
        $v = "$p".Trim()
        if (-not $v) { continue }
        $k = $v.ToLowerInvariant()
        if ($seen.ContainsKey($k)) { continue }
        $seen[$k] = $true
        [void]$out.Add($v)
    }
    return @($out.ToArray())
}

function Invoke-AddDelegate {
    <#
        Add everything in the box, and report on every entry individually.

        A BULK ADD THAT SILENTLY DROPS TWO OF EIGHT IS WORSE THAN NO BULK ADD.
        Each entry lands in exactly one of three groups -- added, already
        there, did not resolve -- and all three are printed. The operator can
        then fix the failures rather than discovering at save time that the
        list is short.

        THE FAILURES ARE LEFT IN THE BOX. Everything that worked is removed
        from it, so pressing Add again retries only what did not, without
        re-typing the ones that did. Clearing the whole box on partial success
        is how a name quietly goes missing.
    #>
    if (-not $script:Cp) { $txtOut.Text = 'Look a resource up first.'; return }

    $candidates = @(Split-DelegateInput -Text $txtAdd.Text)
    if (@($candidates).Count -eq 0) { return }

    $addedNames = New-Object System.Collections.ArrayList
    $dupNames   = New-Object System.Collections.ArrayList
    $failed     = New-Object System.Collections.ArrayList
    $leftover   = New-Object System.Collections.ArrayList

    # One directory lookup per entry, so a long paste is a slow one. The cursor
    # says so; a window that simply stops responding does not.
    $oldCursor = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        foreach ($c in $candidates) {
            $res = Resolve-DelegatePrincipal -Text $c
            if (-not $res.Ok) {
                [void]$failed.Add("$c")
                [void]$leftover.Add("$c")
                continue
            }

            # Compared on identity, not on the string that would be written --
            # see Get-DelegateKey. Otherwise the same person already on the
            # resource as a canonical name is not recognised, and both go on
            # the list. Re-read inside the loop because it grows as we go.
            $already = @($script:Working | ForEach-Object { Get-DelegateKey -Entry $_ })
            if ($already -contains $res.Address.ToLowerInvariant()) {
                [void]$dupNames.Add("$($res.Display) <$($res.Address)>")
                continue
            }

            [void]$script:Working.Add([PSCustomObject]@{
                Raw = $res.Address; Address = $res.Address; Display = $res.Display; Kind = $res.Kind; Resolved = $true
            })
            [void]$addedNames.Add("$($res.Display) <$($res.Address)> [$($res.Kind)]")
        }
    } finally {
        $form.Cursor = $oldCursor
    }

    $txtAdd.Text = (@($leftover) -join [Environment]::NewLine)
    Update-DelegateList

    $L = New-Object System.Collections.ArrayList
    [void]$L.Add("$(@($candidates).Count) entr(y/ies) processed: $(@($addedNames).Count) added, $(@($dupNames).Count) already on the list, $(@($failed).Count) not resolved.")
    if (@($addedNames).Count -gt 0) {
        [void]$L.Add('')
        [void]$L.Add('ADDED to the pending list:')
        foreach ($n in $addedNames) { [void]$L.Add("    $n") }
    }
    if (@($dupNames).Count -gt 0) {
        [void]$L.Add('')
        [void]$L.Add('ALREADY THERE, skipped:')
        foreach ($n in $dupNames) { [void]$L.Add("    $n") }
    }
    if (@($failed).Count -gt 0) {
        [void]$L.Add('')
        [void]$L.Add('DID NOT RESOLVE -- left in the box so you can correct them:')
        foreach ($n in $failed) { [void]$L.Add("    $n") }
        [void]$L.Add('A bare name is retried against the tenant''s primary domain. If that is not the right domain for these people, type the full address.')
    }
    if (@($addedNames).Count -gt 0) {
        [void]$L.Add('')
        [void]$L.Add('NOT SAVED YET -- press Save, which rewrites the whole list.')
    }
    $txtOut.Text = ($L -join [Environment]::NewLine)
}

$btnAdd.Add_Click({ try { Invoke-AddDelegate } catch { Show-HandlerError -ErrorRecord $_ -Context 'Add delegates' } })
$txtAdd.Add_KeyDown({
    try {
        # CTRL+ENTER, not Enter. The box is multiline now, so Enter has a job of
        # its own -- taking it would make a list impossible to type.
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter -and $_.Control) {
            # Otherwise Windows plays the ding and the form looks broken.
            $_.SuppressKeyPress = $true
            Invoke-AddDelegate
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Add delegates' }
})

# The dropdown only edits a pending value, so nothing is written here -- but the
# explanation under it has to keep up with the selection.
$cboMode.Add_SelectedIndexChanged({
    try { Update-ModeLabel } catch { Show-HandlerError -ErrorRecord $_ -Context 'Automate processing' }
})

$btnRemove.Add_Click({
    try {
        $idx = @($lstDelegates.SelectedIndices)
        if (@($idx).Count -eq 0) { $txtOut.Text = 'Select one or more delegates to remove.'; return }
        foreach ($i in @($idx | Sort-Object -Descending)) { $script:Working.RemoveAt($i) }
        Update-DelegateList
        $txtOut.Text = "Removed $(@($idx).Count) from the pending list. NOT SAVED YET."
        # The PENDING mode, not the stored one: the dropdown may already be set
        # to AutoUpdate for a resource that is not on it yet, and saving those
        # two changes together is what would leave it unbookable.
        if (@($script:Working).Count -eq 0 -and (Get-PendingMode) -eq 'AutoUpdate') {
            $txtOut.Text = "$($txtOut.Text)$([Environment]::NewLine)$([Environment]::NewLine)THE LIST IS NOW EMPTY AND AutoUpdate IS SELECTED. Saving this would leave nobody able to approve a booking, and the resource would appear permanently unavailable."
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove' }
})
$lstDelegates.Add_KeyDown({
    try {
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Delete) { $btnRemove.PerformClick() }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove' }
})

$btnReload.Add_Click({ try { Update-DelegateView } catch { Show-HandlerError -ErrorRecord $_ -Context 'Reload' } })
$btnSave.Add_Click({ try { Invoke-DelegateSave } catch { Show-HandlerError -ErrorRecord $_ -Context 'Save' } })
#endregion

# AFTER every control and BEFORE the help button: the bar shifts what is already
# on the form down by 40 and grows the window to match, so anything added later
# would be left behind -- and the '?' positions itself from ClientSize, so it
# has to come after the form has grown.
$bar = Add-ToolTargetBar -Form $form -InitialValue $script:Target -Label 'Resource'
$bar.Button.Add_Click({
    try {
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-DelegateView
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Look Up' }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Resource Delegates' | Out-Null

$form.Add_Shown({
    try { Update-DelegateView } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Resource Delegates' }
})

[void]$form.ShowDialog()
