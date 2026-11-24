#Requires -Version 5.1
<#
.SYNOPSIS
    Layer 1 launcher -- Mailbox / Group lookup + dynamic tool menu.
.DESCRIPTION
    This is the "menu system" half of the launcher/tool-script architecture
    described in REFACTOR-DESIGN.md. It intentionally contains NO business
    logic for any individual tool (no Message Trace, Inbox Rules, Mail
    Attributes, Mailbox Health, Calendar, OOF, or All Attributes code lives
    here). Its job is exactly four things:
      1. Collect the target email address, then auto-detect which
         environment (HQ or Franchise) and Domain Controller it belongs to
         -- see the REVISED 2026-07-23 note below.
      2. Detect the object type (Mailbox / Microsoft 365 Group / on-prem
         Group) via Common.ps1's Resolve-Target -- moved unchanged from the
         monolith's own Resolve-Target function.
      3. Read Tools.json and render one button per entry, ALWAYS -- lookup is
         optional, not a gate. No tool button is hardcoded here -- removing
         an entry from Tools.json removes the button, full stop.
      4. Launch the selected tool script IN THIS SAME POWERSHELL SESSION
         (call operator, not a new process), passing -SmtpAddress /
         -Environment / -DomainController so the tool opens pre-populated
         instead of prompting the operator again. Running in-session is
         deliberate, not incidental: it's what lets a tool see whatever
         Exchange Online connection / ActiveDirectory module are already
         loaded here, exactly like the original monolith always assumed.
         See the comment above Start-ToolScript below for the full
         reasoning and trade-off.

    REVISED 2026-07-11 (design discussion, not a bug report): lookup used to
    be a hard gate -- no buttons existed at all until a target resolved, and
    only the buttons matching its type were built. That meant "I already
    know I want Message Trace" still required a Lookup click first, even
    though every tool script already falls back to its own prompt dialog
    (Resolve-ToolTarget -> Show-TargetPromptDialog in Common.ps1) when
    launched with no -SmtpAddress. So the gate was pure friction, not a
    capability the tool scripts actually needed AdminConsole to provide.
    Now: every button in Tools.json renders at startup, always clickable.
    Once a lookup resolves a known type, buttons whose ObjectTypes doesn't
    include it are greyed out (Enabled = $false) with a tooltip explaining
    why, rather than removed -- clicking one cold, or after a lookup that
    didn't match, still works exactly as before: the tool launches with
    whatever address is known (possibly none) and prompts for its own
    target if it needs one. If a lookup fails outright or doesn't resolve to
    one of the three known types, nothing is greyed out (there's no
    positively-detected type to compare against), so every button stays
    live. See Build-ToolButtons / Update-ToolButtonState below.
    KNOWN GAP, not addressed here: GroupEditor.ps1 is the only tool script
    that validates the object type it was actually handed and shows a clean
    rejection message if it's wrong (see its own header). The other nine
    tool scripts assume Mailbox by convention and don't self-check -- feeding
    one a group's address (by launching it cold and typing a group's address
    into its own prompt, or by clicking a greyed-out button which is still
    technically possible via keyboard/accessibility paths) will likely
    surface a raw Exchange cmdlet error instead of a friendly one. That gap
    already existed any time a tool was run standalone by hand; this change
    doesn't introduce it, just makes it slightly more reachable. Worth a
    follow-up pass giving each tool the same type-check GroupEditor already
    has, tracked separately from this change.

    Deliberate behavior change from the old monolith (documented, not
    accidental): the old single Run click also fetched and displayed the
    full mailbox/group "CURRENT STATE" summary (size, delegate counts, owner/
    member lists) inline, because permission-editing controls lived directly
    on that same form. That summary + editing now lives in
    MailboxPermissions.ps1 / GroupEditor.ps1 instead -- AdminConsole itself
    only tells you WHAT was found (type + address), not its current state,
    because rendering that summary is display/business logic that belongs in
    a tool script per the "no business logic in the launcher" requirement.
    See REFACTOR-DESIGN.md Section 8 for the full trade-off discussion --
    this is easy to revisit if you'd rather have the summary back inline.

    REVISED 2026-07-21 (operator request -- two button sections + fixed DC,
    redone same day after an earlier version of this same change moved the
    fixed-DC constant into Common.ps1 and broke disable-intake.ps1, which
    also dot-sources Common.ps1 and picked up the new shared variable in a
    way that produced a wrong-DC error there. This time the constant stays
    LOCAL to this file -- see $script:DefaultDC below -- so nothing else
    that dot-sources Common.ps1 is affected by it):
      - The single "Tools" FlowLayoutPanel is now two: "Admin Tools" (sized
        for a 5-wide x 4-tall grid, 20 buttons) and "Daily Scripts" (5-wide x
        3-tall, 15 buttons). Both are still FlowLayoutPanels, not
        TableLayoutPanels -- same proven auto-wrap mechanism as before, just
        sized so 5 buttons of the existing 150x32 (Margin 4) size land
        exactly one row wide before wrapping, and AutoScroll stays on as a
        safety net if a section ever holds more than its nominal capacity
        (buttons still all reachable, just not in a clean rectangle).
        Which panel a tool lands in is driven by Tools.json's "Section"
        field ("Admin Tools" or "Daily Scripts") -- see Get-ToolSection /
        Build-ToolButtons below. Any entry with no Section (older Tools.json
        files) or an unrecognized value defaults to "Admin Tools", the same
        "be permissive on read" convention Get-ToolDefinitions already uses
        for a missing Order. ToolManager.ps1 and Tools.json themselves are
        NOT touched by this change -- the operator already has working
        copies of both with Section support.
      - The Domain Controller selector (label + ComboBox) is commented out,
        not deleted -- everyone now connects to the same DC, so there's
        nothing for the operator to choose. $script:DefaultDC below is the
        one line to change to repoint the whole team at a different DC;
        every place that used to read $cboDC.SelectedItem reads
        $script:DefaultDC instead. Uncommenting the control block and
        swapping those reads back is the entire revert if per-operator DC
        selection is ever wanted again -- nothing else about the lookup flow
        changed.
      - Form grew from 780x560 to fit both sections without scrolling the
        whole window; content controls above the tool sections were widened
        to match (740 -> 860) and shifted up to close the gap left by the
        removed DC row.

    REVISED 2026-07-21 1800 (operator request -- layout polish + Clear
    button, same session as the two-section change above):
      - $txtOut (the result/error box) doubled in height, 70 -> 140, so
        long error text (Format-ErrorDetail output, multi-line resolution
        errors) needs less scrolling. Everything below it (both tool-section
        labels and panels) shifted down by the same 70px, gaps between
        controls unchanged from the prior revision.
      - $flowAdminTools / $flowDailyScripts widened 800 -> 860 to match the
        width of every other control on the form (lblWarn, grpEnv, txtInput,
        txtOut all already used 860). The two flow panels were the one place
        still using the old 800 figure, which is what left a visible grey
        strip of unused form background to their right -- widening them
        closes that gap. The "5 columns fit before wrapping" math still
        holds (5*158=790 either way; the extra width is inert margin inside
        the bordered panel, not a 6th column) -- see each panel's own
        comment for the column/row arithmetic.
      - Form height trimmed 800 -> 780 (now that content is taller from the
        $txtOut change, 800 would have left an even bigger dead strip at the
        bottom than before; 780 is sized to the new bottom-most control,
        $flowDailyScripts, plus a small margin -- not a scroll-everything
        shrink, the two tool panels keep their own AutoScroll as a safety
        net if a section ever exceeds its nominal button count).
      - Added a "Clear" button between Look Up and Close (Look Up / Clear /
        Close, left to right) so clearing a lookup no longer requires
        blanking $txtInput by hand and re-clicking Look Up. Clear resets
        $txtInput, $lblResult, and $txtOut, drops $script:LastSmtp /
        $script:LastADServer / $script:LastType back to $null, and calls
        Update-ToolButtonState -DetectedType $null to release any grey-out
        from the last lookup -- i.e. it's the same "blank input" reset path
        $btnRun.Add_Click already had, exposed as its own button instead of
        only reachable by clearing the box and clicking Look Up again.

    REVISED 2026-07-23 (operator request -- third button section, "Connect
    Scripts"): Tools.json's Section field now accepts a third value,
    "Connect Scripts", alongside "Admin Tools" and "Daily Scripts". It renders
    as its own single-row FlowLayoutPanel ($flowConnectScripts), sized for 5
    buttons wide x 1 row tall (same 158px-per-button math as the other two
    panels), placed above Admin Tools since connecting is typically the first
    thing an operator does in a session -- everything below it (lblAdminTools,
    flowAdminTools, lblDailyScripts, flowDailyScripts) shifted down by the
    same 82px this section occupies (20px label + 2px gap + 52px panel + 8px
    gap to the next label), and the form grew 780 -> 862 to fit without
    scrolling the whole window. Get-ToolSection now recognizes "Connect
    Scripts" as its own bucket instead of folding it into the "anything not
    Daily Scripts defaults to Admin Tools" catch-all; Build-ToolButtons and
    Update-ToolButtonState were extended the same way every other section
    already worked (filter AllTools by section, populate the panel, include
    its controls in the grey-out sweep) -- no new pattern, just a third
    instance of the existing one. ToolManager.ps1 and Tools.json were updated
    to know about the new Section value; see ToolManager.ps1's own header for
    that half of the change. No Connect Scripts tool entries are seeded here
    -- the section exists and is ready, but which tools (if any) belong in it
    is left to the operator via ToolManager.ps1.

    REVISED 2026-07-23, later (operator request -- auto-detect HQ vs
    Franchise, drop the manual Environment picker): the operator confirmed
    account names (mailboxes and groups both) are globally unique across the
    HQ and Franchise AD forests -- there is never a "bjsmith in HQ" AND a
    separate "bjsmith in Franchise". That makes "does an AD object with this
    mail address exist in forest X" a completely reliable way to determine
    which environment an account belongs to, so there's no longer any reason
    to make the operator pick Franchise/HQ by hand before every lookup.
      - The Environment radio group ($grpEnv / $radFranchise / $radHQ) is
        commented out, not deleted -- same "leave a revert path" convention
        already used for the Domain Controller selector below. Uncommenting
        it is NOT a full revert by itself though: Find-TargetEnvironment,
        the new $btnRun.Add_Click body, and Start-ToolScript below would all
        need to go back to reading $radHQ.Checked (via the old
        Get-SelectedEnvLabel, also removed) instead of $script:LastEnvironment
        to fully restore manual selection.
      - New function Find-TargetEnvironment does the detection: it runs
        Get-ADObject -Filter "mail -eq '<address>'" against HQ first (using
        Common.ps1's $script:HQDomain -- forest-wide, so any available DC
        answers, and HQ is the smaller of the two forests) and, only if that
        comes back empty, against Franchise (using $script:DefaultDC below,
        the same fixed DC already in use since the 2026-07-21 DC-selector
        removal). Whichever forest answers first sets Environment + ADServer
        for the rest of the lookup and every tool launched afterward -- if
        neither forest has a matching object, $btnRun.Add_Click reports
        "Not found in HQ or Franchise" instead of guessing or defaulting to
        one side.
      - Resolve-Target itself is UNCHANGED -- it still does its own
        EXO-first (Get-Mailbox / Get-UnifiedGroup) then on-prem
        (Get-ADGroup) type detection exactly as before. The only thing that
        changed is WHICH $ADServer gets handed to it: previously whatever
        the operator had selected on the radio, now whatever
        Find-TargetEnvironment actually confirmed the account lives in.
      - $script:LastEnvironment (alongside the existing $script:LastSmtp /
        $script:LastADServer / $script:LastType) replaces the old
        Get-SelectedEnvLabel function -- Start-ToolScript now passes
        $script:LastEnvironment straight through to -Environment instead of
        reading a radio button that no longer exists. $btnClear.Add_Click
        resets it to $null along with the other three, same as before.
      - KNOWN GAP, not addressed here: a mailbox with no on-prem AD
        counterpart at all (a fully cloud-only account, never synced) won't
        be found by Find-TargetEnvironment in either forest and will report
        "Not found in HQ or Franchise" even though it may genuinely exist as
        a mailbox in Exchange Online. This suite is a hybrid on-prem+EXO
        environment where that's expected to be rare/nonexistent, but it's a
        real edge case worth knowing about, not silently handled.
      - Scope, per operator instruction: this change touches ONLY
        AdminConsole.ps1. The separate HQ and Franchise Disable intake
        windows, and the separate New User intake windows, keep their own
        manual Franchise/HQ selection unchanged -- nothing about them was
        touched. Common.ps1 (Get-ADServerString, Resolve-Target,
        Show-TargetPromptDialog, Resolve-ToolTarget) is also completely
        unchanged, so every other tool script that dot-sources it, and every
        tool launched standalone (still prompted via Show-TargetPromptDialog
        with its own manual Environment radios), behaves exactly as before --
        echoing the lesson from the 2026-07-21 DefaultDC-in-Common.ps1
        incident, nothing new was added to shared state that other scripts
        could be affected by.

    REVISED 2026-07-23, layout polish (operator feedback on a screenshot):
      - Every full-width control (lblWarn, txtInput, lblResult, txtOut, and
        the three tool-section labels/panels) narrowed 860 -> 800. The three
        FlowLayoutPanels only ever needed 790px for their 5-column button
        rows (5*158); the extra 70px of panel width was dead space visible
        as an empty strip to the right of every button row -- narrowing to
        800 keeps a small buffer without the wasted strip. Form width
        trimmed 900 -> 840 to match; nothing else about the vertical layout
        changed.
      - $lblWarn.Text simplified to "If not already connected use the
        connect scripts below to connect." -- the previous text described
        the now-removed Environment auto-detect mechanics in detail, which
        was more explanation than the label needed; the Connect Scripts
        section below it (added earlier the same day) is the actual answer
        to "how do I connect."
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10
    Revised : 2026-07-21 1800 (doubled result box, widened tool panels,
              trimmed form height, added Clear button)
    Revised : 2026-07-23 (added "Connect Scripts" third button section --
              see REVISED note above)
    Revised : 2026-07-23, later (auto-detect HQ vs Franchise via AD lookup,
              removed the manual Environment radio group -- see REVISED
              note above)
    Revised : 2026-07-23, layout polish (narrowed panels, simplified warning
              label text -- see REVISED note above)
#>
param(
    [string]$ToolsConfigPath = (Join-Path $PSScriptRoot 'Tools.json')
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Tools.json load -------------------------------------------------------
function Get-ToolDefinitions {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Tools.json not found at '$Path'. Use ToolManager.ps1 to create one, or copy the seed Tools.json shipped with this console."
    }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $tools = @($raw | ConvertFrom-Json -ErrorAction Stop)
    } catch {
        throw "Tools.json at '$Path' could not be parsed as JSON: $($_.Exception.Message)"
    }
    return @($tools | Sort-Object { if ($_.PSObject.Properties['Order']) { [int]$_.Order } else { 999 } })
}
#endregion

#region Main window ------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Admin Console"
$form.Size = New-Object System.Drawing.Size(840, 808)
# REVISED 2026-07-24 (operator request) -- always open at the top-left of the
# screen. Tools launched from here open at the BOTTOM-left, so the console and
# whatever it launched never sit on top of each other. Placement happens just
# before ShowDialog at the bottom of this file, because the form's final Height
# has to be settled before a bottom-anchored position can be calculated.
$form.StartPosition = "Manual"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

# === STATUS ROW -- the very top of the window (REVISED 2026-07-24) ===========
# Connection state and tenant now sit above everything else, because they are
# the first things worth knowing when the window opens. Each service gets its
# OWN label rather than sharing one: a single WinForms Label can only carry one
# ForeColor, so ExO and SPO could not be coloured independently while they were
# one string.
$lblExo = New-Object System.Windows.Forms.Label
$lblExo.Location = New-Object System.Drawing.Point(12, 8)
$lblExo.Size = New-Object System.Drawing.Size(150, 20)
$lblExo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblExo.Text = "ExO: ..."
$form.Controls.Add($lblExo)

$lblSpo = New-Object System.Windows.Forms.Label
$lblSpo.Location = New-Object System.Drawing.Point(170, 8)
$lblSpo.Size = New-Object System.Drawing.Size(150, 20)
$lblSpo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblSpo.Text = "SPO: ..."
$form.Controls.Add($lblSpo)

# Moved up here from beside the Connect Scripts label, so the indicator and the
# button that refreshes it stay together.
$btnConnRefresh = New-Object System.Windows.Forms.Button
$btnConnRefresh.Text = "Refresh"
$btnConnRefresh.Location = New-Object System.Drawing.Point(330, 5)
$btnConnRefresh.Size = New-Object System.Drawing.Size(90, 24)
$btnConnRefresh.Add_Click({ Update-ConnectionButtonColors })
$form.Controls.Add($btnConnRefresh)

# Which tenant Exchange Online is actually connected to. Shows the VANITY
# domain (expresspros.com) rather than the tenant's default accepted domain,
# which here is expresspros.onmicrosoft.com -- see Get-ConnectedTenantDomain's
# REVISED note in Common.ps1 for why the default is the wrong thing to show.
$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Location = New-Object System.Drawing.Point(430, 8)
$lblTenant.Size = New-Object System.Drawing.Size(382, 20)
$lblTenant.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblTenant.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

# === Connection warning, now on its own row below the status ================
$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "If not already connected use the connect scripts below to connect."
$lblWarn.Location = New-Object System.Drawing.Point(12, 30)
$lblWarn.Size = New-Object System.Drawing.Size(800, 18)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

# REVISED 2026-07-24 -- green when known, grey when not.
#
# The earlier version turned RED whenever the connected tenant differed from
# $script:MailDomain, which fired a false alarm: this tenant's DEFAULT accepted
# domain is expresspros.onmicrosoft.com, which is entirely normal and not a
# problem. Get-ConnectedTenantDomain now prefers the vanity domain, and the
# mismatch warning is gone -- it was flagging a non-issue.
function Update-TenantLabel {
    $d = Get-ConnectedTenantDomain
    if ($d) {
        $lblTenant.Text = "Tenant: $d"
        $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
    } else {
        $lblTenant.Text = "Tenant: not connected"
        $lblTenant.ForeColor = [System.Drawing.Color]::DimGray
    }
}

# Environment radio group -- COMMENTED OUT 2026-07-23, later (operator
# request): account names are unique across the HQ and Franchise forests, so
# Find-TargetEnvironment below (an AD lookup, HQ first then Franchise)
# reliably determines environment per-account with no operator input needed.
# Left in place, not deleted, in case manual per-operator selection is ever
# wanted again -- uncommenting this block is only PART of that revert though;
# see the header REVISED 2026-07-23, later note for what else would need to
# change back (Find-TargetEnvironment, $btnRun.Add_Click, Start-ToolScript).
# $grpEnv = New-Object System.Windows.Forms.GroupBox
# $grpEnv.Text = "Environment (governs on-prem AD group lookups only)"
# $grpEnv.Location = New-Object System.Drawing.Point(12, 46)
# $grpEnv.Size = New-Object System.Drawing.Size(860, 50)
# $form.Controls.Add($grpEnv)
# $radFranchise = New-Object System.Windows.Forms.RadioButton
# $radFranchise.Text = "Franchise"
# $radFranchise.Location = New-Object System.Drawing.Point(10, 20)
# $radFranchise.AutoSize = $true
# $radFranchise.Checked = $true
# $grpEnv.Controls.Add($radFranchise)
# $radHQ = New-Object System.Windows.Forms.RadioButton
# $radHQ.Text = "HQ"
# $radHQ.Location = New-Object System.Drawing.Point(120, 20)
# $radHQ.AutoSize = $true
# $grpEnv.Controls.Add($radHQ)

# Domain Controller selector -- COMMENTED OUT 2026-07-21 (operator request):
# everyone connects to the same DC now, so there's nothing left to choose.
# Left in place, not deleted, in case per-operator DC selection is wanted
# again later -- uncomment this block AND swap every $script:DefaultDC read
# below (Find-TargetEnvironment, $btnRun.Add_Click) back to $cboDC.SelectedItem
# to fully revert.
# $lblDC = New-Object System.Windows.Forms.Label
# $lblDC.Text = "Domain Controller (Franchise only; HQ is forest-wide):"
# $lblDC.Location = New-Object System.Drawing.Point(12, 100)
# $lblDC.AutoSize = $true
# $form.Controls.Add($lblDC)
# $cboDC = New-Object System.Windows.Forms.ComboBox
# $cboDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
# $script:FranchiseDCs | ForEach-Object { [void]$cboDC.Items.Add($_) }
# $cboDC.SelectedIndex = 0
# $cboDC.Location = New-Object System.Drawing.Point(320, 97)
# $cboDC.Size = New-Object System.Drawing.Size(400, 20)
# $cboDC.Font = New-Object System.Drawing.Font("Consolas", 9)
# $form.Controls.Add($cboDC)

# Fixed shared DC, replacing the commented-out selector above. This is the
# ONE line to change to repoint the whole team at a different DC.
# DELIBERATELY LOCAL to this file, not Common.ps1 -- an earlier version of
# this change put the equivalent constant in Common.ps1 so every dot-sourced
# tool could share it, but that meant EVERY script that dot-sources
# Common.ps1 (including disable-intake.ps1) picked up a new variable it
# never asked for, and it broke disable-intake.ps1's own DC handling in a
# way that was hard to track down. Keeping this local means only
# AdminConsole.ps1 is affected by it; if SendToDL.ps1 / SendAsDL.ps1 or any
# other standalone tool wants the same fixed DC, give it its own local copy
# of this same line rather than centralizing it again.
$script:DefaultDC = 'fran-vm-dc01.franchises.expresspersonnel.com'

$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Enter mailbox or group email address ($($script:MailDomain) optional):"
$lblInput.Location = New-Object System.Drawing.Point(12, 50)
$lblInput.AutoSize = $true
$form.Controls.Add($lblInput)
$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(12, 70)
$txtInput.Size = New-Object System.Drawing.Size(800, 20)
$txtInput.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtInput)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Look Up"
$btnRun.Location = New-Object System.Drawing.Point(12, 98)
$btnRun.Size = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnRun)
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object System.Drawing.Point(130, 98)
$btnClear.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClear)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(228, 98)
$btnClose.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)
$form.CancelButton = $btnClose

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Not looked up yet."
$lblResult.Location = New-Object System.Drawing.Point(12, 136)
$lblResult.Size = New-Object System.Drawing.Size(800, 20)
$lblResult.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblResult)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.Location = New-Object System.Drawing.Point(12, 160)
$txtOut.Size = New-Object System.Drawing.Size(800, 140)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtOut.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($txtOut)

# REVISED 2026-07-24 -- back to a plain full-width label. The live ExO/SPO
# status it briefly carried, and the Refresh button that sat beside it, both
# moved to the status row at the very top of the window (operator request), so
# this is just a section heading again.
$lblConnectScripts = New-Object System.Windows.Forms.Label
$lblConnectScripts.Text = "Connect Scripts"
$lblConnectScripts.Location = New-Object System.Drawing.Point(12, 308)
$lblConnectScripts.Size = New-Object System.Drawing.Size(800, 20)
$form.Controls.Add($lblConnectScripts)

# Added 2026-07-23 -- see header REVISED note. Same FlowLayoutPanel pattern
# as $flowAdminTools/$flowDailyScripts below, sized for exactly 5 columns
# (158px each, 5*158=790 < 800 panel width) x 1 row (32 + 4*2 = 40px < 52
# panel height) before wrapping/scrolling -- single row only, per operator
# request. AutoScroll stays on as the same safety net the other two panels
# use if a section ever grows past its nominal capacity.
#
# REVISED 2026-07-23, later -- narrowed 860 -> 800 (operator feedback: the
# extra 70px past the 5-column content width was pure dead space inside the
# bordered panel, visible as a wide empty strip to the right of every button
# row). Every other full-width control on the form narrowed the same amount,
# and the form itself shrank 900 -> 840 to match -- see header REVISED note.
$flowConnectScripts = New-Object System.Windows.Forms.FlowLayoutPanel
$flowConnectScripts.Location = New-Object System.Drawing.Point(12, 330)
$flowConnectScripts.Size = New-Object System.Drawing.Size(800, 52)
$flowConnectScripts.AutoScroll = $true
$flowConnectScripts.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($flowConnectScripts)

$lblAdminTools = New-Object System.Windows.Forms.Label
$lblAdminTools.Text = "Admin Tools"
$lblAdminTools.Location = New-Object System.Drawing.Point(12, 390)
$lblAdminTools.Size = New-Object System.Drawing.Size(800, 20)
$form.Controls.Add($lblAdminTools)

# Dynamic button surface -- FlowLayoutPanel auto-wraps buttons to new rows,
# so this doesn't need to know in advance how many tools are configured.
# Sized to fit exactly 5 columns (button 150 + Margin 4*2 = 158px each,
# 5*158=790 < 800 panel width) x 4 rows (32 + 4*2 = 40px each,
# 4*40=160 < 172 panel height) before wrapping/scrolling. AutoScroll stays
# on as a safety net if more than 20 tools ever land in this section.
#
# REVISED 2026-07-23, later -- narrowed 860 -> 800; see $flowConnectScripts'
# own REVISED note above for why (dead space to the right of the buttons).
$flowAdminTools = New-Object System.Windows.Forms.FlowLayoutPanel
$flowAdminTools.Location = New-Object System.Drawing.Point(12, 412)
$flowAdminTools.Size = New-Object System.Drawing.Size(800, 172)
$flowAdminTools.AutoScroll = $true
$flowAdminTools.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($flowAdminTools)

$lblDailyScripts = New-Object System.Windows.Forms.Label
$lblDailyScripts.Text = "Daily Scripts"
$lblDailyScripts.Location = New-Object System.Drawing.Point(12, 592)
$lblDailyScripts.Size = New-Object System.Drawing.Size(800, 20)
$form.Controls.Add($lblDailyScripts)

# Same sizing logic as $flowAdminTools above, for 3 rows instead of 4
# (3*40=120 < 132 panel height). Narrowed 860 -> 800 for the same reason.
$flowDailyScripts = New-Object System.Windows.Forms.FlowLayoutPanel
$flowDailyScripts.Location = New-Object System.Drawing.Point(12, 614)
$flowDailyScripts.Size = New-Object System.Drawing.Size(800, 132)
$flowDailyScripts.AutoScroll = $true
$flowDailyScripts.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($flowDailyScripts)

# One shared ToolTip instance services every tool button -- SetToolTip is
# called per-button in Update-ToolButtonState to explain why a button is
# greyed out (or cleared when it isn't).
$toolTip = New-Object System.Windows.Forms.ToolTip
#endregion

#region State + lookup ----------------------------------------------------------
$script:LastSmtp = $null
$script:LastType = $null
$script:LastADServer = $null
$script:LastEnvironment = $null
$script:AllTools = @()

try {
    $script:AllTools = Get-ToolDefinitions -Path $ToolsConfigPath
} catch {
    $script:AllTools = @()
    $txtOut.Text = $_.Exception.Message
}

# Added 2026-07-23, later -- see header REVISED note. Auto-detects which AD
# forest a mailbox/group's on-prem object lives in: account names are unique
# across HQ and Franchise (operator-confirmed), so "an AD object with this
# mail address exists here" is a reliable, unambiguous stand-in for "this
# account belongs to this environment" -- no separate identity mapping is
# needed. HQ is checked first (the smaller of the two forests, and
# $script:HQDomain from Common.ps1 is forest-wide, so any available DC
# answers -- no fixed-DC risk the way Franchise has); Franchise is only
# checked if HQ comes back empty, using the same $script:DefaultDC already
# in use since the 2026-07-21 DC-selector removal. Returns $null if the
# address isn't found in either forest -- the caller reports that as "not
# found in HQ or Franchise" rather than silently defaulting to one side.
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
    $franServer = Get-ADServerString -Environment 'Franchise' -DomainController $script:DefaultDC
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

# CHANGED 2026-07-10 (field report): this originally used Start-Process
# powershell.exe to launch each tool as a fully separate process, so a hang
# or crash in one tool could never take down AdminConsole or another open
# tool window. In practice this broke every tool that depends on a live
# connection or an already-imported module -- Connect-ExchangeOnline and
# Import-Module ActiveDirectory are per-session state, and a brand-new
# powershell.exe process starts with neither, no matter what's connected in
# the window AdminConsole itself is running in. The monolith never had this
# problem because it WAS the one session the operator connected in (see the
# warning label above: "Connect to Exchange Online before running").
#
# Fix: launch the tool script with the call operator (&), not Start-Process,
# so it runs in AdminConsole's own PowerShell session/runspace and sees
# whatever is already connected/imported there. The call operator (not dot-
# sourcing) is deliberate too: & still gives the tool script its own child
# scope, so its $script:-scoped variables/functions (CurrentRules,
# CurrentTraceResults, MailAttribsCurrentADUser, etc.) don't leak into or
# collide with AdminConsole's own state, or with the next tool launched
# after it -- only the process-level state (loaded modules, live EXO
# connection) is actually shared, which is exactly what was needed.
#
# Trade-off, and it's a real one: a tool window now runs on AdminConsole's
# own UI thread, so AdminConsole can't be clicked again -- and a second tool
# can't be launched -- until the current tool's window is closed, and an
# unhandled crash inside a tool would take AdminConsole down with it too.
# Given the choice between "isolated but broken" and "shared session, one
# tool window at a time," the working session wins. If losing multi-window
# parallelism becomes a real problem in practice, the fix is a per-tool
# Connect-IfNeeded helper in Common.ps1 (each tool reconnects for itself)
# combined with going back to Start-Process -- not attempted here since it's
# a bigger change than this bug report calls for.
function Start-ToolScript {
    param($Tool)
    $scriptPath = Join-Path $PSScriptRoot $Tool.Script
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "'$($Tool.Name)' points to '$($Tool.Script)', which does not exist under this console's folder.`n`nThis is expected if this is a reduced-tier distribution that doesn't include every tool -- otherwise, check Tools.json (ToolManager.ps1) for a stale/typo'd path.",
            'Tool Script Not Found', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    try {
        & $scriptPath -SmtpAddress $script:LastSmtp -Environment $script:LastEnvironment -DomainController $script:LastADServer
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Launching $($Tool.Name)"
    }
}

# Set of object types AdminConsole/Resolve-Target actually knows how to
# detect. Anything outside this set (most notably 'NotFound', but also plain
# $null before any lookup has happened) means "no positively-detected type
# to compare a tool's ObjectTypes against" -- see Update-ToolButtonState.
$script:KnownObjectTypes = @('Mailbox', 'UnifiedGroup', 'OnpremGroup')

# Which panel a tool belongs in -- Tools.json "Section" field. Missing
# (older Tools.json files written before this existed) or unrecognized
# values default to "Admin Tools", matching Get-ToolDefinitions' own
# "be permissive on read" handling of a missing Order.
function Get-ToolSection {
    param($Tool)
    $s = "$($Tool.Section)".Trim()
    if ($s -eq 'Daily Scripts') { return 'Daily Scripts' }
    if ($s -eq 'Connect Scripts') { return 'Connect Scripts' }
    return 'Admin Tools'
}

# Shared by both sections below -- builds one panel's buttons from its own
# filtered tool list. Factored out so Admin Tools and Daily Scripts can't
# drift into two slightly-different button-construction code paths.
function Add-ToolButtonsToPanel {
    param($Panel, $Tools, [string]$EmptyText)
    if (@($Tools).Count -eq 0) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $EmptyText
        $lbl.AutoSize = $true
        $Panel.Controls.Add($lbl)
        return
    }
    foreach ($tool in $Tools) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $tool.Name
        $btn.Size = New-Object System.Drawing.Size(150, 32)
        $btn.Margin = New-Object System.Windows.Forms.Padding(4)
        $btn.Tag = $tool
        # NOTE: no .GetNewClosure() here -- that would run this scriptblock in
        # a new, isolated session state that only inherits captured variables,
        # not script-local functions, so Start-ToolScript itself becomes
        # invisible to it ("not a valid cmdlet" at click time). Not needed
        # anyway: WinForms Add_Click scriptblocks get $this bound to whichever
        # control actually raised the event, so each button already resolves
        # its own .Tag correctly with no closure capture required.
        # ADDED 2026-07-24 -- after a Connect Scripts button runs, re-probe and
        # repaint so the indicator reflects reality without the operator having
        # to click Refresh. Only for that section: re-probing after every
        # Admin/Daily tool would add a blocking SPO call to actions that have
        # nothing to do with connecting.
        #
        # Start-ToolScript invokes the tool with the call operator, which is
        # synchronous, so by the time it returns a connect script has finished
        # and its session is established -- which is exactly why this can be a
        # plain sequential call and not a timer or event.
        $btn.Add_Click({
            Start-ToolScript -Tool $this.Tag
            if ((Get-ToolSection $this.Tag) -eq 'Connect Scripts') {
                Update-ConnectionButtonColors
            }
        })
        $Panel.Controls.Add($btn)
    }
}

# Builds both button surfaces ONCE from $script:AllTools -- every tool gets a
# button, always, regardless of any lookup. This is the ENTIRE "dynamic
# button system" -- no switch/case of hardcoded tool names anywhere. Greying
# out buttons that don't match a resolved type is Update-ToolButtonState's
# job below, run separately (and repeatedly, once per lookup) against these
# same button instances -- the surfaces themselves are never cleared/rebuilt
# again after this runs, so button identity (and any tooltip state) persists.
function Build-ToolButtons {
    $flowConnectScripts.Controls.Clear()
    $flowAdminTools.Controls.Clear()
    $flowDailyScripts.Controls.Clear()
    $connectTools = @($script:AllTools | Where-Object { (Get-ToolSection $_) -eq 'Connect Scripts' })
    $adminTools = @($script:AllTools | Where-Object { (Get-ToolSection $_) -eq 'Admin Tools' })
    $dailyTools = @($script:AllTools | Where-Object { (Get-ToolSection $_) -eq 'Daily Scripts' })
    Add-ToolButtonsToPanel -Panel $flowConnectScripts -Tools $connectTools -EmptyText "No Connect Scripts found in Tools.json. Use ToolManager.ps1 to add some."
    Add-ToolButtonsToPanel -Panel $flowAdminTools -Tools $adminTools -EmptyText "No Admin Tools found in Tools.json. Use ToolManager.ps1 to add some."
    Add-ToolButtonsToPanel -Panel $flowDailyScripts -Tools $dailyTools -EmptyText "No Daily Scripts found in Tools.json. Use ToolManager.ps1 to add some."
}

#region NEW 2026-07-24 -- colour the Connect Scripts buttons by live state ----
# The button you click to fix a connection is now also the light that tells you
# it's broken -- green when that service is connected, red when it isn't.
#
# HOW A BUTTON IS MATCHED TO A SERVICE: by its Name in Tools.json, via the
# patterns below. There is no new Tools.json field, deliberately -- adding one
# would mean touching ToolManager.ps1 and re-saving every existing entry, for
# something a name match handles. A Connect button whose name matches NOTHING
# here is left its normal colour rather than guessed at, so entries like
# "Connect All" or "PIM Up" don't get a misleading red.
#
# To adjust matching, edit the regexes here -- one place, no other changes.
$script:ConnectButtonServices = @(
    [PSCustomObject]@{ Service = 'ExO'; Pattern = 'exo|exchange' }
    [PSCustomObject]@{ Service = 'SPO'; Pattern = 'spo|sharepoint' }
)

# Excel-style light fills with dark text rather than saturated red/green with
# white -- easier to read at button size, and consistent with the light-green
# highlight already used for the ServiceNow Queue button.
$script:ConnClrOk   = [System.Drawing.Color]::FromArgb(198, 239, 206)
$script:ConnClrBad  = [System.Drawing.Color]::FromArgb(255, 199, 206)
$script:ConnClrNone = [System.Drawing.SystemColors]::Control

function Get-ConnectButtonService {
    # Returns 'ExO', 'SPO', or $null for a Connect Scripts tool.
    param($Tool)
    $name = "$($Tool.Name)"
    foreach ($svc in $script:ConnectButtonServices) {
        if ($name -imatch $svc.Pattern) { return $svc.Service }
    }
    return $null
}

function Update-ConnectionButtonColors {
    <#
        Probes each service once, then paints every matching Connect button.

        -SkipSpo exists because Test-SpoConnectionState makes a REAL service
        call (SPO has no local state cmdlet), which blocks this single-threaded
        form for its duration. Startup and post-connect refreshes accept that
        pause deliberately; nothing polls on a timer, which would freeze the
        form at random moments mid-task.
    #>
    param([switch]$SkipSpo)

    $prevCursor = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # Reconnecting to a different tenant mid-session is exactly when the
        # tenant label must not be stale, and this runs on startup, on Refresh,
        # and right after any Connect button -- so drop the cache and re-read.
        Clear-ConnectedTenantDomainCache
        Update-TenantLabel

        $exo = Test-ExoConnectionState
        $spo = if ($SkipSpo) { $null } else { Test-SpoConnectionState }

        foreach ($ctrl in @($flowConnectScripts.Controls)) {
            if (-not $ctrl.Tag) { continue }   # the "no tools found" fallback label
            $svc = Get-ConnectButtonService -Tool $ctrl.Tag
            if (-not $svc) {
                $ctrl.BackColor = $script:ConnClrNone
                continue
            }
            $state = switch ($svc) { 'ExO' { $exo } 'SPO' { $spo } }
            if (-not $state) { continue }      # SPO skipped -- leave as-is

            # BackColor is ignored under the default FlatStyle (System), so
            # switch to Flat for the buttons we actually colour. Left alone on
            # every other panel's buttons so their appearance doesn't change.
            $ctrl.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $ctrl.BackColor = if ($state.Connected) { $script:ConnClrOk } else { $script:ConnClrBad }
            $toolTip.SetToolTip($ctrl, "$svc -- $($state.Detail)")
        }

        # REVISED 2026-07-24 -- each service now has its own label in the top
        # status row, so they can be coloured independently. A single Label
        # carries only one ForeColor, which is why the previous combined
        # "[ ExO: ... | SPO: ... ]" string could not be green-and-red at once.
        $green = [System.Drawing.Color]::FromArgb(0, 128, 0)
        $red   = [System.Drawing.Color]::FromArgb(200, 0, 0)

        if ($exo.Connected) {
            $lblExo.Text = 'ExO: Connected';     $lblExo.ForeColor = $green
        } else {
            $lblExo.Text = 'ExO: Not Connected'; $lblExo.ForeColor = $red
        }
        $toolTip.SetToolTip($lblExo, "$($exo.Detail)")

        if (-not $spo) {
            # Skipped this pass -- grey, because "not checked" is not the same
            # fact as "not connected".
            $lblSpo.Text = 'SPO: not checked'
            $lblSpo.ForeColor = [System.Drawing.Color]::DimGray
            $toolTip.SetToolTip($lblSpo, 'Not checked on this pass -- click Refresh.')
        } elseif ($spo.Connected) {
            $lblSpo.Text = 'SPO: Connected';     $lblSpo.ForeColor = $green
            $toolTip.SetToolTip($lblSpo, "$($spo.Detail)")
        } else {
            $lblSpo.Text = 'SPO: Not Connected'; $lblSpo.ForeColor = $red
            $toolTip.SetToolTip($lblSpo, "$($spo.Detail)")
        }
    } catch {
        # A probe failing must never stop the console from opening.
        $lblExo.Text = 'ExO: unavailable'
        $lblExo.ForeColor = [System.Drawing.Color]::DimGray
        $lblSpo.Text = 'SPO: unavailable'
        $lblSpo.ForeColor = [System.Drawing.Color]::DimGray
    } finally {
        $form.Cursor = $prevCursor
    }
}
#endregion

# Toggles Enabled/tooltip on the already-built buttons in BOTH panels --
# never adds, removes, or reorders a button. $DetectedType outside
# $script:KnownObjectTypes (covers $null before any lookup, and
# 'NotFound'/anything else a failed or no-match lookup produces) means every
# button is re-enabled: there's no positively-known type to declare a
# mismatch against, so nothing is greyed.
function Update-ToolButtonState {
    param([string]$DetectedType)
    $showDimming = ($DetectedType -and ($script:KnownObjectTypes -contains $DetectedType))
    foreach ($ctrl in (@($flowConnectScripts.Controls) + @($flowAdminTools.Controls) + @($flowDailyScripts.Controls))) {
        if (-not $ctrl.Tag) { continue }   # the "no tools found" fallback label, if present
        $tool = $ctrl.Tag
        $matches = @($tool.ObjectTypes) -contains $DetectedType
        if ($showDimming -and -not $matches) {
            $ctrl.Enabled = $false
            $toolTip.SetToolTip($ctrl, "'$($tool.Name)' applies to: $(@($tool.ObjectTypes) -join ', '). The last lookup resolved to '$DetectedType', so this is disabled for it. Look up a matching object to enable it -- or clear the input box and re-run Look Up with nothing entered to release the grey-out on every tool.")
        } else {
            $ctrl.Enabled = $true
            $toolTip.SetToolTip($ctrl, '')
        }
    }
}

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $txtOut.Text = ''
    $lblResult.Text = 'Looking up...'
    try {
        $smtp = Get-TargetSmtp -Text $txtInput.Text
        if (-not $smtp) {
            # Blank input doubles as "clear the last lookup" -- releases any
            # grey-out from a previous lookup instead of just erroring, since
            # there's no separate Clear button for that.
            $script:LastSmtp = $null
            $script:LastADServer = $null
            $script:LastEnvironment = $null
            $script:LastType = $null
            $lblResult.Text = 'Enter an email address first. (Tools grey-out has been cleared.)'
            Update-ToolButtonState -DetectedType $null
            return
        }
        $lblResult.Text = 'Looking up (HQ first, then Franchise)...'
        $envResult = Find-TargetEnvironment -Smtp $smtp
        if (-not $envResult.Environment) {
            $script:LastSmtp = $null
            $script:LastADServer = $null
            $script:LastEnvironment = $null
            $script:LastType = $null
            $lblResult.Text = "Not found in HQ or Franchise -- $smtp"
            if ($envResult.Errors.Count -gt 0) {
                $txtOut.Text = ($envResult.Errors -join "`r`n`r`n")
            }
            Update-ToolButtonState -DetectedType $null
            return
        }
        $target = Resolve-Target -Smtp $smtp -ADServer $envResult.ADServer
        $script:LastSmtp = $smtp
        $script:LastADServer = $envResult.ADServer
        $script:LastEnvironment = $envResult.Environment
        $script:LastType = $target.Type
        switch ($target.Type) {
            'Mailbox'      { $lblResult.Text = "Found: Mailbox ($($envResult.Environment)) -- $smtp" }
            'UnifiedGroup' { $lblResult.Text = "Found: Microsoft 365 Group ($($envResult.Environment)) -- $smtp" }
            'OnpremGroup'  { $lblResult.Text = "Found: On-Prem Group ($($envResult.Environment)) -- $smtp" }
            default {
                $lblResult.Text = "AD object found in $($envResult.Environment) but not confirmed as a Mailbox, Microsoft 365 Group, or on-prem AD group -- $smtp"
                if ($target.ResolutionErrors.Count -gt 0) {
                    $txtOut.Text = ($target.ResolutionErrors -join "`r`n`r`n")
                }
            }
        }
        Update-ToolButtonState -DetectedType $target.Type
    } catch {
        $lblResult.Text = 'Lookup failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Admin Console lookup'
        # Nothing positively resolved -- release any grey-out from a prior
        # lookup rather than leaving it stale against a failed attempt.
        Update-ToolButtonState -DetectedType $null
    } finally {
        $btnRun.Enabled = $true
    }
})
$btnClear.Add_Click({
    # Same reset $btnRun.Add_Click already performs when handed blank input
    # (see the "blank input doubles as clear" branch above) -- exposed as
    # its own button so clearing doesn't require blanking the box by hand
    # and clicking Look Up again.
    $txtInput.Text = ''
    $lblResult.Text = 'Not looked up yet.'
    $txtOut.Text = ''
    $script:LastSmtp = $null
    $script:LastADServer = $null
    $script:LastEnvironment = $null
    $script:LastType = $null
    Update-ToolButtonState -DetectedType $null
})
$btnClose.Add_Click({ $form.Close() })

# Build the full button surface up front -- no lookup has happened yet, so
# every tool starts enabled (see the header REVISED 2026-07-11 note above).
Build-ToolButtons
Update-ToolButtonState -DetectedType $null

# ADDED 2026-07-24 -- paint the Connect buttons once at startup so their colour
# is meaningful the moment the console appears. This includes the blocking SPO
# probe, which is acceptable here (the form is still being built, so there is no
# work to interrupt) and is the same trade-off the web host makes on page load.
Update-ConnectionButtonColors

# Top-left placement, done last so the form's final size is known. See the
# $form.StartPosition note at the top of this region.
Set-FormScreenCorner -Form $form -Corner 'TopLeft'

[void]$form.ShowDialog()
#endregion