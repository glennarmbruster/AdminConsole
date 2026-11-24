#Requires -Version 5.1
<#
    Retention-Policy.ps1
    Which retention policy a mailbox is on, read from CustomAttribute8.

    THE POLICY LIVES IN A CUSTOM ATTRIBUTE, not in RetentionPolicy. That is a
    local convention rather than an Exchange feature, so nothing validates it
    and nothing keeps it in step with anything else -- which is exactly why a
    tool that reads it consistently is worth having.

    -Environment IS ACCEPTED AND USED, and passed to every Write-*Log call.
#>
# TOOLKIT-HELP-RUNS: Reads CustomAttribute8 from a mailbox. Read-only.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Which retention policy a mailbox is on.
#
# Type any of these and press Check:
#   the alias            jsmith
#   the first part       john.smith
#   the whole address    john.smith@contoso.com
#
# THE ANSWER COMES FROM CustomAttribute8. An empty attribute is not missing
# data -- it means the mailbox is on the standard policy, which is why a blank
# reads as "Default 3 Year Retention Policy" rather than as an error.
#
# THAT DISTINCTION MATTERS BOTH WAYS. A blank attribute and a mailbox that does
# not exist are completely different answers, and this tool never reports one as
# the other: a mailbox it cannot find is reported as not found, never as the
# default policy.
#
# IT IS A LABEL, NOT AN ENFORCED SETTING. CustomAttribute8 is free text that
# something else acts on. This tool reports what is written there; it does not
# prove any policy is actually applied to the mailbox.
#
# Read-only. It reports and changes nothing.
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

$script:DefaultPolicy = 'Default 3 Year Retention Policy'
$script:Rows = New-Object System.Collections.ArrayList

#region Lookup ------------------------------------------------------------------
function Get-RetentionRows {
    <#
        Find the mailbox, then read CustomAttribute8.

        TWO ATTEMPTS, AND THE SECOND IS NOT A RETRY. An exact -Identity lookup
        is fast and unambiguous, and it fails outright on a partial name. Only
        then is a wildcard filter used -- doing it the other way round would
        return three people called Smith when the operator typed an exact alias
        that matches one.

        THE WILDCARD SEARCH CAN RETURN SEVERAL, and all of them are shown rather
        than the first. Picking one silently is how the wrong person's retention
        gets quoted on a legal hold question.
    #>
    param([Parameter(Mandatory)][string]$Identity)

    $r = [PSCustomObject]@{ Ok = $false; Reason = ''; Rows = @(); Matched = 'exact' }
    $found = @()

    try {
        $found = @(Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
            Identity = $Identity; ErrorAction = 'Stop' })
    } catch {
        # Not found by identity is expected for a partial name and is not an
        # error worth showing -- fall through to the wildcard search.
        $found = @()
    }

    if (@($found).Count -eq 0) {
        $r.Matched = 'search'
        $wild = "*$($Identity.Trim('*'))*"
        try {
            $found = @(Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
                Filter        = "Alias -like '$wild' -or DisplayName -like '$wild' -or PrimarySmtpAddress -like '$wild'"
                ResultSize    = 50
                WarningAction = 'SilentlyContinue'
                ErrorAction   = 'Stop'
            })
        } catch {
            $r.Reason = Format-ErrorDetail -ErrorRecord $_ -Context "Get-Mailbox search for '$Identity'"
            return $r
        }
    }

    if (@($found).Count -eq 0) {
        # NOT FOUND IS NOT THE DEFAULT POLICY. Reporting the default here would
        # answer a legal-hold question about a mailbox that does not exist.
        $r.Reason = "No mailbox matches '$Identity'. That is not the same as being on the default policy -- nothing was found to read."
        return $r
    }

    $rows = New-Object System.Collections.ArrayList
    foreach ($m in $found) {
        $raw = "$($m.CustomAttribute8)"
        # TRIMMED BEFORE TESTING. A single space in the attribute is not a
        # policy name, and treating it as one would report a blank-looking value
        # as though somebody had set it deliberately.
        $isBlank = [string]::IsNullOrWhiteSpace($raw)
        [void]$rows.Add([PSCustomObject]@{
            DisplayName = "$($m.DisplayName)"
            Alias       = "$($m.Alias)"
            Address     = "$($m.PrimarySmtpAddress)"
            Type        = "$($m.RecipientTypeDetails)"
            Policy      = $(if ($isBlank) { $script:DefaultPolicy } else { $raw.Trim() })
            Source      = $(if ($isBlank) { 'default (CustomAttribute8 is empty)' } else { 'CustomAttribute8' })
            RawValue    = $raw
        })
    }

    $r.Rows = @($rows.ToArray())
    $r.Ok   = $true
    return $r
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Retention Policy'
$form.ClientSize      = New-Object System.Drawing.Size(940, 600)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$lblWho          = New-Object System.Windows.Forms.Label
$lblWho.Text     = 'Alias, name or address'
$lblWho.Location = New-Object System.Drawing.Point(14, 52)
$lblWho.Size     = New-Object System.Drawing.Size(150, 20)
[void]$form.Controls.Add($lblWho)

$txtWho          = New-Object System.Windows.Forms.TextBox
$txtWho.Location = New-Object System.Drawing.Point(170, 49)
$txtWho.Size     = New-Object System.Drawing.Size(330, 22)
[void]$form.Controls.Add($txtWho)

$btnCheck           = New-Object System.Windows.Forms.Button
$btnCheck.Text      = 'Check'
$btnCheck.Location  = New-Object System.Drawing.Point(512, 47)
$btnCheck.Size      = New-Object System.Drawing.Size(110, 26)
$btnCheck.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnCheck)

$lblHint           = New-Object System.Windows.Forms.Label
$lblHint.Text      = 'A partial name is fine -- every match is listed rather than guessing one.'
$lblHint.Location  = New-Object System.Drawing.Point(170, 74)
$lblHint.Size      = New-Object System.Drawing.Size(600, 18)
$lblHint.ForeColor = [System.Drawing.Color]::Gray
$lblHint.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
[void]$form.Controls.Add($lblHint)

# THE ANSWER, BIG. This tool exists to answer one question and the answer should
# not have to be hunted for in a grid.
$lblAnswer           = New-Object System.Windows.Forms.Label
$lblAnswer.Location  = New-Object System.Drawing.Point(14, 100)
$lblAnswer.Size      = New-Object System.Drawing.Size(910, 30)
$lblAnswer.Font      = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$lblAnswer.Anchor    = 'Top,Left,Right'
[void]$form.Controls.Add($lblAnswer)

$grid                     = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(14, 136)
$grid.Size                = New-Object System.Drawing.Size(910, 200)
$grid.ReadOnly            = $true
$grid.AllowUserToAddRows  = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode       = 'FullRowSelect'
$grid.MultiSelect         = $false
$grid.Anchor              = 'Top,Left,Right'
[void]$form.Controls.Add($grid)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(14, 346)
$txtOut.Size       = New-Object System.Drawing.Size(910, 204)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(804, 560)
$btnClose.Size      = New-Object System.Drawing.Size(120, 28)
$btnClose.FlatStyle = 'Flat'
$btnClose.Anchor    = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
[void]$form.Controls.Add($btnClose)
#endregion

#region Display -----------------------------------------------------------------
function Invoke-Check {
    $who = "$($txtWho.Text)".Trim()
    if (-not $who) { $txtOut.Text = 'Enter an alias, a name or an address.'; return }

    $btnCheck.Enabled = $false
    $lblAnswer.Text = 'Looking ...'
    $lblAnswer.ForeColor = [System.Drawing.Color]::Gray
    $form.Refresh()

    try {
        $res = Get-RetentionRows -Identity $who
        $script:Rows.Clear()
        foreach ($r in @($res.Rows)) { [void]$script:Rows.Add($r) }

        $t = New-Object System.Data.DataTable
        foreach ($c in @('DisplayName', 'Alias', 'Address', 'Type', 'Policy', 'Source')) { [void]$t.Columns.Add($c) }
        foreach ($r in $script:Rows) { [void]$t.Rows.Add($r.DisplayName, $r.Alias, $r.Address, $r.Type, $r.Policy, $r.Source) }
        $grid.DataSource = $t

        # A mailbox on something other than the default is the interesting case,
        # so it is the one that is coloured.
        foreach ($gr in $grid.Rows) {
            if ("$($gr.Cells['Policy'].Value)" -ne $script:DefaultPolicy) {
                $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGoldenrod
            }
            $gr.Tag = 'SemanticColour'
        }

        $L = New-Object System.Collections.ArrayList
        if (-not $res.Ok) {
            $lblAnswer.Text = 'Not found'
            $lblAnswer.ForeColor = [System.Drawing.Color]::Firebrick
            [void]$L.Add($res.Reason)
            [void]$L.Add('')
            [void]$L.Add('Check the spelling, or try a shorter fragment -- a partial name is searched')
            [void]$L.Add('against alias, display name and primary address.')
            $txtOut.Text = ($L -join [Environment]::NewLine)
            Write-CompletedLog -Action 'Report' -Target (Get-LogTarget $who 'Retention Policy') -Result 'Success' `
                -Detail "No mailbox matched '$who'." -Environment $Environment
            return
        }

        if (@($script:Rows).Count -eq 1) {
            $one = $script:Rows[0]
            $lblAnswer.Text = $one.Policy
            $lblAnswer.ForeColor = $(if ($one.Policy -eq $script:DefaultPolicy) { [System.Drawing.Color]::ForestGreen } else { [System.Drawing.Color]::DarkGoldenrod })

            [void]$L.Add("$($one.DisplayName)  <$($one.Address)>")
            [void]$L.Add("Alias    : $($one.Alias)")
            [void]$L.Add("Type     : $($one.Type)")
            [void]$L.Add('')
            [void]$L.Add("POLICY   : $($one.Policy)")
            [void]$L.Add("Read from: $($one.Source)")
            [void]$L.Add('')
            if ($one.Policy -eq $script:DefaultPolicy) {
                [void]$L.Add('CustomAttribute8 is empty on this mailbox, which is what the standard')
                [void]$L.Add('policy looks like. It is not missing data and nothing needs setting.')
            } else {
                [void]$L.Add('This mailbox has an explicit value in CustomAttribute8, so it is NOT on')
                [void]$L.Add('the standard policy. Somebody set that deliberately.')
                [void]$L.Add('')
                [void]$L.Add("Exact stored value: '$($one.RawValue)'")
                [void]$L.Add('')
                [void]$L.Add('IT IS A LABEL, NOT AN ENFORCED SETTING. CustomAttribute8 is free text that')
                [void]$L.Add('something else acts on -- this shows what is written there, not that any')
                [void]$L.Add('policy is actually being applied.')
            }
        } else {
            $lblAnswer.Text = "$(@($script:Rows).Count) mailboxes match -- pick the right one"
            $lblAnswer.ForeColor = [System.Drawing.Color]::DarkGoldenrod
            [void]$L.Add("'$who' matched $(@($script:Rows).Count) mailboxes. Every one is listed above.")
            [void]$L.Add('')
            [void]$L.Add('NO ANSWER IS GIVEN FOR A PARTIAL MATCH. Quoting the first result would be a')
            [void]$L.Add('guess, and retention is the kind of question that gets asked about legal')
            [void]$L.Add('holds. Narrow the search until one mailbox comes back.')
            [void]$L.Add('')
            foreach ($r in $script:Rows) { [void]$L.Add(("   {0,-34} {1,-28} {2}" -f $r.Address, $r.Alias, $r.Policy)) }
        }
        $txtOut.Text = ($L -join [Environment]::NewLine)

        Write-CompletedLog -Action 'Report' -Target (Get-LogTarget $who 'Retention Policy') -Result 'Success' `
            -Detail "Matched $(@($script:Rows).Count) mailbox(es) for '$who' via $($res.Matched)." -Environment $Environment
    } catch {
        $lblAnswer.Text = 'Failed'
        $lblAnswer.ForeColor = [System.Drawing.Color]::Firebrick
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Retention Policy'
        Write-FailedLog -Action 'Report' -Target (Get-LogTarget $who 'Retention Policy') `
            -Detail "$($_.Exception.Message)" -Environment $Environment
    } finally {
        $btnCheck.Enabled = $true
    }
}
#endregion

$btnCheck.Add_Click({ try { Invoke-Check } catch { Show-HandlerError -ErrorRecord $_ -Context 'Check' } })
# Enter in the box runs the check -- this is a one-field tool and reaching for
# the mouse to run it is friction for no reason.
$txtWho.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        try { Invoke-Check } catch { Show-HandlerError -ErrorRecord $_ -Context 'Check' }
    }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Retention Policy' | Out-Null

$form.Add_Shown({
    try {
        if ($SmtpAddress) { $txtWho.Text = "$SmtpAddress" }
        $lblAnswer.Text = ''
        $txtOut.Text = 'Enter an alias, the first part of an address, or a whole address, then press Check.' +
            [Environment]::NewLine + [Environment]::NewLine +
            "Environment: $(if ($Environment) { $Environment } else { '(none passed)' })" +
            [Environment]::NewLine + [Environment]::NewLine +
            'A blank CustomAttribute8 means the standard policy and is reported as' + [Environment]::NewLine +
            "'$script:DefaultPolicy'. A mailbox that cannot be found is reported as not found," + [Environment]::NewLine +
            'never as the default -- those are different answers.'
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Retention Policy' }
})

[void]$form.ShowDialog()
