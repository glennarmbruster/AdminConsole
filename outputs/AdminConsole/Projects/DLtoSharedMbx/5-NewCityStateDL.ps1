#Requires -Version 5.1
# TOOLKIT-VARIANT: OnPrem
# TOOLKIT-REQUIRES: OwnProcess  -- New-DistributionGroup exists in BOTH Exchanges
#   and creates a completely different object in each. On-premises it lands in
#   the AD OU named below; in Exchange Online there is no OU and the group is
#   cloud-only. Sharing a process with a connected ExO session is how a
#   franchise DL gets created in the wrong directory with no error.
# ---------------------------------------------------------------------------
<#
    5-NewCityStateDL.ps1
    Create the Staffing / Sales distribution groups for a franchise office.

    Companion to 6-NewCityState.ps1, which creates CLOUD shared mailboxes with
    no OU. This one creates ON-PREMISES distribution groups in a named AD OU,
    which is a different object in a different directory -- hence a separate
    script rather than a switch on that one.

    Author   : Glenn Armbruster
    Revised  : 2026-08-10
#>
# TOOLKIT-HELP-BEGIN
# Creates the two distribution groups a new franchise office needs.
#
# Give it the OU in canonical form -- the same shape ADUC shows:
#   franchises.expresspersonnel.com/Franchises/2765-F67/F67_Users
#
# Type the city and state with no space (HuntingtonWV) and tick the groups you
# want. It creates:
#   Staffing.HuntingtonWV
#   Sales.HuntingtonWV
#
# Names, aliases and addresses are all derived from that one box, so the three
# cannot disagree with each other.
#
# ON-PREMISES ONLY. It runs in its own window because New-DistributionGroup
# exists in both Exchanges and creates a different object in each.
#
# Preview first. Preview reads only -- it checks the OU exists and whether the
# groups are already there, and writes nothing.
# TOOLKIT-HELP-END

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Common.ps1 lives at the TOOLKIT ROOT. Walk up rather than counting folders --
# this script sits three levels down and the next project may not.
$toolkitRoot = $PSScriptRoot
while ($toolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $toolkitRoot 'Common.ps1'))) {
    $toolkitRoot = Split-Path -Parent $toolkitRoot
}
if (-not $toolkitRoot) { throw 'Could not locate Common.ps1 above this script.' }
. (Join-Path $toolkitRoot 'Common.ps1')

#region Naming -----------------------------------------------------------------

$script:Prefixes = @('Staffing', 'Sales')

function Get-DLName {
    <#
        One place that builds a name, so Name, Alias and the address cannot
        drift apart.

        The dot is part of the convention -- 'Staffing.HuntingtonWV' -- and the
        alias keeps it. Alias tolerates a dot; if that ever changes, this is the
        single line to alter rather than three call sites.
    #>
    param([string]$Prefix, [string]$CityState)
    return "$Prefix.$(("$CityState").Trim())"
}

function Test-CityState {
    <#
        Rejects what the convention cannot carry, before anything is created.

        A space or a hyphen here produces a group whose alias and address do
        not match its name, and the mismatch is not visible until someone tries
        to mail it. 'Huntington WV' is the obvious typo and the one worth
        catching.
    #>
    param([string]$Value)
    $v = "$Value".Trim()
    if (-not $v)                    { return 'City and state is required.' }
    if ($v -match '\s')             { return 'No spaces -- use HuntingtonWV, not Huntington WV.' }
    if ($v -notmatch '^[A-Za-z0-9]+$') { return 'Letters and numbers only -- no dots, hyphens or punctuation.' }
    if ($v.Length -lt 3)            { return 'That looks too short to be a city and state.' }
    return ''
}

#endregion

#region Directory --------------------------------------------------------------

function Test-TargetOu {
    <#
        Confirms the OU before anything is created.

        New-DistributionGroup with a bad OU fails per group, so without this the
        operator gets the same error twice and has to work out that neither one
        was the group's fault. Asking once, up front, names the actual problem.
    #>
    param([string]$Ou)
    $out = [PSCustomObject]@{ Ok = $false; Name = ''; Reason = '' }
    $o = "$Ou".Trim()
    if (-not $o) { $out.Reason = 'An OU is required.'; return $out }
    try {
        $found = @(Invoke-OnPremCommand -Name 'Get-OrganizationalUnit' -Parameters @{
            Identity = $o; ErrorAction = 'Stop'
        })
        if ($found.Count -eq 0) { $out.Reason = 'not found'; return $out }
        $out.Ok = $true
        $out.Name = "$($found[0].Name)"
    } catch {
        $out.Reason = $_.Exception.Message
    }
    return $out
}

function Get-ExistingDL {
    param([string]$Name)
    try {
        $g = Invoke-OnPremCommand -Name 'Get-DistributionGroup' -Parameters @{
            Identity = $Name; ErrorAction = 'Stop'
        }
        return $g
    } catch { return $null }
}

#endregion

#region UI ---------------------------------------------------------------------

function Show-NewCityStateDLForm {

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'New Franchise Distribution Groups -- on-premises'
    $form.Size = New-Object System.Drawing.Size(860, 620)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblOu = New-Object System.Windows.Forms.Label
    $lblOu.Text = 'Target OU (canonical, as ADUC shows it):'
    $lblOu.Location = New-Object System.Drawing.Point(12, $y)
    $lblOu.Size = New-Object System.Drawing.Size(500, 18)
    $form.Controls.Add($lblOu); $y += 20

    $txtOu = New-Object System.Windows.Forms.TextBox
    $txtOu.Location = New-Object System.Drawing.Point(12, $y)
    $txtOu.Size = New-Object System.Drawing.Size(820, 24)
    $txtOu.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOu); $y += 28

    $lblOuEg = New-Object System.Windows.Forms.Label
    $lblOuEg.Text = 'e.g.  franchises.expresspersonnel.com/Franchises/2765-F67/F67_Users'
    $lblOuEg.Location = New-Object System.Drawing.Point(12, $y)
    $lblOuEg.Size = New-Object System.Drawing.Size(820, 18)
    $lblOuEg.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblOuEg); $y += 30

    $lblCs = New-Object System.Windows.Forms.Label
    $lblCs.Text = 'City and state, no space (HuntingtonWV):'
    $lblCs.Location = New-Object System.Drawing.Point(12, $y)
    $lblCs.Size = New-Object System.Drawing.Size(500, 18)
    $form.Controls.Add($lblCs); $y += 20

    $txtCs = New-Object System.Windows.Forms.TextBox
    $txtCs.Location = New-Object System.Drawing.Point(12, $y)
    $txtCs.Size = New-Object System.Drawing.Size(300, 24)
    $form.Controls.Add($txtCs); $y += 32

    $chkStaffing = New-Object System.Windows.Forms.CheckBox
    $chkStaffing.Text = 'Staffing'
    $chkStaffing.Location = New-Object System.Drawing.Point(12, $y)
    $chkStaffing.Size = New-Object System.Drawing.Size(110, 22)
    $chkStaffing.Checked = $true
    $form.Controls.Add($chkStaffing)

    $chkSales = New-Object System.Windows.Forms.CheckBox
    $chkSales.Text = 'Sales'
    $chkSales.Location = New-Object System.Drawing.Point(130, $y)
    $chkSales.Size = New-Object System.Drawing.Size(110, 22)
    $chkSales.Checked = $true
    $form.Controls.Add($chkSales)

    $lblPreviewName = New-Object System.Windows.Forms.Label
    $lblPreviewName.Location = New-Object System.Drawing.Point(260, ($y + 2))
    $lblPreviewName.Size = New-Object System.Drawing.Size(570, 18)
    $lblPreviewName.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblPreviewName)
    $y += 34

    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = 'Preview'
    $btnPreview.Location = New-Object System.Drawing.Point(12, $y)
    $btnPreview.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnPreview)

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = 'Create'
    $btnCreate.Location = New-Object System.Drawing.Point(150, $y)
    $btnCreate.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnCreate)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Location = New-Object System.Drawing.Point(702, $y)
    $btnClose.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })
    $y += 40

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(12, $y)
    $txtOut.Size = New-Object System.Drawing.Size(820, 330)
    $txtOut.Multiline = $true
    $txtOut.ScrollBars = 'Vertical'
    $txtOut.ReadOnly = $true
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)

    # Live name preview, so the operator sees what will be created before
    # pressing anything. A name assembled at write time is a name nobody read.
    $updateNames = {
        $cs = "$($txtCs.Text)".Trim()
        if (-not $cs) { $lblPreviewName.Text = ''; return }
        $names = @()
        if ($chkStaffing.Checked) { $names += Get-DLName -Prefix 'Staffing' -CityState $cs }
        if ($chkSales.Checked)    { $names += Get-DLName -Prefix 'Sales'    -CityState $cs }
        $lblPreviewName.Text = ($names -join '    ')
    }
    $txtCs.Add_TextChanged($updateNames)
    $chkStaffing.Add_CheckedChanged($updateNames)
    $chkSales.Add_CheckedChanged($updateNames)

    $gather = {
        $err = Test-CityState -Value $txtCs.Text
        if ($err) { return @{ Ok = $false; Msg = $err } }

        $picked = @()
        if ($chkStaffing.Checked) { $picked += 'Staffing' }
        if ($chkSales.Checked)    { $picked += 'Sales' }
        if ($picked.Count -eq 0)  { return @{ Ok = $false; Msg = 'Tick Staffing, Sales, or both.' } }

        $dom = Get-MailDomain
        if (-not $dom) {
            return @{ Ok = $false; Msg = 'No primary domain in Tenant.json -- addresses cannot be built.' }
        }

        $cs = "$($txtCs.Text)".Trim()
        $plan = @()
        foreach ($p in $picked) {
            $n = Get-DLName -Prefix $p -CityState $cs
            $plan += [PSCustomObject]@{ Name = $n; Alias = $n; Smtp = "$n@$dom" }
        }
        return @{ Ok = $true; Ou = "$($txtOu.Text)".Trim(); Plan = $plan }
    }

    $doPreview = {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $g = & $gather
            if (-not $g.Ok) { $txtOut.Text = $g.Msg; return }

            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('PREVIEW -- nothing has been created')
            [void]$sb.AppendLine('')

            $ou = Test-TargetOu -Ou $g.Ou
            if (-not $ou.Ok) {
                [void]$sb.AppendLine("  OU: $($g.Ou)")
                [void]$sb.AppendLine("      NOT USABLE -- $($ou.Reason)")
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine('  Canonical form, the same shape ADUC shows:')
                [void]$sb.AppendLine('    franchises.expresspersonnel.com/Franchises/2765-F67/F67_Users')
                $txtOut.Text = $sb.ToString(); return
            }
            [void]$sb.AppendLine("  OU: $($g.Ou)   [$($ou.Name)]")
            [void]$sb.AppendLine('')

            foreach ($p in $g.Plan) {
                $ex = Get-ExistingDL -Name $p.Name
                if ($ex) {
                    [void]$sb.AppendLine("  EXISTS   $($p.Name)   <$($ex.PrimarySmtpAddress)>  -- will be skipped")
                } else {
                    [void]$sb.AppendLine("  create   $($p.Name)   <$($p.Smtp)>")
                }
            }
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('  All three of Name, Alias and address come from the city/state box,')
            [void]$sb.AppendLine('  so they cannot disagree.')
            $txtOut.Text = $sb.ToString()
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Previewing'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $doCreate = {
        $g = & $gather
        if (-not $g.Ok) { $txtOut.Text = $g.Msg; return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $ou = Test-TargetOu -Ou $g.Ou
            if (-not $ou.Ok) {
                $txtOut.Text = "OU '$($g.Ou)' is not usable -- $($ou.Reason). Nothing created."
                return
            }

            $names = ($g.Plan | ForEach-Object { $_.Name }) -join "`r`n  "
            $a = [System.Windows.Forms.MessageBox]::Show(
                ("Create these distribution groups ON-PREMISES?" + "`r`n`r`n  " + $names +
                 "`r`n`r`nIn OU:" + "`r`n  $($g.Ou)"),
                'Confirm', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('RESULT')
            [void]$sb.AppendLine('')
            foreach ($p in $g.Plan) {
                if (Get-ExistingDL -Name $p.Name) {
                    [void]$sb.AppendLine("  SKIPPED  $($p.Name) -- already exists")
                    continue
                }
                try {
                    # One call per group, not a batch. A half-applied pair is
                    # easier to finish than to diagnose, and the log then names
                    # exactly which one landed.
                    Invoke-OnPremCommand -Name 'New-DistributionGroup' -Parameters @{
                        Name               = $p.Name
                        Alias              = $p.Alias
                        PrimarySmtpAddress = $p.Smtp
                        OrganizationalUnit = $g.Ou
                        ErrorAction        = 'Stop'
                    } | Out-Null

                    Write-CompletedLog -Action 'Create' -Target $p.Smtp -Result 'Success' `
                        -Detail "Created on-premises distribution group '$($p.Name)' in $($g.Ou)"
                    [void]$sb.AppendLine("  CREATED  $($p.Name)   <$($p.Smtp)>")
                } catch {
                    Write-FailedLog -Action 'Create' -Target $p.Smtp -ErrorRecord $_ `
                        -Detail "Create distribution group '$($p.Name)' in $($g.Ou)"
                    [void]$sb.AppendLine("  FAILED   $($p.Name) -- $($_.Exception.Message)")
                }
            }
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('  Created on-premises. Allow a sync cycle before they appear in the cloud.')
            $txtOut.Text = $sb.ToString()
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Creating the groups'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnPreview.Add_Click($doPreview)
    $btnCreate.Add_Click($doCreate)

    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'New Franchise DLs' | Out-Null
    Set-ConsoleTheme -Form $form
    $form.Add_Shown({ $txtOu.Focus() })
    [void]$form.ShowDialog()
}

#endregion

Show-NewCityStateDLForm
