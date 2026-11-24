<#
.SYNOPSIS
    WinForms tool to grant or remove OneDrive site collection admin access for custodians.

.DESCRIPTION
    Manages SharePoint Online site collection admin rights across all combinations of one or
    more disabled users and one or more custodians (many-to-many). Every custodian in the
    Custodians box receives (or loses) IsSiteCollectionAdmin on every OneDrive site derived
    from the Users box.

    Lookup key: email address (documented exception to the sAMAccountName default).
    Target OneDrive sites are identified by the disabled user's email address, which is used
    to construct the personal-site URL. Local-part alone is accepted; the script silently
    appends @expresspros.com before constructing the URL.

    OneDrive URL format:
        https://expresspros-my.sharepoint.com/personal/<first_last>_expresspros_com

    Actions (mutually exclusive radio buttons):
      Grant Site Collection Admin — calls Set-SPOUser -IsSiteCollectionAdmin $true for each
                                    custodian on each user's OneDrive.
      Remove Site Collection Admin — calls Set-SPOUser -IsSiteCollectionAdmin $false for each
                                     custodian on each user's OneDrive.

    What this script does NOT change:
      - No other site permissions, sharing settings, or document library permissions are modified.
      - No Exchange, Active Directory, or Azure AD / Entra ID attributes are touched.
      - No OneDrive files or folders are moved, deleted, or shared.
      - IsSiteCollectionAdmin on other existing admins is never cleared.

    Pre-flight: users and custodians are validated for non-empty input before any change is made.
    The OneDrive URL is constructed from the email prefix — if the URL is incorrect, Set-SPOUser
    will return an error which is isolated per site and surfaced in the Result box.

    Cmdlet provenance: Set-SPOUser is taken directly from the working script (treated as verified
    per Section 2). Get-SPOUser is used for post-action verification (standard SPO module cmdlet).

    Environment: SharePoint Online only.
    Connection is assumed already established — no Connect-SPOService inside this script.

.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-07-01
    Revised         : 2026-07-01 0000
    Run from        : Standard PowerShell with Microsoft.Online.SharePoint.PowerShell module loaded
    Required modules: Microsoft.Online.SharePoint.PowerShell
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Config  ---------------------------------------------------------------

$Domain              = "@expresspros.com"
$MySiteBase          = "https://expresspros-my.sharepoint.com/personal"
$CompletedTaskFolder = "E:\Scripts\M365\Logs\Completed"
$CompletedTaskFile   = Join-Path $CompletedTaskFolder "CompletedTasks.csv"

#endregion

#region Helpers  --------------------------------------------------------------

function Get-AddressList {
    param([string]$Raw)

    if ([string]::IsNullOrWhiteSpace($Raw)) { return @() }

    $tokens = $Raw -split "[\r\n,;]+"

    $clean = foreach ($t in $tokens) {
        $name = $t.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($name -notmatch '@') { $name = $name + $Domain }
        $name
    }

    return @($clean | Sort-Object -Unique)
}

function Get-OneDriveUrl {
    param([string]$Email)
    $localPart = ($Email -split '@')[0]
    $urlSafe   = $localPart -replace '\.', '_'
    return "$MySiteBase/${urlSafe}_expresspros_com"
}

function Write-CompletedLog {
    param([string]$Task)
    if (-not (Test-Path -LiteralPath $CompletedTaskFolder)) {
        New-Item -Path $CompletedTaskFolder -ItemType Directory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $CompletedTaskFile)) {
        "Task,DateTime" | Out-File -FilePath $CompletedTaskFile -Encoding utf8
    }
    $dt = Get-Date -Format "yyyy-MM-dd HHmm"
    if ($Task -and $Task.Trim() -ne "") {
        Add-Content -Path $CompletedTaskFile -Value ('"{0}","{1}"' -f $Task.Trim(), $dt)
    }
}

#endregion

#region Parsing  --------------------------------------------------------------

function Get-ParsedInputs {
    param(
        [string]$UserRaw,
        [string]$CustodianRaw,
        [string]$Action
    )

    $users      = Get-AddressList -Raw $UserRaw
    $custodians = Get-AddressList -Raw $CustodianRaw

    return [PSCustomObject]@{
        Users      = $users
        Custodians = $custodians
        Action     = $Action
        OpCount    = ($users.Count * $custodians.Count)
    }
}

#endregion

#region Actions  --------------------------------------------------------------

function Invoke-OneDriveAccess {
    param(
        [string[]]$Users,
        [string[]]$Custodians,
        [string]$Action   # "Grant" | "Remove"
    )

    $log = New-Object System.Text.StringBuilder

    $isAdmin    = ($Action -eq "Grant")
    $actionText = if ($isAdmin) { "Grant Site Collection Admin" } else { "Remove Site Collection Admin" }
    $verb       = if ($isAdmin) { "Granted" } else { "Removed" }

    [void]$log.AppendLine("=== OneDrive Access Manager ===")
    [void]$log.AppendLine("Action       : $actionText")
    [void]$log.AppendLine("User(s)      : $($Users -join ', ')")
    [void]$log.AppendLine("Custodian(s) : $($Custodians -join ', ')")
    [void]$log.AppendLine("")

    $succeeded = [System.Collections.Generic.List[PSCustomObject]]::new()
    $failed    = [System.Collections.Generic.List[string]]::new()

    foreach ($u in $Users) {
        $siteUrl = Get-OneDriveUrl -Email $u

        [void]$log.AppendLine("--- $u ---")
        [void]$log.AppendLine("    Site : $siteUrl")
        [void]$log.AppendLine("")

        foreach ($c in $Custodians) {
            try {
                # Set-SPOUser (SharePoint Online) — sets or clears IsSiteCollectionAdmin
                Set-SPOUser -Site $siteUrl -LoginName $c -IsSiteCollectionAdmin $isAdmin -ErrorAction Stop | Out-Null
                [void]$log.AppendLine("  [OK]   $verb site collection admin : $c -> $u")
                $succeeded.Add([PSCustomObject]@{ User = $u; Custodian = $c })
            } catch {
                [void]$log.AppendLine("  [FAIL] $c -> $u : $($_.Exception.Message)")
                $failed.Add("$c -> $u")
            }
        }

        # Verification: read back current site collection admins for this site
        [void]$log.AppendLine("")
        [void]$log.AppendLine("  Verify — site collection admins on $u :")
        try {
            $admins = Get-SPOUser -Site $siteUrl -ErrorAction Stop |
                      Where-Object { $_.IsSiteCollectionAdmin -eq $true }
            if ($admins) {
                foreach ($a in $admins) {
                    [void]$log.AppendLine("    $($a.LoginName)")
                }
            } else {
                [void]$log.AppendLine("    (none)")
            }
        } catch {
            [void]$log.AppendLine("    (error reading admins: $($_.Exception.Message))")
        }

        [void]$log.AppendLine("")
    }

    [void]$log.AppendLine(("Summary : {0} succeeded, {1} failed." -f $succeeded.Count, $failed.Count))

    # Logging: one row per user (target object), per Section 21
    # Grammar: <Verb> OneDriveAdmin <user> (custodians: c1+c2)
    if ($succeeded.Count -gt 0) {
        $custList = $Custodians -join "+"
        foreach ($u in $Users) {
            $uSuccess = @($succeeded | Where-Object { $_.User -eq $u })
            if ($uSuccess.Count -gt 0) {
                Write-CompletedLog -Task "$verb OneDriveAdmin $u (custodians: $custList)"
            }
        }
    }

    return $log.ToString()
}

#endregion

#region WinForms UI  ----------------------------------------------------------

# Shared state between Parse and Run
$script:Parsed     = $null
$script:ParsedFrom = $null

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text            = "OneDrive Access Manager"
$form.Size            = New-Object System.Drawing.Size(760, 740)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

# --- Warning label (always first — red bold, Section 4) ---
$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Text      = "Make sure you are connected to SharePoint Online before running."
$lblWarn.Location  = New-Object System.Drawing.Point(12, 10)
$lblWarn.AutoSize  = $true
$lblWarn.Font      = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

# --- Disabled User(s) label + box ---
$lblUser          = New-Object System.Windows.Forms.Label
$lblUser.Text     = "Disabled User(s) — one per line (enter email address, @expresspros.com optional):"
$lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(12, 38)
$form.Controls.Add($lblUser)

$txtUser               = New-Object System.Windows.Forms.TextBox
$txtUser.Multiline     = $true
$txtUser.ScrollBars    = "Vertical"
$txtUser.Size          = New-Object System.Drawing.Size(722, 100)
$txtUser.Location      = New-Object System.Drawing.Point(12, 58)
$txtUser.Font          = New-Object System.Drawing.Font("Consolas", 9)
$txtUser.AcceptsReturn = $true
$form.Controls.Add($txtUser)

# --- Custodian(s) label + box ---
$lblCustodian          = New-Object System.Windows.Forms.Label
$lblCustodian.Text     = "Custodian(s) to grant/remove access — one per line (enter email address, @expresspros.com optional):"
$lblCustodian.AutoSize = $true
$lblCustodian.Location = New-Object System.Drawing.Point(12, 170)
$form.Controls.Add($lblCustodian)

$txtCustodian               = New-Object System.Windows.Forms.TextBox
$txtCustodian.Multiline     = $true
$txtCustodian.ScrollBars    = "Vertical"
$txtCustodian.Size          = New-Object System.Drawing.Size(722, 100)
$txtCustodian.Location      = New-Object System.Drawing.Point(12, 190)
$txtCustodian.Font          = New-Object System.Drawing.Font("Consolas", 9)
$txtCustodian.AcceptsReturn = $true
$form.Controls.Add($txtCustodian)

# --- Parsed preview label + box ---
$lblParsed          = New-Object System.Windows.Forms.Label
$lblParsed.Text     = "Parsed preview — review before clicking Run:"
$lblParsed.AutoSize = $true
$lblParsed.Location = New-Object System.Drawing.Point(12, 302)
$form.Controls.Add($lblParsed)

$txtParsed            = New-Object System.Windows.Forms.TextBox
$txtParsed.Multiline  = $true
$txtParsed.ScrollBars = "Vertical"
$txtParsed.ReadOnly   = $true
$txtParsed.BackColor  = [System.Drawing.Color]::WhiteSmoke
$txtParsed.Size       = New-Object System.Drawing.Size(722, 75)
$txtParsed.Location   = New-Object System.Drawing.Point(12, 322)
$txtParsed.Font       = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtParsed)

# --- Action group ---
$grpAction          = New-Object System.Windows.Forms.GroupBox
$grpAction.Text     = "Action"
$grpAction.Size     = New-Object System.Drawing.Size(722, 80)
$grpAction.Location = New-Object System.Drawing.Point(12, 406)
$form.Controls.Add($grpAction)

$rbGrant          = New-Object System.Windows.Forms.RadioButton
$rbGrant.Text     = "Grant Site Collection Admin"
$rbGrant.AutoSize = $true
$rbGrant.Location = New-Object System.Drawing.Point(15, 25)
$rbGrant.Checked  = $true
$grpAction.Controls.Add($rbGrant)

$rbRemove          = New-Object System.Windows.Forms.RadioButton
$rbRemove.Text     = "Remove Site Collection Admin"
$rbRemove.AutoSize = $true
$rbRemove.Location = New-Object System.Drawing.Point(15, 52)
$grpAction.Controls.Add($rbRemove)

# --- Buttons: Parse / Run / Close ---
$btnParse          = New-Object System.Windows.Forms.Button
$btnParse.Text     = "Parse"
$btnParse.Size     = New-Object System.Drawing.Size(80, 28)
$btnParse.Location = New-Object System.Drawing.Point(490, 500)
$form.Controls.Add($btnParse)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = "Run"
$btnRun.Size     = New-Object System.Drawing.Size(80, 28)
$btnRun.Location = New-Object System.Drawing.Point(578, 500)
$form.Controls.Add($btnRun)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = "Close"
$btnClose.Size     = New-Object System.Drawing.Size(80, 28)
$btnClose.Location = New-Object System.Drawing.Point(666, 500)
$form.Controls.Add($btnClose)

# No AcceptButton set — Enter in multiline boxes adds newlines, not trigger Run

# --- Result label + box ---
$lblResult          = New-Object System.Windows.Forms.Label
$lblResult.Text     = "Result:"
$lblResult.AutoSize = $true
$lblResult.Location = New-Object System.Drawing.Point(12, 542)
$form.Controls.Add($lblResult)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.ReadOnly   = $true
$txtOut.Size       = New-Object System.Drawing.Size(722, 148)
$txtOut.Location   = New-Object System.Drawing.Point(12, 562)
$txtOut.Font       = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOut)

# --- Helper: return selected action string ---
function Get-SelectedAction {
    if ($rbGrant.Checked)  { return "Grant" }
    if ($rbRemove.Checked) { return "Remove" }
}

# --- Helper: populate the parsed preview box ---
function Update-ParsedPreview {
    param($Parsed)
    $actionText = if ($Parsed.Action -eq "Grant") { "Grant Site Collection Admin" } else { "Remove Site Collection Admin" }
    $txtParsed.Lines = @(
        "Action           : $actionText"
        "Users       ($($Parsed.Users.Count)) : $($Parsed.Users -join ', ')"
        "Custodians  ($($Parsed.Custodians.Count)) : $($Parsed.Custodians -join ', ')"
        "Total operations : $($Parsed.OpCount)"
    )
}

# --- Parse button: preview without executing ---
$btnParse.Add_Click({
    $btnParse.Enabled = $false
    $btnRun.Enabled   = $false
    try {
        $action = Get-SelectedAction
        $p      = Get-ParsedInputs -UserRaw $txtUser.Text -CustodianRaw $txtCustodian.Text -Action $action

        if ($p.Users.Count -lt 1) {
            $txtOut.Text = "Please enter at least one disabled user."
            return
        }
        if ($p.Custodians.Count -lt 1) {
            $txtOut.Text = "Please enter at least one custodian."
            return
        }

        $script:Parsed     = $p
        $script:ParsedFrom = $txtUser.Text + "|" + $txtCustodian.Text + "|" + $action

        Update-ParsedPreview -Parsed $p
        $txtOut.Text = "Parsed OK. Review the preview above, then click Run to proceed."
    } catch {
        $txtOut.Text = "Parse error: $($_.Exception.Message)"
    } finally {
        $btnParse.Enabled = $true
        $btnRun.Enabled   = $true
    }
})

# --- Run button: auto-parses if stale, confirms, then executes ---
$btnRun.Add_Click({
    $btnRun.Enabled   = $false
    $btnParse.Enabled = $false
    try {
        $action   = Get-SelectedAction
        $stateKey = $txtUser.Text + "|" + $txtCustodian.Text + "|" + $action

        # Silently re-parse if input changed since last Parse click
        if (-not $script:Parsed -or $stateKey -ne $script:ParsedFrom) {
            $script:Parsed     = Get-ParsedInputs -UserRaw $txtUser.Text -CustodianRaw $txtCustodian.Text -Action $action
            $script:ParsedFrom = $stateKey
            Update-ParsedPreview -Parsed $script:Parsed
        }

        $p = $script:Parsed

        if ($p.Users.Count -lt 1) {
            $txtOut.Text = "Please enter at least one disabled user."
            return
        }
        if ($p.Custodians.Count -lt 1) {
            $txtOut.Text = "Please enter at least one custodian."
            return
        }

        $actionText = if ($p.Action -eq "Grant") { "Grant Site Collection Admin" } else { "Remove Site Collection Admin" }
        $userList   = ($p.Users      | ForEach-Object { "  $_" }) -join "`r`n"
        $custList   = ($p.Custodians | ForEach-Object { "  $_" }) -join "`r`n"

        $confirmMsg = ("Action: {0}`r`n`r`nDisabled User(s) ({1}):`r`n{2}`r`n`r`nCustodian(s) ({3}):`r`n{4}`r`n`r`nTotal operations: {5}`r`n`r`nProceed?" -f `
            $actionText, $p.Users.Count, $userList, $p.Custodians.Count, $custList, $p.OpCount)

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            $confirmMsg,
            "Confirm Action",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $txtOut.Text = "Cancelled."
            return
        }

        $txtOut.Text = "Running — please wait..."
        $form.Refresh()

        $result      = Invoke-OneDriveAccess -Users $p.Users -Custodians $p.Custodians -Action $p.Action
        $txtOut.Text = $result

    } catch {
        $txtOut.Text = "ERROR: $($_.Exception.Message)"
    } finally {
        $btnRun.Enabled   = $true
        $btnParse.Enabled = $true
    }
})

# --- Close button ---
$btnClose.Add_Click({ $form.Close() })

[void]$form.ShowDialog()

#endregion