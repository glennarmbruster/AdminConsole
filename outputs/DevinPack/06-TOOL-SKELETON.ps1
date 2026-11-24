#Requires -Version 5.1

# TOOLKIT-HELP-RUNS: reads one mailbox. Changes nothing.
# TOOLKIT-HELP-BEGIN
# One or two sentences an OPERATOR would find useful, in plain language.
#
# Say what question this answers and when it is the wrong tool. The second half
# matters more than it looks: somebody who reaches for the wrong tool concludes
# the data is missing rather than that they are in the wrong place.
#
# Read-only. Nothing here changes anything.
# TOOLKIT-HELP-END

<#
    <ToolName>.ps1
    One line on what this is.

    Reasoning: ScriptNotes/<ToolName>.ps1.notes.md
        -- OPTIONAL and out of scope for the first pass. The convention is that
        the long reasoning lives beside the toolkit in a notes file while the
        script keeps only short guardrail comments. Do not create the folder
        unless you have something to put in it; an empty convention is worse
        than none.

    THE SHAPE OF THIS FILE IS THE POINT. Every tool in the toolkit follows it,
    so an operator who has read one can read any of them, and a reviewer knows
    where to look for the parts that matter -- the connection guard, the
    handler wrapping, and the audit call.
#>

# -----------------------------------------------------------------------------
# Parameters
# -----------------------------------------------------------------------------
# DECLARE ONLY WHAT THIS TOOL USES. The console asks each script what it
# declares and passes only that. A script declaring SOME of the common
# parameters rejects the others with a binding error, while one declaring no
# param() block at all silently absorbs extras into $args -- which is why a
# fixed-set launcher works right up until somebody writes a normal script.
param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -----------------------------------------------------------------------------
# Shared layer
# -----------------------------------------------------------------------------
# SEARCH UPWARD, NEVER COUNT '..'. A hardcoded climb encodes the depth of the
# folder tree into every file at the bottom of it, and fails in the worst
# available way: the wrong number of levels usually lands on a directory that
# EXISTS, so Join-Path succeeds and the error names a plausible folder instead
# of saying the file moved.
$ToolkitRoot = $PSScriptRoot
while ($ToolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $ToolkitRoot 'Common.ps1'))) {
    $parent = Split-Path -Parent $ToolkitRoot
    # Without this guard a drive root or certain UNC forms loop forever.
    if (-not $parent -or $parent -eq $ToolkitRoot) { $ToolkitRoot = ''; break }
    $ToolkitRoot = $parent
}
if (-not $ToolkitRoot) { throw "Common.ps1 not found above '$PSScriptRoot'." }
. (Join-Path $ToolkitRoot 'Common.ps1')

# -----------------------------------------------------------------------------
# Connection guard
# -----------------------------------------------------------------------------
# A TOOL THAT WRITES ON-PREMISES MUST REFUSE, NOT WARN, when a cloud session is
# live in the same process. Both editions export identically-named cmdlets and
# the last import owns the name, so a tool can report "connected to
# on-premises" truthfully and still send every call to the cloud. See C1/C2.
#
# Delete this block in a read-only tool. Keep it in anything that writes.
#
# A tool that cannot share a process declares it in its own header, in the
# first 40 lines, and the console obeys that rather than a checkbox somebody
# has to remember.
# Started by the console, or double-clicked in Explorer? A tool run on its own
# has no looked-up object, no directory detection and no connection guard, so it
# fails somewhere confusing instead of at the front door. Courtesy, NOT a
# security control -- see the contract note in 08.
if (-not (Test-ConsoleLaunch)) { return }

$exo = Test-ExoConnectionState
if (-not $exo.Connected) {
    [System.Windows.Forms.MessageBox]::Show(
        "Not connected to Exchange Online. Use the Connect button on the console first.`r`n`r`n$($exo.Detail)",
        'Not connected', 'OK', 'Warning') | Out-Null
    return
}

# AND IF THIS TOOL WROTE ON-PREMISES, it would also have to REFUSE -- not warn
# -- when a cloud session is live in this same process, because both editions
# export the same cmdlet names and the last import owns them (C1/C2):
#
#     if ((Test-EmsAvailable) -and (Test-ExoInProcess)) { ...refuse... }
#
# This tool is read-only against the cloud, so the guard is shown and not used.

# -----------------------------------------------------------------------------
# Work
# -----------------------------------------------------------------------------
function Get-ToolData {
    <#
        Keep the work in functions and the UI below. It makes the tool
        readable and the logic testable without a form.
    #>
    param([Parameter(Mandatory)][string]$Identity)

    # RESOLVE ONCE, THEN BIND ON A KEY THAT CANNOT BE RE-INTERPRETED (C4).
    # A returned object's .Identity renders as the AD CN -- no domain, no
    # address -- and passing it back is a fresh forest-wide name lookup that
    # can resolve to a DIFFERENT object, not a bind.
    $mbx = Get-Mailbox -Identity $Identity -ErrorAction Stop
    $key = $mbx.Guid.ToString()

    # ENUMERATING CALLS NEED ResultSize EXPLICITLY (C3). The default is 1000
    # and the short answer is silent. A call that fetches ONE object by
    # identity does not need it.
    $perms = @(Get-MailboxPermission -Identity $key -ResultSize Unlimited -ErrorAction Stop |
               Where-Object { -not $_.IsInherited })

    return [PSCustomObject]@{
        Guid        = $key
        Primary     = "$($mbx.PrimarySmtpAddress)"
        DisplayName = "$($mbx.DisplayName)"
        Permissions = $perms
    }
}

# -----------------------------------------------------------------------------
# UI
# -----------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = '<Tool Name>'
$form.ClientSize = New-Object System.Drawing.Size(840, 560)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

$txtId = New-Object System.Windows.Forms.TextBox
$txtId.Location = New-Object System.Drawing.Point(12, 12)
$txtId.Size = New-Object System.Drawing.Size(300, 24)
$txtId.Text = "$SmtpAddress"
$form.Controls.Add($txtId)

$btnGo = New-Object System.Windows.Forms.Button
$btnGo.Text = 'Look up'
$btnGo.Location = New-Object System.Drawing.Point(320, 11)
$btnGo.Size = New-Object System.Drawing.Size(100, 26)
$form.Controls.Add($btnGo)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(12, 48)
$txtOut.Size = New-Object System.Drawing.Size(816, 500)
$txtOut.Multiline = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.ReadOnly = $true
# WORD WRAP ON, ALWAYS (B6). A real diagnosis once sat entirely in the second
# clause of an error message, off the right edge of a window with wrapping
# disabled, and three rounds of fixes went to the wrong layer. The cost of a
# horizontal scrollbar is that nobody scrolls.
$txtOut.WordWrap = $true
$form.Controls.Add($txtOut)

# -----------------------------------------------------------------------------
# Handlers
# -----------------------------------------------------------------------------
# EVERY HANDLER IS WRAPPED (B1). An exception inside a WinForms handler is
# caught by .NET rather than PowerShell, so the operator gets a CLR crash
# dialog -- a hundred lines of internals, no script name, no line number.
# Show-HandlerError converts it back and writes a Failed audit row, because an
# error the operator dismissed is otherwise the one failure that leaves no
# trace at all.
$btnGo.Add_Click({
    try {
        $id = "$($txtId.Text)".Trim()
        if (-not $id) { return }

        $data = Get-ToolData -Identity $id

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("$($data.DisplayName)  <$($data.Primary)>")
        [void]$sb.AppendLine('')
        foreach ($p in $data.Permissions) {
            [void]$sb.AppendLine("  $($p.User)  --  $($p.AccessRights -join ', ')")
        }
        $txtOut.Text = $sb.ToString()

        # TARGET IS THE PRIMARY SMTP ADDRESS, NEVER A DISPLAY NAME. With
        # several mailboxes in play a display name does not identify what was
        # looked at, let alone what changed.
        Write-CompletedLog -Action 'Read' -Target $data.Primary -Result 'Success' `
            -Detail "$(@($data.Permissions).Count) explicit permission(s)"
    }
    catch {
        # Never an empty catch. Surface it AND record it.
        Show-HandlerError -ErrorRecord $_ -Context 'Mailbox lookup'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox lookup'
        Write-FailedLog -Action 'Read' -Target "$($txtId.Text)".Trim() -ErrorRecord $_
    }
})

# -----------------------------------------------------------------------------
# Show
# -----------------------------------------------------------------------------
Set-ConsoleTheme -Form $form
# BOTTOM-LEFT. The console owns the top-left corner and stays there; a tool
# that also opens top-left lands on top of it and hides the connection
# indicators you may need to read while the tool is running.
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title '<Tool Name>' | Out-Null

[void]$form.ShowDialog()
$form.Dispose()
