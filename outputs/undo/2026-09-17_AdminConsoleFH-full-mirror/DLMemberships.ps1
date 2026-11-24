#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
# CLOUD VERSION, AND NOW THE ONLY VERSION.
#
# AD-DLMemberships.ps1 was the on-premises twin. It has been PARKED to
# DisabledTools\, so there is no longer an AD- variant beside this file and
# nothing here needs a domain controller.
#
# The header used to say the AD- variant was "beside it", which stopped being
# true the moment the file moved. A comment that names a sibling is a
# reference like any other -- it goes stale silently, and the next person
# reads it as current.
# -------------------------------------------------------------------------
<#
    DLMemberships.ps1
    Every distribution list a person belongs to.

    Reasoning: ScriptNotes/DLMemberships.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-19
#>
# TOOLKIT-HELP-BEGIN
# Shows every distribution list a person belongs to.
#
# Answers "why is this person getting this mail" and "what will they stop
# receiving when they leave" -- the second one is worth checking before a
# termination rather than after.
#
# Type an address and press Look Up. It is a single filtered query and should
# come back in a second or two, however many groups the tenant has.
#
# DIRECT MEMBERSHIPS ONLY. If Barney is in "All Staff" because "Accounts" is in
# "All Staff", this lists Accounts and not All Staff. That is a real limitation
# and the reason a person can still receive mail from a list that is not shown
# here.
#
# Microsoft 365 groups and dynamic distribution groups are not included --
# dynamic groups have no stored membership to query, they are evaluated when
# mail is sent.
#
# Read-only. To change membership, use DL Manager or Bulk DL Membership.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: Reads group membership from Exchange Online. Changes nothing.

param(
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

#region Lookup ------------------------------------------------------------------
function Get-MailboxDLMemberships {
    <#
        Which groups this address belongs to.

        ONE SERVER-SIDE QUERY, NOT A SCAN. This used to enumerate every
        distribution group in the tenant and call Get-DistributionGroupMember on
        each one -- roughly 300 round trips in one directory and 4,800 in the other, which
        is why it needed a progress bar, a Cancel button, and a paragraph
        apologising for itself.

        The apology said "Exchange Online has no reverse-membership cmdlet".
        THAT WAS WRONG. Exchange supports a server-side filter on the Members
        property, documented in Microsoft's own filterable-properties reference:

            $dn = (Get-Recipient user).DistinguishedName
            Get-Recipient -Filter "Members -eq '$dn'"

        Members is multi-valued and the filter matches if the property CONTAINS
        the value, which is exactly the question being asked. Exchange evaluates
        it and returns the matching groups in a single call.

        THE DISTINGUISHED NAME IS MANDATORY HERE. The filter does not accept an
        address or a display name for Members -- Microsoft's documentation says
        so explicitly, and passing one returns an empty set rather than an
        error, which is the worst possible failure for a tool whose whole output
        is a list that might legitimately be empty.

        DIRECT MEMBERSHIPS ONLY, same as the scan it replaces. Nested groups are
        not walked; a group that is a member of another group appears, the
        parent does not.

        DYNAMIC DISTRIBUTION GROUPS ARE NOT COVERED, and cannot be. They store
        no membership -- the recipient set is evaluated when mail is sent -- so
        there is nothing for a Members filter to match. Microsoft's filter
        documentation calls this out as an explicit exception.
    #>
    param([Parameter(Mandatory)][string]$Smtp)

    $result = [PSCustomObject]@{
        Groups     = @()
        UserDn     = ''
        FatalError = ''
    }

    # STEP 1: the recipient, for its distinguished name.
    try {
        $r = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop' }
    } catch {
        $result.FatalError = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Recipient (resolving the person)')
        return $result
    }

    $dn = "$($r.DistinguishedName)"
    if (-not $dn) {
        $result.FatalError = "'$Smtp' resolved, but Exchange returned no distinguished name for it, so the membership filter cannot be built."
        return $result
    }
    $result.UserDn = $dn

    # A DN can legitimately contain an apostrophe -- O'Brien, O'Neill. Inside a
    # single-quoted OPATH filter that ends the string early and the filter
    # either errors or, worse, silently matches nothing. Doubling is the escape.
    $dnEscaped = $dn.Replace("'", "''")

    # STEP 2: one filtered query.
    try {
        Write-ToolStep 'Enumerating group memberships ...'
        $groups = @(Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
            Filter      = "Members -eq '$dnEscaped'"
            ResultSize  = 'Unlimited'
            ErrorAction = 'Stop' })
    } catch {
        $result.FatalError = (Format-ErrorDetail -ErrorRecord $_ -Context "Get-Recipient -Filter ""Members -eq ...""")
        return $result
    }

    # PROPERTIES THE FILTER ALREADY RETURNED, and nothing else.
    #
    # The old version also showed each group's owner and member count. Those are
    # not on a Get-Recipient result, and fetching them would mean one more call
    # per group -- reintroducing exactly the round trips this rewrite removes.
    # A fast answer to the question asked beats a slow answer to a broader one;
    # DL Manager shows owners and counts for any group listed here.
    $result.Groups = @($groups | ForEach-Object {
        [PSCustomObject]@{
            DisplayName        = "$($_.DisplayName)"
            PrimarySmtpAddress = "$($_.PrimarySmtpAddress)"
            GroupType          = "$($_.RecipientTypeDetails)"
        }
    })
    return $result
}
#endregion

#region WinForms UI --------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'DL Memberships'
$form.Size          = New-Object System.Drawing.Size(760, 640)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

# --- Address row ------------------------------------------------------------
# ON THE FORM, not a modal prompt before it. Being asked for an address in one
# window and shown the answer in another means a second person means closing
# the whole tool and starting again -- and this is a tool you run several times
# in a row while working a ticket.
$lblWho          = New-Object System.Windows.Forms.Label
$lblWho.Text     = 'Mailbox or address:'
$lblWho.Location = New-Object System.Drawing.Point(12, 15)
$lblWho.AutoSize = $true
$form.Controls.Add($lblWho)

$txtWho          = New-Object System.Windows.Forms.TextBox
$txtWho.Location = New-Object System.Drawing.Point(140, 12)
$txtWho.Size     = New-Object System.Drawing.Size(340, 22)
$txtWho.Font     = New-Object System.Drawing.Font('Consolas', 9)
$txtWho.Anchor   = 'Top,Left,Right'
$form.Controls.Add($txtWho)

$btnLookup          = New-Object System.Windows.Forms.Button
$btnLookup.Text     = 'Look Up'
$btnLookup.Location = New-Object System.Drawing.Point(490, 10)
$btnLookup.Size     = New-Object System.Drawing.Size(110, 26)
$btnLookup.Anchor   = 'Top,Right'
$form.Controls.Add($btnLookup)

# Enter looks up. Typing an address and pressing Enter is what everyone does.
# The help '?' lands at ClientSize-40, so the button stops well short of it.
$form.AcceptButton = $btnLookup

# --- Note -------------------------------------------------------------------
# AUTOSIZE OFF, HEIGHT SET, AND THE TEXT SHORTENED.
#
# The previous note was a 270-character paragraph in a 46px label and the last
# two lines were simply cut off -- including the sentence explaining what the
# tool does NOT cover, which is the part worth reading. The long version now
# lives in the help block behind the '?'; this says the one thing that changes
# how you read the results.
$lblNote           = New-Object System.Windows.Forms.Label
$lblNote.Text      = 'Direct memberships only. A group reached through another group is not listed, and dynamic distribution groups cannot be included -- they store no membership.'
$lblNote.Location  = New-Object System.Drawing.Point(12, 44)
$lblNote.Size      = New-Object System.Drawing.Size(720, 34)
$lblNote.Anchor    = 'Top,Left,Right'
$lblNote.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblNote)

$lblProgress          = New-Object System.Windows.Forms.Label
$lblProgress.Text     = 'Enter an address above.'
$lblProgress.Location = New-Object System.Drawing.Point(12, 82)
$lblProgress.Size     = New-Object System.Drawing.Size(720, 20)
$lblProgress.Anchor   = 'Top,Left,Right'
$lblProgress.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblProgress)

# THE RESULTS BOX STARTS AT x12 LIKE EVERYTHING ELSE. It was inset while the
# labels above it were not, which read as a margin nobody chose.
$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 108)
$txtOut.Size       = New-Object System.Drawing.Size(720, 440)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Top,Left,Right,Bottom'
$form.Controls.Add($txtOut)

$btnCopy          = New-Object System.Windows.Forms.Button
$btnCopy.Text     = 'Copy'
$btnCopy.Location = New-Object System.Drawing.Point(12, 558)
$btnCopy.Size     = New-Object System.Drawing.Size(110, 28)
$btnCopy.Anchor   = 'Left,Bottom'
$form.Controls.Add($btnCopy)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(622, 558)
$btnClose.Size     = New-Object System.Drawing.Size(110, 28)
$btnClose.Anchor   = 'Right,Bottom'
$form.Controls.Add($btnClose)
#endregion

#region Behaviour ----------------------------------------------------------------
function Invoke-Lookup {
    $who = Get-TargetSmtp -Text $txtWho.Text
    if (-not $who) {
        $lblProgress.Text = 'Enter an address above.'
        $txtOut.Text      = ''
        return
    }

    $btnLookup.Enabled = $false
    $form.Cursor       = [System.Windows.Forms.Cursors]::WaitCursor
    $lblProgress.Text  = "Looking up $who ..."
    $form.Refresh()
    try {
        $result = Get-MailboxDLMemberships -Smtp $who

        if ($result.FatalError) {
            $lblProgress.Text = 'Failed.'
            $txtOut.Text      = $result.FatalError
            return
        }

        $groups = @($result.Groups)
        $lblProgress.Text = "$who -- $($groups.Count) group(s)."

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("Person : $who")
        [void]$sb.AppendLine("Groups : $($groups.Count)")
        [void]$sb.AppendLine('')

        if ($groups.Count -eq 0) {
            # AN EMPTY RESULT IS AN ANSWER, and it needs to look like one.
            # A blank box is indistinguishable from a lookup that silently
            # failed, which is the failure mode the DN requirement above can
            # otherwise produce.
            [void]$sb.AppendLine('No direct group memberships found.')
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('If you expected some, remember this lists DIRECT membership only --')
            [void]$sb.AppendLine('mail arriving through a group that contains another group will not')
            [void]$sb.AppendLine('show here, and dynamic distribution groups never will.')
        } else {
            $groups | Sort-Object DisplayName | ForEach-Object {
                [void]$sb.AppendLine("  $($_.DisplayName)  <$($_.PrimarySmtpAddress)>")
                [void]$sb.AppendLine("      $($_.GroupType)")
            }
        }
        $txtOut.Text = $sb.ToString()
    }
    catch {
        $lblProgress.Text = 'Failed.'
        $txtOut.Text      = Format-ErrorDetail -ErrorRecord $_ -Context 'DL membership lookup'
    }
    finally {
        $btnLookup.Enabled = $true
        $form.Cursor       = [System.Windows.Forms.Cursors]::Default
    }
}

$btnLookup.Add_Click({ Invoke-Lookup })
$btnClose.Add_Click({ $form.Close() })
$btnCopy.Add_Click({
    try {
        if ("$($txtOut.Text)".Trim()) { [System.Windows.Forms.Clipboard]::SetText($txtOut.Text) }
    } catch {
        $null = Show-ToolDialog(
            "Copy to clipboard failed: $($_.Exception.Message)", 'Copy',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# An address supplied by the console is looked up immediately -- having typed it
# once in the console's own box, being asked for it again is the thing moving
# the lookup onto this form was meant to remove.
if ("$SmtpAddress".Trim()) {
    $txtWho.Text = "$SmtpAddress".Trim()
    Invoke-Lookup
}

Set-ConsoleTheme -Form $form
Set-FormResizable -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL Memberships' | Out-Null
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
#endregion
