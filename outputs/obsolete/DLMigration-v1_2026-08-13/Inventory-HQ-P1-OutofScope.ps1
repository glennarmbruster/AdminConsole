#Requires -Version 5.1
<#
    Inventory-HQ-P1-OutofScope.ps1
    The other half of the phase 1 picture: the mail-enabled groups in the same
    directory that phase 1 is NOT moving, and why each one is being left.

    WHY THIS EXISTS AT ALL. The in-scope list on its own cannot be checked. A
    group missing from it might be out of scope by decision, or missing because
    a filter was wrong -- and there is no way to tell which from the in-scope
    file alone. Run both, and every mail-enabled group in the directory appears
    in exactly one of them. That is a number that can be reconciled against the
    directory, which is what makes the scope defensible in a change record.

    Read-only. On-premises Exchange and Active Directory only.

    Author   : Glenn Armbruster
    Revised  : 2026-08-13
#>
# TOOLKIT-HELP-BEGIN
# Lists the mail-enabled groups that phase 1 is NOT moving, with the reason for
# each, and writes a CSV to the reports folder.
#
# Three reasons appear:
#
#   Mail-enabled security group   it grants permissions as well as delivering
#                                 mail, so moving it moves access too
#   Room list                     it moves with the resource mailboxes later
#   Dynamic group                 it stores a filter, not a list of members
#
# Dynamic groups show a name and an address and nothing else. That is not a
# gap -- there is no stored membership to count. Getting a number means running
# each group's own filter, and the answer changes on its own as people move
# departments, so it would be stale before the file was read.
#
# Run this alongside the in-scope report. Between them every mail-enabled group
# in the directory is accounted for exactly once, which is how you prove the
# scope is right rather than asserting it.
#
# Read-only. Nothing is changed by running this.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: on-premises Exchange and Active Directory. Exchange Online is not queried.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Three levels up: Projects\<Tenant>\DLMigration\ -> the toolkit root.
$ToolkitRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
. (Join-Path $ToolkitRoot 'Common.ps1')

# FIRST THING AFTER Common LOADS -- see the note in the in-scope script.
if (-not (Test-ConsoleLaunch)) { return }

. (Join-Path $PSScriptRoot '_Inventory-Shared.ps1')

# The mirror image of $InScopeTypes next door. These two lists together must
# cover every mail-enabled group type, or a group falls through both reports
# and nobody finds out until the migration misses it.
#
# VERIFIED AGAINST THE CMDLET, NOT ASSUMED. Get-DistributionGroup returns
# exactly four subtypes -- MailUniversalDistributionGroup, MailNonUniversalGroup,
# MailUniversalSecurityGroup and RoomList. Two are in scope next door, two are
# here, and dynamic groups come from the other cmdlet. Nothing falls through
# and nothing is counted twice.
#
# THE ONE MAIL-ENABLED GROUP TYPE NEITHER REPORT SHOWS is GroupMailbox, the
# Microsoft 365 group. Neither cmdlet returns it, and that is correct here:
# a source-anchor project moves objects that are mastered on-premises, and a
# Microsoft 365 group is mastered in the cloud already. If group writeback is
# ever switched on, the written-back objects are still cloud-mastered and still
# out of scope -- but this is the paragraph to re-read before assuming so.
$OutOfScopeTypes = @(
    'MailUniversalSecurityGroup'
    'RoomList'
)

$ReasonByType = @{
    'MailUniversalSecurityGroup' = 'Mail-enabled security group -- grants permissions as well as mail'
    'RoomList'                   = 'Room list -- moves with the resource mailboxes in a later phase'
    'DynamicDistributionGroup'   = 'Dynamic group -- stores a filter, not a membership'
}

Show-InventoryWindow `
    -Title 'HQ Groups -- Phase 1 Out of Scope' `
    -Prefix 'Inventory-HQ-P1-OutofScope' `
    -Blurb 'Mail-enabled groups phase 1 is not moving, with the reason for each. Run alongside the in-scope report.' `
    -ScriptPath $PSCommandPath `
    -RecipientTypeDetails $OutOfScopeTypes `
    -IncludeDynamic `
    -ReasonMap $ReasonByType `
    -ExtraColumns @( @{ Name = 'OutOfScopeReason'; Header = 'Why It Is Out Of Scope'; Width = 330 } )
