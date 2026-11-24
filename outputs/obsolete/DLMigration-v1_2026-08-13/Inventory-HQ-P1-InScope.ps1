#Requires -Version 5.1
<#
    Inventory-HQ-P1-InScope.ps1
    Lists the distribution groups that phase 1 of the source-anchor project
    will move to Microsoft 365, and writes them to a CSV.

    PHASE 1 IS ONE DIRECTORY AND ONE KIND OF GROUP. Everything else -- the
    other directory, dynamic groups, mail-enabled security groups, room lists
    -- is in the companion report, Inventory-HQ-P1-OutofScope.ps1. Two files
    rather than one report with a column, because the question being asked is
    "what am I moving first", and an answer you have to filter is an answer you
    can get wrong.

    Read-only. Queries on-premises Exchange and Active Directory. Exchange
    Online is never touched -- these groups are all still source-anchored
    on-premises, which is the entire premise of the project.

    Author   : Glenn Armbruster
    Revised  : 2026-08-13
#>
# TOOLKIT-HELP-BEGIN
# Lists the distribution groups that are IN SCOPE for phase 1 of the move to
# Microsoft 365, and writes a CSV to the reports folder.
#
# In scope means all three of these:
#
#   - it lives in the default directory from Tenant.json, not a child of it
#   - it is a distribution group, not a mail-enabled security group
#   - it is not a room list
#
# Room lists are excluded on purpose. They go with the resource mailboxes in a
# later phase.
#
# Columns: group name, primary email, recipient type, AD group type, number of
# members, whether it contains other groups, and the owners.
#
# "Has Nested Groups" matters more than it looks. A group containing another
# group has to be moved in the right order, or the members go missing.
#
# Read-only. Nothing is changed by running this.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: on-premises Exchange and Active Directory. Exchange Online is not queried.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Three levels up: Projects\<Tenant>\DLMigration\ -> the toolkit root.
# Resolve-Path rather than a literal join so the path in any error message is
# the real one rather than one with '..' in the middle of it.
$ToolkitRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
. (Join-Path $ToolkitRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost.
if (-not (Test-ConsoleLaunch)) { return }

. (Join-Path $PSScriptRoot '_Inventory-Shared.ps1')

# THE SCOPE, AS DATA. Written here rather than buried in the query so that the
# one line that decides what phase 1 means can be read without following a call
# into another file.
#
# MailUniversalDistributionGroup  the ordinary case, and almost all of them.
# MailNonUniversalGroup           legacy, from before universal scope was the
#                                 default. Rare, migrates the same way, and
#                                 leaving it out would quietly under-count.
#
# NOT here, and each for its own reason:
#   MailUniversalSecurityGroup    carries permissions as well as membership
#   RoomList                      goes with the resource mailboxes
#   DynamicDistributionGroup      stores a filter, not a membership
$InScopeTypes = @(
    'MailUniversalDistributionGroup'
    'MailNonUniversalGroup'
)

Show-InventoryWindow `
    -Title 'HQ Distribution Groups -- Phase 1 In Scope' `
    -Prefix 'Inventory-HQ-P1-InScope' `
    -Blurb 'Distribution groups in scope for phase 1. Room lists and mail-enabled security groups are excluded.' `
    -ScriptPath $PSCommandPath `
    -RecipientTypeDetails $InScopeTypes
