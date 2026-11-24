#Requires -Version 5.1
<#
    SpamPolicyViewer.ps1
    Shows which anti-spam, anti-phish and outbound spam policies actually apply to one mailbox, and what they do.

    Author   : Glenn Armbruster
    Revised  : 2026-09-12
#>
# TOOLKIT-HELP-RUNS: reads anti-spam, anti-phish and outbound spam policies and rules from Exchange Online. Changes nothing.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Answers "which spam policy applies to this mailbox, and what does it do?"
#
# "Why did this get through" and "why was this blocked" are unanswerable
# until you know which policy won. Exchange picks the policy the same way
# for anti-spam, anti-phish and outbound spam: rules are walked in Priority
# order (0 first), the first ENABLED rule whose conditions match the
# recipient and whose exceptions do not wins, and if nothing matches the
# Default policy applies.
#
# This tool walks that ladder for one address and shows the winning policy,
# the rule that selected it, WHY it matched, the settings that matter, and
# the whole ladder so you can see what would have won instead.
#
# Group-scoped rules (SentToMemberOf / FromMemberOf) are flagged as
# "possibly" -- membership is NOT expanded. Check the group yourself.
#
# Read-only. Writes a CSV of the effective settings only when you ask.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# FIRST THING AFTER Common LOADS. See Test-ConsoleLaunch in Common.ps1.
if (-not (Test-ConsoleLaunch)) { return }

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

# Inbound spam settings shown for the winning policy, in this order. Each
# name was checked against Set-HostedContentFilterPolicy on Learn; a name the
# object does not carry is skipped rather than shown blank.
$InboundSettingNames = @(
    'SpamAction', 'HighConfidenceSpamAction', 'PhishSpamAction',
    'HighConfidencePhishAction', 'BulkSpamAction', 'BulkThreshold',
    'QuarantineRetentionPeriod', 'EnableEndUserSpamNotifications',
    'SpamZapEnabled', 'PhishZapEnabled', 'InlineSafetyTipsEnabled',
    'TestModeAction', 'EnableLanguageBlockList', 'EnableRegionBlockList'
)
# The advanced spam filter (ASF) switches. Only the ones that are not Off
# are shown -- a dozen "Off" rows would bury the one that matters.
$InboundAsfNames = @(
    'IncreaseScoreWithImageLinks', 'IncreaseScoreWithNumericIps',
    'IncreaseScoreWithRedirectToOtherPort', 'IncreaseScoreWithBizOrInfoUrls',
    'MarkAsSpamEmptyMessages', 'MarkAsSpamJavaScriptInHtml',
    'MarkAsSpamFramesInHtml', 'MarkAsSpamObjectTagsInHtml',
    'MarkAsSpamEmbedTagsInHtml', 'MarkAsSpamFormTagsInHtml',
    'MarkAsSpamWebBugsInHtml', 'MarkAsSpamSensitiveWordList',
    'MarkAsSpamSpfRecordHardFail', 'MarkAsSpamFromAddressAuthFail',
    'MarkAsSpamNdrBackscatter', 'MarkAsSpamBulkMail'
)
# Lists shown as a count, not the contents. The contents belong in a
# different tool; here the question is "is there an allow list at all".
$InboundListNames = @(
    'AllowedSenders', 'AllowedSenderDomains', 'BlockedSenders', 'BlockedSenderDomains'
)

# Anti-phish, checked against Set-AntiPhishPolicy on Learn.
$PhishSettingNames = @(
    'Enabled', 'ImpersonationProtectionState', 'EnableSpoofIntelligence',
    'AuthenticationFailAction', 'HonorDmarcPolicy', 'DmarcRejectAction',
    'DmarcQuarantineAction', 'PhishThresholdLevel',
    'EnableMailboxIntelligence', 'EnableMailboxIntelligenceProtection',
    'MailboxIntelligenceProtectionAction',
    'EnableTargetedUserProtection', 'TargetedUserProtectionAction',
    'EnableTargetedDomainsProtection', 'EnableOrganizationDomainsProtection',
    'TargetedDomainProtectionAction',
    'EnableFirstContactSafetyTips', 'EnableUnauthenticatedSender', 'EnableViaTag'
)
$PhishListNames = @('TargetedUsersToProtect', 'TargetedDomainsToProtect')

# Outbound, checked against Set-HostedOutboundSpamFilterPolicy on Learn.
$OutboundSettingNames = @(
    'RecipientLimitExternalPerHour', 'RecipientLimitInternalPerHour',
    'RecipientLimitPerDay', 'ActionWhenThresholdReached', 'AutoForwardingMode',
    'BccSuspiciousOutboundMail', 'NotifyOutboundSpam'
)
#endregion

#region Helpers --------------------------------------------------------------
function Get-PropValue {
    # Property access that tolerates a module build that does not carry the
    # property. Returns $null instead of throwing or emitting a strict-mode
    # error.
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Test-PropExists {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $false }
    return ($null -ne $Object.PSObject.Properties[$Name])
}

function ConvertTo-StringList {
    # Rule conditions come back as MultiValuedProperty or string[]; either way
    # we want plain trimmed strings with any "smtp:" prefix removed.
    param($Value)
    $out = @()
    if ($null -eq $Value) { return $out }
    foreach ($v in @($Value)) {
        $s = "$v".Trim()
        if ($s -match '^(?i)smtp:(.+)$') { $s = $Matches[1] }
        if ($s) { $out += $s }
    }
    return $out
}

function Format-ListPreview {
    param([string[]]$Items, [int]$Max = 3)
    $items = @($Items)
    if ($items.Count -eq 0) { return '' }
    $shown = @($items | Select-Object -First $Max)
    $text = ($shown -join ', ')
    if ($items.Count -gt $Max) { $text += " (+$($items.Count - $Max) more)" }
    return $text
}

function Invoke-PolicyRead {
    # One read per cmdlet, with the same guard pattern every tool uses:
    # the cmdlet must exist (else we are not connected) and ResultSize is
    # passed only when the installed build declares it. The policy cmdlets
    # do not declare it today; the check is cheap insurance.
    param([Parameter(Mandatory)][string]$Cmdlet)

    if (-not (Get-Command -Name $Cmdlet -ErrorAction SilentlyContinue)) {
        throw "$Cmdlet is not available. Connect to Exchange Online first."
    }
    $params = @{ ErrorAction = 'Stop' }
    $declared = Get-CmdletParameterNames -Name $Cmdlet
    if ($declared -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }
    return @(& $Cmdlet @params)
}
#endregion

#region Data -----------------------------------------------------------------
function Resolve-TargetRecipient {
    param([Parameter(Mandatory)][string]$Address)

    if (-not (Get-Command -Name 'Get-Recipient' -ErrorAction SilentlyContinue)) {
        throw 'Get-Recipient is not available. Connect to Exchange Online first.'
    }
    $found = @(Get-Recipient -Identity $Address -ErrorAction Stop)
    if ($found.Count -eq 0) { throw "No recipient found for '$Address'." }

    $r = $found[0]
    if ($found.Count -gt 1) {
        # An alias or display name can be ambiguous. Prefer the exact
        # address match; otherwise take the first and say so.
        $exact = @($found | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $Address })
        if ($exact.Count -gt 0) { $r = $exact[0] }
    }

    $primary = "$($r.PrimarySmtpAddress)"
    if (-not $primary -or $primary -notmatch '@') {
        throw "'$Address' resolved to '$($r.Name)' but it has no SMTP address."
    }
    $domain = $primary.Substring($primary.IndexOf('@') + 1)

    # Every handle a rule author could have typed into SentTo. Exchange
    # stores whatever resolves, so a rule may hold the alias, the name or
    # the DN rather than the address.
    $handles = New-Object System.Collections.Generic.List[string]
    $handles.Add($primary)
    foreach ($n in @('Alias', 'Name', 'DisplayName', 'DistinguishedName', 'Guid', 'Identity')) {
        $v = "$(Get-PropValue -Object $r -Name $n)"
        if ($v) { $handles.Add($v) }
    }
    foreach ($e in (ConvertTo-StringList -Value (Get-PropValue -Object $r -Name 'EmailAddresses'))) {
        if ($e -match '@') { $handles.Add($e) }
    }

    return [PSCustomObject]@{
        PrimarySmtpAddress = $primary
        Domain             = $domain
        DisplayName        = "$($r.DisplayName)"
        RecipientType      = "$(Get-PropValue -Object $r -Name 'RecipientTypeDetails')"
        Handles            = @($handles | Select-Object -Unique)
        Ambiguous          = ($found.Count -gt 1)
    }
}

function Test-RuleAgainstRecipient {
    <#
        Applies one rule's conditions and exceptions to the recipient the
        way the New-*Rule pages describe it:

          * Different condition TYPES use AND -- every non-empty condition
            must be satisfied.
          * Exception types use OR -- any one that matches excludes.

        Group conditions cannot be evaluated without expanding membership,
        so a rule whose only unresolved part is a group comes back as
        'Possible' rather than 'Match' or 'NoMatch'.

        The property names differ between inbound (SentTo...) and outbound
        (From...) rules; the caller passes the right set.
    #>
    param(
        [Parameter(Mandatory)]$Rule,
        [Parameter(Mandatory)]$Recipient,
        [Parameter(Mandatory)][string]$UserProp,
        [Parameter(Mandatory)][string]$GroupProp,
        [Parameter(Mandatory)][string]$DomainProp
    )

    $users   = ConvertTo-StringList -Value (Get-PropValue -Object $Rule -Name $UserProp)
    $groups  = ConvertTo-StringList -Value (Get-PropValue -Object $Rule -Name $GroupProp)
    $domains = ConvertTo-StringList -Value (Get-PropValue -Object $Rule -Name $DomainProp)
    $xUsers   = ConvertTo-StringList -Value (Get-PropValue -Object $Rule -Name "ExceptIf$UserProp")
    $xGroups  = ConvertTo-StringList -Value (Get-PropValue -Object $Rule -Name "ExceptIf$GroupProp")
    $xDomains = ConvertTo-StringList -Value (Get-PropValue -Object $Rule -Name "ExceptIf$DomainProp")

    $reasons  = @()
    $verdict  = 'Match'

    if ($users.Count + $groups.Count + $domains.Count -eq 0) {
        return [PSCustomObject]@{ Verdict = 'NoMatch'; Reason = 'Rule has no conditions' }
    }

    # Conditions: AND across types.
    if ($users.Count -gt 0) {
        $hit = @($users | Where-Object { $u = $_; @($Recipient.Handles | Where-Object { $_ -ieq $u }).Count -gt 0 })
        if ($hit.Count -gt 0) { $reasons += "$UserProp lists $($hit[0])" }
        else { return [PSCustomObject]@{ Verdict = 'NoMatch'; Reason = "$UserProp does not include this recipient" } }
    }
    if ($domains.Count -gt 0) {
        $hit = @($domains | Where-Object { $_ -ieq $Recipient.Domain })
        if ($hit.Count -gt 0) { $reasons += "$DomainProp lists $($hit[0])" }
        else { return [PSCustomObject]@{ Verdict = 'NoMatch'; Reason = "$DomainProp does not include $($Recipient.Domain)" } }
    }
    if ($groups.Count -gt 0) {
        $verdict = 'Possible'
        $reasons += "$GroupProp = $(Format-ListPreview -Items $groups) -- membership not checked"
    }

    # Exceptions: OR across types. A definite exception beats everything.
    if ($xUsers.Count -gt 0) {
        $hit = @($xUsers | Where-Object { $u = $_; @($Recipient.Handles | Where-Object { $_ -ieq $u }).Count -gt 0 })
        if ($hit.Count -gt 0) {
            return [PSCustomObject]@{ Verdict = 'NoMatch'; Reason = "excluded: ExceptIf$UserProp lists $($hit[0])" }
        }
    }
    if ($xDomains.Count -gt 0) {
        $hit = @($xDomains | Where-Object { $_ -ieq $Recipient.Domain })
        if ($hit.Count -gt 0) {
            return [PSCustomObject]@{ Verdict = 'NoMatch'; Reason = "excluded: ExceptIf$DomainProp lists $($hit[0])" }
        }
    }
    if ($xGroups.Count -gt 0) {
        $verdict = 'Possible'
        $reasons += "ExceptIf$GroupProp = $(Format-ListPreview -Items $xGroups) -- might exclude, membership not checked"
    }

    return [PSCustomObject]@{ Verdict = $verdict; Reason = ($reasons -join '; ') }
}

function Get-RuleTargetingText {
    param($Rule, [string]$UserProp, [string]$GroupProp, [string]$DomainProp)
    $parts = @()
    foreach ($n in @($UserProp, $GroupProp, $DomainProp)) {
        $v = ConvertTo-StringList -Value (Get-PropValue -Object $Rule -Name $n)
        if ($v.Count -gt 0) { $parts += "$n`: $(Format-ListPreview -Items $v)" }
    }
    if ($parts.Count -eq 0) { return '(no conditions)' }
    return ($parts -join ' | ')
}

function Get-RuleExceptionText {
    param($Rule, [string]$UserProp, [string]$GroupProp, [string]$DomainProp)
    $parts = @()
    foreach ($n in @("ExceptIf$UserProp", "ExceptIf$GroupProp", "ExceptIf$DomainProp")) {
        $v = ConvertTo-StringList -Value (Get-PropValue -Object $Rule -Name $n)
        if ($v.Count -gt 0) { $parts += "$n`: $(Format-ListPreview -Items $v)" }
    }
    return ($parts -join ' | ')
}

function Resolve-EffectivePolicy {
    <#
        Walks one policy type's ladder for the recipient.

        Returns the winning policy, the rule that picked it (or $null for
        the default), the reason, any group-scoped rules that sat ABOVE the
        winner and might have won instead, and the ladder rows for the grid.
    #>
    param(
        [Parameter(Mandatory)][string]$TypeLabel,
        [Parameter(Mandatory)]$Recipient,
        [Parameter(Mandatory)][AllowEmptyCollection()]$Policies,
        [Parameter(Mandatory)][AllowEmptyCollection()]$Rules,
        [Parameter(Mandatory)][string]$PolicyRefProp,
        [Parameter(Mandatory)][string]$UserProp,
        [Parameter(Mandatory)][string]$GroupProp,
        [Parameter(Mandatory)][string]$DomainProp
    )

    $policyByName = @{}
    foreach ($p in @($Policies)) { $policyByName["$($p.Name)"] = $p }

    $ordered = @($Rules | Sort-Object { [int](Get-PropValue -Object $_ -Name 'Priority') })

    $winner     = $null
    $winnerRule = $null
    $winnerWhy  = ''
    $possibles  = @()
    $ladder     = New-Object System.Collections.Generic.List[object]

    foreach ($rule in $ordered) {
        $state   = "$(Get-PropValue -Object $rule -Name 'State')"
        $enabled = ($state -ieq 'Enabled')
        $polName = "$(Get-PropValue -Object $rule -Name $PolicyRefProp)"
        $policy  = $null
        if ($policyByName.ContainsKey($polName)) { $policy = $policyByName[$polName] }

        $test = Test-RuleAgainstRecipient -Rule $rule -Recipient $Recipient `
                    -UserProp $UserProp -GroupProp $GroupProp -DomainProp $DomainProp

        $outcome = ''
        if (-not $enabled) {
            $outcome = 'Skipped -- rule disabled'
        } elseif ($null -eq $winner) {
            if ($test.Verdict -eq 'Match') {
                $winner = $policy; $winnerRule = $rule; $winnerWhy = $test.Reason
                $outcome = 'WINS -- ' + $test.Reason
            } elseif ($test.Verdict -eq 'Possible') {
                $possibles += [PSCustomObject]@{ Rule = "$($rule.Name)"; Policy = $polName; Reason = $test.Reason }
                $outcome = 'Possibly -- ' + $test.Reason
            } else {
                $outcome = 'No -- ' + $test.Reason
            }
        } else {
            $outcome = 'Not reached -- a higher rule already won'
        }

        $ladder.Add([PSCustomObject][ordered]@{
            Type       = $TypeLabel
            Priority   = "$([int](Get-PropValue -Object $rule -Name 'Priority'))"
            Policy     = $polName
            Rule       = "$($rule.Name)"
            IsDefault  = $false
            Enabled    = $enabled
            Targeting  = (Get-RuleTargetingText -Rule $rule -UserProp $UserProp -GroupProp $GroupProp -DomainProp $DomainProp)
            Exceptions = (Get-RuleExceptionText -Rule $rule -UserProp $UserProp -GroupProp $GroupProp -DomainProp $DomainProp)
            Outcome    = $outcome
        })
    }

    # The default policy has no rule. It sits at the bottom of the ladder
    # and catches whatever nothing above it claimed.
    $default = @($Policies | Where-Object { [bool](Get-PropValue -Object $_ -Name 'IsDefault') })
    $defaultPolicy = $null
    if ($default.Count -gt 0) { $defaultPolicy = $default[0] }

    $defaultOutcome = 'Not reached -- a rule above it won'
    if ($null -eq $winner) {
        if ($null -ne $defaultPolicy) {
            $winner    = $defaultPolicy
            $winnerWhy = 'No enabled rule matched, so the default policy applies'
            if ($possibles.Count -gt 0) { $defaultOutcome = 'WINS unless a "Possibly" rule above applies via group' }
            else { $defaultOutcome = 'WINS -- nothing above matched' }
        } else {
            $defaultOutcome = 'No default policy was returned'
        }
    }
    if ($null -ne $defaultPolicy) {
        $defEnabled = Get-PropValue -Object $defaultPolicy -Name 'Enabled'
        if ($null -eq $defEnabled) { $defEnabled = $true }
        $ladder.Add([PSCustomObject][ordered]@{
            Type       = $TypeLabel
            Priority   = '(last)'
            Policy     = "$($defaultPolicy.Name)"
            Rule       = '(default -- no rule)'
            IsDefault  = $true
            Enabled    = [bool]$defEnabled
            Targeting  = 'Everyone not claimed above'
            Exceptions = ''
            Outcome    = $defaultOutcome
        })
    }

    # Policies with neither a rule nor the default flag are orphans: they
    # exist but can never apply. Worth seeing on the ladder.
    $referenced = @($ordered | ForEach-Object { "$(Get-PropValue -Object $_ -Name $PolicyRefProp)" })
    foreach ($p in @($Policies)) {
        $pn = "$($p.Name)"
        if ([bool](Get-PropValue -Object $p -Name 'IsDefault')) { continue }
        if ($referenced -contains $pn) { continue }
        $ladder.Add([PSCustomObject][ordered]@{
            Type       = $TypeLabel
            Priority   = '-'
            Policy     = $pn
            Rule       = '(no rule)'
            IsDefault  = $false
            Enabled    = $false
            Targeting  = 'Nobody -- policy has no rule'
            Exceptions = ''
            Outcome    = 'Cannot apply'
        })
    }

    $winnerRuleName = ''
    if ($null -ne $winnerRule) { $winnerRuleName = "$($winnerRule.Name)" }
    $winnerName = ''
    if ($null -ne $winner) { $winnerName = "$($winner.Name)" }

    return [PSCustomObject]@{
        TypeLabel   = $TypeLabel
        Policy      = $winner
        PolicyName  = $winnerName
        RuleName    = $winnerRuleName
        Why         = $winnerWhy
        Possibles   = $possibles
        Ladder      = @($ladder)
        PolicyCount = @($Policies).Count
        RuleCount   = @($Rules).Count
    }
}

function Get-SettingRows {
    # Two-column rows for the effective-settings grid. Section rows carry the
    # policy name so the grid reads top to bottom without a third column.
    param(
        [string]$Section,
        $Policy,
        [string[]]$Names,
        [string[]]$AsfNames,
        [string[]]$ListNames
    )
    $rows = New-Object System.Collections.Generic.List[object]
    $label = '(none)'
    if ($null -ne $Policy) { $label = "$($Policy.Name)" }
    $rows.Add([PSCustomObject]@{ Setting = "[$Section]"; Value = $label; Section = $Section })
    if ($null -eq $Policy) { return @($rows) }

    foreach ($n in @($Names)) {
        if (-not (Test-PropExists -Object $Policy -Name $n)) { continue }
        $rows.Add([PSCustomObject]@{ Setting = $n; Value = "$(Get-PropValue -Object $Policy -Name $n)"; Section = $Section })
    }
    foreach ($n in @($AsfNames)) {
        if (-not (Test-PropExists -Object $Policy -Name $n)) { continue }
        $v = "$(Get-PropValue -Object $Policy -Name $n)"
        if ($v -and $v -ine 'Off') {
            $rows.Add([PSCustomObject]@{ Setting = $n; Value = $v; Section = $Section })
        }
    }
    foreach ($n in @($ListNames)) {
        if (-not (Test-PropExists -Object $Policy -Name $n)) { continue }
        $items = ConvertTo-StringList -Value (Get-PropValue -Object $Policy -Name $n)
        $rows.Add([PSCustomObject]@{ Setting = "$n (count)"; Value = "$($items.Count)"; Section = $Section })
    }
    return @($rows)
}

function Get-SpamPolicyReport {
    param(
        [Parameter(Mandatory)][string]$Address,
        [scriptblock]$Progress
    )

    & $Progress "Resolving $Address..." 5
    $recipient = Resolve-TargetRecipient -Address $Address

    & $Progress 'Reading anti-spam policies and rules...' 20
    $spamPolicies = Invoke-PolicyRead -Cmdlet 'Get-HostedContentFilterPolicy'
    $spamRules    = Invoke-PolicyRead -Cmdlet 'Get-HostedContentFilterRule'

    & $Progress 'Reading anti-phish policies and rules...' 45
    $phishPolicies = Invoke-PolicyRead -Cmdlet 'Get-AntiPhishPolicy'
    $phishRules    = Invoke-PolicyRead -Cmdlet 'Get-AntiPhishRule'

    & $Progress 'Reading outbound spam policies and rules...' 70
    $outPolicies = Invoke-PolicyRead -Cmdlet 'Get-HostedOutboundSpamFilterPolicy'
    $outRules    = Invoke-PolicyRead -Cmdlet 'Get-HostedOutboundSpamFilterRule'

    & $Progress 'Working out which policy wins...' 85

    # Inbound and anti-phish rules are recipient-scoped (SentTo...); outbound
    # rules are sender-scoped (From...). Same ladder, different property
    # names -- verified on New-HostedContentFilterRule, New-AntiPhishRule and
    # New-HostedOutboundSpamFilterRule.
    $spam = Resolve-EffectivePolicy -TypeLabel 'Inbound spam' -Recipient $recipient `
                -Policies $spamPolicies -Rules $spamRules -PolicyRefProp 'HostedContentFilterPolicy' `
                -UserProp 'SentTo' -GroupProp 'SentToMemberOf' -DomainProp 'RecipientDomainIs'
    $phish = Resolve-EffectivePolicy -TypeLabel 'Anti-phish' -Recipient $recipient `
                -Policies $phishPolicies -Rules $phishRules -PolicyRefProp 'AntiPhishPolicy' `
                -UserProp 'SentTo' -GroupProp 'SentToMemberOf' -DomainProp 'RecipientDomainIs'
    $outbound = Resolve-EffectivePolicy -TypeLabel 'Outbound spam' -Recipient $recipient `
                -Policies $outPolicies -Rules $outRules -PolicyRefProp 'HostedOutboundSpamFilterPolicy' `
                -UserProp 'From' -GroupProp 'FromMemberOf' -DomainProp 'SenderDomainIs'

    $settings = New-Object System.Collections.Generic.List[object]
    foreach ($row in (Get-SettingRows -Section 'Inbound spam policy' -Policy $spam.Policy `
                        -Names $InboundSettingNames -AsfNames $InboundAsfNames -ListNames $InboundListNames)) { $settings.Add($row) }
    foreach ($row in (Get-SettingRows -Section 'Anti-phish policy' -Policy $phish.Policy `
                        -Names $PhishSettingNames -AsfNames @() -ListNames $PhishListNames)) { $settings.Add($row) }
    foreach ($row in (Get-SettingRows -Section 'Outbound spam policy' -Policy $outbound.Policy `
                        -Names $OutboundSettingNames -AsfNames @() -ListNames @())) { $settings.Add($row) }

    $ladder = New-Object System.Collections.Generic.List[object]
    foreach ($r in $spam.Ladder)     { $ladder.Add($r) }
    foreach ($r in $phish.Ladder)    { $ladder.Add($r) }
    foreach ($r in $outbound.Ladder) { $ladder.Add($r) }

    return [PSCustomObject]@{
        Recipient = $recipient
        Spam      = $spam
        Phish     = $phish
        Outbound  = $outbound
        Settings  = @($settings)
        Ladder    = @($ladder)
    }
}

function Format-ReportSummary {
    param($Result)
    $sb = New-Object System.Text.StringBuilder
    $r  = $Result.Recipient

    [void]$sb.AppendLine("Mailbox : $($r.PrimarySmtpAddress)  ($($r.DisplayName), $($r.RecipientType))")
    [void]$sb.AppendLine("Domain  : $($r.Domain)")
    if ($r.Ambiguous) {
        [void]$sb.AppendLine('NOTE: the address you typed matched more than one recipient; the exact address match was used.')
    }
    [void]$sb.AppendLine('')

    foreach ($e in @($Result.Spam, $Result.Phish, $Result.Outbound)) {
        $line = "$($e.TypeLabel): "
        if (-not $e.PolicyName) {
            $line += 'no policy could be determined (no matching rule and no default policy returned).'
        } elseif ($e.RuleName) {
            $line += "'$($e.PolicyName)' applies, selected by rule '$($e.RuleName)' because $($e.Why)."
        } else {
            $line += "'$($e.PolicyName)' applies. $($e.Why)."
        }
        [void]$sb.AppendLine($line)
        foreach ($p in @($e.Possibles)) {
            [void]$sb.AppendLine("  Possibly instead: rule '$($p.Rule)' -> '$($p.Policy)' ($($p.Reason)).")
        }
        [void]$sb.AppendLine("  Ladder: $($e.RuleCount) rule(s), $($e.PolicyCount) policy/policies.")
    }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('How this was decided: rules are walked in Priority order (0 first). The first ENABLED')
    [void]$sb.AppendLine('rule whose conditions all match (user, domain) and whose exceptions do not wins.')
    [void]$sb.AppendLine('If nothing matches, the Default policy applies. Inbound spam and anti-phish rules')
    [void]$sb.AppendLine('target the recipient; outbound spam rules target the sender.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Group-scoped rules (SentToMemberOf / FromMemberOf) are listed but membership was NOT')
    [void]$sb.AppendLine('expanded. A "Possibly" line above means a group rule sits higher than the winner and')
    [void]$sb.AppendLine('would take over if this mailbox is a member. Check the group before concluding.')
    [void]$sb.AppendLine('Domain matching uses the primary SMTP domain only.')
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Spam Policy Viewer -- which policy applies to this mailbox'
$form.Size            = New-Object System.Drawing.Size(900, 730)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblAddr          = New-Object System.Windows.Forms.Label
$lblAddr.Text     = 'Mailbox:'
$lblAddr.Location = New-Object System.Drawing.Point(12, 19)
$lblAddr.Size     = New-Object System.Drawing.Size(60, 18)
$form.Controls.Add($lblAddr)

$txtAddr          = New-Object System.Windows.Forms.TextBox
$txtAddr.Text     = "$SmtpAddress"
$txtAddr.Location = New-Object System.Drawing.Point(76, 16)
$txtAddr.Size     = New-Object System.Drawing.Size(380, 24)
$form.Controls.Add($txtAddr)

$btnLookup          = New-Object System.Windows.Forms.Button
$btnLookup.Text     = 'Look up'
$btnLookup.Location = New-Object System.Drawing.Point(468, 12)
$btnLookup.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnLookup)

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(590, 12)
$btnExport.Size     = New-Object System.Drawing.Size(120, 30)
$btnExport.Enabled  = $false
$form.Controls.Add($btnExport)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(762, 12)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 52)
$lblStatus.Size     = New-Object System.Drawing.Size(860, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 72 -Width 860 -Height 12

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'SpamPolicy' `
                -X 12 -Y 90 -Width 860 -StatusLabel $lblStatus

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, 130)
$tabs.Size     = New-Object System.Drawing.Size(860, 330)
$form.Controls.Add($tabs)

$tabEffective      = New-Object System.Windows.Forms.TabPage
$tabEffective.Text = 'Effective settings'
$tabs.TabPages.Add($tabEffective)

$tabLadder      = New-Object System.Windows.Forms.TabPage
$tabLadder.Text = 'All policies (the ladder)'
$tabs.TabPages.Add($tabLadder)

$gridSettings                       = New-Object System.Windows.Forms.DataGridView
$gridSettings.Dock                  = 'Fill'
$gridSettings.ReadOnly              = $true
$gridSettings.AllowUserToAddRows    = $false
$gridSettings.AllowUserToDeleteRows = $false
$gridSettings.RowHeadersVisible     = $false
$gridSettings.SelectionMode         = 'FullRowSelect'
$gridSettings.AutoSizeColumnsMode   = 'Fill'
$gridSettings.AutoGenerateColumns   = $false
$colSetting = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSetting.HeaderText = 'Setting'; $colSetting.DataPropertyName = 'Setting'; $colSetting.FillWeight = 45
$colValue   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colValue.HeaderText   = 'Value';   $colValue.DataPropertyName   = 'Value';   $colValue.FillWeight   = 55
[void]$gridSettings.Columns.Add($colSetting)
[void]$gridSettings.Columns.Add($colValue)
$tabEffective.Controls.Add($gridSettings)

$gridLadder                       = New-Object System.Windows.Forms.DataGridView
$gridLadder.Dock                  = 'Fill'
$gridLadder.ReadOnly              = $true
$gridLadder.AllowUserToAddRows    = $false
$gridLadder.AllowUserToDeleteRows = $false
$gridLadder.RowHeadersVisible     = $false
$gridLadder.SelectionMode         = 'FullRowSelect'
$gridLadder.AutoSizeColumnsMode   = 'Fill'
$gridLadder.AutoGenerateColumns   = $false
foreach ($spec in @(
    @{ H = 'Type';       P = 'Type';       W = 9  },
    @{ H = 'Priority';   P = 'Priority';   W = 6  },
    @{ H = 'Policy';     P = 'Policy';     W = 16 },
    @{ H = 'Rule';       P = 'Rule';       W = 14 },
    @{ H = 'Default';    P = 'IsDefault';  W = 6  },
    @{ H = 'Enabled';    P = 'Enabled';    W = 6  },
    @{ H = 'Targeting';  P = 'Targeting';  W = 20 },
    @{ H = 'Exceptions'; P = 'Exceptions'; W = 12 },
    @{ H = 'Outcome';    P = 'Outcome';    W = 18 }
)) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.HeaderText = $spec.H; $c.DataPropertyName = $spec.P; $c.FillWeight = $spec.W
    [void]$gridLadder.Columns.Add($c)
}
$tabLadder.Controls.Add($gridLadder)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 470)
$txtOut.Size       = New-Object System.Drawing.Size(860, 205)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

# Result of the last lookup. Kept at file scope so Export can reach it; the
# grids hold display copies only.
$script:LastReport = $null

$btnLookup.Add_Click({
    $addr = $txtAddr.Text.Trim()
    if (-not $addr) {
        $lblStatus.Text = 'Type a mailbox address first.'
        return
    }
    $btnLookup.Enabled = $false
    $btnExport.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    $gridSettings.DataSource = $null
    $gridLadder.DataSource   = $null
    $script:LastReport = $null
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Spam policy lookup'
        $report = Get-SpamPolicyReport -Address $addr -Progress $progress
        Complete-ProgressReporter -Activity 'Spam policy lookup' -Bar $pbar

        $script:LastReport = $report

        # ArrayList so the grid gets an IList it can bind to; a raw
        # PowerShell array binds to nothing.
        $settingsList = New-Object System.Collections.ArrayList
        foreach ($row in $report.Settings) { [void]$settingsList.Add($row) }
        $gridSettings.DataSource = $settingsList

        $ladderList = New-Object System.Collections.ArrayList
        foreach ($row in $report.Ladder) { [void]$ladderList.Add($row) }
        $gridLadder.DataSource = $ladderList

        $txtOut.Text = Format-ReportSummary -Result $report
        $lblStatus.Text = "Done. Inbound: $($report.Spam.PolicyName)  |  Anti-phish: $($report.Phish.PolicyName)  |  Outbound: $($report.Outbound.PolicyName)"
        $btnExport.Enabled = $true

        Write-CompletedLog -Task "Spam policy lookup -- $($report.Recipient.PrimarySmtpAddress)" -Result 'Success'
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Spam policy lookup'
    } finally {
        Complete-ProgressReporter -Activity 'Spam policy lookup'
        $form.Cursor       = $prev
        $btnLookup.Enabled = $true
    }
})

$btnExport.Add_Click({
    try {
        $report = $script:LastReport
        if ($null -eq $report -or @($report.Settings).Count -eq 0) {
            $lblStatus.Text = 'Nothing to export -- run a lookup first.'
            return
        }
        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $safe  = ($report.Recipient.PrimarySmtpAddress -replace '[^A-Za-z0-9.@_-]', '_')
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "SpamPolicy_$($safe)_$stamp.csv"

        $rows = foreach ($s in $report.Settings) {
            [PSCustomObject][ordered]@{
                Mailbox = $report.Recipient.PrimarySmtpAddress
                Section = $s.Section
                Setting = $s.Setting
                Value   = $s.Value
            }
        }
        $rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text = "Exported $(@($rows).Count) row(s) to $csv"
    } catch {
        Show-HandlerError -ErrorRecord $_ -Context 'Export spam policy CSV'
    }
})

$form.Add_Shown({
    try {
        if ($txtAddr.Text.Trim()) { $btnLookup.PerformClick() }
    } catch {
        Show-HandlerError -ErrorRecord $_ -Context 'Spam policy auto-lookup'
    }
})

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Spam Policy Viewer' | Out-Null

[void]$form.ShowDialog()
#endregion
