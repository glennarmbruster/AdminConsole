#Requires -Version 5.1
<#
    ProofpointPodQuery.ps1
    Query functions for Proofpoint PoD Log API integration with Unified Trace
    
    This module provides query functions that integrate with Unified-Trace.ps1
    to add Proofpoint PoD Log API data to the unified message trace.
    
    Author: Alan Smith (with Glenn Armbruster's Unified Trace)
    Date: August 27, 2026
#>

<#
    Get-ProofpointPodRows
    Query Proofpoint PoD Log API and return rows for Unified Trace
    
    Parameters:
    - From: Sender email address to search for
    - To: Recipient email address to search for
    - Start: Start time for query
    - End: End time for query
    
    Returns: Adds rows to $script:Rows via Add-TraceRow
#>
function Get-ProofpointPodRows {
    param(
        [string]$From,
        [string]$To,
        [datetime]$Start,
        [datetime]$End
    )
    
    # Get Proofpoint credentials from the credential cache
    $cred = Get-ApiSecret -Service 'ProofpointPod' -Purpose 'Unified Trace'
    if (-not $cred.Ok) {
        [void]$script:Notes.Add("Proofpoint PoD -- $($cred.Reason)")
        return
    }
    
    # Extract cluster ID and API token from credential
    $clusterId = $cred.Credential.UserName
    $apiToken = $cred.Credential.GetNetworkCredential().Password
    
    if (-not $clusterId -or -not $apiToken) {
        [void]$script:Notes.Add("Proofpoint PoD -- Credentials incomplete (missing cluster ID or API token)")
        return
    }
    
    try {
        # Call Python script to query Proofpoint
        $pythonScript = Join-Path $PSScriptRoot 'proofpoint_mcp.py'
        
        if (-not (Test-Path $pythonScript)) {
            [void]$script:Notes.Add("Proofpoint PoD -- Query script not found at $pythonScript")
            return
        }
        
        # Set environment variables for Python script
        $env:PROOFPOINT_CLUSTER_ID = $clusterId
        $env:PROOFPOINT_JWT_TOKEN = $apiToken
        
        # Build query parameters
        $queryParams = @{
            sender = if ($From) { $From } else { $null }
            recipient = if ($To) { $To } else { $null }
            hours_back = [int](($End - $Start).TotalHours)
        }
        
        # Run Python query
        $output = & python $pythonScript @queryParams 2>&1
        
        # Parse output and add rows
        $lines = $output -split "`n"
        $currentEmail = $null
        
        foreach ($line in $lines) {
            # Look for email headers in output
            if ($line -match "^From:\s+(.+)$") {
                if ($currentEmail) {
                    # Add previous email to rows
                    Add-TraceRow -When $currentEmail.Time -Source 'Proofpoint' `
                        -SenderAddress $currentEmail.From -Recipient $currentEmail.To `
                        -Subject $currentEmail.Subject -Outcome $currentEmail.Status `
                        -Detail "PoD Log API" -SizeBytes 0
                }
                
                $currentEmail = @{
                    From = $matches[1]
                    To = ""
                    Subject = ""
                    Time = $End
                    Status = "CONTINUE"
                }
            }
            elseif ($line -match "^To:\s+(.+)$" -and $currentEmail) {
                $currentEmail.To = $matches[1]
            }
            elseif ($line -match "^Subject:\s+(.+)$" -and $currentEmail) {
                $currentEmail.Subject = $matches[1]
            }
            elseif ($line -match "^Time:\s+(.+)$" -and $currentEmail) {
                try {
                    $currentEmail.Time = [datetime]$matches[1]
                } catch {
                    $currentEmail.Time = $End
                }
            }
            elseif ($line -match "^Status:\s+(.+)$" -and $currentEmail) {
                $currentEmail.Status = $matches[1]
            }
        }
        
        # Add last email if exists
        if ($currentEmail) {
            Add-TraceRow -When $currentEmail.Time -Source 'Proofpoint' `
                -SenderAddress $currentEmail.From -Recipient $currentEmail.To `
                -Subject $currentEmail.Subject -Outcome $currentEmail.Status `
                -Detail "PoD Log API" -SizeBytes 0
        }
        
        [void]$script:Notes.Add("Proofpoint PoD -- Query completed successfully")
        
    } catch {
        [void]$script:Notes.Add("Proofpoint PoD -- Query failed: $($_.Exception.Message)")
    }
}

# Export-ModuleMember THROWS WHEN THIS FILE IS DOT-SOURCED, which is how it is
# actually loaded -- AvePoint-Dashboard does '. (Join-Path $PSScriptRoot ...)'.
# The cmdlet is only legal inside a module, so calling it unguarded turns a
# working library into a load-time error. Dot-sourcing already puts every
# function in the caller's scope, so there is nothing to export in that case.
if ($MyInvocation.MyCommand.ScriptBlock.Module) {
    Export-ModuleMember -Function @('Get-ProofpointPodRows')
}
