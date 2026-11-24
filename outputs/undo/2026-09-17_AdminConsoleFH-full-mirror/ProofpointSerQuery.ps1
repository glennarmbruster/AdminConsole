#Requires -Version 5.1
<#
    ProofpointSerQuery.ps1
    Query functions for Proofpoint Secure Email Relay (SER) API integration with Unified Trace
    
    This module provides query functions that integrate with Unified-Trace.ps1
    to add Proofpoint SER API data to the unified message trace.
    
    Credentials stored in environment variables:
    - PROOFPOINT_SER_KEY
    - PROOFPOINT_SER_SECRET
    
    Author: Alan Smith (with Glenn Armbruster's Unified Trace)
    Date: August 27, 2026
#>

<#
    Get-ProofpointSerRows
    Query Proofpoint SER API and return rows for Unified Trace
    
    Parameters:
    - From: Sender email address to search for
    - To: Recipient email address to search for
    - Start: Start time for query
    - End: End time for query
    
    Returns: Adds rows to $script:Rows via Add-TraceRow
#>
function Get-ProofpointSerRows {
    param(
        [string]$From,
        [string]$To,
        [datetime]$Start,
        [datetime]$End
    )
    
    # Get SER credentials from environment variables
    $serKey = $env:PROOFPOINT_SER_KEY
    $serSecret = $env:PROOFPOINT_SER_SECRET
    
    if (-not $serKey -or -not $serSecret) {
        [void]$script:Notes.Add("Proofpoint SER -- Credentials not configured in environment variables")
        return
    }
    
    try {
        # Proofpoint SER API endpoint
        $endpoint = "https://api.proofpoint.com/v1/message/search"
        
        # Build query parameters
        $queryParams = @{
            key = $serKey
            secret = $serSecret
            from = if ($From) { $From } else { $null }
            to = if ($To) { $To } else { $null }
            startTime = $Start.ToString("yyyy-MM-ddTHH:mm:ssZ")
            endTime = $End.ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        
        # Remove null values
        $queryParams = $queryParams.GetEnumerator() | Where-Object { $_.Value } | 
            ForEach-Object { @{ $_.Key = $_.Value } } | 
            Measure-Object -Sum | ForEach-Object { $_.Sum }
        
        # Make API call
        $response = Invoke-RestMethod -Uri $endpoint -Method Get -Body $queryParams -ErrorAction Stop
        
        # Parse response and add rows
        if ($response.messages) {
            foreach ($msg in $response.messages) {
                Add-TraceRow -When $msg.timestamp -Source 'Proofpoint SER' `
                    -Sender $msg.sender -Recipient ($msg.recipients -join ', ') `
                    -Subject $msg.subject -Outcome $msg.status `
                    -Detail "SER API" -SizeBytes 0
            }
            [void]$script:Notes.Add("Proofpoint SER -- Query completed successfully ($(@($response.messages).Count) messages)")
        } else {
            [void]$script:Notes.Add("Proofpoint SER -- No messages found")
        }
        
    } catch {
        [void]$script:Notes.Add("Proofpoint SER -- Query failed: $($_.Exception.Message)")
    }
}

# Export-ModuleMember THROWS WHEN THIS FILE IS DOT-SOURCED, which is how it is
# actually loaded -- AvePoint-Dashboard does '. (Join-Path $PSScriptRoot ...)'.
# The cmdlet is only legal inside a module, so calling it unguarded turns a
# working library into a load-time error. Dot-sourcing already puts every
# function in the caller's scope, so there is nothing to export in that case.
if ($MyInvocation.MyCommand.ScriptBlock.Module) {
    Export-ModuleMember -Function @('Get-ProofpointSerRows')
}
