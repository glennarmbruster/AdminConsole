#Requires -Version 5.1
<#
    SendGridQuery.ps1
    Query functions for SendGrid API integration with Unified Trace
    
    This module provides query functions that integrate with Unified-Trace.ps1
    to add SendGrid email data to the unified message trace.
    
    Credentials stored in environment variables:
    - SENDGRID_API_KEY
    
    Author: Alan Smith (with Glenn Armbruster's Unified Trace)
    Date: August 27, 2026
#>

<#
    Get-SendGridRows
    Query SendGrid API and return rows for Unified Trace
    
    Parameters:
    - From: Sender email address to search for
    - To: Recipient email address to search for
    - Start: Start time for query
    - End: End time for query
    
    Returns: Adds rows to $script:Rows via Add-TraceRow
#>
function Get-SendGridRows {
    param(
        [string]$From,
        [string]$To,
        [datetime]$Start,
        [datetime]$End
    )
    
    # Get SendGrid API key from environment variables
    $apiKey = $env:SENDGRID_API_KEY
    
    if (-not $apiKey) {
        [void]$script:Notes.Add("SendGrid -- API key not configured in environment variables")
        return
    }
    
    try {
        # SendGrid API endpoint for message activity
        $endpoint = "https://api.sendgrid.com/v3/messages"
        
        # Build query parameters
        $query = ""
        if ($From) { $query += "from:$From " }
        if ($To) { $query += "to:$To " }
        
        # Add time range (SendGrid uses Unix timestamps)
        $startUnix = [int][double]::Parse((Get-Date $Start -UFormat %s))
        $endUnix = [int][double]::Parse((Get-Date $End -UFormat %s))
        $query += "before_timestamp:$endUnix after_timestamp:$startUnix"
        
        # Build headers with API key
        $headers = @{
            "Authorization" = "Bearer $apiKey"
            "Content-Type" = "application/json"
        }
        
        # Build request body
        $body = @{
            query = $query
            limit = 100
        } | ConvertTo-Json
        
        # Make API call
        $response = Invoke-RestMethod -Uri $endpoint -Method Post -Headers $headers -Body $body -ErrorAction Stop
        
        # Parse response and add rows
        if ($response.messages) {
            foreach ($msg in $response.messages) {
                # Extract sender and recipient from message
                $sender = $msg.from.email
                $recipient = $msg.to[0].email
                $subject = $msg.subject
                $timestamp = [datetime]::UnixEpoch.AddSeconds($msg.timestamp)
                
                # Determine status
                $status = "SENT"
                if ($msg.status -eq "bounce") { $status = "BOUNCE" }
                elseif ($msg.status -eq "dropped") { $status = "DROPPED" }
                elseif ($msg.status -eq "delivered") { $status = "DELIVERED" }
                elseif ($msg.status -eq "deferred") { $status = "DEFERRED" }
                elseif ($msg.status -eq "click") { $status = "CLICK" }
                elseif ($msg.status -eq "open") { $status = "OPEN" }
                
                Add-TraceRow -When $timestamp -Source 'SendGrid' `
                    -Sender $sender -Recipient $recipient `
                    -Subject $subject -Outcome $status `
                    -Detail "SendGrid API" -SizeBytes 0
            }
            [void]$script:Notes.Add("SendGrid -- Query completed successfully ($(@($response.messages).Count) messages)")
        } else {
            [void]$script:Notes.Add("SendGrid -- No messages found")
        }
        
    } catch {
        $errorMsg = $_.Exception.Message
        if ($_.Exception.Response.StatusCode -eq 401) {
            [void]$script:Notes.Add("SendGrid -- Authentication failed (401). Check API key.")
        } else {
            [void]$script:Notes.Add("SendGrid -- Query failed: $errorMsg")
        }
    }
}

# Export-ModuleMember THROWS WHEN THIS FILE IS DOT-SOURCED, which is how it is
# actually loaded -- AvePoint-Dashboard does '. (Join-Path $PSScriptRoot ...)'.
# The cmdlet is only legal inside a module, so calling it unguarded turns a
# working library into a load-time error. Dot-sourcing already puts every
# function in the caller's scope, so there is nothing to export in that case.
if ($MyInvocation.MyCommand.ScriptBlock.Module) {
    Export-ModuleMember -Function @('Get-SendGridRows')
}
