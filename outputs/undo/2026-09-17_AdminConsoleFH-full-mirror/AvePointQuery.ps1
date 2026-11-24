#Requires -Version 5.1
<#
    AvePointQuery.ps1
    Query functions for AvePoint Graph API integration with AdminConsole
    
    This module provides query functions for AvePoint Executive Dashboard
    and other AvePoint capabilities.
    
    Credentials stored in environment variables:
    - AVEPOINT_CLIENT_ID
    - AVEPOINT_CLIENT_SECRET
    
    Author: Alan Smith (with Glenn Armbruster's AdminConsole)
    Date: August 27, 2026
#>

<#
    Get-AvePointAccessToken
    Obtain access token from AvePoint Graph API
    
    Returns: Access token or $null if failed
#>
function Get-AvePointAccessToken {
    param()
    
    $clientId = $env:AVEPOINT_CLIENT_ID
    $clientSecret = $env:AVEPOINT_CLIENT_SECRET
    
    if (-not $clientId -or -not $clientSecret) {
        Write-Error "AvePoint credentials not configured in environment variables"
        return $null
    }
    
    try {
        $endpoint = "https://api.avepoint.com/oauth/token"
        
        $body = @{
            grant_type = "client_credentials"
            client_id = $clientId
            client_secret = $clientSecret
        }
        
        $response = Invoke-RestMethod -Uri $endpoint -Method Post -Body $body -ErrorAction Stop
        
        if ($response.access_token) {
            return $response.access_token
        } else {
            Write-Error "Failed to obtain AvePoint access token"
            return $null
        }
    } catch {
        Write-Error "AvePoint authentication failed: $($_.Exception.Message)"
        return $null
    }
}

<#
    Get-AvePointExecutiveDashboard
    Query AvePoint Executive Dashboard data
    
    Returns: Dashboard data or $null if failed
#>
function Get-AvePointExecutiveDashboard {
    param()
    
    $token = Get-AvePointAccessToken
    if (-not $token) { return $null }
    
    try {
        $endpoint = "https://api.avepoint.com/v1/executive-dashboard"
        
        $headers = @{
            "Authorization" = "Bearer $token"
            "Content-Type" = "application/json"
        }
        
        $response = Invoke-RestMethod -Uri $endpoint -Method Get -Headers $headers -ErrorAction Stop
        
        return $response
    } catch {
        Write-Error "Failed to retrieve AvePoint Executive Dashboard: $($_.Exception.Message)"
        return $null
    }
}

<#
    Get-AvePointTenants
    Query AvePoint managed tenants
    
    Returns: List of tenants or $null if failed
#>
function Get-AvePointTenants {
    param()
    
    $token = Get-AvePointAccessToken
    if (-not $token) { return $null }
    
    try {
        $endpoint = "https://api.avepoint.com/v1/tenants"
        
        $headers = @{
            "Authorization" = "Bearer $token"
            "Content-Type" = "application/json"
        }
        
        $response = Invoke-RestMethod -Uri $endpoint -Method Get -Headers $headers -ErrorAction Stop
        
        return $response.tenants
    } catch {
        Write-Error "Failed to retrieve AvePoint tenants: $($_.Exception.Message)"
        return $null
    }
}

<#
    Get-AvePointComplianceStatus
    Query AvePoint compliance status for a tenant
    
    Parameters:
    - TenantId: AvePoint tenant ID
    
    Returns: Compliance status or $null if failed
#>
function Get-AvePointComplianceStatus {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TenantId
    )
    
    $token = Get-AvePointAccessToken
    if (-not $token) { return $null }
    
    try {
        $endpoint = "https://api.avepoint.com/v1/tenants/$TenantId/compliance"
        
        $headers = @{
            "Authorization" = "Bearer $token"
            "Content-Type" = "application/json"
        }
        
        $response = Invoke-RestMethod -Uri $endpoint -Method Get -Headers $headers -ErrorAction Stop
        
        return $response
    } catch {
        Write-Error "Failed to retrieve AvePoint compliance status: $($_.Exception.Message)"
        return $null
    }
}

<#
    Get-AvePointSites
    Query AvePoint managed SharePoint sites
    
    Parameters:
    - TenantId: AvePoint tenant ID (optional)
    
    Returns: List of sites or $null if failed
#>
function Get-AvePointSites {
    param(
        [string]$TenantId
    )
    
    $token = Get-AvePointAccessToken
    if (-not $token) { return $null }
    
    try {
        if ($TenantId) {
            $endpoint = "https://api.avepoint.com/v1/tenants/$TenantId/sites"
        } else {
            $endpoint = "https://api.avepoint.com/v1/sites"
        }
        
        $headers = @{
            "Authorization" = "Bearer $token"
            "Content-Type" = "application/json"
        }
        
        $response = Invoke-RestMethod -Uri $endpoint -Method Get -Headers $headers -ErrorAction Stop
        
        return $response.sites
    } catch {
        Write-Error "Failed to retrieve AvePoint sites: $($_.Exception.Message)"
        return $null
    }
}

# Export-ModuleMember THROWS WHEN THIS FILE IS DOT-SOURCED, which is how it is
# actually loaded -- AvePoint-Dashboard does '. (Join-Path $PSScriptRoot ...)'.
# The cmdlet is only legal inside a module, so calling it unguarded turns a
# working library into a load-time error. Dot-sourcing already puts every
# function in the caller's scope, so there is nothing to export in that case.
if ($MyInvocation.MyCommand.ScriptBlock.Module) {
    Export-ModuleMember -Function @(
        'Get-AvePointAccessToken',
        'Get-AvePointExecutiveDashboard',
        'Get-AvePointTenants',
        'Get-AvePointComplianceStatus',
        'Get-AvePointSites'
    )
}
