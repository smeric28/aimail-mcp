<#
.SYNOPSIS
    Creates Azure AD app registration using Microsoft Graph PowerShell
    
.DESCRIPTION
    This script creates an Azure AD app registration with necessary permissions
    for email access (read-only + drafts) using the modern Microsoft.Graph module
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$AppName = "OpenCode Email MCP"
)

$ErrorActionPreference = "Stop"

Write-Host "Installing Microsoft.Graph module if needed..." -ForegroundColor Cyan
try {
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}
catch {
    Write-Host "Microsoft.Graph module not found. Installing..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
    Import-Module Microsoft.Graph.Applications
    Import-Module Microsoft.Graph.Authentication
}

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Application.ReadWrite.All"

# Get tenant info
$context = Get-MgContext
$tenantId = $context.TenantId

Write-Host "Creating app registration: $AppName" -ForegroundColor Cyan

# Define required permissions
$requiredResourceAccess = @{
    ResourceAppId  = "00000003-0000-0000-c000-000000000000"  # Microsoft Graph
    ResourceAccess = @(
        @{
            Id   = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"  # User.Read
            Type = "Scope"
        },
        @{
            Id   = "570282fd-fa5c-430d-a7fd-fc8e98c4ab4c"  # Mail.Read
            Type = "Scope"
        },
        @{
            Id   = "e2a3a72e-5f79-4c64-b1b1-878b674786c9"  # Mail.ReadWrite
            Type = "Scope"
        }
    )
}

# Create the app registration
$app = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMyOrg" -RequiredResourceAccess $requiredResourceAccess -Web @{ RedirectUris = @("http://localhost:3000/callback") }

Write-Host "App created with ID: $($app.AppId)" -ForegroundColor Green

# Create client secret
$passwordCred = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
    DisplayName = "OpenCode MCP Secret"
}

# Output configuration
Write-Host "`n=== CONFIGURATION ===" -ForegroundColor Green
Write-Host "Application (Client) ID: $($app.AppId)"
Write-Host "Directory (Tenant) ID: $tenantId"
Write-Host "Client Secret: $($passwordCred.SecretText)"
Write-Host ""

# Save configuration (without secret)
$config = @{
    ApplicationId = $app.AppId
    TenantId      = $tenantId
    ClientSecret  = $passwordCred.SecretText
}

$configPath = "$PSScriptRoot\..\config\azure\app-config.json"
$config | ConvertTo-Json | Out-File -FilePath $configPath -Force

Write-Host "Configuration saved to: $configPath" -ForegroundColor Yellow
Write-Host "IMPORTANT: Store the Client Secret securely!" -ForegroundColor Red
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "1. Grant admin consent in Azure Portal for the app permissions"
Write-Host "2. Update C:\codebase\aimail\config\mcp-servers.json with these credentials"
Write-Host "3. Run .\build-and-register.ps1 to complete setup"

# Disconnect
Disconnect-MgGraph