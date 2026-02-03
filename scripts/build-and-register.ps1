<#
.SYNOPSIS
    Build and register the MCP server with OpenCode
#>

$ErrorActionPreference = "Stop"

Set-Location -Path "$PSScriptRoot\..\mcp-server"

Write-Host "Installing dependencies..." -ForegroundColor Cyan
npm install

Write-Host "Building TypeScript..." -ForegroundColor Cyan
npm run build

Write-Host "`nMCP server built successfully!" -ForegroundColor Green
Write-Host "Location: $PSScriptRoot\..\mcp-server\dist\index.js`n"

# Copy config to OpenCode
Write-Host "Copying configuration to OpenCode..." -ForegroundColor Cyan
$opencodeConfigPath = "$env:USERPROFILE\.opencode\mcp-servers\microsoft-email.json"
Copy-Item "$PSScriptRoot\..\config\mcp-servers.json" $opencodeConfigPath -Force

Write-Host "`nConfiguration copied to: $opencodeConfigPath" -ForegroundColor Yellow
Write-Host "Edit the file to add your Azure AD credentials to the env section." -ForegroundColor Yellow

Write-Host "`n=== NEXT STEPS ===" -ForegroundColor Green
Write-Host "1. Edit: $opencodeConfigPath"
Write-Host "2. Add your Azure AD credentials to the env section"
Write-Host "3. Restart OpenCode"
Write-Host "4. Test with: 'Search my emails from this week'"