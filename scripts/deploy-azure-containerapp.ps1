<#
.SYNOPSIS
    Build and deploy the FBI-MCP-O365 remote server to Azure Container Apps.

.DESCRIPTION
    Creates (or updates) an Azure Container App that hosts the MCP server over
    HTTPS so it can be added as a Claude custom connector on iPhone, desktop,
    and web. Uses `az containerapp up` to build from the Dockerfile and deploy.

    Prerequisites:
      - Azure CLI (`az`) logged in:  az login
      - The containerapp extension:  az extension add --name containerapp
      - An Entra app registration (client id + secret) with the Graph scopes
        from scripts/update-azure-app-graph-scopes.ps1 and admin consent granted.

.EXAMPLE
    ./deploy-azure-containerapp.ps1 `
        -ResourceGroup fbi-mcp-o365-rg `
        -Location eastus `
        -ClientId <app-id> -ClientSecret <secret> -TenantId <tenant-guid> `
        -PublicBaseUrl https://fbi-mcp-o365.fireballz.ai
#>

param(
    [Parameter(Mandatory = $true)][string]$ResourceGroup,
    [Parameter(Mandatory = $false)][string]$Location = "eastus",
    [Parameter(Mandatory = $false)][string]$AppName = "fbi-mcp-o365",
    [Parameter(Mandatory = $false)][string]$EnvName = "fbi-mcp-o365-env",
    [Parameter(Mandatory = $true)][string]$ClientId,
    [Parameter(Mandatory = $true)][string]$ClientSecret,
    [Parameter(Mandatory = $true)][string]$TenantId,
    [Parameter(Mandatory = $true)][string]$PublicBaseUrl
)

$ErrorActionPreference = "Stop"
$sourcePath = Join-Path $PSScriptRoot "..\mcp-server"

Write-Host "Ensuring resource group '$ResourceGroup' in $Location..." -ForegroundColor Cyan
az group create --name $ResourceGroup --location $Location | Out-Null

Write-Host "Deploying Container App '$AppName' (builds from Dockerfile)..." -ForegroundColor Cyan
# `az containerapp up` builds the image (via ACR) and creates the app + env.
az containerapp up `
    --name $AppName `
    --resource-group $ResourceGroup `
    --location $Location `
    --environment $EnvName `
    --source $sourcePath `
    --ingress external `
    --target-port 8080 | Out-Null

Write-Host "Setting environment variables and secrets..." -ForegroundColor Cyan
az containerapp secret set `
    --name $AppName --resource-group $ResourceGroup `
    --secrets "azure-client-secret=$ClientSecret" | Out-Null

az containerapp update `
    --name $AppName --resource-group $ResourceGroup `
    --set-env-vars `
        "MCP_TRANSPORT=http" `
        "PORT=8080" `
        "AZURE_CLIENT_ID=$ClientId" `
        "AZURE_TENANT_ID=$TenantId" `
        "PUBLIC_BASE_URL=$PublicBaseUrl" `
        "AZURE_CLIENT_SECRET=secretref:azure-client-secret" | Out-Null

$fqdn = az containerapp show --name $AppName --resource-group $ResourceGroup --query "properties.configuration.ingress.fqdn" -o tsv

Write-Host "`n=== DEPLOYED ===" -ForegroundColor Green
Write-Host "Container App FQDN: https://$fqdn"
Write-Host "MCP endpoint:       $PublicBaseUrl/mcp"
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Cyan
Write-Host "1. Point PublicBaseUrl ($PublicBaseUrl) at https://$fqdn (custom domain or use the FQDN directly)."
Write-Host "2. In the Entra app registration, add redirect URI: $PublicBaseUrl/auth/callback (type: Web)."
Write-Host "3. Add the connector in Claude:  Settings > Connectors > Add custom connector > $PublicBaseUrl/mcp"
