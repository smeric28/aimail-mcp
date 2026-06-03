<#
.SYNOPSIS
    Add the full Microsoft 365 connector Graph scopes (and the hosted redirect
    URI) to an EXISTING Entra app registration, then grant admin consent.

.DESCRIPTION
    The FBI-MCP-O365 connector mirrors the Microsoft 365 connector surface with
    write access: mail, calendar, contacts, OneDrive/SharePoint files, and To
    Do tasks. This script patches an existing app registration to request all
    of the required delegated scopes and adds the remote OAuth callback URL.

    Run scripts/setup-azure-ad-graph.ps1 first only if the app does not exist.

.EXAMPLE
    ./update-azure-app-graph-scopes.ps1 -AppId <application-client-id> -PublicBaseUrl https://fbi-mcp-o365.fireballz.ai
#>

param(
    [Parameter(Mandatory = $true)][string]$AppId,
    [Parameter(Mandatory = $false)][string]$PublicBaseUrl = "https://fbi-mcp-o365.fireballz.ai",
    [Parameter(Mandatory = $false)][switch]$GrantAdminConsent
)

$ErrorActionPreference = "Stop"

Import-Module Microsoft.Graph.Applications -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Connect-MgGraph -Scopes "Application.ReadWrite.All", "DelegatedPermissionGrant.ReadWrite.All"

$graphResourceId = "00000003-0000-0000-c000-000000000000"  # Microsoft Graph

# Delegated scopes the connector needs (must match mcp-server/src/config/scopes.ts).
$scopeNames = @(
    "User.Read",
    "Mail.ReadWrite",
    "Mail.ReadWrite.Shared",
    "Mail.Send",
    "Mail.Send.Shared",
    "Calendars.ReadWrite",
    "Calendars.ReadWrite.Shared",
    "Contacts.ReadWrite",
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
    "Tasks.ReadWrite",
    "Chat.ReadWrite",
    "People.Read"
)

# Resolve scope name -> permission ID dynamically from the Graph service
# principal so we never depend on hard-coded (and possibly stale) GUIDs.
Write-Host "Resolving delegated permission IDs from the Graph service principal..." -ForegroundColor Cyan
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphResourceId'"
$scopeMap = @{}
foreach ($s in $graphSp.Oauth2PermissionScopes) { $scopeMap[$s.Value] = $s.Id }

$resourceAccess = @()
foreach ($name in $scopeNames) {
    if (-not $scopeMap.ContainsKey($name)) { throw "Graph delegated scope '$name' not found." }
    $resourceAccess += @{ Id = $scopeMap[$name]; Type = "Scope" }
}

$app = Get-MgApplication -Filter "appId eq '$AppId'"
if (-not $app) { throw "App registration with appId $AppId not found." }

Write-Host "Updating required resource access (delegated Graph scopes)..." -ForegroundColor Cyan
Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @(@{
    ResourceAppId  = $graphResourceId
    ResourceAccess = $resourceAccess
})

# Ensure the hosted OAuth callback redirect URI is present.
$callback = "$($PublicBaseUrl.TrimEnd('/'))/auth/callback"
$existing = @($app.Web.RedirectUris)
if ($existing -notcontains $callback) {
    Write-Host "Adding redirect URI: $callback" -ForegroundColor Cyan
    Update-MgApplication -ApplicationId $app.Id -Web @{ RedirectUris = ($existing + $callback) }
}

Write-Host "`nConfigured delegated scopes:" -ForegroundColor Green
$scopeNames | Sort-Object | ForEach-Object { Write-Host "  - $_" }

if ($GrantAdminConsent) {
    Write-Host "`nGranting admin consent..." -ForegroundColor Cyan
    Write-Host "Open this URL signed in as a tenant admin to grant consent for all users:" -ForegroundColor Yellow
    $ctx = Get-MgContext
    Write-Host "https://login.microsoftonline.com/$($ctx.TenantId)/adminconsent?client_id=$AppId"
} else {
    Write-Host "`nRe-run with -GrantAdminConsent (or use the Azure Portal > API permissions > Grant admin consent)." -ForegroundColor Yellow
}

Disconnect-MgGraph
