<#
.SYNOPSIS
    Test all MCP server workflows
#>

Write-Host "=== Microsoft Email MCP Server Test Suite ===`n" -ForegroundColor Cyan

$Tests = @(
    @{
        Name = "Search Emails"
        Query = "Find unread emails from this week"
        Expected = "Returns matching emails"
    },
    @{
        Name = "Read Email"
        Query = "Get the details of the most recent email"
        Expected = "Returns email content"
    },
    @{
        Name = "Create Draft"
        Query = "Create a draft in aiDrafts folder to test@example.com about Test Subject"
        Expected = "Draft created, no send"
    }
)

foreach ($Test in $Tests) {
    Write-Host "Test: $($Test.Name)" -ForegroundColor Yellow
    Write-Host "  Query: $($Test.Query)"
    Write-Host "  Expected: $($Test.Expected)`n"
}

Write-Host "`nManual verification steps:" -ForegroundColor Cyan
Write-Host "1. Check aiDrafts folder in Outlook for new draft"
Write-Host "2. Verify draft appears with correct content"
Write-Host "3. Attempt to send - should require manual action"
Write-Host "4. Try to delete an email - should fail (read-only)"
"