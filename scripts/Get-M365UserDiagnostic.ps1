param(
    [Parameter(Mandatory)]
    [string]$UserPrincipalName
)

# Check Microsoft Graph connection
$context = Get-MgContext

if (-not $context) {
    Write-Host "Not connected to Microsoft Graph." -ForegroundColor Yellow
    Write-Host "Connecting with User.Read.All permission..."
    Connect-MgGraph -Scopes "User.Read.All"
}

# Retrieve the user from Microsoft Graph
$user = Get-MgUser -UserId $UserPrincipalName -Property Id,DisplayName,UserPrincipalName,AccountEnabled

# Display basic diagnostic information
$user | Select-Object Id, DisplayName, UserPrincipalName, AccountEnabled | Format-List