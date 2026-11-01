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
$user = Get-MgUser -UserId $UserPrincipalName -Property Id,DisplayName,UserPrincipalName,AccountEnabled,UsageLocation,AssignedLicenses

# Display diagnostic header
Write-Host ""
Write-Host "Microsoft 365 User Diagnostic" -ForegroundColor Cyan
Write-Host "=============================" -ForegroundColor Cyan
Write-Host "Target: $($user.UserPrincipalName)"
Write-Host ""

# Display basic user information
Write-Host "IDENTITY" -ForegroundColor Cyan
Write-Host "--------"

Write-Host "[PASS] User found: $($user.DisplayName)" -ForegroundColor Green
Write-Host "[INFO] User ID: $($user.Id)"

# Evaluate account status
if ($user.AccountEnabled -eq $true) {
    Write-Host "[PASS] Account is enabled." -ForegroundColor Green
}
else {
    Write-Host "[WARNING] Account is disabled." -ForegroundColor Yellow
}

# Evaluate usage location
if ([string]::IsNullOrWhiteSpace($user.UsageLocation)) {
    Write-Host "[WARNING] Usage location is missing." -ForegroundColor Yellow
}
else {
    Write-Host "[PASS] Usage location: $($user.UsageLocation)" -ForegroundColor Green
}

# Evaluate licensing
Write-Host ""
Write-Host "LICENSING" -ForegroundColor Cyan
Write-Host "---------"

$licenseCount = $user.AssignedLicenses.Count

if ($licenseCount -gt 0) {
    Write-Host "[PASS] $licenseCount license(s) assigned." -ForegroundColor Green

    foreach ($license in $user.AssignedLicenses) {
        Write-Host "[INFO] SKU ID: $($license.SkuId)"
    }
}
else {
    Write-Host "[WARNING] No licenses are assigned." -ForegroundColor Yellow
}

# Retrieve group memberships
Write-Host ""
Write-Host "GROUP MEMBERSHIP" -ForegroundColor Cyan
Write-Host "----------------"

$groups = Get-MgUserMemberOf -UserId $UserPrincipalName -All

if ($groups.Count -gt 0) {
    Write-Host "[INFO] User belongs to $($groups.Count) directory object(s)." -ForegroundColor Cyan
}
else {
    Write-Host "[INFO] No group memberships found."
}