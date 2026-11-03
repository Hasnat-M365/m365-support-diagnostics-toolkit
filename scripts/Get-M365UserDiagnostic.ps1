param(
    [Parameter(Mandatory)]
    [string]$UserPrincipalName
)

# Define required Microsoft Graph permissions
$requiredScopes = @(
    "User.Read.All",
    "UserAuthenticationMethod.Read.All"
)

# Check current Microsoft Graph connection
$context = Get-MgContext

if (-not $context) {

    Write-Host "Not connected to Microsoft Graph." -ForegroundColor Yellow
    Write-Host "Connecting to Microsoft Graph..."

    Connect-MgGraph -Scopes $requiredScopes
}
else {

    $missingScopes = $requiredScopes |
        Where-Object { $_ -notin $context.Scopes }

    if ($missingScopes) {

        Write-Host "Microsoft Graph connection is missing required permissions." -ForegroundColor Yellow
        Write-Host "Reconnecting with required permissions..."

        Connect-MgGraph -Scopes $requiredScopes
    }
    else {
        Write-Host "[PASS] Microsoft Graph connection verified." -ForegroundColor Green
    }
}

# Retrieve the user from Microsoft Graph
try {
    $user = Get-MgUser `
        -UserId $UserPrincipalName `
        -Property Id,DisplayName,UserPrincipalName,AccountEnabled,UsageLocation,AssignedLicenses `
        -ErrorAction Stop
}
catch {
    Write-Host ""
    Write-Host "[ERROR] User not found in Microsoft Entra ID." -ForegroundColor Red
    Write-Host "Target: $UserPrincipalName"
    exit 1
}

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

    $tenantSkus = Get-MgSubscribedSku -All

    foreach ($license in $user.AssignedLicenses) {

        $matchedSku = $tenantSkus |
            Where-Object { $_.SkuId -eq $license.SkuId }

        if ($matchedSku) {
            Write-Host "[INFO] License: $($matchedSku.SkuPartNumber)"
        }
        else {
            Write-Host "[INFO] Unknown SKU: $($license.SkuId)"
        }
    }
}
else {
    Write-Host "[WARNING] No licenses are assigned." -ForegroundColor Yellow
}

# Retrieve group memberships
Write-Host ""
Write-Host "GROUP MEMBERSHIP" -ForegroundColor Cyan
Write-Host "----------------"

$memberships = Get-MgUserMemberOf -UserId $UserPrincipalName -All

$groupIds = $memberships |
    Where-Object {
        $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group'
    } |
    Select-Object -ExpandProperty Id

if ($groupIds.Count -gt 0) {

    Write-Host "[INFO] User belongs to $($groupIds.Count) group(s)." -ForegroundColor Cyan

    foreach ($groupId in $groupIds) {
        $group = Get-MgGroup -GroupId $groupId
        Write-Host "[INFO] Group: $($group.DisplayName)"
    }
}
else {
    Write-Host "[INFO] No group memberships found."
}

# Retrieve authentication methods
Write-Host ""
Write-Host "AUTHENTICATION METHODS" -ForegroundColor Cyan
Write-Host "----------------------"

try {
    $authMethods = Get-MgUserAuthenticationMethod `
        -UserId $UserPrincipalName `
        -ErrorAction Stop

    if ($authMethods.Count -gt 0) {

        Write-Host "[PASS] $($authMethods.Count) authentication method(s) registered." -ForegroundColor Green

        foreach ($method in $authMethods) {

            $methodType = $method.AdditionalProperties.'@odata.type'

            switch ($methodType) {

                "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                    Write-Host "[INFO] Microsoft Authenticator"
                }

                "#microsoft.graph.phoneAuthenticationMethod" {
                    Write-Host "[INFO] Phone authentication"
                }

                "#microsoft.graph.fido2AuthenticationMethod" {
                    Write-Host "[INFO] FIDO2 security key"
                }

                "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                    Write-Host "[INFO] Windows Hello for Business"
                }

                "#microsoft.graph.softwareOathAuthenticationMethod" {
                    Write-Host "[INFO] Software OATH token"
                }

                "#microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                    Write-Host "[INFO] Temporary Access Pass"
                }

                "#microsoft.graph.passwordAuthenticationMethod" {
                    Write-Host "[INFO] Password"
                }

                default {
                    Write-Host "[INFO] Other authentication method: $methodType"
                }
            }
        }
    }
    else {
        Write-Host "[WARNING] No authentication methods found." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "[ERROR] Unable to retrieve authentication methods." -ForegroundColor Red
    Write-Host "[INFO] Verify Microsoft Graph permissions."
}