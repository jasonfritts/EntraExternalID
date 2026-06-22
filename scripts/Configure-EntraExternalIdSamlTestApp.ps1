[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$TenantDomain,

    [Parameter(Mandatory = $true)]
    [string]$ExternalIdDomain,

    [Parameter()]
    [string]$AppDisplayName = "SAML SP Tester",

    [Parameter()]
    [string]$BaseUrl = "http://localhost:3000",

    [Parameter()]
    [switch]$UseDeviceCode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "[STEP] $Message" -ForegroundColor Cyan
}

function Ensure-Module {
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Step "Installing PowerShell module: $Name"
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module $Name -ErrorAction Stop
}

function Trim-TrailingSlash {
    param([Parameter(Mandatory = $true)][string]$Value)

    if ($Value.EndsWith("/")) {
        return $Value.Substring(0, $Value.Length - 1)
    }

    return $Value
}

function Get-ThumbprintFromCustomKeyIdentifier {
    param([Parameter(Mandatory = $true)]$CustomKeyIdentifier)

    $bytes = $null

    if ($CustomKeyIdentifier -is [byte[]]) {
        $bytes = $CustomKeyIdentifier
    }
    elseif ($CustomKeyIdentifier -is [string]) {
        if ([string]::IsNullOrWhiteSpace($CustomKeyIdentifier)) {
            return ""
        }

        $normalized = $CustomKeyIdentifier.Trim()
        try {
            $bytes = [System.Convert]::FromBase64String($normalized)
        }
        catch {
            $bytes = $null
        }
    }

    if (-not $bytes -or $bytes.Length -eq 0) {
        return ""
    }

    return ([System.BitConverter]::ToString($bytes)).Replace("-", "")
}

$BaseUrl = Trim-TrailingSlash -Value $BaseUrl

$requiredScopes = @(
    "Application.ReadWrite.All",
    "Directory.ReadWrite.All"
)

Write-Step "Ensuring Microsoft.Graph modules are installed"
Ensure-Module -Name Microsoft.Graph.Authentication
Ensure-Module -Name Microsoft.Graph.Applications
Ensure-Module -Name Microsoft.Graph.Identity.SignIns

Write-Step "Connecting to Microsoft Graph"
if ($UseDeviceCode) {
    Connect-MgGraph -TenantId $TenantId -Scopes $requiredScopes -UseDeviceCode -NoWelcome
}
else {
    Connect-MgGraph -TenantId $TenantId -Scopes $requiredScopes -NoWelcome
}

Write-Step "Creating app registration: $AppDisplayName"
$app = New-MgApplication `
    -DisplayName $AppDisplayName `
    -SignInAudience "AzureADMyOrg"

Write-Step "Creating service principal for the app"
$sp = New-MgServicePrincipal -AppId $app.AppId

$entityId = "urn:saml-sp-test:$($app.AppId)"
$replyUrl = "$BaseUrl/acs"
$signOnUrl = "$BaseUrl/"
$metadataUrl = "https://$ExternalIdDomain/$TenantDomain/federationmetadata/2007-06/federationmetadata.xml?appid=$($app.AppId)"

Write-Step "Configuring Identifier URI (Entity ID) on app registration"
Update-MgApplication -ApplicationId $app.Id -IdentifierUris @($entityId)

Write-Step "Configuring URLs on app registration"
$appPatch = @{
    web = @{
        redirectUris = @($replyUrl)
        homePageUrl  = $signOnUrl
    }
}

Invoke-MgGraphRequest `
    -Method PATCH `
    -Uri "https://graph.microsoft.com/v1.0/applications/$($app.Id)" `
    -ContentType "application/json" `
    -Body ($appPatch | ConvertTo-Json -Depth 10)

Write-Step "Configuring enterprise app SAML SSO settings"
$spPatch = @{
    preferredSingleSignOnMode = "saml"
    appRoleAssignmentRequired = $false
    samlSingleSignOnSettings = @{
        relayState = ""
    }
}

Invoke-MgGraphRequest `
    -Method PATCH `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)" `
    -ContentType "application/json" `
    -Body ($spPatch | ConvertTo-Json -Depth 10)

Write-Step "Ensuring a SAML token-signing certificate exists"
$spAfterPatch = Get-MgServicePrincipal -ServicePrincipalId $sp.Id -Property "id,displayName,keyCredentials,preferredTokenSigningKeyThumbprint"

$existingSigningKey = $null
if ($spAfterPatch.KeyCredentials) {
    $existingSigningKey = $spAfterPatch.KeyCredentials |
        Where-Object { $_.Type -eq "AsymmetricX509Cert" -and $_.Usage -eq "Verify" -and $_.CustomKeyIdentifier } |
    Sort-Object EndDateTime -Descending |
    Select-Object -First 1
}

if (-not $existingSigningKey) {
    $certEndDate = (Get-Date).ToUniversalTime().AddYears(2).ToString("o")
    $certRequestBody = @{
        displayName = "CN=$AppDisplayName"
        endDateTime = $certEndDate
    }

    Invoke-MgGraphRequest `
        -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/addTokenSigningCertificate" `
        -ContentType "application/json" `
        -Body ($certRequestBody | ConvertTo-Json -Depth 10)

    # Graph response shape can vary, so always re-read keyCredentials after certificate creation.
    $spAfterCert = Get-MgServicePrincipal -ServicePrincipalId $sp.Id -Property "keyCredentials"
    $existingSigningKey = $spAfterCert.KeyCredentials |
        Where-Object { $_.Type -eq "AsymmetricX509Cert" -and $_.Usage -eq "Verify" -and $_.CustomKeyIdentifier } |
        Sort-Object EndDateTime -Descending |
        Select-Object -First 1

    if (-not $existingSigningKey) {
        throw "Token signing certificate was requested but no usable signing key was found on the service principal."
    }
}

$thumbprint = Get-ThumbprintFromCustomKeyIdentifier -CustomKeyIdentifier $existingSigningKey.customKeyIdentifier

if ([string]::IsNullOrWhiteSpace($thumbprint)) {
    throw "Unable to determine token-signing certificate thumbprint from Graph response."
}

if ($thumbprint.Length -lt 1 -or $thumbprint.Length -gt 256) {
    throw "Computed thumbprint has invalid length: $($thumbprint.Length)."
}

if (-not $spAfterPatch.PreferredTokenSigningKeyThumbprint -or $spAfterPatch.PreferredTokenSigningKeyThumbprint -ne $thumbprint) {
    Write-Step "Setting preferred SAML token-signing certificate thumbprint"
    $thumbprintPatch = @{
        preferredTokenSigningKeyThumbprint = $thumbprint
    }

    Invoke-MgGraphRequest `
        -Method PATCH `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)" `
        -ContentType "application/json" `
        -Body ($thumbprintPatch | ConvertTo-Json -Depth 10)
}

Write-Step "Retrieving updated app and service principal"
$updatedApp = Get-MgApplication -ApplicationId $app.Id
$updatedSp = Get-MgServicePrincipal -ServicePrincipalId $sp.Id -Property "id,preferredTokenSigningKeyThumbprint"

Write-Host ""
Write-Host "Configuration complete." -ForegroundColor Green
Write-Host ""
Write-Host "Use these values in Entra portal and test app:" -ForegroundColor Yellow
Write-Host "  Client ID:            $($updatedApp.AppId)"
Write-Host "  App Object ID:        $($updatedApp.Id)"
Write-Host "  Enterprise App ID:    $($updatedSp.Id)"
Write-Host "  Identifier (Entity):  $entityId"
Write-Host "  Reply URL (ACS):      $replyUrl"
Write-Host "  Sign-on URL:          $signOnUrl"
Write-Host "  Metadata URL:         $metadataUrl"
Write-Host "  Signing Thumbprint:   $($updatedSp.PreferredTokenSigningKeyThumbprint)"
Write-Host ""
Write-Host "Enter these in the local SAML test app UI:" -ForegroundColor Yellow
Write-Host "  domain:               $ExternalIdDomain"
Write-Host "  tenantDomain:         $TenantDomain"
Write-Host "  clientId:             $($updatedApp.AppId)"
Write-Host ""
Write-Host "Notes:" -ForegroundColor Yellow
Write-Host "  - User assignment required is set to false for easier testing."
Write-Host "  - Signature validation is still handled by your app logic (inspection app currently does not validate)."
Write-Host ""
Write-Host "Done." -ForegroundColor Green
