<#
.SYNOPSIS
#################

.DESCRIPTION
Long description

.EXAMPLE
An example

.NOTES
General notes
#>

Connect-AzureAD

$aadApplication = New-AzureADApplication -DisplayName "ScubaGear" -IdentifierUris "http://mtsdemoapp.contoso.com" -HomePage "http://mtsdemo.contoso.com"


#Seetings API Permission to Application
-Description


#Add Secret to Application

#Add Certificate to Application


# Connectez-vous à votre compte Azure
Connect-MgGraph -Scopes "Application.ReadWrite.All", "Directory.ReadWrite.All"

# Définissez des variables pour le nom de l'application, l'URI de redirection et la description
$appName = "ScubaGear"
$redirectUri = "https://localhost"
$appDescription = "Scan Scuba Gear - CISAGOV"

# Créez l'application Azure AD avec une description
$app = New-MgApplication -DisplayName $appName -Web @{RedirectUris = @($redirectUri)} -Description $appDescription

# Créez un principal de service pour l'application
$sp = New-MgServicePrincipal -AppId $app.AppId

# Définissez les permissions API requises (exemple avec Microsoft Graph API)
$graphApiId = "00000003-0000-0000-c000-000000000000" # ID de l'API Microsoft Graph
$scope = "User.Read" # Permission requise

# Créez une requête pour ajouter les permissions API à l'application
$requiredResourceAccess = @{
    ResourceAppId = $graphApiId
    ResourceAccess = @(@{Id = (Get-MgServicePrincipalAppRole -ServicePrincipalId $graphApiId | Where-Object {$_.Value -eq $scope}).Id; Type = "Scope"})
}

# Affectez les permissions API à l'application
Update-MgApplication -ApplicationId $app.AppId -RequiredResourceAccess @($requiredResourceAccess)

# Consentement administrateur (requiert des permissions administrateur)
Update-MgServicePrincipalDelegatedPermissionGrant -ServicePrincipalId $sp.Id -Scope $scope

Write-Output "L'application '$appName' a été créée avec succès avec la description '$appDescription' et les permissions API ont été affectées."



$cert = New-SelfSignedCertificate -CertStoreLocation "cert:\CurrentUser\My" `
  -Subject "CN=exampleappScriptCert" `
  -KeySpec KeyExchange
$keyValue = [System.Convert]::ToBase64String($cert.GetRawCertData())

$sp = New-AzADServicePrincipal -DisplayName exampleapp `
  -CertValue $keyValue `
  -EndDate $cert.NotAfter `
  -StartDate $cert.NotBefore
Sleep 20
New-AzRoleAssignment -RoleDefinitionName Reader -ServicePrincipalName $sp.AppId

