<#
.SYNOPSIS
#################

.DESCRIPTION
Long description

.EXAMPLE
An example

.NOTES
General notes

https://github.com/cisagov/ScubaGear
https://cisagov.github.io/ScubaGear/

#>

$NeededModules = @()
$NeededModules =+ "ScubaGear"
#$NeededModules =+ "OPAforSCuBA"
$NeededModules | ForEach-Object {
    If (-not (Get-InstalledModule -Name $_)) { Install-Module -Name $_}
}

Initialize-SCuBA


#Create A Service Principal

#Add API permissions
#Create A Certificate
#associate the Certificate With rthe SPN
#Determining the TCertificate Thumbprint

#Add PowerPlatform Registration

Add-PowerAppsAccount -Endpoint prod -TenantID 22f22c70-de09-4d21-b82f-af8ad73391d9

New-PowerAppManagementApp -ApplicationId abcdef0123456789abcde01234566789 