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



#https://github.com/silverhack/monkey365

#https://silverhack.github.io/monkey365/

#Cloning Repository
#https://github.com/silverhack/monkey365/archive/refs/heads/main.zip

#Unblock Module
Get-ChildItem -Recurse C:\_Cultura\_Tools\monkey365 | Unblock-File

#Import Module
Import-Module C:\_Cultura\_Tools\monkey365 -Force


#Creation Application

#Definition des permissions
#Entra : Global Reader, Security Reader
#MS365 :  Global Reader, Sharepoint Admin


#Directory.Read.All
#Policy.Read.All
#UserAuthenticationMethod.read.All
#Reader on All Subscriptions
#Sites.FullControl.All
#Excahnge.ManageAsApp
#Global Reader


#Creation d'un Certificat KeyVault
#Export du Certificat
#Inscription dans application.
<#
$param = @{
    Instance = 'Microsoft365';
    Analysis = 'SharePointOnline';
    PromptBehavior = 'SelectAccount';
    IncludeEntraID = $true;
    ExportTo = 'PRINT';
}
$assets = Invoke-Monkey365 @param
#>

Import-Module C:\_Cultura\_Tools\monkey365 -Force

$DateStamp = Get-Date -Format "yyyyMMdd"

$param = @{
    Instance = 'Azure';
    Analysis = 'All';
    PromptBehavior = 'SelectAccount';
    ExportTo = @("PRINT","JSON","HTML");
    Subscriptions = 'e0a33254-35e0-498f-a858-114a6df2d886'; #Techniques Infrastructures
    TenantID = '37ddb62e-1d49-42b1-aacc-e08f83d1253d' #Cultura
    #SaveProject = "C:\_Cultura\_Tmp\Monkey365-$DateStamp"
}

$myAzMonkey = Invoke-Monkey365 @param

$param = @{
    ClientId = '00000000-0000-0000-0000-000000000000';
    certificate = 'C:\monkey365\testapp.pfx';
    CertFilePassword = ("MySuperCertSecret" | ConvertTo-SecureString -AsPlainText -Force);
    Instance = 'Microsoft365';
    Analysis = 'SharePointOnline';
    Subscriptions = '00000000-0000-0000-0000-000000000000';
    TenantID = '00000000-0000-0000-0000-000000000000';
    ExportTo = @("JSON","HTML");

}





$params = @{
    Instance = 'Microsoft365';
    IncludeEntraID = $true;
#    DeviceCode = $true
    TenantID = '37ddb62e-1d49-42b1-aacc-e08f83d1253d'; #Cultura
    Analysis = "ExchangeOnline","SharePointOnline","Purview","MicrosoftTeams","Microsoft365";
    ExportTo = @("PRINT","JSON","HTML");
    Threads = 4;
    #SaveProject = 'C:\_Cultura\_Tmp\Monkey365';
    Compress = $true
}


$myMS365Monkey = Invoke-Monkey365 @params

<#
#TenantID = '00000000-0000-0000-00-000000000000';
#37ddb ----- 1253d # 2 autres?? cheops et claranet?

#Subscriptions = '00000000-0000-0000-0000-000000000000';
#XXXXX Subscriptions ... :(
#https://silverhack.github.io/monkey365/
-SaveProject
-Compress
-ForceMSALDesktop

-AllSubscriptions $true ### !!! attention au role necessaire sur toutes !!!
#>





