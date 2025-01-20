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

Invoke-Monkey365
Import-Module C:\_Cultura\_Tools\monkey365 -Force
Invoke-Monkey365 -Environment AzurePublic -IncludeEntraID -OutDir "C:\_Cultura\_Tmp" -ExportTo JSON,HTML -Compress -Threads 4 -AuditorName "Stephane Giraud" -ForceMSALDesktop -Instance Azure -WriteLog -Analysis All -Subscriptions All


Invoke-Monkey365 -Environment AzurePublic -IncludeEntraID -OutDir "C:\_Cultura\_Tmp" -ExportTo JSON,HTML -Compress -Threads 6 -AuditorName "Stephane Giraud" -ForceAuth -Instance Microsoft365 -WriteLog -Analysis ExchangeOnline,Microsoft365,MicrosoftTeams,Purview,SharePointOnline


