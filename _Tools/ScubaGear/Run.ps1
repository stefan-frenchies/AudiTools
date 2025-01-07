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

# Requires Version 5
Clear-Host
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$MyScriptName = [String]([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path))

$ToolName = "ScubaGear"
$ClientName = "cultura"

If ( ! (Test-Path -Path "$PSScriptRoot\..\_Tmp\$ToolName") ) {New-Item "$PSScriptRoot\..\_Tmp" -Name "$ToolName" -ItemType Directory -Force | Out-Null}
$DataFolder = (Get-Item -Path "$PSScriptRoot\..\_Tmp\$ToolName\").FullName

$NeededModules = @("ScubaGear")
$NeededModules | ForEach-Object {
    If (-not (Get-InstalledModule -Name $_)) {
        Write-Host "Installation Module $_" -ForegroundColor Cyan
        Install-Module -Name $_ -Force}
}
Initialize-SCuBA

$DateStamp = Get-Date -Format "yyyyMMdd"
Invoke-SCuBA -ProductNames * `
-LogIn $true  `
-DisconnectOnExit  `
-OutFolderName "$ToolName" `
-OutPath "$DataFolder" `
-OutJsonFileName "$ClientName-$DateStamp" `
-OutReportName "$ClientName-$DateStamp" `
-OutRegoFileName "$ClientName" `
-MergeJson `
-Quiet
#-OPAPath "C:\Users\admin-sgiraud-t0\.scubagear\Tools" `

#Compress Files... and copy to Export?

