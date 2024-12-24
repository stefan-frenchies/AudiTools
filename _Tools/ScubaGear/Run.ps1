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


$MyScriptName = [String]([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path))

$ToolName = "ScubaGear"
$ClientName = "cultura"

$datafolder = (New-Item "$PSScriptRoot\..\_Tmp" -Name "$ToolName" -ItemType Directory -Force | Out-Null).FullName



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
-OutFolderName "ScubaGear" `
-OutPath "$datafolder" `
-OutJsonFileName "$ClientName-$DateStamp" `
-OutReportName "$ClientName-$DateStamp" `
-OutRegoFileName "$ClientName" `
-MergeJson `
-Quiet
#-OPAPath "C:\Users\admin-sgiraud-t0\.scubagear\Tools" `

