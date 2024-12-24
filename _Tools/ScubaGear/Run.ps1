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
$TempFolder = "ScubaGear"
$datafolder = (New-Item "$PSScriptRoot\..\_Tmp" -Name "$TempFolder" -ItemType Directory -Force | Out-Null).FullName


$NeededModules = @("ScubaGear")
$NeededModules | ForEach-Object {
    If (-not (Get-InstalledModule -Name $_)) {
        Write-Host "Installation Module $_" -ForegroundColor Cyan
        Install-Module -Name $_ -Force}
}
Initialize-SCuBA

$TimeStamp = Get-Date -Format "yyyyMMdd"
Invoke-SCuBA -ProductNames * `
-LogIn $true  `
-DisconnectOnExit  `
-OutFolderName "ScubaGear" `
-OutPath "$datafolder" `
-OutJsonFileName "cultura-$TimeStamp" `
-OutReportName "cultura-$TimeStamp" `
-OutRegoFileName "cultura" `
-MergeJson `
-Quiet
#-OPAPath "C:\Users\admin-sgiraud-t0\.scubagear\Tools" `

