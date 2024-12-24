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

param (
    [Parameter(Mandatory=$false, Position=0)]
    [ValidateScript({Test-Path $_})]
    [string]$Search = (Get-Location).Path,
    [Parameter(Mandatory=$false, Position=1)]
    [ValidateRange(1, 9)]
    [Int]$Depth = 1,
    [Parameter(Mandatory=$false, Position=2)]
    [ValidateScript({ if ($ZipName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -eq -1) { return $true } else  { throw "The ZipName specified is not Valid" } })]
    [String]$ZipName = "AllMyScubaReports",
    [Parameter(Mandatory=$false, Position=0)]
    [ValidateScript({Test-Path $_})]
    [string]$Folder = (Get-Location).Path
)

$AppName = "Scuba"
$TimeStamp = "$((Get-Date).ToString("yyyyMMddTHHmmss"))"
try {
    "$TimeStamp : Launch $Search, $Depth, $ZipName,$Folder" | Out-File "$Folder\..\_Logs\$AppName-Collect.log" -Encoding utf8 -Force
}
catch {
    Write-Warning "Could Not Create LogFile, Please Set Folder Parameter to a writable Path"
    EXIT
}


$MyAppReportFiles = New-Object 'System.Collections.Generic.List[System.Object]'
 
Write-Verbose "Searching json in $Search For $Depth Level" 
Get-ChildItem -Path "$Search" -Depth $Depth -Name "*.json" | ForEach-Object {
    $MyFullName = "$Search\$_"
    Write-Verbose "Getting Infos About $MyFullName" 
    $obj = [PSCustomObject]@{
        FullName    = [String]("$MyFullName")
        FolderPath  = [String](([System.IO.Directory]::GetParent("$MyFullName")).FullName)
        BaseName    = [String]([System.IO.Path]::GetFileNameWithoutExtension("$MyFullName"))
        Extension   = [String]([System.IO.Path]::GetExtension("$MyFullName"))
    }
    $MyAppReportFiles.Add($Obj)
}


$SearchFrom = [String]((Get-Item -Path "$(($MyAppReportFiles | Sort-Object FolderPath | Select-Object -First 1).FolderPath)").BaseName)

"$TimeStamp : $AppName Found Files :" | Out-File "$Folder\..\_Logs\$AppName-Collect.log" -Encoding utf8 -Append
$MyAppReportFiles | Out-File "$Folder\..\_Logs\$AppName-Collect.log" -Encoding utf8 -Append
" !!!! Script Will Obfuscate your Folder Paths on the Export !!!!" | Out-File "$Folder\..\_Logs\$AppName-Collect.log" -Encoding utf8 -Append

$PotentialsOutputFolders = $MyAppReportFiles.FolderPath | Select-Object -Unique


If ($PotentialsOutputFolders.Count -ge 1) {
    Write-Host "Collecting $AppName Reports Files Found in $($PotentialsOutputFolders.Count) Folders" -ForegroundColor Cyan
    New-Item -Path "$Folder\..\_tmp" -Name "$AppName-tmp" -ItemType Directory | Out-Null
    $i=0
    $PotentialsOutputFolders | ForEach-Object {
        $MyFolder = $_
        $i=$i+1
        Write-Verbose "Copying $MyFolder Files to $i" 
        New-Item -Path "$Folder\..\_tmp\$AppName-tmp" -Name $i -ItemType Directory -Force| Out-Null
        ($MyAppReportFiles | Where-Object FolderPath -eq "$MyFolder").FullName | Copy-Item -Destination "$Folder\..\_tmp\$AppName-tmp\$i"
        "$MyFolder Files copied to $i" | Out-File "$Folder\..\_Logs\$AppName-Collect.log" -Encoding utf8 -Append
    }
    try {
        Write-Verbose "Compressing $Folder\..\_tmp\$AppName-tmp in $ZipName-$SearchFrom-$TimeStamp.zip" 
        Compress-Archive "$Folder\..\_tmp\$AppName-tmp\*" -DestinationPath "$Folder\..\_Exports\$ZipName-$SearchFrom-$TimeStamp.zip" -ErrorAction Stop
        "Temp folder $Folder\..\_tmp\$AppName-tmp Compressed to $Folder\..\_Exports\$ZipName-$SearchFrom-$TimeStamp.zip" | Out-File "$Folder\..\_Logs\$AppName-Collect.log" -Encoding utf8 -Append
        Write-Host "Temp folder $Folder\..\_tmp\$AppName-tmp deleted" -ForegroundColor Cyan
        "Temp folder $Folder\..\_tmp\$AppName-tmp deleted" | Out-File "$Folder\..\_Logs\$AppName-Collect.log" -Encoding utf8 -Append
        Remove-Item -Path "$Folder\..\_tmp\$AppName-tmp" -Recurse -Force | Out-Null
        Write-Host "All Things Done Well - Collect is in $Folder\..\_Exports\$ZipName-$SearchFrom-$TimeStamp.zip" -ForegroundColor Green
        Write-Host "Feel Free to Open Zip or Review Log in $Folder\..\_Logs\$AppName-Collect.log" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Could Not Generate ZipFiles"
        "Could Not Generate ZipFiles" | Out-File "$Folder\..\_Logs\$AppName-Collect.log" -Encoding utf8 -Append
        Write-Host "Please Compress Folder $Folder\..\_tmp\$AppName-tmp Manually" -ForegroundColor Cyan
        "Please Compress Folder $Folder\..\_tmp\$AppName-tmp Manually" | Out-File "$Folder\..\_Logs\$AppName-Collect.log" -Encoding utf8 -Append
    }    
    
} else {
    Write-Host "Could Not Find any $AppName Like Files From $Search With a Depth of $Depth Subfolders" -ForegroundColor Cyan
    "Could Not Find any $AppName Like Files From $Search With a Depth of $Depth Subfolders" | Out-File "$Folder\..\_Logs\$AppName-Collect.log" -Encoding utf8 -Append
}


#.\ScubaCollect.ps1 -Search "C:\_Cultura\_Tmp\ScubaGear" -Verbose