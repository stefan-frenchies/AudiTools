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

Clear-Host
Add-Type -AssemblyName PresentationCore,PresentationFramework
Add-Type -AssemblyName System.IO.Compression.FileSystem

$AppName = "PingCastle Report Import by Stef@n"
$ImportsPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Imports\").FullName)
$ToolName = "PingCastle"
$DatasPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Datas\$ToolName").FullName)
$ImagesPath = [String]((Get-Item -Path "$PSScriptRoot\images").FullName)


If (Test-Path -Path "$ImportsPath\Tmp" -PathType Container) { Remove-Item -Path "$ImportsPath\Tmp" -Recurse -Force -Confirm:$false }
New-Item -Path "$ImportsPath" -Name "Tmp" -ItemType Directory -Force | Out-Null
New-Item -Path "$ImportsPath" -Name "_ZipFiles" -ItemType Directory -Force | Out-Null

$MyZipFiles = Get-ChildItem -Path "$ImportsPath\*" -Include "RunPC_*.zip" -ErrorAction SilentlyContinue

$Imported=@()
$ClientReview = 0

If ($MyZipFiles.Count -lt 1) {
    #Write-Warning "No Zip Files in Import Folders"
    $MessageBody = "No PingCastle Zip Files in Import Folders `n`n Do You Want to open it"
    $Result = [System.Windows.MessageBox]::Show($MessageBody,$AppName,[System.Windows.MessageBoxButton]::YesNo,[System.Windows.MessageBoxImage]::Warning)
    If ($Result -eq 6) {Invoke-Item "$ImportsPath"}
    Exit
}
$ZipFileErrors = @()
ForEach ($MyZipFile in $MyZipFiles) {
    $MyZipFileName = $MyZipFile.BaseName
    Write-Output "Traitement de $MyZipFileName"
    $MyTempFolder = $MyZipFileName.Split("_")[1]
    $MyZipFileNameFull = "$MyZipFileName.zip"
    $MyZipFileNameFullPath = "$ImportsPath\$MyZipFileNameFull"
    #$MyZipFileNameFullPath = "C:\_Metsys\_Scripts\AllMyPCReports-20240701-20240709T122934.zip"
    #$MyZipFileNameFullPath = "C:\_Metsys\_Imports\AllMyPCReports-_Metsys-20240712T122448.zip"
    [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null
    $MyZip = [IO.Compression.ZipFile]::OpenRead($MyZipFileNameFullPath)

    $PCFiles = New-Object 'System.Collections.Generic.List[System.Object]'
    $MyZip.Entries | ForEach-Object {
        $obj = [PSCustomObject]@{
            FullName    = [String]($($_.FullName))
            FolderName  = [string]($MyTempFolder)
            BaseName    = [String]($($_.Name))
        }
        $PCFiles.Add($Obj)
    }

   
    #CreationSubFolders
    $PCFiles.FolderName | Select-Object -unique | ForEach-Object {New-Item -ItemType Directory -Path "$ImportsPath\Tmp" -Name "$($_)" -Force | Out-Null}
    

    
    $MyPCXMLFiles = $PCFiles | Where-Object BaseName -like "ad_hc_*.xml"
    Write-Verbose "Analysing $($MyPCXMLFiles.Name.Count) Files"
#    $MyPCTXTFiles = $MyZip.Entries | Where-Object Fullname -like "ad_*.txt" 
    $MyPCHTMLFiles = $MyZip.Entries | Where-Object Fullname -like "ad_hc_*.html"
    If (-not $MyPCXMLFiles) {
        Write-Warning "No PingCastle XML in Zip File $MyZipFileName"
        $ZipFileErrors += $MyZipFile.FullName
        $MyZip.Dispose()
        Move-Item -Path "$ImportsPath\$MyZipFileNameFull" -Destination "$ImportsPath\_ZipFiles\INVALID_$MyZipFileNameFull" -Force
    } else {
        Write-Verbose "Analysing $($MyPCXMLFiles.BaseName.Count) XML Files"

        $MyPCXMLFiles | ForEach-Object {
            #Remove-Item -Path "$ImportsPath\Tmp" -Filter "*.xml" -Force
            #[IO.Compression.ZipFileExtensions]::ExtractToFile( $_.FullName, "$ImportsPath\Tmp\$($_.Name")
            $XMLFuNZip = [String]($_.FullName)
            $XMLFoNZip = [String]($_.FolderName)
            $XMLFiNZip = [String]($_.BaseName)

            $MyZip.Entries | Where-Object FullName -eq "$XMLFuNZip" | ForEach-Object {
                Write-Verbose "Extracting $XMLFiNZip to $XMLFoNZip as $XMLFiNZip"
                [IO.Compression.ZipFileExtensions]::ExtractToFile( $_, "$ImportsPath\Tmp\$XMLFoNZip\$XMLFiNZip")
            }
            #$MyPCXMLFullFileName = $_.FullName
            #$MyPCXMLFileName = $_.BaseName
            #$MyPCXMLSubFolderName = $_.FolderName #[string](($_.FullName).Split("/"))[0]
            Write-Verbose "Reading $ImportsPath\Tmp\$XMLFoNZip\$XMLFiNZip"
            [xml]$PCReportInfos = Get-Content "$ImportsPath\Tmp\$XMLFoNZip\$XMLFiNZip"
            $MyPCReportDomain = "$($PCReportInfos.HealthcheckData.DomainFQDN)"
            $MyPCReportDate = ([DateTime]$($PCReportInfos.HealthcheckData.GenerationDate)).ToString('yyyyMMdd')
            Write-Verbose "Analyse $XMLFiNZip"
            if ($MyPCReportDomain -and $MyPCReportDate) {
                If (-not (Test-Path -Path "$DatasPath\$MyPCReportDomain\$MyPCReportDate" -PathType Container)) {
                    New-Item -Path "$DatasPath\$MyPCReportDomain" -Name "$MyPCReportDate" -ItemType Directory -Force | Out-Null
                    Write-Host "Creation du dossier PingCastle $MyPCReportDomain - $MyPCReportDate"

                    Move-Item -Path "$ImportsPath\Tmp\$XMLFoNZip\$XMLFiNZip" -Destination "$DatasPath\$MyPCReportDomain\$MyPCReportDate" -Force
                    <#
                    $MyZip.Entries | Where-Object Fullname -like "$XMLFoNZip\ad_*$MyPCReportDomain.txt" | ForEach-Object {
                        [IO.Compression.ZipFileExtensions]::ExtractToFile( $_, "$ReportsPath\$MyPCReportDomain\$MyPCReportDate\$($_.Name)")
                    }
                    #> 
                    $MyZip.Entries | Where-Object Fullname -like "ad_hc_$MyPCReportDomain*.html" | ForEach-Object {
                        [IO.Compression.ZipFileExtensions]::ExtractToFile( $_, "$DatasPath\$MyPCReportDomain\$MyPCReportDate\$($_.Name)")
                    } 

                } else {
                    Write-Warning "$MyPCReportDate Datas Folder For $MyPCReportDomain Exist - No Copy !!"
                }
            } else {
                Write-Verbose "$XMLFiNZip Invalid PingCastle File"
                #$ZipFileErrors += "$MyZipFileName"
            }

        }
        $MyZip.Dispose()
        Move-Item -Path "$ImportsPath\$MyZipFileNameFull" -Destination "$ImportsPath\_ZipFiles\$MyZipFileNameFull" -Force
    }
    
}
