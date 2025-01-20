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


Add-Type -AssemblyName PresentationCore,PresentationFramework
Add-Type -AssemblyName System.IO.Compression.FileSystem

Import-Module -Name "psexcel"

$AppName = "PurpleKnight Report Import by Stef@n"
$ImportsPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Imports").FullName)
#$ImportsPath = "C:\_Metsys\_Imports"
$DatasPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Datas\PurpleKnight").FullName)


If (Test-Path -Path "$ImportsPath\Tmp" -PathType Container) { Remove-Item -Path "$ImportsPath\Tmp" -Recurse -Force -Confirm:$false }
New-Item -Path "$ImportsPath" -Name "Tmp" -ItemType Directory -Force | Out-Null
New-Item -Path "$ImportsPath" -Name "_ZipFiles" -ItemType Directory -Force | Out-Null

$MyZipFiles = Get-ChildItem -Path "$ImportsPath\*" -Include "AllMySPKReports-*.zip" -ErrorAction SilentlyContinue

$Imported=@()
$ClientReview = 0

If ($MyZipFiles.Count -lt 1) {
    #Write-Warning "No Zip Files in Import Folders"
    $MessageBody = "No PurpleKnight Zip Files in Import Folders `n`n Do You Want to open it"
    $Result = [System.Windows.MessageBox]::Show($MessageBody,$AppName,[System.Windows.MessageBoxButton]::YesNo,[System.Windows.MessageBoxImage]::Warning)
    If ($Result -eq 6) {Invoke-Item "$ImportsPath"}
    Exit
}


$ZipFileErrors = @()
ForEach ($MyZipFile in $MyZipFiles) {
    $MyZipFileName = $MyZipFile.BaseName
    Write-Output "Traitement de $MyZipFileName"
    $MyZipFileNameFull = "$MyZipFileName.zip"
    $MyZipFileNameFullPath = "$ImportsPath\$MyZipFileNameFull"
    #$MyZipFileNameFullPath = "C:\_Metsys\_Imports\AllMySPKReports-07_18_2024_12_31_25-20240807T154638.zip"
    [Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null
    $MyZip = [IO.Compression.ZipFile]::OpenRead($MyZipFileNameFullPath)

    $PKFiles = New-Object 'System.Collections.Generic.List[System.Object]'
    $MyZip.Entries | ForEach-Object {
        $obj = [PSCustomObject]@{
            FullName    = [String]($($_.FullName))
            FolderName  = [string](($_.FullName).Split("\"))[0]
            BaseName    = [String]($($_.Name))
        }
        $PKFiles.Add($Obj)
    }
    #$PKFiles.FolderName | Select-Object  -unique | ForEach-Object {New-Item -ItemType Directory -Path "$ImportsPath\Tmp" -Name "$($_)" -Force | Out-Null}

    <#
    $MyPKPDFFiles = $MyZip.Entries | Where-Object Name -like "Security_Assessment_Report_*.pdf"
    $MyPKHTMLFiles = $MyZip.Entries | Where-Object Name -like "Security_Assessment_Report_*.html"
    $MyPKXLSFiles = $MyZip.Entries | Where-Object Name -like "Security_Assessment_Report_*.xlsx"
    $MyPKCSVFiles = $MyZip.Entries | Where-Object Name -like "Security_Assessment_Report_*.csv"

    
    $MyPKPDFFiles = $PKFiles | Where-Object BaseName -like "Security_Assessment_Report_*.pdf"
    $MyPKHTMLFiles = $PKFiles | Where-Object BaseName -like "Security_Assessment_Report_*.html"
    $MyPKXLSFiles = $PKFiles | Where-Object BaseName -like "Security_Assessment_Report_*.xlsx"
    $MyPKCSVFiles = $PKFiles | Where-Object BaseName -like "Security_Assessment_Report_*.csv"
    #>
    $MyPKPDFFiles = $MyZip.Entries | Where-Object Name -like "Security_Assessment_Report_*.pdf"
    $MyPKHTMFiles = $MyZip.Entries | Where-Object Name -like "Security_Assessment_Report_*.html"
    $MyPKXLSFiles = $MyZip.Entries | Where-Object Name -like "Security_Assessment_Report_*.xlsx"
    #$MyPKCSVFiles = $MyZip.Entries | Where-Object Name -like "Security_Assessment_Report_*.csv"
    $MyPKConfXMLFile = $MyZip.Entries | Where-Object Name -eq "Scripts.config.xml"

    Write-Verbose "Analysing $($PKFiles.Count) Zip Files"
#    $MyPCTXTFiles = $MyZip.Entries | Where-Object Fullname -like "ad_*.txt" 
#    $MyPCHTMLFiles = $MyZip.Entries | Where-Object Fullname -like "ad_hc_*.html"
    If (-not $MyPKXLSFiles ) {
        Write-Warning "No PurpleKnight XLS Files in Zip File $MyZipFileName"
        $ZipFileErrors += $MyZipFile.FullName
        $MyZip.Dispose()
        Move-Item -Path "$ImportsPath\$MyZipFileNameFull" -Destination "$ImportsPath\_ZipFiles\INVALID_$MyZipFileNameFull" -Force
    } else {
        Write-Host "Analysing $($PKFiles.Count) Files"
        <#
        $MyPKPDFFiles | ForEach-Object { [IO.Compression.ZipFileExtensions]::ExtractToFile( $_, "$ImportsPath\Tmp\$($_.FolderName)\$($_.Name)")}
        $MyPKHTMFiles | ForEach-Object { [IO.Compression.ZipFileExtensions]::ExtractToFile( $_, "$ImportsPath\Tmp\$($_.FolderName)\$($_.Name)")}
        $MyPKXLSFiles | ForEach-Object { [IO.Compression.ZipFileExtensions]::ExtractToFile( $_, "$ImportsPath\Tmp\$($_.FolderName)\$($_.Name)")}
        #>
        $MyZip.Dispose()
        Expand-Archive -Path "$MyZipFileNameFullPath" -DestinationPath "$ImportsPath\Tmp"

        #pause
        $PKFiles | Select-Object FolderName -Unique | ForEach-Object {
            $MyFolder = [String]($_.FolderName)

            $MyXLSReport = Get-ChildItem -Path "$ImportsPath\Tmp\$MyFolder" -Filter "Security_Assessment_Report_*.xlsx"
            $MyPDFReport = Get-ChildItem -Path "$ImportsPath\Tmp\$MyFolder" -Filter "Security_Assessment_Report_*.pdf"
            $MyHTMReport = Get-ChildItem -Path "$ImportsPath\Tmp\$MyFolder" -Filter "Security_Assessment_Report_*.html"
            $MYPKReportSummary = Import-XLSX -Path "$($MyXLSReport.FullName)" -Sheet "Assessment summary" -Header "Name","Value" -FirstRowIsData
            
            #$MYPKReportDomains = Import-XLSX -Path "C:\_Metsys\AllMySPKReports-1-20240722T152809\1\Security_Assessment_Report_02_07_2024_11_24_52.xlsx" -Sheet "AD domains" -Header "Name","Value" -FirstRowIsData
            #$MYPKReportRisks = Import-XLSX -Path "C:\_Metsys\AllMySPKReports-1-20240722T152809\1\Security_Assessment_Report_02_07_2024_11_24_52.xlsx" -Sheet "Indicators results"
            $MyPKReportDomain = [String]($MYPKReportSummary | Where-Object Name -eq "AD: forest name").Value
            $MyPKReportDate = ([DateTime]$(($MYPKReportSummary | Where-Object Name -eq "Generated date").Value)).ToString('yyyyMMdd')
            #Close-Excel -Path "$($MyXLSReport.FullName)"
            if ($MyPKReportDomain -and $MyPKReportDate) {
                If (-not (Test-Path -Path "$DatasPath\$MyPKReportDomain\$MyPKReportDate\$($MyXLSReport.BaseName)" -PathType Leaf)) {
                    New-Item -Path "$DatasPath\$MyPKReportDomain" -Name "$MyPKReportDate" -ItemType Directory -Force | Out-Null
                    Write-Verbose "Creation du dossier PupleKnight $MyPKReportDomain - $MyPKReportDate"
                    If ($MyXLSReport) {Move-Item -Path "$($MyXLSReport.FullName)" -Destination "$DatasPath\$MyPKReportDomain\$MyPKReportDate" -Force }
                    If ($MyPDFReport) {Move-Item -Path "$($MyPDFReport.FullName)" -Destination "$DatasPath\$MyPKReportDomain\$MyPKReportDate" -Force}
                    If ($MyHTMReport) {Move-Item -Path "$($MyHTMReport.FullName)" -Destination "$DatasPath\$MyPKReportDomain\$MyPKReportDate" -Force}
                } else {
                    Write-Warning "$($MyXLSReport.BaseName) Report For $MyPKReportDomain Exist - No Copy !!"
                }
            } else {
                Write-Verbose "Invalid PurpleKnight File"
                #$ZipFileErrors += "$MyZipFileName"
            }
            #$MyZip.Dispose()
            
        }
        #Write-Host "Movin ZIp $MyZipFileNameFullPath"
        Move-Item -Path "$MyZipFileNameFullPath" -Destination "$ImportsPath\_ZipFiles\$MyZipFileNameFull" -Force -ErrorAction SilentlyContinue
    }

}

