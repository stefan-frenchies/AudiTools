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

$DatasPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Datas\PurpleKnight").FullName)
#$$ReportsPath = "C:\_Metsys\_Reports"
$RemediatePath = [String]((Get-Item -Path "$PSScriptRoot\..\_Remediate").FullName)


Import-Module -Name "psexcel"



Get-ChildItem -Path "$DatasPath" -Filter "Security_Assessment_Report_*.xlsx" -Recurse | ForEach-Object {

    
    $MyPKXLSFilePath = [String]($($_.FullName))
    Write-Host "Please Wait For Analysing $MyPKXLSFilePath" -ForegroundColor Cyan
    #$MyPKXLSFilePath = "C:\_Metsys\_Reports\PurpleKnight\cultura.intra\20240718\Security_Assessment_Report_18_07_2024_12_17_09.xlsx"
    $MYPKReportSummary = Import-XLSX -Path "$MyPKXLSFilePath" -Sheet "Assessment summary"  -FirstRowIsData -Header "Name","Value"

    $MyReportXMLPath = "$(Split-Path -Path "$MyPKXLSFilePath" -Parent)\SecurityAssessmentReport_$((Get-Item -Path "$MyPKXLSFilePath").Directory.Name).xml"

    If (Test-Path -Path "$MyReportXMLPath" -PathType Leaf) {
        Write-Warning "$MyReportXMLPath Exist, Bypass..."
    } else {

        $xml = New-Object System.Xml.XmlDocument

        $newXMLNode = $xml.CreateElement("Report")
        $ReportNode = $xml.AppendChild($newXMLNode)
    
        $newXMLNode = $xml.CreateElement("Info")
        $InfoNode = $ReportNode.AppendChild($newXMLNode)
    
        
        $MYPKReportSummary | ForEach-Object {
            $MyProp = [String]($(-join (($_.Name) -split '[^a-zA-Z]')))
            $MyValue = [String]($_.Value)
            $newXMLNode = $xml.CreateElement("$MyProp")
            $MyXMLNode = $InfoNode.AppendChild($newXMLNode)
            $MyXMLNode.InnerText = [String]($MyValue)
    
            #$obj | Add-Member -NotePropertyName "$MyProp" -NotePropertyValue "$([String]$($MyValue))"
    
        }
    
    
    
        $MYPKReportIndicators = Import-XLSX -Path "$MyPKXLSFilePath" -Sheet "Indicators results"
        $MyIOES = $MYPKReportIndicators | Where-Object Status -eq "IOE Found"
    
        Write-Host "Analysing $($MyIOES.Count) Risk"
    
    
        #$MyIOEProps = $MyIOES | Get-Member -MemberType "NoteProperty"#).Name
        $MyIOEProps = @("Name","SI version","ShortName","Status","Description","Target","Category","Severity","Score","Number of results","Result message")

        $MyIOES | ForEach-Object {
            $newRiskNode = $xml.CreateElement("Risk")
            $RiskNode = $ReportNode.AppendChild($newRiskNode)
    
            $MyIOE = $_
            Write-Host "Analysing $($MyIOE.ShortName)"
            $MyIOEProps | ForEach-Object {
                $MyProp = [String]($(-join ($($_) -split '[^a-zA-Z0-9]')))
                
                #$MyValue = [String]($MyIOE.("$($_.Name)"))
                $MyValue = [String]($MyIOE.("$($_)"))
                $newRiskInfo = $xml.CreateElement("$MyProp")
                $RiskInfoNode = $RiskNode.AppendChild($newRiskInfo)
                $RiskInfoNode.InnerText = [String]($MyValue)
                #$obj | Add-Member -NotePropertyName "$MyProp" -NotePropertyValue "$([String]$($MyValue))"
    
         
            }
            If ($MyIOE.Result -like "SI*") {
                Write-Host "Details $($MyIOE.Result)" -ForegroundColor Cyan
                $MYPKReportIndicatorDetail = Import-XLSX -Path "$MyPKXLSFilePath" -Sheet "$($MyIOE.Result)"
    
                $MyRiskDetailsProps = $MYPKReportIndicatorDetail | Get-Member -MemberType "NoteProperty"
                $MYPKReportIndicatorDetail | ForEach-Object {
    
                    $newRiskDetailNode = $xml.CreateElement("Detail")
                    $RiskDetailNode = $RiskNode.AppendChild($newRiskDetailNode)
    
                    $MyDetail = $_
    
                    $MyRiskDetailsProps | ForEach-Object {
                        $MyProp = [String]($(-join (($_.Name) -split '[^a-zA-Z0-9]')))
                        $MyValue = [String]($MyDetail.("$($_.Name)"))
                        $newRiskDetailInfo = $xml.CreateElement("$MyProp")
                        $RiskDetailInfo = $RiskDetailNode.AppendChild($newRiskDetailInfo)
                        $RiskDetailInfo.InnerText = [String]($MyValue)
                        #$obj | Add-Member -NotePropertyName "$MyProp" -NotePropertyValue "$([String]$($MyValue))"
            
                 
                    }
    
        
                }
    
            } else { Write-Host "No Details For $($MyIOE.ShortName) " -ForegroundColor Green}
        }
    
        $xml.Save("$MyReportXMLPath")
    }

}
