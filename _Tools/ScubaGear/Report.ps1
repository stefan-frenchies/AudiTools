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
Param(
    [Parameter(Mandatory=$false, Position=0)]
    [ValidateRange(0, [int]::MaxValue)]
    [Int] $Last = 0
)

$filepath = "C:\_Cultura\_Tmp"

#$MyCMDB =  Get-Content ".\$ScubaGear.json" | ConvertFrom-Json

$MyReportsHistory = New-Object 'System.Collections.Generic.List[System.Object]'
$MyReportsCounts = New-Object 'System.Collections.Generic.List[System.Object]'
$MyReportsGroups = New-Object 'System.Collections.Generic.List[System.Object]'
$MyReportsDetails = New-Object 'System.Collections.Generic.List[System.Object]'

$MyJsonFiles = Get-ChildItem -Path $filepath -File -Filter "cultura*.json" -Recurse


$MyDataFolders = ( Get-ChildItem -Path $filepath -File -Filter "cultura*.json" -Recurse).Directory.FullName | Select-Object -Unique


$MyDatas = Get-Content -Raw ($MyJsonFiles | Where-Object Name -like "cultura-*.json").FullName | ConvertFrom-Json
$MyResults = Get-Content -Raw ($MyJsonFiles | Where-Object Name -eq "cultura.json").FullName | ConvertFrom-Json

#$MyResults | Where-Object -not RequirementMet #Requirement False

$MyProducts = New-Object PSObject
$MyDatas.MetaData.ProductsAssessed | ForEach-Object { $MyProducts | Add-Member -MemberType NoteProperty -Name "$($MyDatas.MetaData.ProductAbbreviationMapping."$_")" -Value "$_"}

#$MesTests = @("Manual","Passes","Errors","Failures","Warnings")


$MyReportCounts = [PSCustomObject]@{
    "Date"   = [String]([datetime]$($MyDatas.MetaData.TimestampZulu)).ToString("yyyyMMdd")
}

($MyDatas.Summary | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
    $MyProd = [String]($_)
    ($MyDatas.Summary."$MyProd" | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
        $MyVal = [String]($_)
        $MyCount = [Int]($MyDatas.Summary."$MyProd"."$MyVal")
        $MyReportCounts | Add-Member -MemberType NoteProperty -Name "$MyProd-$MyVal" -Value "$MyCount"

    }
}

#AddToHistory?
$MyReportsCounts.Add($MyReportCounts)

$MyReportInfos = New-Object 'System.Collections.Generic.List[System.Object]'
$MyReportInfos = [PSCustomObject]@{
    "TenantId"        = [string]$($MyDatas.MetaData.TenantId)
    "DisplayName"    = [string]$($MyDatas.MetaData.DisplayName)
    "DomainName"   = [string]$($MyDatas.MetaData.DomainName)
    "Tool"   = [string]$($MyDatas.MetaData.Tool)
    "ToolVersion"   = [string]$($MyDatas.MetaData.ToolVersion)
    "Date"   = [String]([datetime]$($MyDatas.MetaData.TimestampZulu)).ToString("yyyyMMdd") #[datetime]$($MyDatas.MetaData.TimestampZulu)
    

}

#AddToHistory?
$MyReportsHistory.Add($MyReportInfos)

$MyGroupInfos = New-Object 'System.Collections.Generic.List[System.Object]'
#$MyReportFullDetails = New-Object 'System.Collections.Generic.List[System.Object]'
$MyReportsDetails = New-Object 'System.Collections.Generic.List[System.Object]'

($MyDatas.Summary | Get-Member -MemberType NoteProperty).Name | ForEach-Object { 
    $MyProd = [String]($_)
    Write-Verbose "dealing With $MyProd"
    $MyDatas.Results."$MyProd"  | ForEach-Object {
        $Mygroup = $_
        $Obj = [PSCustomObject]@{
            "Id"        = [string]$($MyProd)
            "GroupName" = [String]($Mygroup.GroupName)
            "GroupNumber" = [String]($Mygroup.GroupNumber)
            "GroupReferenceURL" = [String]($Mygroup.GroupReferenceURL)
            "Date" = [String]([datetime]$($MyDatas.MetaData.TimestampZulu)).ToString("yyyyMMdd")  #[datetime]$($MyDatas.MetaData.TimestampZulu)
        }
        $MyGroupInfos.Add($Obj)
    }
    #AddToHistory?
    $MyReportsGroups.Add($MyGroupInfos)

    $MyDatas.Results."$MyProd".Controls | ForEach-Object {
        $MyDetails = $_
        Write-Verbose "Dealing With $($MyDetails."Control ID")"
        $Obj2 = [PSCustomObject]@{
            "ControlID"        = [String]($MyDetails."Control ID")
            "Criticality"    = [String]($MyDetails.Criticality)
            "Details"   = [String]($MyDetails.Details)
            "Requirement"   = [String]($MyDetails.Requirement)
            "Result"   = [String]($MyDetails.Result)
            }
        $reqMet = ($MyResults | Where-Object PolicyId -eq "$($MyDetails.("Control Id"))").RequirementMet
        $Obj2 | Add-Member -MemberType NoteProperty -Name "RequirementMet" -Value "$([String]($reqMet))"

        $Obj2 | Add-Member -MemberType NoteProperty -Name "Category" -Value "$([String]($MyProducts."$MyProd"))"
        $Obj2 | Add-Member -MemberType NoteProperty -Name "Model" -Value "$([String]($Obj.GroupName))"
        $Obj2 | Add-Member -MemberType NoteProperty -Name "Tool" -Value "$([String]$($MyDatas.MetaData.Tool))"
        $Obj2 | Add-Member -MemberType NoteProperty -Name "Date" -Value "$([String]([datetime]$($MyDatas.MetaData.TimestampZulu)).ToString("yyyyMMdd"))" #"$([datetime]$($MyDatas.MetaData.TimestampZulu))"

        $MyReportsDetails.Add($Obj2)
    }
}

#AddToHistory?
#$MyReportsFullDetails.Add($MyReportDetails)


$NbReportselected = [Int]($MyReportsHistory.Count)
If ($Last) {$NbReportselected = $Last}



$TSDate = [String]([datetime]$($MyDatas.MetaData.TimestampZulu)).ToString("yyyyMMdd")


$SourcesPath = "$PSScriptRoot"
$ReportPath = "$SourcesPath\_Reports"
#Start Auditor informations
$AuditorColor = "Blue"
$AuditorCompany = "iFrenchies"
$AuditorLogo = "$SourcesPath\images\logo$AuditorCompany.png"
$AuditorURL = "https://www.ifrenchies.eu"

$ConsultantName = "Stephane Giraud"
$ConsultantPhone = "+33695985004"
$ConsultantMail = "sgiraud@ifrenchies.eu"

$ToolName = ($MyReportsHistory | Select-Object -First 1).Tool
#Adaptation ToolName


$TooLogo = "$SourcesPath\images\$ToolName.png"
$ToolUrl = "https://cisagov.github.io/ScubaGear/"

#Start Client Info
$ClientName = ($MyReportsHistory | Select-Object -First 1).DisplayName
$FQDN = ($MyReportsHistory | Select-Object -First 1).DomainName
$FQDNId = ($MyReportsHistory | Select-Object -First 1).Tenant
$ClientLogo = "$SourcesPath\images\logo$ClientName.png"
$ClientContact = "Kevin Coupé"
$ClientPhone = "+33123456789"
$ClientMail = "kcoupe-ext@cultura.fr"
#End Client Info

#Start Report infos
$HeaderColor1 = "Yellow"
$HeaderColor2 = "Blue"
$HTMLReportFile = "$ReportPath\$ToolName-$ClientName.html"
#End Report infos


$MyFQDNDomain = $FQDN
$ReportsDate = @()
$MyReportsHistory | Sort-Object Date -Unique| ForEach-Object { $ReportsDate += "$($_.Date)"}
$NbReports = $ReportsDate.Count

New-HTML -TitleText "$ClientName - Scuba Gear Analysis of $FQDN" -Author "$ConsultantName" -Encoding UTF8 {
    Enable-HTMLFeature -Feature FontsAwesome

    New-HTMLFooter  {
        New-HTMLText -Text "&copy; $(Get-Date -Format "yyyy") - <font color=""$HeaderColor1""><a href =""$AuditorURL"" target=_blank>$AuditorCompany</a></font>" -Color Blue -Alignment center
    }

    New-HTMLTabStyle  -Transition -LinearGradient -SelectorColor Blue -SelectorColorTarget AliceBlue -FontSize 15 -SlimTabs
    New-HTMLTab -Name "Synthesis" -IconBrands hubspot -IconColor Blue {


        New-HTMLContent -HeaderText 'Informations' -HeaderTextSize 22 -HeaderTextColor $HeaderColor1 -HeaderTextAlignment center -HeaderBackGroundColor $HeaderColor2 {
            New-HTMLColumn -Width 33% {
                New-HTMLImage -Source "$AuditorLogo" -Height "160" -Inline -UrlLink "$AuditorURL" -Target _blank
                New-HTMLFontIcon -IconSolid address-book
                New-HTMLHeading h2 -HeadingText "$($AuditorCompany.ToUpper())"
                New-HTMLFontIcon -IconSolid address-card 
                New-HTMLText -Text "$ConsultantName"
                New-HTMLFontIcon -IconSolid phone
                New-HTMLText -Text "$ConsultantPhone"
                New-HTMLFontIcon -IconSolid envelope
                New-HTMLText -Text "$ConsultantMail"
            } -AlignContentText center
    
            New-HTMLColumn -Width 33% {
                New-HTMLText  -Text "$FQDN" -Color $HeaderColor2 -FontVariant small-caps -FontWeight bold -FontSize 42
                New-HTMLFontIcon -IconSolid info-circle -IconSize 20
                New-HTMLText  -Text "$NbReports Report(s) Available" -Color $HeaderColor2 -FontSize 20 -Alignment center
                New-HTMLFontIcon -IconSolid calendar-week -IconSize 20
                New-HTMLText  -Text "Synthesis From $($ReportsDate[0]) To $($ReportsDate[$NbReportselected -1])" -Color $HeaderColor2 -FontSize 20 -Alignment center
                New-HTMLImage -Source "$TooLogo" -Width "160" -Inline -UrlLink "$ToolURL" -Target _blank
    
            } -AlignContentText center
            
            New-HTMLColumn -Width 33% {
                New-HTMLImage -Source "$ClientLogo" -Height "160" -Inline
                New-HTMLFontIcon -IconSolid address-book
                New-HTMLHeading h2 -HeadingText "$($ClientName.ToUpper())"
                New-HTMLFontIcon -IconSolid address-card
                New-HTMLText -Text "$ClientContact"
                New-HTMLFontIcon -IconSolid phone
                New-HTMLText -Text "$ClientPhone"
                New-HTMLFontIcon -IconSolid envelope
                New-HTMLText -Text "$ClientMail"                
            } -AlignContentText center
        }

    }
    $ReportsDate | Select-Object -First $NbReportselected |  ForEach-Object {
        $MyDate = [String]$_
        New-HTMLTab -Name "$MyDate" { #-IconBrands acquisitions-incorporated -IconColor Yellow {
            $HTMLReportInfos =  $MyReportsHistory | Where-Object Date -eq "$MyDate"
            $HTMLReportDetails = $MyReportsDetails | Where-Object Date -eq "$MyDate"
            $HTMLReportCounts = $MyReportsCounts | Where-Object Date -eq "$MyDate"
            $HTMLReportGroups = $MyReportsGroups | Where-Object Date -eq "$MyDate"
  
            New-HTMLContent -HeaderText 'Informations' -HeaderTextSize 22 -HeaderTextColor $HeaderColor1 -HeaderTextAlignment center -HeaderBackGroundColor $HeaderColor2 {
  
                New-HTMLFontIcon -IconSolid address-book
                New-HTMLHeading h2 -HeadingText "$($AuditorCompany.ToUpper())"
                New-HTMLFontIcon -IconSolid address-card 
                New-HTMLText -Text "$ConsultantName"
                New-HTMLFontIcon -IconSolid phone
                New-HTMLText -Text "$ConsultantPhone"
                New-HTMLFontIcon -IconSolid envelope
                New-HTMLText -Text "$ConsultantMail"
            
            }
  
            New-HTMLContent -HeaderText 'Baselines Recommendations' -HeaderTextSize 22 -HeaderTextColor $HeaderColor1 -HeaderTextAlignment center -HeaderBackGroundColor $Or {
  
              $DataReport = $MyReportsDetails | Where-Object Date -EQ $MyDate 
              $MyCols = @("ControlID","Requirement","Criticality","Category","Model")
              New-HTMLTable -Title "Indicators Found" -DataTable ($DataReport | Sort-Object Score) -IncludeProperty $MyCols -HideFooter -HideButtons -DisablePaging {
                  <#
                  $LevelColors | ForEach-Object {
                      $MyLevelIndex = $LevelColors.IndexOf("$($_)")
                      New-HTMLTableCondition -Name "Level" -ComparisonType string -Operator eq -Value "$($MyLevelIndex)" -BackgroundColor $LevelColors[$($MyLevelIndex)] -Row
                  }
                      #>
                  #New-HTMLTableCondition -Name "Score" -ComparisonType string -Operator eq -Value $MaxScore -Color Green -Row
                  
              }
            }
        }
  
    }

} -FilePath "$HTMLReportFile"

Invoke-Item "$HTMLReportFile"
<#




$MyData.Results

    AAD
    Defender
    EXO
    PowerPlatform
    SharePoint
    Teams

        GroupName
        GroupNumber
        GroupReferenceURL

        Controls

            Control ID
            Requirement
            Result
            Criticality
            Details


$MyResults
    ActualValue
    Commandlet
    Criticality
    PolicyId
    ReportDetails
    RequirementMet

    MS.TEAMS.7.2v1
#>