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

$filepath = "C:\Users\s.giraud-ext\Cultura\Team Infra Microsoft - Documents\Audit\Infos\Scuba_2024_12_13_10_39_57"
$filepath = "C:\Users\s.giraud-ext\Cultura\Team Infra Microsoft - Documents\ScubaGear"

#$MyCMDB =  Get-Content ".\$ScubaGear.json" | ConvertFrom-Json

$MyReportsHistory = New-Object 'System.Collections.Generic.List[System.Object]'
$MyReportsCounts = New-Object 'System.Collections.Generic.List[System.Object]'
$MyReportsGroups = New-Object 'System.Collections.Generic.List[System.Object]'
$MyReportsDetails = New-Object 'System.Collections.Generic.List[System.Object]'


$MyAppReportFiles = New-Object 'System.Collections.Generic.List[System.Object]'

Get-ChildItem -Path $filepath -File -Filter "*.json"  -Depth 1 | ForEach-Object {
    $MyFullName = $_
    Write-Verbose "Getting Infos About $MyFullName"
    $FolderName  = [String](([System.IO.Directory]::GetParent("$MyFullName")).BaseName)
    $DateString = $FolderName.Substring($FolderName.Split("_")[0].Length + 1)
    
    $obj = [PSCustomObject]@{
        FullName    = [String]("$MyFullName")
        FolderPath  = [String](([System.IO.Directory]::GetParent("$MyFullName")).FullName)
        FolderParent  = [String](([System.IO.Directory]::GetParent("$MyFullName")).Parent.FullName)
        FolderName  = [String](([System.IO.Directory]::GetParent("$MyFullName")).BaseName)
        FolderDate = ([Datetime]::ParseExact($DateString, 'yyyy_MM_dd_hh_mm_ss', $null)).ToString("yyyyMMdd")
        BaseName    = [String]([System.IO.Path]::GetFileNameWithoutExtension("$MyFullName"))
        Extension   = [String]([System.IO.Path]::GetExtension("$MyFullName"))
    }
    $MyAppReportFiles.Add($Obj)
}


$MyAppReportFiles | Group-Object FolderName | ForEach-Object {
    $MyGroupFiles = $_.Group
    $MyGroupDate = $MyGroupFiles.FolderDate | Select-Object -Unique
    Write-Verbose "Analysing Files $MyGroupDate"
    $MyDatas = Get-Content -Raw "$(($MyGroupFiles | Where-Object BaseName -eq "cultura-$MyGroupDate").FullName)" | ConvertFrom-Json
    $MyResults = Get-Content -Raw "$(($MyGroupFiles | Where-Object BaseName -eq "cultura").FullName)" | ConvertFrom-Json

    $MyProducts = New-Object PSObject
    $MyDatas.MetaData.ProductsAssessed | ForEach-Object { $MyProducts | Add-Member -MemberType NoteProperty -Name "$($MyDatas.MetaData.ProductAbbreviationMapping."$_")" -Value "$_"}
    


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
        "Tool"   = [string]($MyDatas.MetaData.Tool)
        "ToolVersion"   = [string]$($MyDatas.MetaData.ToolVersion)
        #"Date" = [datetime]$($MyDatas.MetaData.TimestampZulu)
        "Date" = [String]([datetime]$($MyDatas.MetaData.TimestampZulu)).ToString("yyyyMMdd")
    }

    #AddToHistory?
    $MyReportsHistory.Add($MyReportInfos)

    $MyGroupInfos = New-Object 'System.Collections.Generic.List[System.Object]'
    $MyReportDetails = New-Object 'System.Collections.Generic.List[System.Object]'

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
                #"Date" = [datetime]$($MyDatas.MetaData.TimestampZulu)
                "Date" = [String]([datetime]$($MyDatas.MetaData.TimestampZulu)).ToString("yyyyMMdd")
            }
            $MyGroupInfos.Add($Obj)
        }
        #AddToHistory?
        $MyReportsGroups.Add($MyGroupInfos)

        $MyDatas.Results."$MyProd".Controls | ForEach-Object {
            $MyDetails = $_
            Write-Verbose "Dealing With $($MyDetails."Control ID")"
            $reqMet = ($MyResults | Where-Object PolicyId -eq "$($MyDetails.("Control Id"))").RequirementMet

            $Obj2 = [PSCustomObject]@{
                "ControlID"        = [String]($MyDetails."Control ID")
                "Criticality"    = [String]($MyDetails.Criticality)
                "Details"   = [String]($MyDetails.Details)
                "Requirement"   = [String]($MyDetails.Requirement)
                "Result"   = [String]($MyDetails.Result)
                "RequirementMet" = [Boolean]($reqMet)
                "Category" = [String]($MyProducts.("$MyProd"))
                "Model" = [String]($Obj.GroupName)
                "Tool" = [String]($MyDatas.MetaData.Tool)
                "Date" = [String](($MyDatas.MetaData.TimestampZulu).ToString("yyyyMMdd"))
                }
            
            $MyReportDetails.Add($Obj2)
            #$Obj2 | Add-Member -MemberType NoteProperty -Name "RequirementMet" -Value "$([String]($reqMet))"
            #$Obj2 | Add-Member -MemberType NoteProperty -Name "Categorie" -Value [String]($MyProducts.("$MyProd"))
            #$Obj2 | Add-Member -MemberType NoteProperty -Name "Modele" -Value [String]($Obj.GroupName)
            #$Obj2 | Add-Member -MemberType NoteProperty -Name "Tool" -Value [String]$($MyDatas.MetaData.Tool)
            #$Obj2 | Add-Member -MemberType NoteProperty -Name "Date" -Value "$([datetime]$($MyDatas.MetaData.TimestampZulu))"
            #$Obj2 | Add-Member -MemberType NoteProperty -Name "Date" -Value [String]([datetime]$($MyDatas.MetaData.TimestampZulu)).ToString("yyyyMMdd")
            #$MyReportDetails.Add($Obj2)
        }

    }

        #AddToHistory?
        $MyReportsDetails.Add($MyReportDetails)
}


#$TSDate = [String]([datetime]$($MyDatas.MetaData.TimestampZulu)).ToString("yyyyMMdd")


$SourcesPath = "$PSScriptRoot\..\..\_Sources"
$AuditorColor = "Blue"
$AuditorCompany = "iFrenchies"
$AuditorLogo = "$SourcesPath\images\logo$AuditorCompany.png"
$AuditorURL = "https://www.ifrenchies.eu"

$ToolName = ($MyReportsHistory | Select-Object -First 1).Tool
$TooLogo = "$SourcesPath\images\logo$ToolName.png"
$ToolUrl = "https://github.com/cisagov/ScubaGear"

$ConsultantName = "Stephane Giraud"
$ConsultantPhone = "+33 695985004"
$ConsultantMail = "sgiraud@ifrenchies.eu"

$HeaderColor1 = "Yellow"
$HeaderColor2 = "Blue"


$ClientName = ($MyReportsHistory | Select-Object -First 1).DisplayName
$FQDN = ($MyReportsHistory | Select-Object -First 1).DomainName
$FQDNId =($MyReportsHistory | Select-Object -First 1).TenantId
$ClientLogo = "$SourcesPath\images\logo$ClientName.png"
$ClientContact = "Kevin Coupé"
$ClientPhone = "+33 123456789"
$ClientMail = "kcoupe-ext@cultura.fr"

$HTMLReportFile = ".\Scubagear-Cultura.html"

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
                New-HTMLText  -Text "$NbReports Report(s)" -Color $HeaderColor2 -FontSize 20 -Alignment center
                New-HTMLFontIcon -IconSolid calendar-week -IconSize 20
                New-HTMLText  -Text "From $($ReportsDate[0]) To $($ReportsDate[$NbPCReports-1])" -Color $HeaderColor2 -FontSize 20 -Alignment center
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

    $ReportsDate | ForEach-Object {
        $MyDate = [String]$_
        New-HTMLTab -Name "$MyDate" { #-IconBrands acquisitions-incorporated -IconColor Yellow {
            $HTMLReportInfos =  $MyReportsHistory | Where-Object Date -eq "$MyDate"
            $HTMLReportDetails = $MyReportDetails | Where-Object Date -eq "$MyDate"
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



            New-HTMLContent -HeaderText 'Graphs & Stats' -HeaderTextSize 22 -HeaderTextColor $HeaderColor2 -HeaderTextAlignment center -HeaderBackGroundColor $HeaderColor1 {
                $AllProducts = $MyProducts | Get-Member -MemberType NoteProperty
                $AllProducts.Name | ForEach-Object {
                    $MyProduct = "$_"
                    New-HTMLColumn -Width (100/$AllProducts.Count)% {


                        New-HTMLChart {
                            New-ChartPie -Name "Evaluated" -Value $([Int]([Int]($InfosPKReport.ADEvaluated) + [Int]($InfosPKReport.EntraIDEvaluated) + [Int]($InfosPKReport.OktaEvaluated)) / $Allindicators * 100)
                            New-ChartPie -Name "Failed to run" -Value $([Int]([Int]($InfosPKReport.ADFailedtorun) + [Int]($InfosPKReport.EntraIDFailedtorun) + [Int]($InfosPKReport.OktaFailedtorun))/ $Allindicators * 100)
                            New-ChartPie -Name "Canceled" -Value $([Int]([Int]($InfosPKReport.ADCanceled) + [Int]($InfosPKReport.EntraIDCanceled) + [Int]($InfosPKReport.OktaCanceled)) / $Allindicators * 100)
                            New-ChartPie -Name "Not selected" -Value $([Int]([Int]($InfosPKReport.ADNotselected) + [Int]($InfosPKReport.EntraIDNotselected) + [Int]($InfosPKReport.OktaNotselected)) / $Allindicators *100)
                            New-ChartPie -Name "Not relevant" -Value $([Int]([Int]($InfosPKReport.ADNotrelevant) + [Int]($InfosPKReport.EntraIDNotrelevant) + [Int]($InfosPKReport.OktaNotrelevant)) / $Allindicators *100)

                        } -Title "$($MyProducts.("$MyProduct")) Indicators Analysis" -TitleAlignment center -TitleColor $HeaderColor1 -Height 250
                    }    
                }
    
            }
        }
    }


} -FilePath "$HTMLReportFile"

Invoke-Item "$HTMLReportFile"
<#

ControlID      : MS.TEAMS.8.1v1
Criticality    : Should/3rd Party
Details        : A custom product can be used to fulfill this policy requirement. If a custom product is used, a 3rd party
                 assessment tool or manually review is needed to ensure compliance. If you are using Defender for Office 365
                 to implement this policy, ensure that when running ScubaGear defender is in the ProductNames parameter.
                 Then, manually review the corresponding Defender for Office 365 policy that fulfills the requirements of
                 this policy. See <a
                 href="https://github.com/cisagov/ScubaGear/blob/v1.3.0/PowerShell/ScubaGear/baselines/teams.md#msteams81v1"
                 target="_blank">Secure Configuration Baseline policy</a> for instructions on manual check.
Requirement    : URL comparison with a blocklist SHOULD be enabled.
Result         : N/A
RequirementMet : False
Category       : Microsoft Teams
Model          : Link Protection
Tool           : ScubaGear
Date           : 20241215



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