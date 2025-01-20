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
    [string]$Domain = "",
    [Parameter(Mandatory=$false, Position=1)]
    [switch] $AllDates
)

$AuditorCompany = "Metsys"
$AuditorURL = "https://www.metsys.fr"
$AuditorLogo = "https://frenchies.net/audit/metsys.png"
$ConsultantName = "Stephane Giraud"
$ConsultantMail = "stephane.giraud@metsys.fr"
$ConsultantPhone = "+33674080981"


$ClientName = "CULTURA"
$ClientLogo = "https://frenchies.net/audit/logocultura.png"
$ClientContact = "Kevin Coupe"
$ClientMail = "k.coupe-ext@cultura.fr"
$ClientPhone = "+33647905144"


$ToolName = "PingCastle"
$DatasPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Datas\$ToolName").FullName)
$ImagesPath = [String]((Get-Item -Path "$PSScriptRoot\images").FullName)
New-Item -Path "$PSScriptRoot\..\_Reports" -Name "$ToolName" -ItemType Directory -Force | Out-Null
$ReportsPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Reports\$ToolName").FullName)
$RemediatePath = [String]((Get-Item -Path "$PSScriptRoot\..\_Remediate").FullName)
$SourcesPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Sources").FullName)

#$ReportsPath  = "C:\_Metsys\_Reports\PingCastle"

#$MyPCReportsXMLFiles = Get-ChildItem -Path "$ReportsPath" -Include "ad_hc_*.xml" -Recurse


$PCXMLFileInfos = New-Object 'System.Collections.Generic.List[System.Object]'
Get-ChildItem -Path "$DatasPath" -Include "ad_hc_*.xml" -Recurse | ForEach-Object {
    $MyPCReportXMLFilePath = [String]($($_.FullName))
    [xml]$PCXMLInfos = Get-Content "$MyPCReportXMLFilePath"
    $MyPCReportDomain = "$($PCXMLInfos.HealthcheckData.DomainFQDN)"
    $MyPCReportDate = ([DateTime]$($PCXMLInfos.HealthcheckData.GenerationDate)).ToString('yyyyMMdd')
    $obj = [PSCustomObject]@{
        XMLFile = [String]($MyPCReportXMLFilePath)
        Date    = [String]($MyPCReportDate)
        Domain  = [String]($MyPCReportDomain)
    }
    $PCXMLFileInfos.Add($Obj)
}

$AllDomains = $PCXMLFileInfos.Domain | Select-Object -Unique


If ($Domain -eq "*") {
    $AllDomains = $True
    $ReportOn = $PCXMLFileInfos
} else {
    If ($AllDomains -contains "$Domain") {
        $MyDomain = "$Domain"
        $ReportOnDomain = $PCXMLFileInfos | Where-Object Domain -eq "$Domain"
    } else {
        If ($AllDomains.Count -gt 1) {
            $DomainsArr = @()
            $i=0
            $AllDomains | ForEach-Object {
                $DomainsArr += [System.Management.Automation.Host.ChoiceDescription]::new("&$i-$($_)")
                $i=$i+1
            }
            $result = $host.ui.PromptForChoice("Report on ...", "Choose Domain ?", [System.Management.Automation.Host.ChoiceDescription[]]($DomainsArr), 0)
            $MyDomain = ($DomainsArr[$result].Label).Split("-")[1]
            
        } else {
            $MyDomain = [String]($AllDomains)
        }
        $ReportOnDomain = $PCXMLFileInfos | Where-Object Domain -eq "$MyDomain"
    }
    If ( -not $AllDates) {
        $DatesArr = @()
        $i=0
        $PCXMLFileInfos | Where-Object Domain -eq "$MyDomain" | Select-Object Date -Unique | Sort-Object Date -Descending| ForEach-Object {
            $DatesArr += [System.Management.Automation.Host.ChoiceDescription]::new("&$i-$($_.Date)")
            $i=$i+1
        }
        $result = $host.ui.PromptForChoice("Report on $MyDomain", "Choose Date ?", [System.Management.Automation.Host.ChoiceDescription[]]($DatesArr), 0)
        $MyDate = ($DatesArr[$result].Label).Split("-")[1]
        $ReportOn = $ReportOnDomain | Where-Object Date -eq "$MyDate"
    } else {
        $ReportOn = $ReportOnDomain
    }
}

$DomainReports = $ReportOn | Select-Object Domain -Unique
$DateReports = $ReportOn | Select-Object Date -Unique

Write-Host "Reporting on $($DomainReports.Domain.Count) Domains Over $($DateReports.Date.Count) Date(s)"
#$DomainReports.Domain
#Write-Host "for $($DateReports.Date.Count) Date(s)"
#$DateReports.Date


$Violet = "#783CBD"
$VioletClair = "#C7ADE5"
$Noir = "#3D3834"
$Gris = "#C9D1D1"
$Or = "#BC9C16"
$OrClair = "#F3DCAC"


$LevelColors = @()
$LevelColors += "White"
$LevelColors += "Red"
$LevelColors += "Orange"
$LevelColors += "Yellow"
$LevelColors += "Blue"
$LevelColors += "Green"

$DomainFunctionnaLevels = @()
$DomainFunctionnaLevels += "Windows 2000 Native"
$DomainFunctionnaLevels += "Windows 2000 Mixed"
$DomainFunctionnaLevels += "Windows 2003"
$DomainFunctionnaLevels += "Windows 2008"
$DomainFunctionnaLevels += "Windows 2008 R2"
$DomainFunctionnaLevels += "Windows 2012"
$DomainFunctionnaLevels += "Windows 2012 R2"
$DomainFunctionnaLevels += "Windows 201"

$ForestFunctionnaLevels = @()
$ForestFunctionnaLevels += "Windows 2000"
$ForestFunctionnaLevels += "Windows 2003 interim"
$ForestFunctionnaLevels += "Windows 2003"
$ForestFunctionnaLevels += "Windows 2008"
$ForestFunctionnaLevels += "Windows 2008 R2"
$ForestFunctionnaLevels += "Windows 2012"
$ForestFunctionnaLevels += "Windows 2012 R2"
$ForestFunctionnaLevels += "Windows 2016"



[xml]$PCRulesXml = Get-Content "$PSScriptRoot\PingCastleRules.xml"
#######################
$PCUniCat = $PCRulesXml.ArrayOfExportedRule.ExportedRule.Category | Select-Object -Unique
$PCUnicModel = $PCRulesXml.ArrayOfExportedRule.ExportedRule.Model | Select-Object -Unique

$PCModelByCat  = New-Object 'System.Collections.Generic.List[System.Object]'
$PCUniCat | ForEach-Object {
    $MaCat = "$_"
    $MesModels = ($PCRulesXml.ArrayOfExportedRule.ExportedRule | Where-Object Category -eq "$MaCat" | Select-Object Model -Unique).Model
    ForEach ($MonModel in $MesModels) {
        #Write-Host "$MaCat - $MonModel"
        $obj = [PSCustomObject]@{
            Category    = [String]($MaCat)
            Model       = [String]($MonModel)
        }
        $PCModelByCat.Add($Obj)
    }
}
$MaxModelNb = (($PCModelByCat | Group-Object Category) | Measure-Object -Maximum Count).Maximum
$MyModelNbByCat = ($PCModelByCat | Group-Object Category) | Select-Object Name, Count

$MyRiskModelArray =  New-Object 'System.Collections.Generic.List[System.Object]'
For ($m=0; $m -lt $MaxModelNb; $m++) {
    $obj = [PSCustomObject]@{}
    $PCUniCat | ForEach-Object {
        $MaCat = "$_"
        $SearchedModel = ($PCModelByCat | Where-Object Category -eq "$MaCat" | Select-Object -index $m).Model
        $obj | Add-Member -NotePropertyName "$MaCat" -NotePropertyValue "$([String]$($SearchedModel))"
    }
    $MyRiskModelArray.Add($Obj)
}


#######################

$ReportOn | ForEach-Object {
   $PCXMLReportFile = "$($_.XMLFile)"
    [xml]$PCHCD = Get-Content -Path "$PCXMLReportFile"
    $SelectedDomain = "$($PCHCD.HealthcheckData.DomainFQDN)"
    $MyPCReportDate = ([DateTime]$($PCHCD.HealthcheckData.GenerationDate)).ToString('yyyyMMdd')
    $MyPCReportFullDate = ([DateTime]$($PCHCD.HealthcheckData.GenerationDate)).ToString('dddd dd MMM yyyy')
    $ReportFileName = "PingCastleAnalysis_$SelectedDomain-$MyPCReportDate.html"

    Write-Host "Reporting For PingCasle Analysis of $MyPCReportFullDate From $PCXMLReportFile" -ForegroundColor Blue

    $MyPCReportRules = $PCHCD.HealthcheckData.RiskRules.HealthcheckRiskRule

    $PCPoints = ($MyPCReportRules.Points | Measure-Object -Sum).Sum

    $DataPCRules = New-Object 'System.Collections.Generic.List[System.Object]'
    $MyPCReportRules | ForEach-Object {
        $obj = [PSCustomObject]@{
                Level       = [int] ($PCRulesXml.ArrayOfExportedRule.ExportedRule | Where-Object RiskId -like "$($_.RiskId)").MaturityLevel
                Points      = [string]($_.Points)
                Category    = [string]($_.Category)
                Model       = [string]($_.Model)
                RiskId      = [string]($_.RiskId)
                Rationale   = [string]($_.Rationale)
            }
        $DataPCRules.Add($Obj)
    }

    #$DataPCRules | Sort-Object Level

    $MyPCCategories = $MyPCReportRules.Category | Select-Object -Unique
    $MyPCModels = $MyPCReportRules.Model | Select-Object -Unique

    $DataRulesByCat = New-Object 'System.Collections.Generic.List[System.Object]'
    $MyPCCategories  | ForEach-Object {
        $MyCat = "$_"
        $MyRulesMatched = (($MyPCReportRules | Where-Object Category -eq "$MyCat").Points | Measure-Object -Sum)
        $obj = [PSCustomObject]@{
            Category = [string]($MyCat)
            NbRules = [int]($MyRulesMatched.Count)
            Points = [int]($MyRulesMatched.Sum)
            
        }
    $DataRulesByCat.Add($Obj)
    }


    $DataRulesByModel = New-Object 'System.Collections.Generic.List[System.Object]'
    $MyPCModels  | ForEach-Object {
        $MyModel = "$_"
        $MyRulesMatched = (($MyPCReportRules | Where-Object Model -eq "$MyModel").Points | Measure-Object -Sum)
        $MyCat = ($PCModelByCat | Where-Object Model -eq "$MyModel" ).Category
        $obj = [PSCustomObject]@{
            Model = [string]($MyModel)
            Category = [String]($MyCat)
            NbRules = [int]($MyRulesMatched.Count)
            Points = [int]($MyRulesMatched.Sum)
        }
    $DataRulesByModel.Add($Obj)

    }


    $DataRulesByCatModel = New-Object 'System.Collections.Generic.List[System.Object]'

    For ($m=0; $m -lt $MaxModelNb; $m++) {
        $obj = [PSCustomObject]@{}
        $PCUniCat | ForEach-Object {
            $MaCat = "$_"
            $SearchedModel = ($PCModelByCat | Where-Object Category -eq "$MaCat" | Select-Object -index $m).Model
            $MonModelInfos = $DataRulesByModel | Where-Object Model -eq "$SearchedModel"
            $obj | Add-Member -NotePropertyName "$MaCat" -NotePropertyValue "$([String]$($MonModelInfos.Points))"
        }
        $DataRulesByCatModel.Add($Obj)
        Start-Sleep -Milliseconds 100
    }

    New-Item -Path "$ReportsPath" -Name "$SelectedDomain" -ItemType Directory -Force | Out-Null
    New-HTML -TitleText "PingCastle Analysis Report" -Author "$ConsultantName" -Encoding UTF8 <#-FavIcon "$MyFavIcon"#>  {

        Enable-HTMLFeature -Feature FontsAwesome
        New-HTMLFooter -HTMLContent { "<center>&copy; $(Get-Date -Format "yyyy") - <font color=""$Violet""><a href =""$AuditorURL"" target=_blank>$AuditorCompany</a></font></center>" }
        

        New-HTMLContent -HeaderText 'Informations' -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
            New-HTMLColumn -Width 33% {
                New-HTMLImage -Source "$AuditorLogo" -Width "240" -Inline -UrlLink "$AuditorURL" -Target _blank
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

                New-HTMLText  -Text "$SelectedDomain" -Color $Violet -FontVariant small-caps -FontWeight bold -FontSize 42
                New-HTMLFontIcon -IconSolid info-circle -IconSize 22
                New-HTMLText  -Text "PingCastle v$($PCHCD.HealthCheckData.EngineVersion)" -Color $Violet -FontSize 22 -Alignment center
                New-HTMLFontIcon -IconSolid calendar-day -IconSize 22
                New-HTMLText  -Text "$MyPCReportFullDate" -Color $Violet -FontSize 22 -Alignment center
                New-HTMLImage -Source "$ImagesPath\PCLogo.png" -Width "160" -Inline -UrlLink "https://www.pingcastle.com/" -Target _blank
    
            } -AlignContentText center
            
            New-HTMLColumn -Width 33% {
                New-HTMLImage -Source "$ClientLogo" -Width "240" -Inline
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


        
        #New-HTMLText -Text "PingCastle v$($PCHCD.HealthCheckData.EngineVersion) - Report Date : $MyPCReportFullDate" -Alignment center

        New-HTMLContent -HeaderText "Maturity" -HeaderTextSize 22 -HeaderTextColor $Violet -HeaderTextAlignment center -HeaderBackGroundColor $Or {
            New-HTMLColumn -Width 25% -BackgroundColor "$($LevelColors[$PCHCD.HealthcheckData.MaturityLevel])" {
                New-HTMLText -FontSize 200 -FontWeight bold -Text "$($PCHCD.HealthcheckData.MaturityLevel)" -Color "$($LevelColors[0])" -Alignment center
            } -AlignContentText center
            New-HTMLColumn -Width 50% -BackgroundColor Grey -AlignContentText center {
                New-HTMLText -FontSize 13 -FontWeight bold -Color "$($LevelColors[1])" -Text "Critical weaknesses and misconfigurations pose an immediate threat to all hosted resources. Corrective actions should be taken as soon as possible"
                New-HTMLText -FontSize 13 -FontWeight bold -Color "$($LevelColors[2])" -Text "Configuration and management weaknesses put all hosted resources at risk of a short-term compromise. Corrective actions should be carefully planned and implemented shortly"
                New-HTMLText -FontSize 13 -FontWeight bold -Color "$($LevelColors[3])" -Text "The Active Directory infrastructure does not appear to have been weakened from what default installation settings provide"
                New-HTMLText -FontSize 13 -FontWeight bold -Color "$($LevelColors[4])" -Text "The Active Directory infrastructure exhibits an enhanced level of security and management"
                New-HTMLText -FontSize 13 -FontWeight bold -Color "$($LevelColors[5])" -Text "The Active Directory infrastructure correctly implements the latest state-of-the-art administrative model and security features"

                New-HTMLChart -Height 150 {
                    New-ChartToolbar -Download
                    #New-ChartBarOptions -Gradient

                    $Points = [int]($PCHCD.HealthcheckData.GlobalScore)
                    if ($Points -gt 75) { $DisColor = "Red"} elseif ($Points -gt 50 -and $Points -le 75) {$DisColor = "DarkOrange"} elseif ($Points -gt 25 -and $Points -le 50) {$DisColor = "Yellow"} elseif ($Points -le 25) {$DisColor = "Green"} else {$DisColor = "White"}
                    New-ChartBar -Name "Global Score" -Value $Points -Color $Violet
                    $Points = [int]($PCHCD.HealthcheckData.StaleObjectsScore)
                    if ($Points -gt 75) { $DisColor = "Red"} elseif ($Points -gt 50 -and $Points -le 75) {$DisColor = "DarkOrange"} elseif ($Points -gt 25 -and $Points -le 50) {$DisColor = "Yellow"} elseif ($Points -le 25) {$DisColor = "Green"} else {$DisColor = "White"}
                    New-ChartBar -Name "StaleObjects Score" -Value $Points -Color "$DisColor"
                    $Points = [int]($PCHCD.HealthcheckData.PrivilegiedGroupScore)
                    if ($Points -gt 75) { $DisColor = "Red"} elseif ($Points -gt 50 -and $Points -le 75) {$DisColor = "DarkOrange"} elseif ($Points -gt 25 -and $Points -le 50) {$DisColor = "Yellow"} elseif ($Points -le 25) {$DisColor = "Green"} else {$DisColor = "White"}
                    New-ChartBar -Name "PrivilegiedGroup Score" -Value $Points -Color "$DisColor"
                    $Points = [int]($PCHCD.HealthcheckData.TrustScore)
                    if ($Points -gt 75) { $DisColor = "Red"} elseif ($Points -gt 50 -and $Points -le 75) {$DisColor = "DarkOrange"} elseif ($Points -gt 25 -and $Points -le 50) {$DisColor = "Yellow"} elseif ($Points -le 25) {$DisColor = "Green"} else {$DisColor = "White"}
                    New-ChartBar -Name "Trust Score" -Value $Points -Color "$DisColor"
                    $Points = [int]($PCHCD.HealthcheckData.AnomalyScore)
                    if ($Points -gt 75) { $DisColor = "Red"} elseif ($Points -gt 50 -and $Points -le 75) {$DisColor = "DarkOrange"} elseif ($Points -gt 25 -and $Points -le 50) {$DisColor = "Yellow"} elseif ($Points -le 25) {$DisColor = "Green"} else {$DisColor = "White"}
                    New-ChartBar -Name "Anomaly Score" -Value $Points -Color "$DisColor"
                    
                }

            }

            New-HTMLColumn -Width 25% {

                New-HTMLChart -Height 200 {
                    New-ChartToolbar -Download
                    New-ChartBarOptions -Gradient
                    #New-ChartLegend "Global Total", "StaleObjects Score", "PrivilegiedGroup Score", "Trust Score", "Anomaly Score"
                    New-ChartBar -Name "$($MyPCReportRules.Count) Rules" -Value $PCPoints
                    $DataRulesByCat | ForEach-Object {
                        $MyCat = [String] ($_.Category)
                        $MyRulesNb = [int] ($_.NbRules)
                        $Points = [int] ($_.Points)

                        #$Points = [int]((($DataPCRules | Where-Object Category -eq "$MyCat").Points | Measure-Object -Sum).Sum)
                        if ($Points -gt 75) { $DisColor = "Red"} elseif ($Points -gt 50 -and $Points -le 75) {$DisColor = "DarkOrange"} elseif ($Points -gt 25 -and $Points -le 50) {$DisColor = "Yellow"} elseif ($Points -le 25) {$DisColor = "Green"} else {$DisColor = "White"}
                        New-ChartBar -Name "$MyCat ($MyRulesNb)" -Value $Points -Color $DisColor
                    }
                    
                } -TitleAlignment center -Title "Total Score"


            }
        }


        New-HTMLContent -HeaderText 'Graphs & Stats' -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
      
            New-HTMLColumn -Width 25% {
                New-HTMLChart {
                    for ($i = 1; $i -le 5; $i++) {
                        New-ChartPie -Name "Level $i" -Value $(($DataPCRules | Where-Object Level -eq $i).Count) -Color "$($LevelColors[$i])"
                    }
                } -Title "$($MyPCReportRules.Count) Rules Matched - By Level" -TitleAlignment center -TitleColor $Violet -Height 250
            }

            New-HTMLColumn -Width 25% {
                New-HTMLChart {
                    $DataRulesByCat | ForEach-Object {
                        $MyCat = [String] ($_.Category)
                        $MyRulesNb = [int] ($_.NbRules)
                        #$Points = [int] ($_.Points)
                        New-ChartPie -Name "$MyCat" -Value $MyRulesNb
                    }
                } -Title "$($MyPCReportRules.Count) Rules Matched - By Category" -TitleAlignment center -TitleColor $Violet -Height 250
            }

            New-HTMLColumn -Width 25% {
                New-HTMLChart -ChartSettings {
                    for ($i = 1; $i -le 5; $i++) {
                        New-ChartPie -Name "Level $i" -Value $((($DataPCRules | Where-Object Level -eq $i).Points | Measure-Object -Sum).Sum) -Color "$($LevelColors[$i])"
                    }
                } -Title "$PCPoints Points - By Level" -TitleAlignment center -TitleColor $Violet -Height 250
            }

 
            New-HTMLColumn -Width 25% {
                New-HTMLChart {
                    $DataRulesByCat | ForEach-Object {
                        $MyCat = [String] ($_.Category)
                        #$MyRulesNb = [int] ($_.NbRules)
                        $Points = [int] ($_.Points)
                        New-ChartPie -Name "$MyCat" -Value $Points
                    }
                } -Title "$PCPoints Points - By Category" -TitleAlignment center -TitleColor $Violet -Height 250
            }
            

        }
    
        New-HTMLContent -HeaderText 'Active Directory Information' -HeaderTextSize 22 -HeaderTextColor $Violet -HeaderTextAlignment center -HeaderBackGroundColor $Or {
            New-HTMLColumn -Width 25% {
                New-HTMLHeading h2 -HeadingText "Informations"
                New-HTMLText -Text "Forest : $($PCHCD.HealthcheckData.ForestFQDN)"
                New-HTMLText -Text "Forest Functional Level : $($ForestFunctionnaLevels[$($PCHCD.HealthcheckData.ForestFunctionalLevel)])"
                New-HTMLText -Text "Domain : $($PCHCD.HealthcheckData.DomainFQDN)"
                New-HTMLText -Text "Domain Functional Level : $($DomainFunctionnaLevels[$($PCHCD.HealthcheckData.DomainFunctionalLevel)])"
                New-HTMLText -Text "Schema version : $($PCHCD.HealthcheckData.SchemaVersion)"
                New-HTMLText -Text "Schema Internal version : $($PCHCD.HealthcheckData.SchemaInternalVersion)"
                New-HTMLText -Text "Schema LastChanged : $(([DateTime]$($($PCHCD.HealthcheckData.SchemaLastChanged))).ToString('yyyy-MM-dd'))"
                New-HTMLText -Text "Domain Creation : $(([DateTime]$($($PCHCD.HealthcheckData.DomainCreation))).ToString('yyyy-MM-dd'))" 

                New-HTMLText -Text "Domain SID : $($PCHCD.HealthcheckData.DomainSid)"
                New-HTMLText -Text "RecycleBin Enabled : $($PCHCD.HealthcheckData.IsRecycleBinEnabled)"
                New-HTMLText -Text "$($PCHCD.HealthcheckData.NumberOfDC) Domain Controllers"
                
                

                If ($PCHCD.HealthcheckData.AzureADName -and $PCHCD.HealthcheckData.AzureADId) {
                    New-HTMLHorizontalLine
                    New-HTMLText -Text "Azure AD Detected"
                    New-HTMLText -Text "Tenant Name : $($PCHCD.HealthcheckData.AzureADName)"
                    New-HTMLText -Text "Tenant Id : $($PCHCD.HealthcheckData.AzureADId)"
                }
                

            }
            New-HTMLColumn -Width 75% {
                New-HTMLHeading h2 -HeadingText "Domain Controllers"

                $DataDomainControllers = New-Object 'System.Collections.Generic.List[System.Object]'
                $MyPCDCInfos = $PCHCD.HealthcheckData.DomainControllers.HealthcheckDomainController

                $MyPCDCInfos | ForEach-Object {
                    $IPAddresses = ($_.IP.String) -join ';'
                    $FMSORoles = ($_.FSMO.String) -join ';'
                    $obj = [PSCustomObject]@{
                            DCName                  = [string] $_.DCName
                            CreationDate            = [DateTime]$($_.CreationDate)
                            StartupTime             = [DateTime]$($_.StartupTime)
                            LastComputerLogonDate   = [DateTime]$($_.LastComputerLogonDate)
                            OperatingSystem         = [string] $_.OperatingSystem
                            <#
                            SupportSMB1             = [Boolean] $_.SupportSMB1
                            RemoteSpoolerDetected   = [Boolean] $_.RemoteSpoolerDetected
                            Ownername               = [string]  $_.OwnerName
                            #>
                            IP                      = [string] $IPAddresses
                            RODC                    = [Boolean] $_.RODC
                            FSMO                    = [string] $FMSORoles
                        }
                    $DataDomainControllers.Add($Obj)
                }
                #$DataDomainControllers | Sort-Object DCName
                New-HTMLTable -Title "Domain Controllers Info" -DataTable ($DataDomainControllers | Sort-Object DCName) -HideFooter -HideButtons -DisablePaging -DisableSearch

            }
        }

        New-HTMLContent -HeaderText 'Risk Model' -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
            New-HTMLColumn -Width 80% {
                New-HTMLTable -Title "Rules By Category and Model" -DataTable $MyRiskModelArray  -HideFooter -HideButtons -DisablePaging -DisableInfo -DisableSearch {
                    For ($MyRowIndex=0; $MyRowIndex -lt $MaxModelNb; $MyRowIndex++) {
                        $PCUniCat | ForEach-Object {
                            $MaCat = "$_"
                            $MyColIndex = [int]$($PCUniCat.IndexOf("$MaCat"))
                            $MyModelName = ($MyRiskModelArray | Select-Object -Index $MyRowIndex).("$MaCat")
                            $ModelPoints = ($DataRulesByCatModel | Select-Object -Index $MyRowIndex).("$MaCat")
                            #$ModelPoints = ($DataRulesByModel | Where-Object Model -eq "").Points
                            $ModelRulesNb = ($DataRulesByModel | Where-Object Model -eq "$MyModelName").NbRules
                            switch ($ModelPoints) {
                                { $_ -gt 30 } { $ModelColor = "Red" }
                                { $_ -ge 10 -and $_ -le 30 } { $ModelColor = "Orange" }
                                { $_ -lt 10 } { $ModelColor = "Yellow" }
                                { $_ -eq 0 } { $ModelColor = "Blue" }
                                { $_ -eq $null -or [string]::IsNullOrEmpty($_) } { $ModelColor = "White" }
                            }
    
                            If ($MyModelName -ne "") {
                                If ($ModelRulesNb) {
                                    $MyLine = "$MyModelName : $ModelPoints Points on $ModelRulesNb Rule(s)"
                                } else {
                                    $MyLine = "$MyModelName"
                                }
                            } else {
                                $MyLine = ""
                            }

                            If ($ModelColor -eq "Blue") {
                                New-HTMLTableContent -ColumnIndex $($MyColIndex+1) -RowIndex $($MyRowIndex+1) -Text "$MyLine" -FontSize 14 -Color White -BackGroundColor $ModelColor #-Text "$SearchedModel"
                            } else {
                                New-HTMLTableContent -ColumnIndex $($MyColIndex+1) -RowIndex $($MyRowIndex+1) -Text "$MyLine" -FontSize 14 -BackGroundColor $ModelColor #-Text "$SearchedModel"
                            } 
                        }
                    }
    #>                
                }
                

            }
            New-HTMLColumn -Width 20% -AlignContentText center{
                New-HTMLText -Text "Color Explain" -FontSize 18 -FontStyle italic
                New-HTMLText -Text "No Alert on that model" -BackGroundColor White -FontSize 18
                New-HTMLText -Text "score is 0 - no risk identified but some improvements detected" -BackGroundColor Blue -FontSize 18 -Color White
                New-HTMLText -Text "score between 1 and 10  - a few actions have been identified" -BackGroundColor Yellow -FontSize 18
                New-HTMLText -Text "score between 10 and 30 - rules should be looked with attention" -BackGroundColor Orange -FontSize 18
                New-HTMLText -Text "score higher than 30 - major risks identified" -BackGroundColor Red -FontSize 18
            }
        }

        If ( Test-Path -Path "$SourcesPath\PCMitre.xml" -PathType Leaf) {
            [xml]$PCMitreXml = Get-Content "$SourcesPath\PCMitre.xml"


            $MitreTec = $PCMitreXml.PCMitre.Mitre.Technic | Select-Object -Unique
            $MitreMit = $PCMitreXml.PCMitre.Mitre.Mitigation | Select-Object -Unique

            $DataPCMitre = New-Object 'System.Collections.Generic.List[System.Object]'

            $MyPCReportRules | ForEach-Object {
                $RiskIdTmp = $_.RiskID
                $obj = [PSCustomObject]@{
                        MitreTec    = [string] ($PCMitreXml.PCMitre.Mitre | Where-Object RiskId -eq "$RiskIdTmp").Technic
                        MitreMit    = [string] ($PCMitreXml.PCMitre.Mitre | Where-Object RiskId -eq "$($_.RiskId)").Mitigation
                        RiskId      = [string] $_.RiskId
                    }
                $DataPCMitre.Add($Obj)
            }


            New-HTMLContent -HeaderText 'MITRE ATT&CK' -HeaderTextSize 22 -HeaderTextColor $Violet -HeaderTextAlignment center -HeaderBackGroundColor $Or {
                New-HTMLColumn -Width 50% {
                    New-HTMLHeading h2 -HeadingText "Techniques"

                    $MitreTec | Sort-Object | ForEach-Object {
                        $Nb = $($DataPCMitre | Where-Object MitreTec -eq "$_").Count
                        If ($Nb) { New-HTMLText -Text "$_ : $Nb"}
                    }

                }
                New-HTMLColumn -Width 50% {
                    New-HTMLHeading h2 -HeadingText "Mitigations"

                    $MitreMit | Sort-Object | ForEach-Object {
                        $Nb = $($DataPCMitre | Where-Object MitreMit -eq "$_").Count
                        If ($Nb) { New-HTMLText -Text "$_ : $Nb"}
                    }

                }
            }
    
        }
 
        New-HTMLContent -HeaderText 'Risks' -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
            New-HTMLTable -Title "Matched Rules" -DataTable ($DataPCRules | Sort-Object Level) -HideFooter -HideButtons -DisablePaging {
                for ($i = 1; $i -le 5; $i++) {
                    New-HTMLTableCondition -Name "Level" -ComparisonType number -Operator eq -Value $i -BackgroundColor  "$($LevelColors[$i])" -Row
                }
            }
        }
    } -FilePath "$ReportsPath\$SelectedDomain\$ReportFileName"
}

