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
    [string]$Domain = ""
    )


$Violet = "#783CBD"
$VioletClair = "#C7ADE5"
$Noir = "#3D3834"
$Gris = "#C9D1D1"
$Or = "#BC9C16"
$OrClair = "#F3DCAC"


$LevelColors = @()
$LevelColors += "White"
$LevelColors += "#f12828"
$LevelColors += "#ff6a00"
$LevelColors += "#ffd800"
$LevelColors += "#00aaff"
$LevelColors += "#83e043"

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
$SourcesPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Sources").FullName)
#####
#$ReportsPath  = "C:\_Metsys\_Reports\PingCastle"
#$SourcesPath = "C:\_Metsys\_Sources"

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
$HistoryOnDomain = $PCXMLFileInfos | Where-Object Domain -eq "$MyDomain" | Sort-Object Date

Write-Host "History Reporting From $($HistoryOnDomain.Date.Count) Reports On $MyDomain"


$HistoryReportFileName = "$ReportsPath\$MyDomain\PingCastleHistory-$MyDomain.html"

[xml]$PCRulesXml = Get-Content "$PsScriptRoot\PingCastleRules.xml"
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

$DataPCRules = New-Object 'System.Collections.Generic.List[System.Object]'
$DataScoreInfos = New-Object 'System.Collections.Generic.List[System.Object]'

$HistoryOnDomain.XMLFile | ForEach-Object {
    [xml]$MyReportInfos = Get-Content "$($_)"


    #[xml]$MyReportInfos = Get-Content "C:\_Metsys\_Reports\PingCastle\cultura.intra\20240701\ad_hc_cultura.intra.xml"

    $SelectedDomain = "$($MyReportInfos.HealthcheckData.DomainFQDN)"
    $MyPCReportDate = ([DateTime]$($MyReportInfos.HealthcheckData.GenerationDate)).ToString('yyyyMMdd')
    $MyPCReportFullDate = ([DateTime]$($MyReportInfos.HealthcheckData.GenerationDate)).ToString('dddd dd MMM yyyy')

    Write-Host "Reading $SelectedDomain Report of $MyPCReportDate : $($_)"

    $MyPCReportRules = $MyReportInfos.HealthcheckData.RiskRules.HealthcheckRiskRule
    $PCPoints = ($MyPCReportRules.Points | Measure-Object -Sum).Sum
    $MyPCReportRules | ForEach-Object {
        $Obj = [PSCustomObject]@{
                ReportDate  = [String]($MyPCReportDate)
                Level       = [int] ($PCRulesXml.ArrayOfExportedRule.ExportedRule | Where-Object RiskId -like "$($_.RiskId)").MaturityLevel
                Points      = [string]($_.Points)
                Category    = [string]($_.Category)
                Model       = [string]($_.Model)
                RiskId      = [string]($_.RiskId)
                Rationale   = [string]($_.Rationale)
            }
        $DataPCRules.Add($Obj)
    }


    $StaleObjectsScoreInfo = (($MyReportInfos.HealthcheckData.RiskRules.HealthcheckRiskRule | Where-Object Category -eq "StaleObjects").Points | Measure-Object -Sum)   
    $TrustScoreInfo = (($MyReportInfos.HealthcheckData.RiskRules.HealthcheckRiskRule | Where-Object Category -eq "Trust").Points | Measure-Object -Sum)
    $PrivilegiedGroupScoreInfo = (($MyReportInfos.HealthcheckData.RiskRules.HealthcheckRiskRule | Where-Object Category -eq "PrivilegedAccounts").Points | Measure-Object -Sum)
    $AnomalyScoreInfo = (($MyReportInfos.HealthcheckData.RiskRules.HealthcheckRiskRule | Where-Object Category -eq "Anomalies").Points | Measure-Object -Sum)

    $Obj = [PSCustomObject]@{
        ReportDate    = [String]($MyPCReportDate)
        MaturityLevel   = [Int]($MyReportInfos.HealthcheckData.MaturityLevel)
        GlobalScore = [Int]($MyReportInfos.HealthcheckData.GlobalScore)
        StaleObjectsScore   = [Int]($MyReportInfos.HealthcheckData.StaleObjectsScore)
        PrivilegiedGroupScore   = [Int]($MyReportInfos.HealthcheckData.PrivilegiedGroupScore)
        TrustScore  = [Int]($MyReportInfos.HealthcheckData.TrustScore)
        AnomalyScore    = [Int]($MyReportInfos.HealthcheckData.AnomalyScore)

        StaleObjectsScoreTotal  = [Int]($StaleObjectsScoreInfo.Sum)
        PrivilegiedGroupScoreTotal  = [Int]($PrivilegiedGroupScoreInfo.Sum)
        TrustScoreTotal = [Int]($TrustScoreInfo.Sum)
        AnomalyScoreTotal = [Int]($AnomalyScoreInfo.Sum)
        GlobalScoreTotal    = [Int]([Int]($AnomalyScoreInfo.Sum) + [Int]($PrivilegiedGroupScoreInfo.Sum) + [Int]($TrustScoreInfo.Sum) + [Int]($StaleObjectsScoreInfo.Sum))
        
    }
    $DataScoreInfos.Add($Obj)


}


$ReportsDate = @()
$DataPCRules | Select-Object ReportDate -Unique | Sort-Object ReportDate | ForEach-Object { $ReportsDate += "$($_.ReportDate)"}
$NbPCReports =  [int]($ReportsDate.Count)


$MyRiskCategories = $DataPCRules.Category | Select-Object -Unique | Sort-Object
$MyRiskModels = $DataPCRules.Model | Select-Object -Unique | Sort-Object
$MyRiskLevels = $DataPCRules.Level | Select-Object -Unique | Sort-Object
$MyRiskIds = $DataPCRules.RiskId | Select-Object -Unique | Sort-Object

$MyRulesByCatInfos = New-Object 'System.Collections.Generic.List[System.Object]'
$ReportsDate | ForEach-Object {
    $MaDate = [String]($_)
    ForEach ($MyInfo in $MyRiskCategories) {
        $MyRulesMatched =  (($DataPCRules | Where-Object { $_.Category -eq "$MyInfo" -and $_.ReportDate -eq "$Madate"}).Points | Measure-Object -Sum)
        $obj = [PSCustomObject]@{
            Info    = [String]($MyInfo)
            Points      = [Int]($MyRulesMatched.Sum)
            NbRules     = [Int]($MyRulesMatched.Count)
            Date        = [String]($MaDate) 
        }
        $MyRulesByCatInfos.Add($Obj)
    }
}

$MyRulesByLevelInfos = New-Object 'System.Collections.Generic.List[System.Object]'
$ReportsDate | ForEach-Object {
    $MaDate = [String]($_)
    ForEach ($MyInfo in $MyRiskLevels) {
        $MyRulesMatched =  (($DataPCRules | Where-Object { $_.Level -eq "$MyInfo" -and $_.ReportDate -eq "$Madate"}).Points | Measure-Object -Sum)
        $Obj = [PSCustomObject]@{
            Info       = [String]($MyInfo)
            Points      = [Int]($MyRulesMatched.Sum)
            NbRules     = [Int]($MyRulesMatched.Count)
            Date        = [String]($MaDate) 
        }
        $MyRulesByLevelInfos.Add($Obj)
    }
}

$MyRulesByModelInfos = New-Object 'System.Collections.Generic.List[System.Object]'
$ReportsDate | ForEach-Object {
    $MaDate = [String]($_)
    ForEach ($MyInfo in $MyRiskModels) {
        $MyRulesMatched =  (($DataPCRules | Where-Object { $_.Model -eq "$MyInfo" -and $_.ReportDate -eq "$Madate"}).Points | Measure-Object -Sum)
        $Obj = [PSCustomObject]@{
            Info       = [String]($MyInfo)
            Points      = [Int]($MyRulesMatched.Sum)
            NbRules     = [Int]($MyRulesMatched.Count)
            Date        = [String]($MaDate) 
        }
        $MyRulesByModelInfos.Add($Obj)
    }
}

$MyRulesByIdInfos = New-Object 'System.Collections.Generic.List[System.Object]'
$ReportsDate | ForEach-Object {
    $MaDate = [String]($_)
    ForEach ($MyInfo in $MyRiskIds) {
        $MyRulesMatched =  (($DataPCRules | Where-Object { $_.RiskId -eq "$MyInfo" -and $_.ReportDate -eq "$Madate"}).Points | Measure-Object -Sum)
        $Obj = [PSCustomObject]@{
            Info       = [String]($MyInfo)
            Points      = [Int]($MyRulesMatched.Sum)
            NbRules     = [Int]($MyRulesMatched.Count)
            Date        = [String]($MaDate) 
        }
        $MyRulesByIdInfos.Add($Obj)
    }
}

$MyRulesByDateInfos = New-Object 'System.Collections.Generic.List[System.Object]'
$ReportsDate | ForEach-Object {
    $MaDate = [String]($_)
    $MyRulesMatched =  (($DataPCRules | Where-Object ReportDate -eq "$Madate").Points | Measure-Object -Sum)
    $Obj = [PSCustomObject]@{
        Points      = [Int]($MyRulesMatched.Sum)
        NbRules     = [Int]($MyRulesMatched.Count)
        Date        = [String]($MaDate) 
    }
    $MyRulesByDateInfos.Add($Obj)
}

$RulesTable = New-Object 'System.Collections.Generic.List[System.Object]'
$MyRiskIds | ForEach-Object {
    $MyIdRisk = [String]($_)
    $MyRiskDetails = ($PCRulesXml.ArrayOfExportedRule.ExportedRule | Where-Object RiskId -eq "$MyIdRisk")
    $Obj = [PSCustomObject]@{
        Level       = [Int]($MyRiskDetails.MaturityLevel)
        Category    = [String]($MyRiskDetails.Category)
        Model       = [String]($MyRiskDetails.Model)
        Id          = [String]($MyIdRisk)
    }

    $MyScoreEvol =@()
    $FormerVal = 0
    ForEach ( $MyScoreDate in ($ReportsDate) ) {
        #$MyIndex = $ReportsDate.IndexOf("$MyScoreDate")
        $MyScoreDateData = $MyRulesByIdInfos | Where-Object {($_.Date -eq "$MyScoreDate") -and ($_.Info -eq "$MyIdRisk")}
        If ($MyScoreDateData.NbRules -gt 0) {
            $LastDatePoints = [Int]($($MyScoreDateData.Points))
            $MyCustVal = [Int]($($($MyScoreDateData.Points) - $FormerVal))
            $FormerVal = [Int]$($MyScoreDateData.Points)
        } else {
            $LastDatePoints = [string]("")
            $MyCustVal = [string]("")
            $FormerVal = 0
        }
        $MyScoreEvol += "$MyCustVal"
        #$obj | Add-Member -NotePropertyName "$MyScoreDate" -NotePropertyValue ([string]$($MyCustVal))
    }

    $Obj | Add-Member -NotePropertyName "LastScore" -NotePropertyValue ([string]$($LastDatePoints))

    switch ($MyCustVal) {
        { $_ -is [int] -and $_ -gt 0 } { $MyTrend = "Bad" }
        { $_ -is [int] -and $_ -lt 0 } { $MyTrend = "Happy" }
        { $_ -is [int] -and $_ -eq 0 } { $MyTrend = "Sad" }
        { $_ -eq $null -or [string]::IsNullOrEmpty($_) } { $MyTrend = "Good" }
    }

    $Obj | Add-Member -NotePropertyName "Trend" -NotePropertyValue ([string]$($MyTrend))

    $ReportsDate | ForEach-Object {
        $MyInfoDate = [String]$_
        $MyValIndex = $($ReportsDate).IndexOf($_)
        $MyInfoVal = $MyScoreEvol[$MyValIndex]
        $obj | Add-Member -NotePropertyName "$MyInfoDate" -NotePropertyValue ([string]$($MyInfoVal))
    }

    $RulesTable.Add($Obj)
}


######### Report Generation ############################
New-Item -Path "$ReportsPath" -Name "$SelectedDomain" -ItemType Directory -Force | Out-Null
New-HTML -TitleText "PingCastle History Analysis Report" -Author "$ConsultantName" -Encoding UTF8   {
    Enable-HTMLFeature -Feature FontsAwesome
      
    New-HTMLFooter -HTMLContent { "<center>&copy; $(Get-Date -Format "yyyy") - <font color=""$Violet""><a href =""$AuditorURL"" target=_blank>$AuditorCompany</a></font></center>" }
       
    New-HTMLContent -HeaderText 'Informations' -HeaderTextSize 22 -HeaderTextColor $Violet -HeaderTextAlignment center -HeaderBackGroundColor $Or {
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
            New-HTMLText  -Text "$NbPCReports Report(s)" -Color $Violet -FontSize 22 -Alignment center
            New-HTMLFontIcon -IconSolid calendar-week -IconSize 22
            New-HTMLText  -Text "From $($ReportsDate[0]) To $($ReportsDate[$NbPCReports-1])" -Color $Violet -FontSize 22 -Alignment center
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

    New-HTMLContent -HeaderText "Trends" -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
        $EvolColor =  "Grey"
        New-HTMLColumn -Width 25% -AlignContentText center {
            $MyVal = $DataScoreInfos | Sort-Object ReportDate
            $LastVal = [Int]($MyVal[-1]).GlobalScore
            $FormerVal = [Int]($MyVal[-2]).GlobalScore
            $Evol = [Int]($($LastVal - $FormerVal))
            switch ($true) {
                { $Evol -gt 0 } { $EvolStatus = "&#128545;" } #&#128545;
                { $Evol -lt 0 } { $EvolStatus = "&#128522;" } #smile-beam
                { $Evol -eq 0 } { $EvolStatus = "&#128530;" } #meh-rolling-eyes
                default { $EvolStatus = "&#128540;" } #grin-tongue-wink
            }
    
            $Evol2 = $LastVal
            switch ($Evol2) {
                0 { $EvolColor = "Blue" }
                { $_ -lt 10 } { $EvolColor = "Yellow" }
                { $_ -ge 10 -and $_ -le 30 } { $EvolColor = "Orange" }
                default { $EvolColor = "Red" }
            }
            New-HTMLContent -AlignContent center -HeaderTextAlignment center -HeaderText "$EvolStatus" -HeaderBackGroundColor "$EvolColor" -HeaderTextSize 42{
                New-HTMLText -Text "PingCastle Score : $LastVal" -FontSize 33 -Alignment center
            }

        }
        New-HTMLColumn -Width 25% -AlignContentText center{
            $MyVal = $DataScoreInfos | Sort-Object ReportDate
            $LastVal = [Int]($MyVal[-1]).MaturityLevel
            $FormerVal = [Int]($MyVal[-2]).MaturityLevel
            $Evol = [Int]($($LastVal - $FormerVal))
            switch ($true) {
                { $Evol -gt 0 } { $EvolStatus = "&#128545;"} #&#128545;
                { $Evol -lt 0 } { $EvolStatus = "&#128522;"} #smile-beam
                { $Evol -eq 0 } { $EvolStatus = "&#128530;"} #meh-rolling-eyes
                default { $EvolStatus = "&#128540;"} #grin-tongue-wink
            }
    
    
            New-HTMLContent -AlignContent center -HeaderTextAlignment center -HeaderText "$EvolStatus" -HeaderBackGroundColor $LevelColors[$LastVal] -HeaderTextSize 42{
                New-HTMLText -Text "Maturity Level : $LastVal" -FontSize 33 -Alignment center
            }
        }
        New-HTMLColumn -Width 25% -AlignContentText center{
            $MyVal = $DataScoreInfos | Sort-Object ReportDate
            $LastVal = [Int]($MyVal[-1]).GlobalScoreTotal
            $FormerVal = [Int]($MyVal[-2]).GlobalScoreTotal
            $Evol = [Int]($($LastVal - $FormerVal))
            switch ($true) {
                { $Evol -gt 0 } { $EvolStatus = "&#128545;" } #&#128545;
                { $Evol -lt 0 } { $EvolStatus = "&#128522;" } #smile-beam
                { $Evol -eq 0 } { $EvolStatus = "&#128530;" } #meh-rolling-eyes
                default { $EvolStatus = "&#128540;" } #grin-tongue-wink
            }
    
            $Evol2 = $LastVal
            switch ($Evol2) {
                0 { $EvolColor = "Blue" }
                { $_ -lt 10 } { $EvolColor = "Yellow" }
                { $_ -ge 10 -and $_ -le 30 } { $EvolColor = "Orange" }
                default { $EvolColor = "Red" }
            }
            New-HTMLContent -AlignContent center -HeaderTextAlignment center -HeaderText "$EvolStatus" -HeaderBackGroundColor "$EvolColor" -HeaderTextSize 42{
                New-HTMLText -Text "Global Score Level : $LastVal" -FontSize 33 -Alignment center
            }
        }
        New-HTMLColumn -Width 25% -AlignContentText center{
            $LastVal = [Int]($MyRulesByDateInfos | Where-Object Date -eq $ReportsDate[-1]).NbRules
            $FormerVal = [Int]($MyRulesByDateInfos | Where-Object Date -eq $ReportsDate[-2]).NbRules
            $Evol = [Int]($LastVal - $FormerVal)
            switch ($true) {
                { $Evol -gt 0 } { $EvolStatus = "&#128545;" } #&#128545;
                { $Evol -lt 0 } { $EvolStatus = "&#128522;" } #smile-beam
                { $Evol -eq 0 } { $EvolStatus = "&#128530;" } #meh-rolling-eyes
                default { $EvolStatus = "&#128540;" } #grin-tongue-wink
            }
  
            $Evol2 = $LastVal
            switch ($Evol2) {
                0 { $EvolColor = "Blue" }
                { $_ -lt 10 } { $EvolColor = "Yellow" }
                { $_ -ge 10 -and $_ -le 30 } { $EvolColor = "Orange" }
                default { $EvolColor = "Red" }
            }
            New-HTMLContent -AlignContent center -HeaderTextAlignment center -HeaderText "$EvolStatus" -HeaderBackGroundColor "$EvolColor" -HeaderTextSize 42{
                New-HTMLText -Text "Matched Rules : $LastVal" -FontSize 33 -Alignment center
            }

        }
    }
    New-HTMLContent -HeaderText "Maturity" -HeaderTextSize 22 -HeaderTextColor $Violet -HeaderTextAlignment center -HeaderBackGroundColor $Or {
        
        New-HTMLChart -Title 'Maturity Level Evolution' -TitleAlignment center {
            New-ChartToolbar -Download
            New-ChartAxisX -Names $ReportsDate

            $MyMatchingRules = $DataScoreInfos | Sort-Object ReportDate
            New-ChartLine -Name "MaturityLevel" -Value $($MyMatchingRules.MaturityLevel) -Cap square -Curve smooth -Dash 1 #-Color $LevelColors[$LastVal]
            New-ChartAxisY -Show -MinValue 0 -MaxValue 5
        }
        

    }

    New-HTMLContent -HeaderText "Scores" -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {

        New-HTMLColumn -AlignContentText center -Width 50% {

            New-HTMLChart -Title 'Real Global Scores' -TitleAlignment center {
                New-ChartToolbar -Download
                    
                New-ChartAxisX -Names $ReportsDate
                #New-ChartLine -Name 'Global Score' -Value $GlobalScoreTotalTab -Cap square -Curve smooth -Color $DisColor
                $MinVal = [Int]($DataScoreInfos | Sort-Object GlobalScoreTotal | Select-Object -First 1).GlobalScoreTotal
                $MaxVal = [Int]($DataScoreInfos | Sort-Object GlobalScoreTotal | Select-Object -Last 1).GlobalScoreTotal
                $MyMatchingRules = $DataScoreInfos | Sort-Object ReportDate
                New-ChartLine -Name "GlobalScoreTotal" -Value $($MyMatchingRules.GlobalScoreTotal) -Cap square -Curve smooth #-Color $EvolColor
                New-ChartAxisY -Show -MinValue $MinVal -MaxValue $MaxVal
            }
        }

        New-HTMLColumn -AlignContentText center -Width 50% {

            New-HTMLChart -Title 'PingCastle Official' -TitleAlignment center {
                New-ChartToolbar -Download
                New-ChartAxisX  -Names $ReportsDate #-Type datetime
                $MyMatchingRules = $DataScoreInfos | Sort-Object ReportDate

                New-ChartLine -Name "Anomaly" -Value $($MyMatchingRules.("AnomalyScore")) -Cap square -Curve smooth
                New-ChartLine -Name "Privilegied Group" -Value $($MyMatchingRules.("PrivilegiedGroupScore")) -Cap square -Curve smooth
                New-ChartLine -Name "Stale Objects" -Value $($MyMatchingRules.("StaleObjectsScore")) -Cap square -Curve smooth
                New-ChartLine -Name "Trusts" -Value $($MyMatchingRules.("TrustScore")) -Cap square -Curve smooth
                New-ChartLine -Name "Global Score" -Value $($MyMatchingRules.("GlobalScore")) -Cap square -Curve smooth
                New-ChartAxisY -Show -MinValue 0 -MaxValue 100
            }

        }


        
    }

    New-HTMLContent -HeaderText "Category" -HeaderTextSize 22 -HeaderTextColor $Violet -HeaderTextAlignment center -HeaderBackGroundColor $Or {
        New-HTMLChart -Title 'Score By Category' -TitleAlignment center {
            New-ChartToolbar -Download
            New-ChartAxisX -Names $ReportsDate #-Type datetime
            $MinVal = [Int]($MyRulesByCatInfos | Sort-Object Points | Select-Object -First 1).Points
            $MaxVal = [Int]($MyRulesByCatInfos | Sort-Object Points | Select-Object -Last 1).Points
            $MyRiskCategories | ForEach-Object {
                $MyLineName = [String]($_)
                $MyMatchingRules = $MyRulesByCatInfos | Sort-Object Date | Where-Object Info -eq "$MyLineName"
                New-ChartLine -Name "$MyLineName" -Value $($MyMatchingRules.Points) -Cap square -Curve smooth
            }

            New-ChartAxisY -Show -MinValue $MinVal -MaxValue $MaxVal
    
        }
        New-HTMLChart -Title 'Rules Nb By category' -TitleAlignment center  {
            New-ChartToolbar -Download
            New-ChartAxisX -Names $ReportsDate #-Type datetime
            $MinVal = [Int]($MyRulesByCatInfos | Sort-Object NbRules | Select-Object -First 1).NbRules
            $MaxVal = [Int]($MyRulesByCatInfos | Sort-Object NbRules | Select-Object -Last 1).NbRules
            $MyRiskCategories | ForEach-Object {
                $MyLineName = [String]($_)
                $MyMatchingRules = $MyRulesByCatInfos | Sort-Object Date | Where-Object Info -eq "$MyLineName"
                New-ChartLine -Name "$MyLineName" -Value $($MyMatchingRules.NbRules) -Cap square -Curve smooth
            }
            New-ChartAxisY -Show -MinValue $MinVal -MaxValue $MaxVal
        }  
}

    New-HTMLContent -HeaderText "Levels" -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
            New-HTMLChart -Title 'Score By Levels' -TitleAlignment center  {
                New-ChartToolbar -Download
                New-ChartAxisX -Names $ReportsDate #-Type datetime
                $MinVal = [Int]($MyRulesByLevelInfos | Sort-Object Points | Select-Object -First 1).Points
                $MaxVal = [Int]($MyRulesByLevelInfos | Sort-Object Points | Select-Object -Last 1).Points
                $MyRiskLevels | ForEach-Object {
                    $MyLineName = [String]($_)
                    $MyMatchingRules = $MyRulesByLevelInfos | Sort-Object Date | Where-Object Info -eq "$MyLineName"
                    New-ChartLine -Name "$MyLineName" -Value $($MyMatchingRules.Points) -Cap square -Curve smooth
                }
                New-ChartAxisY -Show -MinValue $MinVal -MaxValue $MaxVal
            } 
            New-HTMLChart -Title 'Rules Nb By Levels' -TitleAlignment center  {
                New-ChartToolbar -Download
                New-ChartAxisX -Names $ReportsDate #-Type datetime
                $MinVal = [Int]($MyRulesByLevelInfos | Sort-Object NbRules | Select-Object -First 1).NbRules
                $MaxVal = [Int]($MyRulesByLevelInfos | Sort-Object NbRules | Select-Object -Last 1).NbRules
                $MyRiskLevels | ForEach-Object {
                    $MyLineName = [String]($_)
                    $MyMatchingRules = $MyRulesByLevelInfos | Sort-Object Date | Where-Object Info -eq "$MyLineName"
                    New-ChartLine -Name "$MyLineName" -Value $($MyMatchingRules.NbRules) -Cap square -Curve smooth
                }
                New-ChartAxisY -Show -MinValue $MinVal -MaxValue $MaxVal
            }  
    }

    New-HTMLContent -HeaderText "Risk Models" -HeaderTextSize 22 -HeaderTextColor $Violet -HeaderTextAlignment center -HeaderBackGroundColor $Or {
            New-HTMLChart -Title 'Score By Risk Models' -TitleAlignment center  {
                New-ChartToolbar -Download
                New-ChartAxisX -Names $ReportsDate #-Type datetime
                $MinVal = [Int]($MyRulesByModelInfos | Sort-Object Points | Select-Object -First 1).Points
                $MaxVal = [Int]($MyRulesByModelInfos | Sort-Object Points | Select-Object -Last 1).Points
                $MyRiskModels | ForEach-Object {
                    $MyLineName = [String]($_)
                    $MyMatchingRules = $MyRulesByModelInfos | Sort-Object Date | Where-Object Info -eq "$MyLineName"
                    New-ChartLine -Name "$MyLineName" -Value $($MyMatchingRules.Points) -Cap square -Curve smooth
                }
                New-ChartAxisY -Show -MinValue $MinVal -MaxValue $MaxVal
            } 
            New-HTMLChart -Title 'Rules Nb By Risk Models' -TitleAlignment center  {
                New-ChartToolbar -Download
                New-ChartAxisX -Names $ReportsDate #-Type datetime
                $MinVal = [Int]($MyRulesByModelInfos | Sort-Object NbRules | Select-Object -First 1).NbRules
                $MaxVal = [Int]($MyRulesByModelInfos | Sort-Object NbRules | Select-Object -Last 1).NbRules
                $MyRiskModels | ForEach-Object {
                    $MyLineName = [String]($_)
                    $MyMatchingRules = $MyRulesByModelInfos | Sort-Object Date | Where-Object Info -eq "$MyLineName"
                    New-ChartLine -Name "$MyLineName" -Value $($MyMatchingRules.NbRules) -Cap square -Curve smooth
                }
                New-ChartAxisY -Show -MinValue $MinVal -MaxValue $MaxVal
            }  
    }

    New-HTMLContent -HeaderText "Rules" -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
        New-HTMLColumn -AlignContentText center -Width 100% {

            New-HTMLChart -Title 'Matched Rules' -TitleAlignment center {
                New-ChartToolbar -Download
                New-ChartAxisX -Names $ReportsDate #-Type datetime
                $MinVal = [Int]($MyRulesByDateInfos | Sort-Object NbRules | Select-Object -First 1).NbRules
                $MaxVal = [Int]($MyRulesByDateInfos | Sort-Object NbRules | Select-Object -Last 1).NbRules
                $MyMatchingRules = $MyRulesByDateInfos | Sort-Object Date
                New-ChartLine -Name 'RulesNb' -Value $($MyMatchingRules.NbRules) -Cap square -Curve smooth -Color $EvolColor
                New-ChartAxisY -Show -MinValue $MinVal -MaxValue $MaxVal
            }



            New-HTMLTable -Title "All Ever Matched Rules" -DataTable $RulesTable -HideFooter -HideButtons -DisablePaging  -FreezeColumnsLeft 6 -ScrollX { #-EnableScroller

                for ($i = 1; $i -le 5; $i++) { New-HTMLTableCondition -Name "Level" -ComparisonType number -Operator eq -Value $i -BackgroundColor  "$($LevelColors[$i])" -Row }

                New-HTMLTableCondition -Name "Trend" -ComparisonType string -Operator ne -Value "" -BackgroundColor White -Color Black -FontSize 30
                New-HTMLTableCondition -Name "LastScore" -ComparisonType number -Operator gt -Value 30 -BackgroundColor Red -Color Black -FontWeight bold -FontSize 22
                New-HTMLTableCondition -Name "LastScore" -ComparisonType number -Operator betweenInclusive -Value 10,30 -BackgroundColor Orange -Color Black -FontWeight bold -FontSize 22
                New-HTMLTableCondition -Name "LastScore" -ComparisonType number -Operator between -Value 0,10 -BackgroundColor Yellow -Color Black -FontWeight bold -FontSize 22
                New-HTMLTableCondition -Name "LastScore" -ComparisonType number -Operator eq -Value 0 -BackgroundColor Blue -Color Black -FontWeight bold -FontSize 22
                New-HTMLTableCondition -Name "LastScore" -ComparisonType string -Operator eq -Value "" -BackgroundColor Green -Color Black -FontWeight bold -FontSize 22


                $ReportsDate | ForEach-Object {
                    New-HTMLTableCondition -Name "$_" -ComparisonType number -Operator gt -Value 0 -BackgroundColor Crimson -Color Black -FontWeight bold -FontSize 12
                    New-HTMLTableCondition -Name "$_" -ComparisonType number -Operator lt -Value 0 -BackgroundColor MediumSeaGreen -Color Black -FontWeight bold -FontSize 12
                    New-HTMLTableCondition -Name "$_" -ComparisonType string -Operator eq -Value "" -BackgroundColor DarkSeaGreen -FontSize 12
                    New-HTMLTableCondition -Name "$_" -ComparisonType number -Operator eq -Value 0 -BackgroundColor LightSalmon -Color Black -FontWeight bold -FontSize 12
                }

                $RulesTable | ForEach-Object {
                    $MyIndexRule = $RulesTable.IndexOf($_) + 1
                    $MonStatut = ($_.Trend)
                    switch ($MonStatut) {
                        #https://www.w3schools.com/charsets/ref_emoji_smileys.asp
                        "Bad" {  New-HTMLTableContent -ColumnName "Trend" -RowIndex $MyIndexRule -Text "&#128545;"}
                        "Happy" {  New-HTMLTableContent -ColumnName "Trend" -RowIndex $MyIndexRule -Text "&#128522;"}
                        "Sad" {  New-HTMLTableContent -ColumnName "Trend" -RowIndex $MyIndexRule -Text "&#128530;"}
                        "Good" {  New-HTMLTableContent -ColumnName "Trend" -RowIndex $MyIndexRule -Text "&#128540;"}
                    }
                }
                #>
            }
#>                
        }


    }

} -ShowHTML -FilePath "$HistoryReportFileName"

######### End Report Generation ########################






