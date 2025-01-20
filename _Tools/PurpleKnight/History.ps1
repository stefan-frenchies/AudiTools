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
$LevelColors += "Blue"
$LevelColors += "Green"


Function ScoreNoteInfo {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateRange(0, 100)]
        [int]$number
    )
    Switch ($number) {
        {$_ -le 43} { $MyNote = "F"; $MyNoteColor = "#a80116" }
        {$_ -le 57 -and $_ -ge 44} { $MyNote = "D-"; $MyNoteColor = "#cc021b" }
        {$_ -le 66 -and $_ -ge 58} { $MyNote = "D"; $MyNoteColor = "#e0011d" }
        {$_ -le 74 -and $_ -ge 67} { $MyNote = "D+"; $MyNoteColor = "#ff2843" }
        {$_ -le 80 -and $_ -ge 75} { $MyNote = "C-"; $MyNoteColor = "#b94302" }
        {$_ -le 85 -and $_ -ge 81} { $MyNote = "C"; $MyNoteColor = "#d04c02" }
        {$_ -le 89 -and $_ -ge 86} { $MyNote = "C+"; $MyNoteColor = "#e75402" }
        {$_ -le 92 -and $_ -ge 90} { $MyNote = "B-"; $MyNoteColor = "#ba8808" }
        {$_ -le 95 -and $_ -ge 93} { $MyNote = "B"; $MyNoteColor = "#d39e16" }
        {$_ -le 97 -and $_ -ge 96} { $MyNote = "B+"; $MyNoteColor = "#eaa801" }
        {$_ -eq 98} { $MyNote = "A-"; $MyNoteColor = "#39b221" }
        {$_ -eq 99} { $MyNote = "A"; $MyNoteColor = "#1b9a00" }
        {$_ -eq 100} { $MyNote = "A+"; $MyNoteColor = "#147400" }
    }
    
    return "$MyNote;$MyNoteColor"
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Function DrawMyScore {
    param (
        [Parameter(Position=0, Mandatory=$true)]
        [int]$Score,
        [Parameter(Position=1, Mandatory=$true)]
        [string]$ImageName,
        [Parameter(Position=2, Mandatory=$true)]
        [string]$Note,
        [Parameter(Position=3, Mandatory=$true)]
        [string]$Color
    )
    $width = 120
    $height= 120
    $image = New-Object System.Drawing.Bitmap $width,$height
    $image.MakeTransparent()
    $graphics = [System.Drawing.Graphics]::FromImage($image)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $centerX = $width /2
    $centerY = $height /2
    $radius = $width /2
    $startAngle = -90
    $sweepAngle = $Score*3.6
    $pen = [System.Drawing.Pens]::Black
    $backgroundColor = [System.Drawing.ColorTranslator]::FromHtml($Color)
    $brush = New-Object System.Drawing.SolidBrush $backgroundColor

    #$brush = [System.Drawing.Brushes]::$Color
    $graphics.DrawEllipse($pen, $centerX - $radius, $centerY - $radius, $radius * 2, $radius * 2)
    $graphics.FillPie($brush, $centerX - $radius, $centerY - $radius, $radius * 2, $radius * 2, $startAngle, $sweepAngle)

    $textColor = [System.Drawing.Color]::Black
    $font = New-Object System.Drawing.Font("Arial", 20)
    $textBrush = New-Object System.Drawing.SolidBrush $textColor
    $stringFormat = New-Object System.Drawing.StringFormat
    $stringFormat.Alignment = [System.Drawing.StringAlignment]::Center
    $stringFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
    $graphics.DrawString($Note, $font, $textBrush, [System.Drawing.RectangleF]::new(0, 0, $width, $height), $stringFormat)

    $image.Save("$DatasPath\$MyPKReportDomain\$MyPKReportDate\$ImageName.png", [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $image.Dispose()

}


$ToolName = "PurpleKnight"

$DatasPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Datas\$ToolName").FullName)
New-Item -Path "$PSScriptRoot\..\_Reports" -Name "$ToolName" -ItemType Directory -Force | Out-Null
$ReportsPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Reports\$ToolName").FullName)
$ImagesPath = [String]((Get-Item -Path "$PSScriptRoot\images").FullName)
#$ReportsPath = "C:\_Metsys\_Reports\PurpleKnight\cultura.intra"

$SourcesPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Sources").FullName)


$PKXMLFileInfos = New-Object 'System.Collections.Generic.List[System.Object]'
Get-ChildItem -Path "$DatasPath" -Include "SecurityAssessmentReport_*.xml" -Recurse | ForEach-Object {
    $MyPKReportXMLFilePath = [String]($($_.FullName))
    [xml]$PKXMLInfos = Get-Content "$MyPKReportXMLFilePath"
    $MyPKReportDomain = "$($PKXMLInfos.Report.Info.ADforestname)"
    $MyPKReportDate = ([DateTime]$($PKXMLInfos.Report.Info.Generateddate)).ToString('yyyyMMdd')
    $obj = [PSCustomObject]@{
        XMLFile = [String]($MyPKReportXMLFilePath)
        Date    = [String]($MyPKReportDate)
        Domain  = [String]($MyPKReportDomain)
    }
    $PKXMLFileInfos.Add($Obj)
}

$AllDomains = $PKXMLFileInfos.Domain | Select-Object -Unique


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
$HistoryOnDomain = $PKXMLFileInfos | Where-Object Domain -eq "$MyDomain" | Sort-Object Date

Write-Host "History Reporting From $($HistoryOnDomain.Date.Count) Reports On $MyDomain"


#$HistoryReportFileName = "C:\_Metsys\_Reports\$MyDomain\PurpleKnightHistory-$MyDomain.html"
$HistoryReportFileName = "$ReportsPath\$MyDomain\PurpleKnightHistory-$MyDomain.html"

$InfosPKReport = New-Object 'System.Collections.Generic.List[System.Object]'
$DataPKIOEs = New-Object 'System.Collections.Generic.List[System.Object]'

$HistoryOnDomain | ForEach-Object {
    $MyPKReportXMLFilePath = [String]($($_.XMLFile))

#    $MyPKReportXMLFilePath = "C:\_Metsys\_Reports\PurpleKnight\cultura.intra\20240718\SecurityAssessmentReport_20240718.xml"

    [xml]$PKReportXML = Get-Content -Path "$MyPKReportXMLFilePath"

    $MyPKReportDomain = "$($PKReportXML.Report.Info.ADforestname)"
    $MyPKReportDate = ([DateTime]$($PKReportXML.Report.Info.Generateddate)).ToString('yyyyMMdd')

    Write-Host "Reading Report For PurpleKnight Analysis of $MyPKReportDate From $MyPKReportXMLFilePath" -ForegroundColor Blue

    $PKReportInfos = $PKReportXML.Report.Info
    $PKReportRisks = $PKReportXML.Report.Risk
    $Obj = [PSCustomObject]@{ReportDate  = [String]($MyPKReportDate)}
    $PKReportInfosProps  = ($PKReportInfos | Get-Member -MemberType "Property").Name

    $PKReportInfosProps | ForEach-Object {
        $MyProp = [String]($_)
        $MyValue = [String]($PKReportInfos.("$MyProp"))
        $obj | Add-Member -NotePropertyName "$MyProp" -NotePropertyValue "$([String]$($MyValue))"
        
    }
    $InfosPKReport.Add($Obj)


    $PKReportRisks | ForEach-Object {
        switch ($_.Severity) {
            "Critical" { $MyLevel = 1 }
            "Warning" { $MyLevel = 2 }
            "Informational" { $MyLevel = 3 }
            Default { $MyLevel = 4}
        }
        $Obj2 = [PSCustomObject]@{
            ReportDate  = [String]($MyPKReportDate)
            ShortName       = [String]$($_.ShortName)
            Name            = [String]$($_.Name)
            Score           = [int]$($_.Score)
            Severity        = [String]$($_.Severity)
            Level           = [Int]$($MyLevel)
            Status          = [String]$($_.Status.replace(" Found",""))
            Target          = [String]$($_.Target)
            #Weight          = [int]$($_.Weight)
            Category        = [String]$($_.Category)
            Description     = [String]$($_.Description)
            Numberofresults = [int]$($_.Numberofresults)
        }
        $DataPKIOEs.Add($Obj2)
    }
}

$ReportsDate = @()
$InfosPKReport | Select-Object ReportDate -Unique | Sort-Object ReportDate | ForEach-Object { $ReportsDate += "$($_.ReportDate)"}
$NbPKReports =  [int]($ReportsDate.Count)


$MyRiskCategories = $DataPKIOEs.Category | Select-Object -Unique | Sort-Object
$MyRiskModels = $DataPKIOEs.Target | Select-Object -Unique | Sort-Object
$MyRiskLevels = $DataPKIOEs.Severity | Select-Object -Unique | Sort-Object
$MyRiskIds = $DataPKIOEs.ShortName | Select-Object -Unique | Sort-Object








$MyRulesByIdInfos = New-Object 'System.Collections.Generic.List[System.Object]'
$ReportsDate | ForEach-Object {
    $MaDate = [String]($_)
    ForEach ($MyInfo in $MyRiskIds) {
        $MyRulesMatched =  (($DataPKIOEs | Where-Object { $_.ShortName -eq "$MyInfo" -and $_.ReportDate -eq "$Madate"}).Score | Measure-Object -Sum)
        $Obj = [PSCustomObject]@{
            Info       = [String]($MyInfo)
            Points      = [Int]($MyRulesMatched.Sum)
            NbRules     = [Int]($MyRulesMatched.Count)
            Date        = [String]($MaDate) 
        }
        $MyRulesByIdInfos.Add($Obj)
    }
}

$RulesTable = New-Object 'System.Collections.Generic.List[System.Object]'
$MyRiskIds | ForEach-Object {
    $MyIdRisk = [String]($_)
    $MyRiskDetails = ($DataPKIOEs | Where-Object ShortName -eq "$MyIdRisk" | Select-Object -First 1)
    $Obj = [PSCustomObject]@{
        Level       = [Int]($MyRiskDetails.Level)
        Category    = [String]($MyRiskDetails.Category)
        Target       = [String]($MyRiskDetails.Target)
        #Weight      = [Int]($MyRiskDetails.Weight) 
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



New-HTML -TitleText "PurpleKnight History Analysis Report" -Author "$ConsultantName" -Encoding UTF8   {
    Enable-HTMLFeature -Feature FontsAwesome
      
    New-HTMLFooter -HTMLContent { "<center>&copy; $(Get-Date -Format "yyyy") - <font color=""$Violet""><a href =""$AuditorURL"" target=_blank>$AuditorCompany</a></font></center>" }
    New-HTMLContent -HeaderText '$MyPKReportDomain' -HeaderTextSize 33 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
    
        New-HTMLColumn -Width 33% {
            New-HTMLImage -Source "$AuditorLogo" -Width "200" -Inline -UrlLink "$AuditorURL"
            New-HTMLFontIcon -IconSolid address-book
            New-HTMLHeading h2 -HeadingText "$($AuditorCompany.ToUpper())"
            New-HTMLFontIcon -IconSolid address-card 
            New-HTMLText -Text "$ConsultantName"
            New-HTMLFontIcon -IconSolid phone
            New-HTMLText -Text "$ConsultantPhone"
            New-HTMLFontIcon -IconSolid envelope
            New-HTMLText -Text "$ConsultantMail"
        } -AlignContentText center

        New-HTMLColumn -Width 33% -BackgroundColor $Violet {

            New-HTMLImage -Source "$ImagesPath\spklogo.png" -Width "160" -Inline -UrlLink "https://www.purple-knight.com" -Target _blank
            New-HTMLFontIcon -IconSolid info-circle -IconSize 22
            New-HTMLText  -Text "Purple Knight v$($InfosPKReport.ToolVersion | Select-Object -Last 1)" -Color $Or -FontSize 22 -Alignment center
            New-HTMLFontIcon -IconSolid calendar-day -IconSize 22
            New-HTMLText  -Text "$($InfosPKReport.Count) Reports From $($InfosPKReport.ReportDate | Select-Object -First 1) To $($InfosPKReport.ReportDate | Select-Object -Last 1)" -Color $Or -FontSize 22 -Alignment center

        } -AlignContentText center

        New-HTMLColumn -Width 33% {
            New-HTMLImage -Source "$ClientLogo" -Width "200" -Inline -UrlLink "https://www.cultura.fr" -Target _blank
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

    $LastInfosPKReport = $InfosPKReport | Sort-Object ReportDate | Select-Object -Last 1
    New-HTMLContent -HeaderText "Audit Targets" -HeaderTextSize 22 -HeaderTextColor $Violet -HeaderTextAlignment center -HeaderBackGroundColor $Or {
        New-HTMLColumn -Width 33% {
            New-HTMLText -Text "Active Directory"
            New-HTMLText -Text "Forest Name : $($LastInfosPKReport.ADforestname)"  
            New-HTMLText -Text "Number of Domains : $($LastInfosPKReport.ADnumberofdomains)" 
            New-HTMLText -Text "Scan Duration : $($LastInfosPKReport.ADscanduration)"
            New-HTMLText -Text "Requested by : $($LastInfosPKReport.ADrequestedby)" 
            New-HTMLText -Text "Evaluated Rules: $($LastInfosPKReport.ADEvaluated)" 
            New-HTMLText -Text "Unrelevant Rules: $($LastInfosPKReport.ADNotrelevant)"
            New-HTMLText -Text "Failed Rules : $($LastInfosPKReport.ADFailedtorun)" 
            New-HTMLText -Text "Not Selected Rules : $($LastInfosPKReport.ADNotselected)" 
        } -AlignContentText center
        New-HTMLColumn -Width 33% {
            New-HTMLText -Text "Entra"
            New-HTMLText -Text "Entra Tenant : $($LastInfosPKReport.EntraIDtenant)"
            New-HTMLText -Text "Evaluated Rules: $($LastInfosPKReport.EntraIDEvaluated)"
            New-HTMLText -Text "Unrelevant Rules: $($LastInfosPKReport.EntraIDNotrelevant)"
            New-HTMLText -Text "Failed Rules : $($LastInfosPKReport.EntraIDFailedtorun)"
            New-HTMLText -Text "Not Selected Rules : $($LastInfosPKReport.EntraIDNotselected)"
        } -AlignContentText center

        New-HTMLColumn -Width 33% {
            New-HTMLText -Text "Okta"
            New-HTMLText -Text "Okta Domain URL : $($LastInfosPKReport.OktadomainURL)"
            New-HTMLText -Text "Evaluated Rules: $($LastInfosPKReport.OktaEvaluated)"
            New-HTMLText -Text "Unrelevant Rules: $($LastInfosPKReport.OktaNotrelevant)"
            New-HTMLText -Text "Failed Rules : $($LastInfosPKReport.OktaFailedtorun)"
            New-HTMLText -Text "Not Selected Rules : $($LastInfosPKReport.OktaNotselected)"

        } -AlignContentText center




    }

    New-HTMLContent -HeaderText "Trends" -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
        $EvolColor =  "Grey"
        New-HTMLColumn -Width 25% -AlignContentText center {
            $MyVal = $InfosPKReport | Sort-Object ReportDate
            $LastVal = [Int]($MyVal[-1]).ADsecurityposturescore
            $FormerVal = [Int]($MyVal[-2]).ADsecurityposturescore
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
                New-HTMLText -Text "AD Score : $LastVal" -FontSize 33 -Alignment center
            }

        }
        $EvolColor =  "Grey"
        New-HTMLColumn -Width 25% -AlignContentText center {
            $MyVal = $InfosPKReport | Sort-Object ReportDate
            
            $LastVal = [Int]($(($MyVal[-1]).ADIOEsfound) + $(($MyVal[-1]).EntraIDIOEsfound) + $(($MyVal[-1]).OktaIOEsfound))
            $FormerVal = [Int]($(($MyVal[-2]).ADIOEsfound) + $(($MyVal[-2]).EntraIDIOEsfound) + $(($MyVal[-2]).OktaIOEsfound))
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
                New-HTMLText -Text "Found Rules : $LastVal" -FontSize 33 -Alignment center
            }

        }
        $EvolColor =  "Grey"
        New-HTMLColumn -Width 25% -AlignContentText center {
            $MyVal = $InfosPKReport | Sort-Object ReportDate
            $LastVal = [Int]($MyVal[-1]).EntraIDsecurityposturescore
            $FormerVal = [Int]($MyVal[-2]).EntraIDsecurityposturescore
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
                New-HTMLText -Text "ENTRA Score : $LastVal" -FontSize 33 -Alignment center
            }

        }
        $EvolColor =  "Grey"
        New-HTMLColumn -Width 25% -AlignContentText center {
            $MyVal = $InfosPKReport | Sort-Object ReportDate
            $LastVal = [Int]($MyVal[-1]).Oktasecurityposturescore
            $FormerVal = [Int]($MyVal[-2]).Oktasecurityposturescore
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
                New-HTMLText -Text "OKTA Score : $LastVal" -FontSize 33 -Alignment center
            }

        }

    }

    New-HTMLContent -HeaderText "Score Evolution" -HeaderTextSize 22 -HeaderTextColor $Violet -HeaderTextAlignment center -HeaderBackGroundColor $Or {
        New-HTMLChart -Title 'Active Directory' -TitleAlignment center {
            New-ChartToolbar -Download
            New-ChartAxisX -Names $ReportsDate
            $MyPkReportInfo = $InfosPKReport | Sort-Object ReportDate
            $MyLines= @("ADcategoryscoreAccountSecurity","ADcategoryscoreADDelegation","ADcategoryscoreADInfrastructureSecurity","ADcategoryscoreGroupPolicySecurity","ADcategoryscoreHybrid","ADcategoryscoreKerberosSecurity")
            $MyLines | ForEach-Object {
                New-ChartLine -Name "$($_)" -Value $($MyPkReportInfo.("$($_)")) -Cap square -Curve smooth -Dash 1 #-Color $LevelColors[$LastVal]
            }
            #New-ChartAxisY -Show -MinValue 0 -MaxValue 5
        }
        New-HTMLChart -Title 'Entra / Okta' -TitleAlignment center {
            New-ChartToolbar -Download
            New-ChartAxisX -Names $ReportsDate
            $MyPkReportInfo = $InfosPKReport | Sort-Object ReportDate
            $MyLines= @("EntraIDcategoryscoreEntraID","EntraIDcategoryscoreHybrid","OktacategoryscoreOkta")
            $MyLines | ForEach-Object {
                New-ChartLine -Name "$($_)" -Value $($MyPkReportInfo.("$($_)")) -Cap square -Curve smooth -Dash 1 #-Color $LevelColors[$LastVal]
            }
            #New-ChartAxisY -Show -MinValue 0 -MaxValue 5
        }    
    }
    New-HTMLContent -HeaderText "Risks" -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
        New-HTMLColumn -AlignContentText center -Width 100% {
            New-HTMLChart -Title 'Evaluated / Found' -TitleAlignment center {
                New-ChartToolbar -Download
                New-ChartAxisX -Names $ReportsDate
                $MyPkReportInfo = $InfosPKReport | Sort-Object ReportDate
                $MyLines= @("ADEvaluated","ADIOEsfound","EntraIDEvaluated","EntraIDIOEsfound","OktaEvaluated","OktaIOEsfound")
                $MyLines | ForEach-Object {
                    New-ChartLine -Name "$($_)" -Value $($MyPkReportInfo.("$($_)")) -Cap square -Curve smooth -Dash 1 #-Color $LevelColors[$LastVal]
                }
                
    
                #New-ChartAxisY -Show -MinValue 0 -MaxValue 5
            }     

            New-HTMLTable -Title "All Ever Matched Rules" -DataTable $RulesTable -HideFooter -HideButtons -DisablePaging  -FreezeColumnsLeft 6 -ScrollX { #-EnableScroller
                for ($i = 1; $i -le 5; $i++) { New-HTMLTableCondition -Name "Level" -ComparisonType number -Operator eq -Value $i -BackgroundColor  "$($LevelColors[$i])" -Row }

                New-HTMLTableCondition -Name "Trend" -ComparisonType string -Operator ne -Value "" -BackgroundColor White -Color Black -FontSize 30
                New-HTMLTableCondition -Name "LastScore" -ComparisonType string -Operator ne -Value "" -FontWeight bold -FontSize 22

                <#
                New-HTMLTableCondition -Name "LastScore" -ComparisonType number -Operator gt -Value 30 -BackgroundColor Red -Color Black -FontWeight bold -FontSize 22
                New-HTMLTableCondition -Name "LastScore" -ComparisonType number -Operator betweenInclusive -Value 10,30 -BackgroundColor Orange -Color Black -FontWeight bold -FontSize 22
                New-HTMLTableCondition -Name "LastScore" -ComparisonType number -Operator between -Value 0,10 -BackgroundColor Yellow -Color Black -FontWeight bold -FontSize 22
                New-HTMLTableCondition -Name "LastScore" -ComparisonType number -Operator eq -Value 0 -BackgroundColor Blue -Color Black -FontWeight bold -FontSize 22
                New-HTMLTableCondition -Name "LastScore" -ComparisonType string -Operator eq -Value "" -BackgroundColor Green -Color Black -FontWeight bold -FontSize 22
#>

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

            }


        }
    }

    
} -ShowHTML -FilePath "$HistoryReportFileName"