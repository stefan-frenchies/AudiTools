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

Clear-Host

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



#$ReportsPath\$MyPKReportDomain\$MyPKReportDate
$ImagesPath = [String]((Get-Item -Path "$PSScriptRoot\images").FullName)
#$ReportsPath = "C:\_Metsys\_Reports"
$RemediatePath = [String]((Get-Item -Path "$PSScriptRoot\..\_Remediate").FullName)
$SourcesPath = [String]((Get-Item -Path "$PSScriptRoot\..\_Sources").FullName)

Write-Host "Analysing Paths , Please Wait..." -ForegroundColor Cyan
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

#########
If ($Domain -eq "*") {
    $AllDomains = $True
    $ReportOn = $PKXMLFileInfos
} else {
    If ($AllDomains -contains "$Domain") {
        $MyDomain = "$Domain"
        $ReportOnDomain = $PKXMLFileInfos | Where-Object Domain -eq "$Domain"
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
        $ReportOnDomain = $PKXMLFileInfos | Where-Object Domain -eq "$MyDomain"
    }
    If ( -not $AllDates) {
        $DatesArr = @()
        $i=0
        $PKXMLFileInfos | Where-Object Domain -eq "$MyDomain" | Select-Object Date -Unique | Sort-Object Date -Descending | ForEach-Object {
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


#########

Write-Host "Reporting on $($DomainReports.Domain.Count) Domains Over $($DateReports.Date.Count) Date(s)"
$ReportOn | ForEach-Object {
    $MyPKReportXMLFilePath = [String]($($_.XMLFile))

#    $MyPKReportXMLFilePath = "C:\_Metsys\_Reports\PurpleKnight\cultura.intra\20240718\SecurityAssessmentReport_20240718.xml"

    [xml]$PKReportXML = Get-Content -Path "$MyPKReportXMLFilePath"

    $MyPKReportDomain = "$($PKReportXML.Report.Info.ADforestname)"
    $MyPKReportDate = ([DateTime]$($PKReportXML.Report.Info.Generateddate)).ToString('yyyyMMdd')
    $ReportFileName = "PurpleKnightAnalysis_$MyPKReportDomain.$MyPKReportDate.html"

    Write-Host "Reporting For PurpleKnight Analysis of $MyPKReportDate From $MyPKReportXMLFilePath" -ForegroundColor Blue

    $PKReportInfos = $PKReportXML.Report.Info
    $PKReportRisks = $PKReportXML.Report.Risk
    $InfosPKReport = New-Object 'System.Collections.Generic.List[System.Object]'
    $Obj = [PSCustomObject]@{ReportDate  = [String]($MyPKReportDate)}
    $PKReportInfosProps  = ($PKReportInfos | Get-Member -MemberType "Property").Name

    $PKReportInfosProps | ForEach-Object {
        $MyProp = [String]($_)
        $MyValue = [String]($PKReportInfos.("$MyProp"))
        $obj | Add-Member -NotePropertyName "$MyProp" -NotePropertyValue "$([String]$($MyValue))"
        
    }
    $InfosPKReport.Add($Obj)

    $DataPKIOEs = New-Object 'System.Collections.Generic.List[System.Object]'
    $PKReportRisks | ForEach-Object {
        switch ($_.Severity) {
            "Critical" { $MyLevel = 1 }
            "Warning" { $MyLevel = 2 }
            "Informational" { $MyLevel = 3 }
            Default { $MyLevel = 4}
        }
        $obj = [PSCustomObject]@{
            ShortName       = [String]$($_.ShortName)
            Name            = [String]$($_.Name)
            Score           = [int]$($_.Score)
            Severity        = [String]$($_.Severity)
            Level           = [Int]$($MyLevel)
            Status          = [String]$($_.Status.replace(" Found",""))
            Target          = [String]$($_.Target)
            Weight          = [int]$($_.Weight)
            Category        = [String]$($_.Category)
            Description     = [String]$($_.Description)
            Numberofresults = [int]$($_.Numberofresults)
        }
        $DataPKIOEs.Add($Obj)
    }


    New-Item -Path "$ReportsPath" -Name "$MyPKReportDomain" -ItemType Directory -Force | Out-Null

    New-HTML -TitleText "Purple Knight Analysis of $MyPKReportDomain By $AuditorCompany" -Author "$ConsultantName" -Encoding UTF8 {
        Enable-HTMLFeature -Feature FontsAwesome
        New-HTMLFooter -HTMLContent { "<center>&copy; $(Get-Date -Format "yyyy") - <font color=""$Violet""><a href =""$AuditorURL"" target=_blank>$AuditorCompany</a></font></center>" }
        
        New-HTMLContent -HeaderText "$MyPKReportDomain" -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
    
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
    
 #               New-HTMLText  -Text "$MyPKReportDomain" -Color $Or -FontVariant small-caps -FontWeight bold -FontSize 42
                New-HTMLImage -Source "$ImagesPath\spklogo.png" -Width "160" -Inline -UrlLink "https://www.purple-knight.com" -Target _blank
                New-HTMLFontIcon -IconSolid info-circle -IconSize 22
                New-HTMLText  -Text "Purple Knight v$($InfosPKReport.ToolVersion)" -Color $Or -FontSize 22 -Alignment center
                New-HTMLFontIcon -IconSolid calendar-day -IconSize 22
                New-HTMLText  -Text "$MyPKReportDate" -Color $Or -FontSize 22 -Alignment center
                #New-HTMLImage -Source "$ImagesPath\spklogo.png" -Width "160" -Inline -UrlLink "https://www.purple-knight.com" -Target _blank
    
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
        New-HTMLContent -HeaderText 'Scores' -HeaderTextSize 22 -HeaderTextColor $Violet -HeaderTextAlignment center -HeaderBackGroundColor $Or {
    
            New-HTMLColumn -Width 33% {
                New-HTMLHeading h2 -HeadingText "Active Directory"
                
                $MyScore = $($InfosPKReport.ADsecurityposturescore)
                $MyScoreInfo = ScoreNoteInfo $MyScore
                $MyScoreNote = $MyScoreInfo.Split(";")[0]
                $MyScoreColor = $MyScoreInfo.Split(";")[1]
                New-HTMLText -Text "$MyScore"  -Color $MyScoreColor -FontSize 22 -FontStyle italic

                DrawMyScore $MyScore  "ActiveDirectory" "$MyScoreNote" "$MyScoreColor"
                New-HTMLImage -Source "$DatasPath\$MyPKReportDomain\$MyPKReportDate\ActiveDirectory.png" -Width "120"
    
                New-HTMLText -Text "$MyPKReportDomain"
    
                New-HTMLText -Text "IOEs found : $($InfosPKReport.ADIOEsfound)"
                New-HTMLText -Text "Failed to run : $($InfosPKReport.ADFailedtorun)" -FontWeight bold -Color Red
    
    
            } -AlignContentText center
    
            New-HTMLColumn -Width 33% {
                New-HTMLHeading h2 -HeadingText "Entra ID"
                $MyScore = $($InfosPKReport.EntraIDsecurityposturescore)
                $MyScoreInfo = ScoreNoteInfo $MyScore
                $MyScoreNote = $MyScoreInfo.Split(";")[0]
                $MyScoreColor = $MyScoreInfo.Split(";")[1]
                New-HTMLText -Text "$MyScore"  -Color $MyScoreColor -FontSize 22 -FontStyle italic
                DrawMyScore $MyScore  "EntraID" "$MyScoreNote" "$MyScoreColor"
                New-HTMLImage -Source "$DatasPath\$MyPKReportDomain\$MyPKReportDate\EntraID.png" -Width "120"
                #New-HTMLText -Text "$MyScore : $MyScoreNote"  -Color $MyScoreColor
    
                New-HTMLText -Text "$($InfosPKReport.EntraIDtenant)"
    
                New-HTMLText -Text "IOEs found : $($InfosPKReport.EntraIDIOEsfound)"
                New-HTMLText -Text "Failed to run : $($InfosPKReport.EntraIDFailedtorun)" -FontWeight bold -Color Red
    
            } -AlignContentText center
    
            New-HTMLColumn -Width 33% {
                New-HTMLHeading h2 -HeadingText "Okta"
                $MyScore = $($InfosPKReport.Oktasecurityposturescore)
                $MyScoreInfo = ScoreNoteInfo $MyScore
                $MyScoreNote = $MyScoreInfo.Split(";")[0]
                $MyScoreColor = $MyScoreInfo.Split(";")[1]
                New-HTMLText -Text "$MyScore"  -Color $MyScoreColor -FontSize 22 -FontStyle italic
                DrawMyScore $MyScore  "Okta" "$MyScoreNote" "$MyScoreColor"
                New-HTMLImage -Source "$DatasPath\$MyPKReportDomain\$MyPKReportDate\Okta.png" -Width "120"
                #New-HTMLText -Text "$MyScore : $MyScoreNote"  -Color $MyScoreColor
    
                New-HTMLText -Text "$($InfosPKReport.OktadomainURL)"
    
                New-HTMLText -Text "IOEs found : $($InfosPKReport.OktaIOEsfound)"
                New-HTMLText -Text "Failed to run : $($InfosPKReport.OktaFailedtorun)" -FontWeight bold -Color Red
    
            } -AlignContentText center
        }
        
        New-HTMLContent -HeaderText 'Graphs & Stats' -HeaderTextSize 22 -HeaderTextColor $Or -HeaderTextAlignment center -HeaderBackGroundColor $Violet {
      
            New-HTMLColumn -Width 25% {
                $Allindicators = [Int]($InfosPKReport.Selectedindicators)
                New-HTMLChart {
                    #New-ChartPie -Name "Selectedindicators $($MyPKReportInfos.Selectedindicators)" -Value [Int]$($MyPKReportInfos.Selectedindicators)
                    New-ChartPie -Name "Evaluated" -Value $([Int]([Int]($InfosPKReport.ADEvaluated) + [Int]($InfosPKReport.EntraIDEvaluated) + [Int]($InfosPKReport.OktaEvaluated)) / $Allindicators * 100)
                    New-ChartPie -Name "Failed to run" -Value $([Int]([Int]($InfosPKReport.ADFailedtorun) + [Int]($InfosPKReport.EntraIDFailedtorun) + [Int]($InfosPKReport.OktaFailedtorun))/ $Allindicators * 100)
                    New-ChartPie -Name "Canceled" -Value $([Int]([Int]($InfosPKReport.ADCanceled) + [Int]($InfosPKReport.EntraIDCanceled) + [Int]($InfosPKReport.OktaCanceled)) / $Allindicators * 100)
                    New-ChartPie -Name "Not selected" -Value $([Int]([Int]($InfosPKReport.ADNotselected) + [Int]($InfosPKReport.EntraIDNotselected) + [Int]($InfosPKReport.OktaNotselected)) / $Allindicators *100)
                    New-ChartPie -Name "Not relevant" -Value $([Int]([Int]($InfosPKReport.ADNotrelevant) + [Int]($InfosPKReport.EntraIDNotrelevant) + [Int]($InfosPKReport.OktaNotrelevant)) / $Allindicators *100)
                } -Title "$Allindicators Indicators Analysis" -TitleAlignment center -TitleColor $Violet -Height 250
            }
    
            New-HTMLColumn -Width 25% {
                New-HTMLChart {
                    $PKReportRisks | Group-Object Severity | ForEach-Object {
                        $MyName = [String] ($_.Name)
                        $MyValue = [int] ($_.Count)
                        #$Points = [int] ($_.Points)
                        New-ChartPie -Name "$MyName" -Value $MyValue
                    }
                } -Title "Indicators Found - By Severity" -TitleAlignment center -TitleColor $Violet -Height 250
            }
    
            New-HTMLColumn -Width 25% {
                New-HTMLChart {
                    $PKReportRisks | Group-Object Category | ForEach-Object {
                        $MyName = [String] ($_.Name)
                        $MyValue = [int] ($_.Count)
                        #$Points = [int] ($_.Points)
                        
                        New-ChartPie -Name "$MyName" -Value $MyValue
                    }
                } -Title "Indicators Found - By Category" -TitleAlignment center -TitleColor $Violet -Height 250
            }
    
    
            New-HTMLColumn -Width 25% {
                New-HTMLChart {
                    $PKReportRisks | Group-Object Target | ForEach-Object {
                        $MyName = [String] ($_.Name)
                        $MyValue = [int] ($_.Count)
                        #$Points = [int] ($_.Points)
                        New-ChartPie -Name "$MyName" -Value $MyValue
                    }
                } -Title "Indicators Found - By Target" -TitleAlignment center -TitleColor $Violet -Height 250
            }
            
    
        }
    
        New-HTMLContent -HeaderText 'Indicators Found' -HeaderTextSize 22 -HeaderTextColor $Violet -HeaderTextAlignment center -HeaderBackGroundColor $Or {
    
            New-HTMLTable -Title "Indicators Found" -DataTable ($DataPKIOEs | Sort-Object Level) -HideFooter -HideButtons -DisablePaging {
                $LevelColors | ForEach-Object {
                    $MyLevelIndex = $LevelColors.IndexOf("$($_)")
                    New-HTMLTableCondition -Name "Level" -ComparisonType string -Operator eq -Value "$($MyLevelIndex)" -BackgroundColor $LevelColors[$($MyLevelIndex)] -Row
                }
                New-HTMLTableCondition -Name "Level" -ComparisonType string -Operator eq -Value "3" -Color White -Row
            }
        }
    } -FilePath "$ReportsPath\$MyPKReportDomain\$ReportFileName" 
   
}

