<#

#>



Param(
    [Parameter(Mandatory=$false, Position=0)]
    [ValidateRange(0, [int]::MaxValue)]
    [Int] $Last = 0
)


Clear-Host
Import-Module Microsoft.Graph.Security

Connect-MgGraph -Scopes SecurityEvents.Read.All -NoWelcome

Write-host "Reading Information"
$SecurityScores = Get-MgSecuritySecureScore -Property *

#$MyLast = $SecurityScores | Select-Object -Last 1


$MyReportsHistory = New-Object 'System.Collections.Generic.List[System.Object]'
$MyReportsDetails = New-Object 'System.Collections.Generic.List[System.Object]'

Write-host "Analysing Information"
$SecurityScores | ForEach-Object {
  $MySecurityScore = $_
  #recuperation Informations tenant
  $MgOrg = Get-MgOrganization -OrganizationId $MySecurityScore.AzureTenantId
  $DisplayName    = [string]($MgOrg.DisplayName)
  $InitialDomainName  = [String](($MgOrg.VerifiedDomains | Where-Object IsInitial -eq $true).Name)
  $DefaultDomainName  = [String](($MgOrg.VerifiedDomains | Where-Object IsDefault -eq $true).Name)
  #Cfeation Objet Infos report
  $MyReportInfos = [PSCustomObject]@{
    "Id"  = [string]$($MySecurityScore.Id)
    "Date" = [String]([datetime]$($MySecurityScore.CreatedDateTime)).ToString("yyyyMMdd")
    "Tenant"  = [string]$($MySecurityScore.AzureTenantId)
    "Controls"  = [Int]$($MySecurityScore.ControlScores.Count)
    "Score" = [Double]$($MySecurityScore.CurrentScore)
    "MaxScore" = [Double]$($MySecurityScore.MaxScore)
    "Tool"   = [string]($MySecurityScore.VendorInformation.Provider)
    "Provider"   = [string]($MySecurityScore.VendorInformation.Vendor)
    "ActiveUserCount" = [Int32]$($MySecurityScore.ActiveUserCount)
    "EnabledServices" = [array]$($MySecurityScore.EnabledServices)
    "LicensedUserCount" = [Int]$($MySecurityScore.LicensedUserCount)
    "DisplayName"    = [string]$($DisplayName)
    "DomainName"   = [string]$($DefaultDomainName)
    "InitialName"   = [string]$($InitialDomainName)
  }

  $MySecurityScore.ControlScores | ForEach-Object {
    $MySecurityScoreDetail = $_
    $MyReportControls = [PSCustomObject]@{
      "Id"  = [string]$($MySecurityScore.AzureTenantId)
      "Date" = [String]([datetime]$($MySecurityScore.CreatedDateTime)).ToString("yyyyMMdd")
      "Tenant"  = [string]($MySecurityScore.AzureTenantId)
      "Category"   = [string]($MySecurityScoreDetail.ControlCategory)
      "Name"   = [string]($MySecurityScoreDetail.ControlName)
      "Description"   = [string]($MySecurityScoreDetail.Description)
      "Score"   = [Double]($MySecurityScoreDetail.Score)
    }
    #AddDetails?
    $MyReportsDetails.Add($MyReportControls)
  }
  #AddToHistory?
  $MyReportsHistory.Add($MyReportInfos)
}
Write-host "Getting Controls Information"
#Recuperation Informations sur les profils de controle de score de sécurité
$SecurityScoreControlProfiles = Get-MgSecuritySecureScoreControlProfile -Property *
$ScoresInfo = New-Object 'System.Collections.Generic.List[System.Object]'

$SecurityScoreControlProfiles | ForEach-Object { 
  $MySecurityScoreControlProfile = $_
  $MyReportControls = [PSCustomObject]@{
    "Title" = [string]$($MySecurityScoreControlProfile.Title)
    "ActionType"  = [string]$($MySecurityScoreControlProfile.ActionType)
    "ActionUrl" = [string]$($MySecurityScoreControlProfile.ActionUrl)
    "AzureTenantId" = [string]$($MySecurityScoreControlProfile.AzureTenantId)
    "ComplianceInformation" = [Array]$($MySecurityScoreControlProfile.ComplianceInformation)
    "ControlCategory" = [string]$($MySecurityScoreControlProfile.ControlCategory)
    "Deprecated"  = [Boolean]$($MySecurityScoreControlProfile.Deprecated)
    "Id"  = [string]$($MySecurityScoreControlProfile.Id)
    "ImplementationCost"  = [string]$($MySecurityScoreControlProfile.ImplementationCost)
    "MaxScore"  = [double]$($MySecurityScoreControlProfile.MaxScore)
    "Rank"  = [Int32]$($MySecurityScoreControlProfile.Rank)
    "Remediation" = [string]$($MySecurityScoreControlProfile.Remediation)
    "RemediationImpact" = [string]$($MySecurityScoreControlProfile.RemediationImpact)
    "Service" = [string]$($MySecurityScoreControlProfile.Service)
    "Threats" = [Array]$($MySecurityScoreControlProfile.Threats)
    "Tier"  = [string]$($MySecurityScoreControlProfile.Tier)
    "UserImpact"  = [string]$($MySecurityScoreControlProfile.UserImpact)
  }
  #AddDetails?
  $ScoresInfo.Add($MyReportControls)
}



$ReportsFullDetails = New-Object 'System.Collections.Generic.List[System.Object]'
$MyReportsDetails | ForEach-Object { 
  $ReportDetails = $_
  $ControlScore = $ReportDetails.Score
  $ControlInfo = $ScoresInfo | Where-Object Id -eq $ReportDetails.Name

  If ($ControlInfo) {
    $MyReportFullDetails = [PSCustomObject]@{
      "Date"  = [String]($ReportDetails.Date)
      "Description" = [String]($ReportDetails.Description)
      "Score" = [Double]($ReportDetails.Score)
      "Name" = [String]($ReportDetails.Name)
      "Category" = [String]($ReportDetails.Category)
      "Completed" = [Boolean](!($($ControlInfo.MaxScore)-$($ReportDetails.Score)))
      "Title" = [String]($ControlInfo.Title)
      "ActionType"  = [String]($ControlInfo.ActionType)
      "ActionUrl" = [String]($ControlInfo.ActionUrl)
      "ComplianceInformation" = [String]($ControlInfo.ComplianceInformation)
      "ControlCategory" = [String]($ControlInfo.ControlCategory)
      "Deprecated"  = [Boolean]($ControlInfo.Deprecated)
      "Id"  = [String]($ControlInfo.Id)
      "Tenant"  = [String]($ControlInfo.Tenant)
      "ImplementationCost"  = [String]($ControlInfo.ImplementationCost)
      "MaxScore"  = [Double]($ControlInfo.MaxScore)
      "Rank"  = [String]($ControlInfo.Title)
      "RemediationImpact" = [String]($ControlInfo.RemediationImpact)
      "Service" = [String]($ControlInfo.Service)
      "Threats" = [String]($ControlInfo.Threats)
      "Tier"  = [String]($ControlInfo.Tier)
      "UserImpact"  = [String]($ControlInfo.UserImpact)
    }
    $ReportsFullDetails.Add($MyReportFullDetails)
  }

}

$SourcesPath = "$PSScriptRoot"
#$SourcesPath = "C:\_iFrenchies\_Apps"

$ReportPath = "$SourcesPath\_Reports"
#Start Auditor informations
$AuditorColor = "Blue"
$AuditorCompany = "iFrenchies"
$AuditorLogo = "$SourcesPath\images\logo$AuditorCompany.png"
$AuditorURL = "https://www.ifrenchies.eu"

$ConsultantName = "Stephane Giraud"
$ConsultantPhone = "+33695985004"
$ConsultantMail = "sgiraud@ifrenchies.eu"

#End Auditor informations

$ToolName = ($MyReportsHistory | Select-Object -First 1).Tool
#Adaptation ToolName
$ToolName = "$($MySecurityScore.VendorInformation.Vendor) $($MySecurityScore.VendorInformation.Provider)"

$TooLogo = "$SourcesPath\images\$ToolName.png"
$AzTenantId = ($SecurityScores.AzureTenantId | Select-Object -Unique)
$ToolUrl = "https://learn.microsoft.com/en-us/defender-xdr/microsoft-secure-score-improvement-actions"

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

$NbReportselected = [Int]($SecurityScores.Count)
If ($Last) {$NbReportselected = $Last}


#Building Reports Datas



#Report Generation
Write-host "Generating report from $NbReportselected Analysis"
$ReportsDate = @()
$MyReportsHistory | Sort-Object Date -Descending -Unique| ForEach-Object { $ReportsDate += "$($_.Date)"}
$NbReports = $ReportsDate.Count


New-HTML -TitleText "$ClientName - $ToolName Analysis of $FQDN" -Author "$ConsultantName" -Encoding UTF8 {
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
              New-HTMLText  -Text "$NbReportselected of $NbReports Report(s) Available" -Color $HeaderColor2 -FontSize 20 -Alignment center
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

      $DataReport = $ReportsFullDetails | Where-Object { $_.Date -EQ $MyDate -and $_.Completed -eq $false}
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
<#
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
#>


          New-HTMLContent -HeaderText 'Microsoft Recommendations' -HeaderTextSize 22 -HeaderTextColor $HeaderColor1 -HeaderTextAlignment center -HeaderBackGroundColor $HeaderColor2 {

            #$DataReport = $ReportsFullDetails | Where-Object Date -EQ $MyDate 
            $MyCols = @("Id","Title","Score","MaxScore","Category","Service","Threats","ActionType","ImplementationCost","RemediationImpact","UserImpact")
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