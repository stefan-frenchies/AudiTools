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

