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

#PSGallery Trusted

Install-Module ORCA
Connect-ExchangeOnline

$ORCAFiles = Invoke-ORCA -ShowSurvey $false -Output @("HTML","JSON","CSV") -OutputOptions @{CSV=@{OutputDirectory="C:\_Cultura\_tmp"};JSON=@{OutputDirectory="C:\_Cultura\_tmp"};HTML=@{DisplayReport=$False;EmbedConfiguration=$True;OutputDirectory="C:\_Cultura\_tmp"}} -AssessmentLevel All -AlternateDNS 8.8.8.8
