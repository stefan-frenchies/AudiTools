<#

#>

Import-Module Microsoft.Graph.Security

Connect-MgGraph -Scopes SecurityEvents.Read.All

$SecurityScores = Get-MgSecuritySecureScore -Property *

$MySecurityScore = $SecurityScores  | Select-Object -Last 1 

$MySecurityScore | fl

ActiveUserCount
AverageComparativeScores
AzureTenantId
ControlScores
CreatedDateTime
CurrentScore
EnabledServices
Id
LicensedUserCount
MaxScore
AdditionalProperties

{
    "id": "String (identifier)",
    "azureTenantId": "String",
    "activeUserCount": "Int32",
    "createdDateTime": "String (timestamp)",
    "currentScore": "Double",
    "enabledServices": ["String"],
    "licensedUserCount": "Int32",
    "maxScore": "Double",
    "averageComparativeScores": [{"@odata.type": "microsoft.graph.averageComparativeScore"}],
    "controlScores": [{"@odata.type": "microsoft.graph.controlScore"}],
    "vendorInformation": {"@odata.type": "microsoft.graph.securityVendorInformation"},
    }

$SecurityScoreControlProfiles = Get-MgSecuritySecureScoreControlProfile -Property *

$MySecurityScoreControlProfile = $SecurityScoreControlProfiles | Select-Object -First 1

$MySecurityScoreControlProfile | fl


ActionType
ActionUrl
AzureTenantId
ComplianceInformation
ControlCategory
ControlStateUpdates
Deprecated
Id
ImplementationCost
LastModifiedDateTime
MaxScore
Rank
Remediation
RemediationImpact
Service
Threats
Tier
Title
UserImpact
VendorInformation
AdditionalProperties

{
    "actionType": "String",
    "actionUrl": "String",
    "azureTenantId": "String",
    "complianceInformation": [{"@odata.type": "microsoft.graph.complianceInformation"}],
    "controlCategory": "String",
    "controlStateUpdates": [{"@odata.type": "microsoft.graph.secureScoreControlStateUpdate"}],
    "deprecated": "Boolean",
    "id": "String (identifier)",
    "implementationCost": "String",
    "lastModifiedDateTime": "String (timestamp)",
    "maxScore": "Double",
    "rank": "Int32",
    "remediation": "String",
    "remediationImpact": "String",
    "service": "String",
    "threats": ["String"],
    "tier": "String",
    "title": "String",
    "userImpact": "String",
    "vendorInformation": {"@odata.type": "microsoft.graph.securityVendorInformation"}
  }