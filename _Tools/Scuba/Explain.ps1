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


# Définir le chemin du fichier JSON d'entrée et du fichier CSV de sortie
$jsonFilePath = "C:\Users\s.giraud-ext\Cultura\Team Infra Microsoft - Documents\Audit\Infos\Scuba_2024_12_13_10_39_57\cultura.json"
$csvFilePath = "C:\Users\s.giraud-ext\Cultura\Team Infra Microsoft - Documents\Audit\Infos\Scuba_2024_12_13_10_39_57\report.csv"

# Fonction pour convertir récursivement un objet JSON en un tableau de hachage
function Convert-JsonToHashtable {
    param (
        [Parameter(Mandatory = $true)]
        [PSObject]$JsonObject
    )

    $result = @()

    foreach ($item in $JsonObject) {
        if ($item.Value -is [PSCustomObject]) {
            $nestedResult = Convert-JsonToHashtable -JsonObject $item.Value
            foreach ($nestedItem in $nestedResult) {
                $result += [PSCustomObject]@{ ($item.Key + '.' + $nestedItem.Key) = $nestedItem.Value }
            }
        } else {
            $result += [PSCustomObject]@{ $item.Key = $item.Value }
        }
    }

    return $result
}

# Lire le fichier JSON
$jsonContent = Get-Content -Path $jsonFilePath -Raw | ConvertFrom-Json

# Convertir le contenu JSON en tableau de hachage
$hashTable = Convert-JsonToHashtable -JsonObject $jsonContent

# Exporter le tableau de hachage en fichier CSV
$hashTable | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Output "Conversion terminée. Le fichier CSV est disponible à l'emplacement : $csvFilePath"
