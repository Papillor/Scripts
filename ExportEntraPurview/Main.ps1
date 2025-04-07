Write-Host "=== DEMARRAGE DU SCRIPT DE FUSION DES ASSIGNATIONS ===" -ForegroundColor Cyan

$entraData = $null
$purviewData = $null

# Appels aux scripts 1 Get-EntraRoles.ps1 et 2 Get-PurviewRoles.ps1
. .\Get-EntraRoles.ps1 -EntraOutput ([ref]$entraData)
. .\Get-PurviewRoles.ps1 -PurviewOutput ([ref]$purviewData)

# Merged des deux variables
$mergedAssignments = @()
if ($entraData) { $mergedAssignments += $entraData }
if ($purviewData) { $mergedAssignments += $purviewData }

if ($mergedAssignments.Count -eq 0) {
    Write-Host "Aucune assignation trouvée. Aucun fichier généré." -ForegroundColor Yellow
    return
}

# Export CSV avec les assignations merged
$exportPath = "$env:USERPROFILE\Desktop\Export_Entra_Purview_Merged.csv"
$mergedAssignments | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
Write-Host "Fusion des assignations exportée avec succes : $exportPath" -ForegroundColor Green

# Appel au script 3 (Get-UserDetails.ps1)
Write-Host "=== DEMARRAGE DU SCRIPT 3 : Récupération des détails utilisateurs ===" -ForegroundColor Cyan
. .\Get-UserDetails.ps1 -ImportCsvPath $exportPath
Write-Host "=== FIN DU TRAITEMENT DU SCRIPT 3 ===" -ForegroundColor Cyan

# Appel au script 4 (Get-UserLicences.ps1)
Write-Host "=== DEMARRAGE DU SCRIPT 4 : Récupération des licences utilisateurs ===" -ForegroundColor Cyan
. .\Get-UserLicences.ps1 -ImportCsvPath $exportPath
Write-Host "=== FIN DU TRAITEMENT DU SCRIPT 4 ===" -ForegroundColor Cyan

# Appel au script 5 (Get-SigninInfos.ps1)
Write-Host "=== DEMARRAGE DU SCRIPT 5 : Récupération des informations de connexion utilisateurs ===" -ForegroundColor Cyan
. .\Get-SigninInfos.ps1 -ImportCsvPath $exportPath
Write-Host "=== FIN DU TRAITEMENT DU SCRIPT 5 ===" -ForegroundColor Cyan
