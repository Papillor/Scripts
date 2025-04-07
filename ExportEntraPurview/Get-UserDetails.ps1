if (-not (Get-Command 'Dotenv' -ErrorAction SilentlyContinue)) {
    Install-Module -Name Dotenv -Force -Scope CurrentUser
}

# ICI CHANGER LE PATH POUR LE .ENV !!

$envPath = "C:\Users\YOURPATH"

if (Test-Path $envPath) {
    $dotenv = Get-Content $envPath | ForEach-Object {
        if ($_ -match "^\s*([^#][^=]*)=(.*)\s*$") {
            [System.Environment]::SetEnvironmentVariable($matches[1], $matches[2])
        }
    }
    Write-Host "Variables d'environnement chargées avec succès depuis le fichier .env" -ForegroundColor Green
} else {
    Write-Host "Le fichier .env n'a pas été trouvé. Vérifiez le chemin d'accès." -ForegroundColor Red
    return
}

# Sélection du fichier CSV
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "Sélectionner le fichier CSV fusionné"
$dialog.Filter = "CSV files (*.csv)|*.csv"
$dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$null = $dialog.ShowDialog()
$ImportCsvPath = $dialog.FileName

if (![string]::IsNullOrWhiteSpace($ImportCsvPath)) {
    Write-Host "Fichier sélectionné : $ImportCsvPath" -ForegroundColor Green
} else {
    Write-Host "Aucun fichier sélectionné. Le processus est annulé." -ForegroundColor Yellow
    return
}

# Authentification Microsoft Graph
try {
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -ErrorAction Stop
    Write-Host "Authentification réussie à Microsoft Graph." -ForegroundColor Green
} catch {
    Write-Host "Erreur d'authentification avec Microsoft Graph : $_" -ForegroundColor Red
    return
}

# Authentification Purview via .env
$PurviewUserName = [System.Environment]::GetEnvironmentVariable('USER_NAME')
$PurviewPassword = [System.Environment]::GetEnvironmentVariable('USER_PASSWORD')
$SecurePurviewPassword = ConvertTo-SecureString $PurviewPassword -AsPlainText -Force
$PurviewCredential = New-Object System.Management.Automation.PSCredential($PurviewUserName, $SecurePurviewPassword)

try {
    Write-Host "Connexion à Purview via Connect-IPPSSession..." -ForegroundColor Yellow
    Connect-IPPSSession -Credential $PurviewCredential -ErrorAction Stop
    Write-Host "Connexion réussie à Purview" -ForegroundColor Green
} catch {
    Write-Host "Erreur lors de la connexion à Purview : $_" -ForegroundColor Red
    return
}

# Import des assignations
$mergedAssignments = Import-Csv -Path $ImportCsvPath
$enrichedAssignments = @()

foreach ($assignment in $mergedAssignments) {
    $UPNs = @() 

    if ($assignment.Assigned -eq $true) {
        try {
            if ($assignment.Source -eq "Entra") {
                if (![string]::IsNullOrWhiteSpace($assignment.UserId)) {
                    try {
                        $user = Get-MgUser -UserId $assignment.UserId -ErrorAction Stop
                        $UPNs += $user.UserPrincipalName
                    } catch {
                        if ($_.Exception.Message -like "*Resource*does not exist*") {
                            $UPNs += "Non concerné (service/app)"
                        } else {
                            $UPNs += "Erreur Graph : $($_.Exception.Message)"
                        }
                    }
                } else {
                    $UPNs += "Aucun utilisateur assigné"
                }
            } elseif ($assignment.Source -eq "Purview") {
                try {
                    $users = Get-User -Identity $assignment.UserName -ErrorAction Stop
                    if ($users -is [array]) {
                        foreach ($u in $users) {
                            $UPNs += $u.UserPrincipalName
                        }
                    } elseif ($users -ne $null) {
                        $UPNs += $users.UserPrincipalName
                    } else {
                        $UPNs += "Utilisateur non trouvé"
                    }
                } catch {
                    $UPNs += "Erreur Purview : $($_.Exception.Message)"
                }
            }
        } catch {
            $UPNs += "Erreur générale : $($_.Exception.Message)"
        }
    } else {
        $UPNs += "Non demandé"
    }

    $record = [ordered]@{
        RoleDisplayName = $assignment.RoleDisplayName
        Assigned        = $assignment.Assigned
        UserName        = $assignment.UserName
        Source          = $assignment.Source
    }

    for ($i = 0; $i -lt $UPNs.Count; $i++) {
        $record["UPN$($i + 1)"] = $UPNs[$i]
    }

    $enrichedAssignments += New-Object PSObject -Property $record
}

# Export 
if ($enrichedAssignments.Count -eq 0) {
    Write-Host "Aucune donnée enrichie trouvée. Aucune exportation effectuée." -ForegroundColor Yellow
    return
}

$dialogExport = New-Object System.Windows.Forms.SaveFileDialog
$dialogExport.Title = "Enregistrer le fichier exporté"
$dialogExport.Filter = "CSV files (*.csv)|*.csv"
$dialogExport.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$dialogExport.FileName = "Assignations_Entra_Purview_Enriched.csv"
$null = $dialogExport.ShowDialog()
$ExportPath = $dialogExport.FileName

if (![string]::IsNullOrWhiteSpace($ExportPath)) {
    $enrichedAssignments | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Détails utilisateurs exportés avec succès : $ExportPath" -ForegroundColor Green
} else {
    Write-Host "Export annulé par l'utilisateur." -ForegroundColor Yellow
}
