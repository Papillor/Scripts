param (
    [string]$ImportCsvPath
)

# Authentification Graph 
try {
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"
        Write-Host "Connecté à Microsoft Graph." -ForegroundColor Green
    }
} catch {
    Write-Host "Erreur de connexion à Microsoft Graph : $_" -ForegroundColor Red
    return
}

# Sélection du CSV si non fourni
if (-not $ImportCsvPath) {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Sélectionner le fichier CSV des utilisateurs enrichis"
    $dialog.Filter = "CSV files (*.csv)|*.csv"
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $null = $dialog.ShowDialog()
    $ImportCsvPath = $dialog.FileName
}

if (-not (Test-Path $ImportCsvPath)) {
    Write-Host "Fichier CSV introuvable." -ForegroundColor Red
    return
}

# Import CSV
$users = Import-Csv -Path $ImportCsvPath
$finalExport = @()

foreach ($user in $users) {
    $upn = $user.UPN1
    if ([string]::IsNullOrWhiteSpace($upn) -and $user.UPN2) {
        $upn = $user.UPN2
    }

    $userId = ""
    $licenseNames = ""
    $licenseIds = ""

    # Ignorer si non applicable
    if ([string]::IsNullOrWhiteSpace($upn) -or $upn -match 'non demandé|non concerné|erreur') {
        $licenseNames = "Non applicable"
    } else {
        try {
            $mgUser = Get-MgUser -UserId $upn -ErrorAction Stop
            $userId = $mgUser.Id
        } catch {
            Write-Host "Impossible de récupérer l'ID pour $upn : $($_.Exception.Message)" -ForegroundColor Yellow
            $licenseNames = "Erreur récupération ID"
        }

        # Si on a l'ID, on tente de récupérer les licences
        if ($userId) {
            try {
                $assignedLicenses = Get-MgUserLicenseDetail -UserId $userId -ErrorAction Stop

                if ($assignedLicenses.Count -gt 0) {
                    $licenseNames = ($assignedLicenses | Select-Object -ExpandProperty SkuPartNumber) -join " / "
                    $licenseIds   = ($assignedLicenses | Select-Object -ExpandProperty SkuId) -join " / "
                } else {
                    $licenseNames = "Aucune licence"
                }
            } catch {
                Write-Host "Erreur récupération des licences pour $upn : $($_.Exception.Message)" -ForegroundColor Yellow
                $licenseNames = "Erreur récupération licences"
            }
        }
    }

    # Ajout à l'export
    $finalExport += [PSCustomObject]@{
        RoleDisplayName = $user.RoleDisplayName
        Assigned        = $user.Assigned
        UserName        = $user.UserName
        Source          = $user.Source
        UPN1            = $user.UPN1
        UPN2            = $user.UPN2
        UserId          = $userId
        Licenses        = $licenseNames
        LicenseIds      = $licenseIds
    }
}

# Export final
$saveDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveDialog.Title = "Enregistrer le fichier enrichi avec les licences"
$saveDialog.Filter = "CSV files (*.csv)|*.csv"
$saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$saveDialog.FileName = "Assignations_Licenses_Enriched.csv"
$null = $saveDialog.ShowDialog()
$ExportPath = $saveDialog.FileName

if ($ExportPath) {
    $finalExport | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Fichier enrichi avec licences exporté : $ExportPath" -ForegroundColor Green
} else {
    Write-Host "Export annulé par l'utilisateur." -ForegroundColor Yellow
}
