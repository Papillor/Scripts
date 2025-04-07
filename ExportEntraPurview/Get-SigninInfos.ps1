param (
    [string]$ImportCsvPath
)

# Importer les modules nécessaires
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Identity.SignIns
Import-Module Microsoft.Graph.Reports

# Authentification Graph
try {
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes "AuditLog.Read.All", "Directory.Read.All", "User.Read.All", "UserAuthenticationMethod.Read.All" -NoWelcome
        Write-Host "Connecté à Microsoft Graph." -ForegroundColor Green
    }
} catch {
    Write-Host "Erreur de connexion à Graph : $_" -ForegroundColor Red
    return
}

# Sélection fichier CSV
if (-not $ImportCsvPath) {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Sélectionner le fichier CSV enrichi avec les IDs"
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
$results = @()

# Récupérer les rapports d'authentification pour tous les utilisateurs
$authMethodsData = Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop

foreach ($user in $users) {
    $userId = $user.UserId
    $lastInteractive = ""
    $lastNonInteractive = ""
    $authMethods = ""
    $groups = ""

    if ([string]::IsNullOrWhiteSpace($userId)) {
        $lastInteractive = "Non applicable"
        $lastNonInteractive = "Non applicable"
        $authMethods = "Non applicable"
        $groups = "Non applicable"
    } else {
        try {
            # Activité de connexion
            $userActivity = Get-MgUser -UserId $userId -Property SignInActivity | Select-Object DisplayName, UserPrincipalName, SignInActivity
            
            if ($userActivity.SignInActivity) {
                $lastInteractive = $userActivity.SignInActivity.LastSignInDateTime
                $lastNonInteractive = $userActivity.SignInActivity.LastNonInteractiveSignInDateTime
            } else {
                $lastInteractive = "Aucune donnée de connexion"
                $lastNonInteractive = "Aucune donnée de connexion"
            }

            # Groupes
            $memberOf = Get-MgUserMemberOf -UserId $userId -All -ErrorAction Stop
            $groups = ($memberOf | ForEach-Object { 
                if ($_.AdditionalProperties["displayName"]) { 
                    $_.AdditionalProperties["displayName"]
                }
            }) -join " / "

            # Méthodes d'authentification
            $userAuth = $authMethodsData | Where-Object { $_.UserPrincipalName -eq $user.UPN1 }

            if ($userAuth) {
                $methodList = $userAuth.MethodsRegistered -join " / "
                $mfaStatus = if ($userAuth.IsMfaRegistered) { "MFA activé" } else { "MFA non activé" }
                $authMethods = "$methodList | $mfaStatus"
            } else {
                $authMethods = "Utilisateur non trouvé dans le rapport"
            }

        } catch {
            Write-Host "Erreur pour l'utilisateur $($user.UPN1) : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    $results += [PSCustomObject]@{
        RoleDisplayName          = $user.RoleDisplayName
        Assigned                 = $user.Assigned
        UserName                 = $user.UserName
        Source                   = $user.Source
        UPN1                     = $user.UPN1
        UPN2                     = $user.UPN2
        UserId                   = $user.UserId
        Licenses                 = $user.Licenses
        LicenseIds               = $user.LicenseIds
        LastInteractiveSignIn    = $lastInteractive
        LastNonInteractiveSignIn = $lastNonInteractive
        AuthenticationMethods    = $authMethods
        Groups                   = $groups
    }
}

# Export final
$saveDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveDialog.Title = "Enregistrer le fichier final enrichi"
$saveDialog.Filter = "CSV files (*.csv)|*.csv"
$saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$saveDialog.FileName = "Export_Complet.csv"
$null = $saveDialog.ShowDialog()
$ExportPath = $saveDialog.FileName

if ($ExportPath) {
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Export final enrichi généré : $ExportPath" -ForegroundColor Green
} else {
    Write-Host "Export annulé." -ForegroundColor Yellow
}