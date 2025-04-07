param (
    [ref]$EntraOutput
)

# Authentification Microsoft Graph

Connect-MgGraph -Scopes "RoleManagement.Read.All", "User.Read.All" -ErrorAction Stop

Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop

Write-Host "[1/5] - Recuperation des roles et assignations Entra..." -ForegroundColor Yellow

try {
    $entraRoles = Get-MgRoleManagementDirectoryRoleDefinition | Select-Object Id, DisplayName
    Write-Host "$($entraRoles.Count) roles Entra recuperes."

    $entraAssignments = Get-MgRoleManagementDirectoryRoleAssignment -All -ErrorAction Stop
    Write-Host "$($entraAssignments.Count) assignations Entra recuperees.`n"

    $assignmentsWithUsers = @()

    foreach ($role in $entraRoles) {
        $roleAssignments = $entraAssignments | Where-Object { $_.RoleDefinitionId -eq $role.Id }

        if ($roleAssignments.Count -eq 0) {
            $assignmentsWithUsers += [PSCustomObject]@{
                RoleDisplayName = $role.DisplayName
                Assigned        = $false
                UserName        = "Aucun utilisateur assigne"
                UserId          = ""
                Source          = "Entra"
            }
            continue
        }

        foreach ($assignment in $roleAssignments) {
            if ($assignment.PrincipalId) {
                try {
                    # Tentative pour recuperer un utilisateur
                    $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction Stop
                    $assignmentsWithUsers += [PSCustomObject]@{
                        RoleDisplayName = $role.DisplayName
                        Assigned        = $true
                        UserName        = $user.DisplayName
                        UserId          = $user.Id
                        Source          = "Entra"
                    }
                } catch {
                    # Si l'utilisateur n'est pas trouve, essayer de recuperer un service/application
                    try {
                        $servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $assignment.PrincipalId -ErrorAction Stop
                        $assignmentsWithUsers += [PSCustomObject]@{
                            RoleDisplayName = $role.DisplayName
                            Assigned        = $true
                            UserName        = $servicePrincipal.DisplayName
                            UserId          = $servicePrincipal.Id
                            Source          = "Entra"
                        }
                    } catch {
                        $assignmentsWithUsers += [PSCustomObject]@{
                            RoleDisplayName = $role.DisplayName
                            Assigned        = $false
                            UserName        = "Service/Utilisateur introuvable"
                            UserId          = ""
                            Source          = "Entra"
                        }
                    }
                }
            }
        }
    }

    $EntraOutput.Value = $assignmentsWithUsers

} catch {
    Write-Host "Erreur lors du traitement Entra : $_" -ForegroundColor Red
}
