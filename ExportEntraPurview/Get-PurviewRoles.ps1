param (
    [ref]$PurviewOutput
)

Add-Type -AssemblyName System.Windows.Forms
Write-Host "`n[2/5] - Recuperation des roles et membres Purview..." -ForegroundColor Yellow

$purviewAssignments = @()

try {
    Connect-IPPSSession

    $purviewRoles = Get-RoleGroup | Select-Object Name

    foreach ($role in $purviewRoles) {
        Write-Host "Traitement du role : $($role.Name)"

        try {
            $members = Get-RoleGroupMember -Identity $role.Name

            if ($members.Count -eq 0) {
                $purviewAssignments += [PSCustomObject]@{
                    RoleDisplayName = $role.Name
                    Assigned        = $false
                    UserName        = "Aucun utilisateur assigne"
                    UserId          = ""
                    Source          = "Purview"
                }
            } else {
                foreach ($member in $members) {
                    $purviewAssignments += [PSCustomObject]@{
                        RoleDisplayName = $role.Name
                        Assigned        = $true
                        UserName        = $member.DisplayName
                        UserId          = $member.Id
                        Source          = "Purview"
                    }
                }
            }

        } catch {
            $purviewAssignments += [PSCustomObject]@{
                RoleDisplayName = $role.Name
                Assigned        = $false
                UserName        = "Erreur de recuperation"
                UserId          = ""
                Source          = "Purview"
            }
            Write-Host "Erreur sur $($role.Name) : $_" -ForegroundColor Red
        }
    }

    $PurviewOutput.Value = $purviewAssignments

} catch {
    Write-Host "Erreur lors du traitement Purview : $_" -ForegroundColor Red
}
