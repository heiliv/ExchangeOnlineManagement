#Conectar a Exchange Online
Import-Module ExchangeOnlineManagement
$LiveCred = Get-Credential
Connect-ExchangeOnline -Credential $LiveCred
Connect-MsolService -Credential $LiveCred
Write-Host "Conexion establecida"
Pause

# Dar opcion de elegir qué accion realizar
$opcion = Read-Host "¿Que accion desea realizar? (1: Migrar usuarios, 2: Crear alias, 3: Eliminar usuarios, 4: Asignar licencias)"

# Realizar accion seleccionada
switch ($opcion) {
    1 {
        # Migrar usuarios
        Import-Csv -Path "C:/csvUsuarios/usuarios.csv" | ForEach-Object {
            Write-Host $_.UserPrincipalName
            New-MsolUser `
            -DisplayName $_.DisplayName `
            -Password $_.Password `
            -UserPrincipalName $_.UserPrincipalName `
            -StrongPasswordRequire $True `
            -ForceChangePassword $True
        }
    }
    2 {
        # Crear alias
        $CSVPath = "C:\csvUsuarios\alias.csv"
        $CSV = Import-Csv -Path $CSVPath
        foreach ($Line in $CSV) {
            $User = $Line.User
            $Alias = $Line.Alias
            Set-Mailbox -Identity $User -Alias $Alias
        }
    }
    3 {
        # Eliminar usuarios
        Import-Csv "C:\csvUsuarios\usuarios.csv" | ForEach-Object {
            Remove-MsolUser -UserPrincipalName $_.UserPrincipalName -Force
        }
    }
    4 {
        # Asignar licencias
        Import-Csv "C:\csvUsuarios\usuarios.csv" | ForEach-Object {
            Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses "reseller-account:O365_BUSINESS_PREMIUM"
        }
    }
    default {
        Write-Host "Opcion invalida, saliendo del script."
    }
}
