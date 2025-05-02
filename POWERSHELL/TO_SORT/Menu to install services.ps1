Function showmenu {
    Clear-Host
    Write-Host "Menu to install services, etc..."
    Write-Host "1. Change team name."
    Write-Host "2. Active Directory."
    Write-Host "3. Internet Information Services."
    Write-Host "4. Restart PC."
    Write-Host "5. Exit menu."
}

Function showsubmenu2 {
    Clear-Host
    Write-Host "Menu Active Directory"-ForegroundColor Green
    Write-Host "2.1. Install Service Active Directory."
    Write-Host "2.2. Create users in the Active Directory."
    Write-Host "2.3. Create organizational unit in the Active Directory."
    Write-Host "2.4. Back to main menu."
}

Function showsubmenu3 {
    Clear-Host
    Write-Host "Menu Internet Information Services."-ForegroundColor Green
    Write-Host "3.1. Install Service IIS."
    Write-Host "3.2. Create website in IIS."
    Write-Host "3.3. Change anonymous authentication settings."
    Write-Host "3.4. Back to main menu."
}

showmenu

while(($inp = Read-Host -Prompt "Choose an option") -ne "5"){

    switch($inp){
        1 {
            Clear-Host
            $teamName = Read-Host "Insert the new team name."
            Clear-Host
            Rename-Computer -NewName "$teamName"
            Write-Host "The new team name will be $teamName " -ForegroundColor Green
            pause;
            break
        }
        2 {
            showsubmenu2
            while(($subinp = Read-Host -Prompt "Elige una opción") -ne "2.4"){
                switch($subinp){
                    2.1 {
                        Clear-Host
                        Write-Host "Installing the Active Directory Service." -ForegroundColor Green 
                        Clear-Host 
                        Write-Host "Fill in the details below for the installation to be successfull." -ForegroundColor Red -BackgroundColor White
                        $DomainMode = Read-Host "Insert the version of the Server to be installed e.g. Win2012R2." 
                        $DomainName = Read-Host "Insert the domain name that will be ending.local." 
                        $NetBiosNameOfDomain = Read-Host "Insert the name of the domain's Netbios."           
                        Clear-Host
                        Write-Host "Creating and installing the Active Directory $DomainName" -ForegroundColor Green
                        Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools
                        Import-Module ADDSDeployment
                        Install-ADDSForest `
                        -CreateDnsDelegation:$false `
                        -DatabasePath "C:\Windows\NTDS" `
                        -DomainMode $DomainMode `
                        -DomainName "$DomainName.local" `
                        -DomainNetbiosName $NetBiosNameDomain `
                        -ForestMode $DomainMode `
                        -InstallDns:$true `
                        -NoRebootOnCompletion:$false `
                        -SysvolPath "C:\Windows\SYSVOL" `
                        -Force:$true
                        pause; 
                        break
                    }
                    2.2{
                        Clear-Host
                        Write-Host "Fill in the required data to create the user!" -ForegroundColor Red -BackgroundColor White
                        $user = Read-Host "Insert the user's name."
                        $userHome = Read-Host "Insert the name of the user you will be logging in with."
                        $DomainName = Read-Host "Insert the domain name."
                        $OU = Read-Host "Insert the name of the Organizational Unit to which the user will belong."
                        $Password = Read-Host "Insert a password for the user."
                        Clear-Host
                        Write-Host "Creating User......" -ForegroundColor Green
                        $Password = (ConvertTo-SecureString "$Password" -AsPlainText -force)
                        New-ADUSer -Name $user -Sam $userHome -Path "OU=$OU,DC=$DomainName,DC=local" -AccountPassword $Password
                        pause;
                        break
                    }
                    2.3{
                        Clear-Host
                        Write-Host "¡Rellene los datos requeridos para crear la Unidad Organizativa!" -ForegroundColor Red -BackgroundColor White
                        $nombreOU = Read-Host "Inserte el nombre de la Unidad Organizativa a crear."
                        $nombredominioou = Read-Host "Inserte el nombre del dominio."
                        Clear-Host
                        Write-Host "Creando Unidad Organizativa......" -ForegroundColor Green
                        New-ADOrganizationalUnit -Name $nombreOU -Path "dc=$nombredominioou,dc=local"
                        pause; 
                        break 
                    }
                    default {
                        Write-Host -ForegroundColor Red -BackgroundColor White "Opción Incorrecta. Selecciona una opción del 2.1 al 2.2"
                        pause
                    }
                }
                showsubmenu2
            }
            break
        }
        3 {
            showsubmenu3
            while(($subinp = Read-Host -Prompt "Elige una opción") -ne "3.4"){
                switch($subinp){
                    3.1 {
                        Clear-Host
                        Write-Host "Instalando el servicio IIS" -ForegroundColor Green
                        Install-WindowsFeature -name Web-Server –IncludeManagementTools
                        pause;
                        break
                        }
                    3.2 {
                            Clear-Host
                            Write-Host "¡Rellene los datos requeridos para crear su página web en IIS!"  -ForegroundColor Red -BackgroundColor White
                            $nombredelaweb = Read-Host "Inserte el nombre de la Website"
                            $mensajepágina = Read-Host "Inserte el mensaje de su página web"
                            $carpetadesuweb = Read-Host "Inserte el nombre de la carpeta de su página web"
                            $puerto = Read-Host "Inserte el puerto para mostrar su página"
                            Clear-Host
                            Write-Host "¡Creando la página web $nombredelaweb!"
                            Import-Module webadministration

                            Set-Location IIS:\AppPools\

                            $web = New-Item C:\$carpetadesuweb\ –ItemType directory -Force

                            "<html>$mensajepágina</html>" | Out-File C:\$carpetadesuweb\index.html

                            $Website = New-Website -Name "$nombredelaweb" -HostHeader "" -Port $puerto -PhysicalPath $web  -ApplicationPool "DefaultAppPool"
                            pause;
                            break
                        }
                     3.3 {
                        Clear-Host
                        Write-Host "Cambiando la configuración de autenticación anónima" -ForegroundColor Green
                        Import-Module WebAdministration
                        $iisServer = Get-WebConfiguration -PSPath "IIS:\"
                        $authenticationSection = $iisServer.GetSection("system.webServer/security/authentication/anonymousAuthentication")
                        $authenticationSection.OverrideMode = "Allow"
                        $authenticationSection.CommitChanges()

                        pause
                        break
                    }
                    default {
                        Write-Host -ForegroundColor Red -BackgroundColor White "Opción Incorrecta. Selecciona una opción del 6.1 al 6.2"
                        pause
                    }
                }
                showsubmenu3
            }
            break
        }
        4 {
            Clear-Host
            Write-Host "Reiniciando equipo" -ForegroundColor Red
            Shutdown -r
            pause;
            break
        }
        default {
            Write-Host -ForegroundColor Red -BackgroundColor White "Opción Incorrecta. Selecciona una opción del 1 al 4"
            pause
        }
    }

    showmenu
}
