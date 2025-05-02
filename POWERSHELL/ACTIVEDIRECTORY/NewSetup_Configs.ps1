#Requires -RunAsAdministrator

#Check ExecutionPolicy is at a minimum of Bypass
    Write-Host 'Validating your ExecutionPolicy...'
    $ep=Get-ExecutionPolicy
    if(($ep -eq 'Unrestricted') -OR ($ep -eq 'Bypass')){
    }else{
        try {
            Write-Host "Your current ExecutionPolicy: $ep"
            Write-Host 'Setting ExecutionPolicy: Bypass'
            Set-ExecutionPolicy -ExecutionPolicy Bypass
        }
        catch {
            Write-Host $_
            Break
        }
    }

#Check for Version 5 and install if it isn't
    Write-Host 'Checking for PowerShell V5...'
    if(($PSVersionTable.PSVersion).Major -ne 5){
        Write-Host 'PowerShell V5 is not installed, installing first. Afterwards, you will need to reboot and then re-run this script.'
        $LocPath = 'C:\Windows\Temp'
        $RemPath = '\\gsm1900\public\EIT Infrastructure Operations\Storage\Software\PowerShell\PS5.0andWMF5.0_RTM'
        $Win72008R2PS5='Win7AndW2K8R2-KB3134760-x64.msu'
        $Win82012R2PS5='Win8.1AndW2K12R2-KB3134758-x64.msu'
        $OSVer = ((Get-WmiObject -class win32_OperatingSystem).Version).substring(0,3)
        if($OSVer -eq '6.1'){
            Copy-Item -Path (Join-Path -Path $RemPath -ChildPath $Win72008R2PS5) -Destination $LocPath
            $invpath = (Join-Path -Path $LocPath -ChildPath $Win72008R2PS5)
            $invcmd = "$invpath /quiet /promptrestart"
        }elseif($OSVer -eq '6.3'){
            Copy-Item -Path (Join-Path -Path $RemPath -ChildPath $Win82012R2PS5) -Destination $LocPath
            $invpath = (Join-Path -Path $LocPath -ChildPath $Win82012R2PS5)
            $invcmd = "$invpath /quiet /promptrestart"
        }
        try {
            Write-Host "Installing from $invpath.  This will take a few mins and will prompt for restart at end."
            Invoke-Expression $invcmd
        }
        catch {
            Write-Warning $_
            Write-Warning 'Unable to install PowerShell V5. Exiting.'
            Break
        }
    }

#Update help
    Write-Host 'Updating help'
    Update-Help -Force -ErrorAction "SilentlyContinue"

#Set the module repo so we can easily install PS5 mods
    $psrep = 'PSGallery'
    $pssrcloc = 'https://www.powershellgallery.com/api/v2/'
    $pstp = 'Trusted'
    $psregrep = (Get-PSRepository).Name
    $compobj=Compare-Object $psrep $psregrep | Where-Object SideIndicator -eq '<=' | Select-Object InputObject

    if($compobj){
        Write-Host 'Setting the PSRepository and Policy'
        Write-Host 'If the NuGet package manager has never been installed, it will ask to install it. Type Y when prompted.'
        try {
            Register-PSRepository -Name PSGallery -SourceLocation 'https://www.powershellgallery.com/api/v2/' -InstallationPolicy Trusted
            Set-PSRepository -Name PSGallery
        }
        catch {
            Write-Warning $_
        }
    }

    if((Get-PSRepository).InstallationPolicy -ne 'Trusted'){
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }

#Remove viprshell if installed
    $remmodlist = 'viprshell'
    $installedmods = (Get-Module -ListAvailable).Name
    $compobj=Compare-Object $remmodlist $installedmods | Where-Object SideIndicator -eq '<=' | Select-Object InputObject
    if($compobj.InputObject -eq 'viprshell'){
        foreach($modobj in $compobj.InputObject){
            try {
                Write-Host "Removing module: $modobj"
                Remove-Module -Name $modobj
                $whereatloc=(Get-ChildItem ENV:PSModulePath).Value -Split ';'
                $remobjloc = Join-Path -Path ($whereatloc | Where-Object {$_ -like "*Program Files*"}) -ChildPath '\viprshell'
                Write-Host "Remove-Item -path $remobjloc"
                #Remove-Item -Path $remobjloc -Recurse
            }
            catch {
                Write-Warning 'Unable to remove module ViPRShell.  Please investigate and remove manually.'
                Write-Warning "Location: $remobjloc"
                Write-Warning $_
            }
        }
    }
#Import PowerShell Modules that are not hosted on PowerShell Gallery
    $impmodlist = 'Posh-ViPR'
    $installedmods=(Get-Module -ListAvailable).Name
    $compobj=Compare-Object $impmodlist $installedmods | Where-Object SideIndicator -eq '<=' | Select-Object InputObject
    if($compobj.InputObject){
        foreach($modobj in $compobj.InputObject){
            try {
                Write-Host "Importing module: $modobj"
                $modobjremloc = '\\gsm1900\public\EIT Infrastructure Operations\Storage\Software\PowerShell\Modules\Posh-ViPR'
                $wheretoloc = ((Get-ChildItem ENV:PSModulePath).Value).Split(';')
                $modobjloc = $wheretoloc | Where-Object {$_ -like "*Program Files*"}
                Copy-Item -Path $modobjremloc -Destination $modobjloc -Recurse
                Import-Module -Name $modobj
            }
            catch {
                Write-Warning 'Unable to import Posh-ViPR.  Check to see if it copied over.'
                Write-Warning "Location: $modobjloc"
                Write-Warning $_
            }

        }
    }else{
        Write-Host 'All Importable PSGallery modules imported'
    }

#Install PSGallery Modules that are hosted on PowerShell Gallery
    $modlist = 'PSLogging','Posh-SSH','PSScriptAnalyzer','ImportExcel','PureStoragePowerShellSDK','OnCommand-Insight','PureStoragePowerShellToolkit'

    $installedmods = (Get-Module -ListAvailable).Name
    $compobj = Compare-Object $modlist $installedmods | Where-Object SideIndicator -eq '<=' | Select-Object InputObject
    if($compobj.InputObject){
        foreach($modobj in $compobj.InputObject){
                try {
                    Write-Host "Installing module: $modobj"
                    Install-Module -Name $modobj
                }
                catch {
                    Write-Warning "Unable to install $modobj"
                    Write-Warning $_
                }
        }
    }else{
        Write-Host 'All PSGallery modules installed'
    }

