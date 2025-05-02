<#
Get-WindowsFeature is a good command to review feature names.
Please read through all of the comments.
This script should be run on the new domain controller. It will
go out and connect to your primary DC remotely.
#>

function CheckFSMO {
    $fsmotransferred = $true
    
    Write-Host "Checking domain forest..." -ForegroundColor Yellow
    
    $adforest = Get-ADForest $domain | Select SchemaMaster, DomainNamingMaster
    $addomain = Get-ADDomain $domain | Select PDCEmulator,RIDMaster,InfrastructureMaster
    
    #Checking Schema
    if($adforest.SchemaMaster -like "$($computername)*") {
        Write-Host "Schema Master was successfully transferred!" -ForegroundColor Green
    }
    else {
        Write-Error "Schema Master was not successfully transferred."
        $fsmotransferred = $false
    }
    #Checking DomainNamingMaster
    if($adforest.DomainNamingMaster -like "$($computername)*") {
        Write-Host "Naming Master was successfully transferred!" -ForegroundColor Green
    }
    else {
        Write-Error "Naming Master was not successfully transferred."
        $fsmotransferred = $false
    }
    #Checking PDCEmulator
    if($addomain.PDCEmulator -like "$($computername)*") {
        Write-Host "PDC Emulator was successfully transferred!" -ForegroundColor Green
    }
    else {
        Write-Error "PDC Emulator was not successfully transferred."
        $fsmotransferred = $false
    }
    #Checking RIDMaster 
    if($addomain.RIDMaster -like "$($computername)*") {
        Write-Host "RID Master was successfully transferred!" -ForegroundColor Green
    }
    else {
        Write-Error "RID Master was not successfully transferred."
        $fsmotransferred = $false
    }
    #Checking InfrastructureMaster
    if($addomain.InfrastructureMaster -like "$($computername)*") {
        Write-Host "Infrastructure Master was successfully transferred!" -ForegroundColor Green
    }
    else {
        Write-Error "Infrastructure Master was not successfully transferred."
        $fsmotransferred = $false
    }
    
    return $fsmotransferred
}

function RemoveCertificateRole ($targetcomputer) {
    $certrole = Invoke-Command -ComputerName $targetcomputer -ScriptBlock { Get-WindowsFeature AD-Certificate }
    
    if($certrole.Installed) {
        Invoke-Command -ComputerName $targetcomputer -ScriptBlock { Remove-WindowsFeature AD-Certificate }
    }
    
    return !$certrole.Installed
}

function RemoveDomainServices ($targetcomputer) {
    # Testing for Domain Services install
    $domainservices = Invoke-Command -ComputerName $targetcomputer -ScriptBlock { Get-WindowsFeature AD-Domain-Services }
    
    $command = {
        Uninstall-ADDSDomainController -LocalAdministratorPassword $domainpass -Confirm:$false
    }
    
    if($domainservices.Installed) {
        Invoke-Command -ComputerName $targetcomputer -ScriptBlock $command
    }
    
    return !$domainservices.Installed
}

$domain         = Read-Host "Domain Name"
$primarydc      = Read-Host "Primary Domain Controller"
$dsrm           = Read-Host -AsSecureString  "DSRM Password" 
$computername   = (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain           

$domainuser     = Read-Host "Domain Admin Username"
$domainpass     = Read-Host -AsSecureString  "Password" 

$domaincreds    = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $domainuser, $domainpass
$dcinstalled    = Get-WindowsFeature AD-Domain-Services

# Should be customized to your environment. As is, gets the static from the first interface.
# Used to set DNS on the old domain server.
# Does not require you to have a static assigned.
$dns            = Read-Host "New Machine's Static IP"

# If Domain Services aren't installed, install it. 
if(!$dcinstalled.Installed) {
    Write-Host "Domain Service are not installed. Now installing..." -ForegroundColor Red
    Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools
    Import-Module ADDSDeployment
    
    # This step will automatically reboot your computer
    
    Write-Warning "You should set your DNS across your network accordingly beforing continuing this script at reboot."
    Write-Warning "This script will automatically set the primary domain controller's DNS."
    
    Install-ADDSDomainController -DomainName $domain `
                                 -InstallDns `
                                 -Credential $domaincreds `
                                 -SafeModeAdministratorPassword $dsrm `
                                 -Confirm:$false # Remove if you want to be asked.
}

Restart-Computer -Force

Write-Host "Domain Services have been successfully installed." -ForegroundColor Green

$session = New-PSSession -ComputerName $primarydc

if (-not($session)) {
    Write-Error "$primarydc could not be contacted. Exiting script."
    Exit
}
else {
    Remove-PSSession $session
    
    Write-Host "Successfully contacted primary DC. Now transferring FSMO roles..." -ForegroundColor Green

    Move-ADDirectoryServerOperationMasterRole `
    -Identity $env:computername `
    -OperationMasterRole 0,1,2,3,4 `
    -Confirm:$false ` # Remove if you want to be asked.
    
    $command = {
        param ($dnsserver)
        <#
        This should be customized in reference to your environment.
        As is, it will set DNS on the first interface.
        #>
        Get-NetIPInterface | Set-DNSClientServerAddress -ServerAddress($dnsserver)
    }
    
    # This will happen without confirmation. See above for changing that if needed.
    Invoke-Command -ComputerName $primarydc -ScriptBlock $command -ArgumentList $dns
    
    if(CheckFSMO)
    {
        # Need to remove Certificate Role first...
        if(RemoveCertificateRole($primarydc))
        {
            Write-Host "Certificate Services has been successfully removed." -ForegroundColor Green
            Write-Warning "Now demoting the remote domain controller..."
            
            if(!(RemoveDomainServices($primarydc))) { Write-Warning "Now demoting the remote domain controller. Upon reboot, your domain will be successfully migrated." }
        }
    }
}