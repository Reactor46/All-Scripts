<#
Get-WindowsFeature is a good command to review feature names.
Please read through all of the comments.
This script should be run on the new domain controller. It will
go out and connect to your primary DC remotely.
#>

$domain         = Read-Host "Domain Name"
$dsrm           = Read-Host -AsSecureString  "DSRM Password"  

# If Domain Services aren't installed, install it. 
if(!$dcinstalled.Installed) {
    Write-Host "Domain Service are not installed. Now installing..." -ForegroundColor Red
    Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools
    Import-Module ADDSDeployment
    Install-ADDSForest -DomainName $domain -InstallDns -SafeModeAdministratorPassword $dsrm
}

Write-Host "Domain Services have been successfully installed." -ForegroundColor Green