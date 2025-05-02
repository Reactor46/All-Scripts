<#
Name: ActivateWinComputersfrAD.ps1
Author: System Center MVP - Steve Buchanan
Date: 2/15/2015
Version: 1.0
Website: www.buchatech.com

Description:
This script can be used to loop through an OU in Active Directory and activate all computers in that OU.
This script will find only computers with "Windows Server" in the name. 
Run this script using: powershell.exe -executionpolicy unrestricted -command .\ActivateWinComputersfrAD.ps1
#>

# Load the Active Directory PowerShell module
Import-Module -Name ActiveDirectory

# Prompt script runner for information to create variables
#$domain = 'fnbm' #Read-host 'Enter domain to be used. Format as such (DOMAINNAME)'
#$computersou = 'OU=WEB Servers - 2012,OU=Production_Servers,OU=Servers,OU=Las_Vegas,DC=fnbm,DC=corp' #Read-host 'Enter the name of the OU to be searched.'
$Productkey = "TPGJW-NRR7Q-323KJ-43YVW-GVHJ8" #Read-host 'Enter product key. Format as such (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)'

# Create a variable that holds all of the computers from Active Directory
$results = (Get-ADComputer -Filter {OperatingSystem -Like 'Windows Server 2012 R2 Data*' -and Name -Like "LASDMZCHAT0*"} -SearchBase "OU=WEB Servers - 2012,OU=Production_Servers,OU=Servers,OU=Las_Vegas,DC=fnbm,DC=corp")

# Loop through the results variable and activate all computers in that variable. 
# NOTE: Dont forget to replace the $key variable with your own Windows key. 

foreach ($i in $results) 
{
$computer = gc env:computername
$service = get-wmiObject -query "select * from SoftwareLicensingService" -computername $i.Name
$service.InstallProductKey($Productkey)
$service.RefreshLicenseStatus()
 }

Write-Host "The following servers have been activated:" -ForegroundColor Green
$results | Format-Table DNSHostName -HideTableHeaders

Pause