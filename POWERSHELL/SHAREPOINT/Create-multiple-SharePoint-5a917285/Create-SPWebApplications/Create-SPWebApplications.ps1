<#
.SYNOPSIS
Create SharePoint Web Applications from a .csv file which contains all definitions.

.DESCRIPTION
Create SharePoint Web Applications from a .csv file which contains all definitions.
Before run the script, create a .csv file which contains all required and optional 
informations for your web application as Name, ApplicationPool, etc.
If you need more options, just modified the .csv and script files for your requirements.
The script verify the ApplicationPoolAccount if it is not registered as a managed accounts
then will take it.

.EXAMPLE

Create-SPWebApplications.ps1 -importfile <Path>
Create-SPWebApplications.ps1.ps1 -importfile "C:\workdir\WebApps.csv"
#>

#NAME: Create-SPWebApplications.ps1
#AUTHOR: Tibor Revesz
#DATE: 07/11/2013

param (
        [Parameter(Mandatory=$true)]
        [string]$importfile=""
      )

# Add SharePoint powershell snapin.
Add-PSSnapin "Microsoft.SharePoint.PowerShell"

# To create a claims-based authentication provider.
$ap = New-SPAuthenticationProvider

# Import all details from a .csv file and create web applications.
# If you use other delimiter, just change it in the script.
Import-Csv -Path $importfile -Delimiter ";" |
    ForEach-Object {
       if (((Get-SPManagedAccount).UserName -notcontains $_.AppPoolUserName) -eq $True)
          {
           $appPoolCred = Get-Credential $_.appPoolUserName
           New-SPManagedAccount -Credential $appPoolCred
          }
          New-SPWebApplication -Name $_.AppName -ApplicationPool $_.AppPoolName`
            -ApplicationPoolAccount $_.AppPoolUserName `
            -DatabaseName $_.DatabaseName -HostHeader $_.HostHeader `
            -URL $_.url -Port $_.port -AuthenticationProvider $ap
    }
