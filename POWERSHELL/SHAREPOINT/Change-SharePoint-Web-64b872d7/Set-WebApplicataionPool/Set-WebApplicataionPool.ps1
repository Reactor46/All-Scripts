#--------------------------------------------------------------------------------------- 
# Name:            Set-WebApplicataionPool.ps1 
# Description:     This script will change SP WebApplication Pools for a Web Application 
#                 
# Usage:        Run the function with the required parameters 
# By:             Ivan Josipovic, Softlanding.ca 
#--------------------------------------------------------------------------------------- 
Function Set-WebApplicataionPool($WebAppURL,$ApplicationPoolName){ 
    $apppool = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.ApplicationPools | where {$_.Name -eq $ApplicationPoolName} 
    if ($apppool -eq $null){ 
        write-host -foreground red "The Application Pool $ApplicationPoolName does not exist!" 
        return 1 
    } 
    $webapp = get-spwebapplication -Identity $WebAppUrl 
    if ($webapp -eq $null){ 
        write-host -foreground red "The Web Application $WebAppUrl does not exist!" 
        return 1 
    } 
    $webapp.Applicationpool = $apppool 
    $webApp.Update() 
    $webApp.ProvisionGlobally() 
    write-host -foreground green "$WebappURL Application Pool has been changed to $ApplicationPoolName" 
    return 0 
} 
 
Set-WebApplicataionPool -WebAppURL "http://sp2010-a:9006" -ApplicationPoolName "SharePoint WebApplications" 