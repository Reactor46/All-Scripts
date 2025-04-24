$ver = $host | select version
if ($ver.Version.Major -gt 1) {$Host.Runspace.ThreadOptions = "ReuseThread"}
Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

##
#This Script Retrieves The Password for a specific SharePoint Managed Account
##

##
#Begin Setting Script Variables
##

$AccountToRetrieve = "Domain\User"

##
#Create Functions
##

Function VerifyTimerJob ($Filter)
{
$Timer = Get-SPTimerJob | ? {$_.displayname -like $Filter}
If ($Timer)
{
$timer.Delete()
}
}

##
#Begin Script
##

#Retrieve the SharePoint farm
$Farm = get-spfarm | select name

#Get the Configuration Database from the farm
$Configdb = Get-SPDatabase | ? {$_.name -eq $Farm.Name.Tostring()}

#Get the Managed Account we defined as a script variable
$ManagedAccount = get-SPManagedAccount $AccountToRetrieve

#Create a new Web Application, with managed account identified
$WebApplication = new-SPWebApplication -Name "Temp Web Application" -url "http://tempwebapplication" -port 80 -AuthenticationProvider (New-SPAuthenticationProvider) -DatabaseServer $Configdb.server.displayname -DatabaseName TempWebApp_DB -ApplicationPool "Password Retrieval" -ApplicationPoolAccount $ManagedAccount -hostheader "tempwebapplication"

#Retrieve the password, assign it to a variable
$Password = cmd.exe /c $env:windir\system32\inetsrv\appcmd.exe list apppool "Password Retrieval" /text:ProcessModel.Password

#Output the Password to the screen for the administrator
Write-Host "Password for Account "  $AccountToRetrieve  " is " $Password

#Set a timer job filter
$Filter = "Unprovisioning *" + $Webapplication.Displayname + "*"

#Clean up any left-over unprovisioning jobs, and delete the web application and associated objects
VerifyTimerJob($Filter)
Remove-SPWebApplication $WebApplication -DeleteIISSite -RemoveContentDatabases -Confirm:$False
VerifyTimerJob($Filter)

#Clean up any left-over web application provisioning jobs
$ProvisionJobs = Get-SPTimerJob | ? {$_.displayname -like "provisioning web application*"}
if ($ProvisionJobs)
{
    foreach ($ProvisionJob in $ProvisionJobs)
    {
        $ProvisionJob.Delete()
    }
}