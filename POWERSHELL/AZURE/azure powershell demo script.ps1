
Import-Module WAPPSCmdlets

# the goal of this script is to create deployments and storage accounts
# in two different Azure datacenters.  We'll upload a basic Azure package to
# each, and then create a Traffic Manager policy to do geographic load balancing.

#complete these variables with the data from your Azure account
$subid = "{your subscription id}"

# optional: add your own certificate thumbprint, either from the file or Azure portal
# $cert = Get-Item cert:\CurrentUser\My\01784B3F26B609044A56AC5B1CFEF287321420F5
$storageaccount = "somedemoaccount"
$storagekey = "{your key}"

$servicename_nc = "{some dns name}NC"   #added NC for the North Central deployment
$servicename_we = "{some dns name}WE"   #added WE for the West Europe deployment

$storagename_nc = "{some storage name}nc"
$storagename_we = "{some storage name}we"

$globaldns = "{some dns name}global"

# persisting Subscription Setings
Set-Subscription -SubscriptionName powershelldemo -Certificate $cert -SubscriptionId $subid 

# setting default Subscription
Set-Subscription -DefaultSubscription powershelldemo

# setting the current subscription to use 
Select-Subscription -SubscriptionName powershelldemo

# save the cert and subscription id for subscriptions
Set-Subscription -SubscriptionName powershelldemo -StorageAccountName $storageaccount -StorageAccountKey $storagekey

# specify the default storage account to use for the subscription
Set-Subscription -SubscriptionName powershelldemo -DefaultStorageAccount $storageaccount

# configure North Central location
New-StorageAccount -ServiceName $storagename_nc -Location "North Central US"	| Get-OperationStatus –WaitToComplete
New-HostedService -ServiceName $servicename_nc -Location "North Central US" | Get-OperationStatus –WaitToComplete

# optionally, you can add certificates to the deployment for SSL or managing the instances through RDP
#Add-Certificate -CertToDeploy "d:\powershell\cert.pfx" -Password {password} -ServiceName $servicename_nc | Get-OperationStatus –WaitToComplete

New-Deployment -serviceName $servicename_nc –StorageAccountName $storagename_nc `
   	-Label MySite `
	-slot staging -package "D:\powershell\package\PowerShellDemoSite.cspkg" –configuration "D:\powershell\package\ServiceConfiguration.Cloud.cscfg" | Get-OperationStatus -WaitToComplete

Get-Deployment -serviceName $servicename_nc -Slot staging `
	| Set-DeploymentStatus -Status Running | Get-OperationStatus -WaitToComplete

Move-Deployment -DeploymentNameInProduction $servicename_nc -ServiceName $servicename_nc -Name MySite

# configure West Europe Central
New-StorageAccount -ServiceName $storagename_we -Location "West Europe" | Get-OperationStatus –WaitToComplete
New-HostedService -ServiceName $servicename_we -Location "West Europe" | Get-OperationStatus –WaitToComplete
#Add-Certificate -CertToDeploy "d:\powershell\cert.pfx" -Password {password} -ServiceName $servicename_we | Get-OperationStatus –WaitToComplete

New-Deployment -serviceName $servicename_we –StorageAccountName $storagename_we `
   	-Label MySite `
	-slot staging -package "D:\powershell\package\PowerShellDemoSite.cspkg" –configuration "D:\powershell\package\ServiceConfiguration.Cloud.cscfg" | Get-OperationStatus -WaitToComplete

Get-Deployment -serviceName $servicename_we -Slot staging `
	| Set-DeploymentStatus -Status Running | Get-OperationStatus -WaitToComplete

Move-Deployment -DeploymentNameInProduction $servicename_we -ServiceName $servicename_we -Name MySite | Get-OperationStatus -WaitToComplete


# increase the number of instances to 2
Get-HostedService -ServiceName $servicename_we | `
              Get-Deployment -Slot Production | `
              Set-RoleInstanceCount -Count 2 -RoleName "WebRole1"


# set up a geo-loadbalance between the two using the Traffic Manager 
$profile = New-TrafficManagerProfile -ProfileName bhitneyps `
		  -DomainName ($globaldns + ".trafficmanager.net")

$endpoints = @()
$endpoints += New-TrafficManagerEndpoint -DomainName ($servicename_we + ".cloudapp.net")
$endpoints += New-TrafficManagerEndpoint -DomainName ($servicename_nc + ".cloudapp.net")
$endpoints += New-TrafficManagerEndpoint -DomainName ("www.structuretoobig.com")

# configure the endpoint Traffic Manager will monitor for service health 
$monitors = @()
$monitors += New-TrafficManagerMonitor –Port 80 –Protocol HTTP –RelativePath /

# create new definition
$createdDefinition = New-TrafficManagerDefinition -ProfileName bhitneyps -TimeToLiveInSeconds 300 `
			-LoadBalancingMethod Performance -Monitors $monitors -Endpoints $endpoints -Status Enabled
			
# enable the profile with the newly created traffic manager definition 
Set-TrafficManagerProfile -ProfileName bhitneyps -Enable -DefinitionVersion $createdDefinition.Version


# cleanup -- remove all instances, storage accounts, and Traffic Manager profile
Remove-TrafficManagerProfile bhitneyps | Get-OperationStatus -WaitToComplete
Remove-Deployment -Slot production -serviceName $servicename_nc
Remove-Deployment -Slot production -serviceName $servicename_we | Get-OperationStatus -WaitToComplete
Remove-HostedService -serviceName $servicename_nc
Remove-HostedService -serviceName $servicename_we | Get-OperationStatus -WaitToComplete
Remove-StorageAccount -StorageAccountName $storagename_nc
Remove-StorageAccount -StorageAccountName $storagename_we | Get-OperationStatus -WaitToComplete