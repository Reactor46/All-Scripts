<#
.SYNOPSIS 
    In-place upgrade an existing Azure cloud service slot

.DESCRIPTION
    This runbook sets up a connection to an Azure subscription by calling Connection-Azure runbook, 
    It assumes the given slot of given cloud service is deployed already. It then performs the inplace upgrade 
    with given packages. It will wait for the role to become ready and running. If combined with multiple 
    upgrade domains, which require at least 2 instances of a given role, it will offer zero downtime during inplace upgrade. 
    If there is only one instance, then the given role will be stopped and restarted after upgrade is complete.

    After upgrade, the workflow checks whether all roles are coming to ready state, until the specified time out period in minute.
    Note the parameter $Slot should be "Staging" or "Production"    

.PARAMETER AzureConnectionName
    Name of the Azure connection setting that was created in the Automation service.
    This connection setting contains the subscription id and the name of the certificate setting that 
    holds the management certificate.

.PARAMETER StorageAccountName
    Name of the Azure storage account that is currently associated with the Azure connection. This is 
    the storage account where build package is uploaded to. A copy of updated cscfg is also uploaded.

.PARAMETER ServiceName
    Name of the Azure cloud service to deploy.

.PARAMETER Slot
    Name of the slot to upgrade.

.PARAMETER NumberOfMinutesTillTimeOut
    as named, till the time to finish checking the service status.

.PARAMETER PackagePath
    Path/Uri of the cspkg file.

.PARAMETER ConfigurationPath
    Path of the cscfg file.

.PARAMETER DeploymentLabel
    Optional. If not specified, it will be service name appended with time stamp.

.EXAMPLE
    Set-AzureCloudServiceInplaceUpgrade -AzureConnectionName "Visual Studio Ultimate with MSDN" -StorageAccountName "MyStorageAccount" `
    -PackagePath "https://mystorage.blob.core.windows.net/mycontainer/MyPackage.cspkg" -ConfigurationPath "C:\Temp\Temp.cscfg" `
    -ServiceName "MyCloudService" -Slot "Staging" -NumberOfMinutesTillTimeOut 60 

.NOTES
    AUTHOR: Jing Chang
    LASTEDIT: April 30, 2014 
#>
workflow Set-AzureCloudServiceInplaceUpgrade
{
    param(
       
                [parameter(Mandatory=$true)]
                [String]$AzureConnectionName,
            
                [parameter(Mandatory=$true)]
                [String]$StorageAccountName,    
            
                # cloud service name for upgrade
                [Parameter(Mandatory = $true)] 
                [String]$ServiceName,

                [Parameter(Mandatory = $true)]
                [String]$Slot,

                [Parameter(Mandatory = $true)]
                [Int32]$NumberOfMinutesTillTimeOut,

                [Parameter(Mandatory = $true)]
                [String]$PackagePath,

                [Parameter(Mandatory = $true)]
                [String]$ConfigurationPath,

                [Parameter(Mandatory = $false)]
                [String]$DeploymentLabel
    )
  

    $Start = [System.DateTime]::Now
    "Starting: " + $Start.ToString("HH:mm:ss.ffffzzz")

    if ($DeploymentLabel -eq $null)
    {
       $DeploymentLabel = $ServiceName + $Start.ToString("HH:mm:ss.ffffzzz")
    }

    # Call the Connect-Azure Runbook to set up the connection to Azure using the Automation connection asset
    Connect-Azure -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName
    Select-AzureSubscription -SubscriptionName $AzureConnectionName
   
    Write-Verbose ("Upgrading slot {0} of {1}" -f $Slot, $ServiceName)
    Set-AzureDeployment -Upgrade -Slot $Slot -Package $Packagepath -Configuration $ConfigurationPath -label $DeploymentLabel -ServiceName $ServiceName -Force
    
    Checkpoint-Workflow
    Connect-Azure -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName
    Select-AzureSubscription -SubscriptionName $AzureConnectionName

    $DeploymentStatus = (Get-AzureDeployment -ServiceName $ServiceName -Slot $Slot).Status
    Write-Verbose ("Upgrade status is {0}" -f $DeploymentStatus)

    $RefreshIntervalInSeconds = 30
           
     # wait for role ready
     $serviceready = $true
     $totalRetries = $NumberOfMinutesTillTimeOut * 2
     $numOfRetries = 0
     while($numOfRetries -le $totalRetries)
     {
               $RoleStatus = Get-AzureRole -ServiceName $Servicename -Slot $Slot -InstanceDetails
               for($i=0; $i -lt $RoleStatus.Count; $i++)
               {
                   if($RoleStatus[$i].InstanceStatus -ne "ReadyRole")
                   {
                       $serviceready = $false
                       Write-Output ("{0} is not Ready yet. It is in {1} state." -f $RoleStatus[$i].InstanceName, $RoleStatus[$i].InstanceStatus)
                   }
                   else
                   {
                       Write-Output ("{0} is Ready" -f $RoleStatus[$i].InstanceName)
                   }
               }
        
               if($serviceready -eq $true)
               {
                   Write-Output "All Roles are ready!" 
                   $numOfRetries = $totalRetries
               }
        
               $numOfRetries++
               if($numOfRetries -ge $totalRetries -and $serviceready -eq $false)
               {
                   Write-Output "Service does not seem to be ready after trying for some time, please check the cloud service from manage.windowsazure.com for upgrade errors."
               }
            
               if ($numOfRetries -lt $totalRetries -and $serviceready -eq $false)
               {
                   Write-Output ("Will recheck in {0} seconds..." -f $RefreshIntervalInSeconds)
                   Start-Sleep -Seconds $RefreshIntervalInSeconds
                   $serviceready = $true
               }
    }
    
    $Finish = [System.DateTime]::Now
    $TotalUsed = $Finish.Subtract($Start).TotalSeconds
   
    Write-Output ("Updated cloud service {0} slot {1} in {2} seconds." -f $ServiceName, $Slot, $TotalUsed)
    "Finished " + $Finish.ToString("HH:mm:ss.ffffzzz")
}