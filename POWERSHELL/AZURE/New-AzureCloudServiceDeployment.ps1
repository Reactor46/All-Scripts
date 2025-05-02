<#
.SYNOPSIS 
    Creates a new deployment to an existing Azure cloud service slot

.DESCRIPTION
    This runbook sets up a connection to an Azure subscription by calling Connection-Azure runbook, 
    and then create a new deployment to the specified cloud service slot. If there is deployment existing 
    on the given slot already, it first deletes the old deployment before creating the new one.
     
.PARAMETER AzureConnectionName
    Name of the Azure connection setting that was created in the Automation service.
    This connection setting contains the subscription id and the name of the certificate setting that 
    holds the management certificate.

.PARAMETER StorageAccountName
    Name of the Azure storage account that is currently associated with the Azure connection. This is 
    the storage account where build package is uploaded to. A copy of updated cscfg is also uploaded.

.PARAMETER Slot
    Name of the deployment slot.

.PARAMETER PackageFile
    Path/Uri of the cspkg file.

.PARAMETER ServiceConfigurationFile
    Path/Uri of the cscfg file.

.PARAMETER ServiceName
    Name of the Azure cloud service to deploy.

.PARAMETER DeploymentLabel
    Optional. If not specified, it will be service name appended with time stamp.

.EXAMPLE
    New-AzureCloudServiceDeployment -AzureConnectionName "Visual Studio Ultimate with MSDN" -StorageAccountName "MyStorageAccount" `
    -Slot "Staging" -PackageFile "https://mystorage.blob.core.windows.net/mycontainer/MyPackage.cspkg" -ServiceName "MyCloudService" `
    -ServiceConfigurationFile C:\Temp\Temp.cscfg


.NOTES
    AUTHOR: Jing Chang
    LASTEDIT: April 30, 2014 
#>

workflow New-AzureCloudServiceDeployment
{
    param(
       
                [parameter(Mandatory=$true)]
                [String]$AzureConnectionName,

                [parameter(Mandatory=$true)]
                [String]$StorageAccountName,

                [parameter(Mandatory=$true)]
                [String]$Slot,

                [parameter(Mandatory=$true)]
                [String]$PackageFile,

                [parameter(Mandatory=$true)]
                [String]$ServiceConfigurationFile,
          
                # cloud service name for the swap
                [Parameter(Mandatory = $true)] 
                [String]$ServiceName,

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
    
    $Deployment = Get-AzureDeployment -Slot $Slot -ServiceName $ServiceName -ErrorAction Ignore
    if ($Deployment -ne $null -AND $Deployment.DeploymentId  -ne $null)
    {
            Write-Output (" Current Status of {0} with {1}: " -f $Slot, $ServiceName, $Deployment.Status)
            
            # suspend the deployment if it's not suspended
            if($Deployment.Status -ne "Suspended")
            {
                Set-AzureDeployment -Status -ServiceName $ServiceName -Slot $Slot -NewStatus "Suspended"
            } 

            $DeleteStatus = Remove-AzureDeployment -ServiceName $ServiceName -Slot $Slot -DeleteVHD -Force -ErrorAction Stop
            Write-Output ("Status of deleting {0}: {1}" -f $ServiceName, $DeleteStatus.OperationStatus)    
    }else
    {
            Write-Verbose ("There is no deployment in {0} slot of {1}." -f $Slot, $ServiceName)
    }

    Checkpoint-Workflow
    Connect-Azure -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName
    Select-AzureSubscription -SubscriptionName $AzureConnectionName

    write-progress -id 3 -activity "Creating New Deployment" -Status "In progress"
    Write-Output "$(Get-Date -f $timeStampFormat) - Creating New Deployment In progress: New-AzureDeployment -Slot $Slot -Package $PackageFile -Configuration $ServiceConfigurationFile -label $DeploymentLabel -ServiceName $servicename -ErrorAction Suspend"
    
    $opstat = New-AzureDeployment -Slot $Slot -Package $PackageFile -Configuration $ServiceConfigurationFile -label $DeploymentLabel -ServiceName $servicename -ErrorAction Stop
    Write-Verbose ("$opstat")
    Checkpoint-Workflow
    Connect-Azure -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName
    Select-AzureSubscription -SubscriptionName $AzureConnectionName

    $completeDeployment = Get-AzureDeployment -ServiceName $servicename -Slot $Slot -ErrorAction Stop
    $completeDeploymentID = $completeDeployment.deploymentid
  
    write-progress -id 3 -activity "Creating New Deployment" -completed -Status "Complete"
    Write-Output "$(Get-Date -f $timeStampFormat) - Creating New Deployment: Complete, Deployment ID: $completeDeploymentID"   
    
    $Finish = [System.DateTime]::Now
    $TotalUsed = $Finish.Subtract($Start).TotalSeconds
}