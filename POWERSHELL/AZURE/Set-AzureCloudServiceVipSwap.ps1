<#
.SYNOPSIS 
    Swap the deployments between production and staging of an Azure cloud service.

.DESCRIPTION
    This runbook sets up a connection to an Azure subscription by calling Connection-Azure runbook, 
    It then checks to see if the staging slot of given cloud service is deployed already. If it is deployed, 
    then swap the deployments between production and staging. This is useful to roll back to the previous 
    deployment kept in staging slot.

.PARAMETER AzureConnectionName
    Name of the Azure connection setting that was created in the Automation service.
    This connection setting contains the subscription id and the name of the certificate setting that 
    holds the management certificate.

.PARAMETER StorageAccountName
    Name of the Azure storage account that is currently associated with the Azure connection. This is 
    the storage account where build package is uploaded to. A copy of updated cscfg is also uploaded.

.PARAMETER ServiceName
    Name of the Azure cloud service to swap.

.EXAMPLE
    Set-AzureCloudServiceVipSwap -AzureConnectionName "Visual Studio Ultimate with MSDN" -StorageAccountName "MyStorageAccount" `
    -ServiceName "MyCloudService"

.NOTES
    AUTHOR: Jing Chang
    LASTEDIT: April 30, 2014 
#>
workflow Set-AzureCloudServiceVipSwap
{
    param(
       
                [parameter(Mandatory=$true)]
                [String]$AzureConnectionName,
                
                [parameter(Mandatory=$true)]
                [String]$StorageAccountName,
            
                # cloud service name for the swap
                [Parameter(Mandatory = $true)] 
                [String]$ServiceName
                
    )

    $Start = [System.DateTime]::Now
    "Starting: " + $Start.ToString("HH:mm:ss.ffffzzz")

    Checkpoint-Workflow
    Connect-Azure -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName
    Select-AzureSubscription -SubscriptionName $AzureConnectionName

    $Deployment = Get-AzureDeployment -Slot "Staging" -ServiceName $ServiceName -ErrorAction Ignore
    if ($Deployment -ne $null -AND $Deployment.DeploymentId  -ne $null)
    {
            Write-Output (" Current Status of staging with {0}: {1}" -f $ServiceName, $Deployment.Status)
            
            $MoveStatus = Move-AzureDeployment -ServiceName $ServiceName
            Write-Output ("Vip swap of {0} status: {1}" -f $ServiceName, $MoveStatus.OperationStatus)    
    }else
    {
           throw ("There is no deployment in staging slot of {0} to swap." -f $ServiceName)
    }
    
    # stop staging after vipswap
    $Deployment = Get-AzureDeployment -Slot "Staging" -ServiceName $ServiceName -ErrorAction Ignore
    if ($Deployment -ne $null -AND $Deployment.DeploymentId  -ne $null)
    {
            # suspend the deployment if it's not suspended
            if($Deployment.Status -ne "Suspended")
            {
                $SuspensionStatus = Set-AzureDeployment -Status -ServiceName $ServiceName -Slot "Staging" -NewStatus "Suspended"
                Write-Output ("Suspended staging slot of {0} with status: {1}" -f $ServiceName, $SuspensionStatus.OperationStatus) 
            } 
    }

    $Finish = [System.DateTime]::Now
    $TotalUsed = $Finish.Subtract($Start).TotalSeconds
   
    Write-Output ("VIP swapped cloud service {0} in {1} seconds." -f $ServiceName, $TotalUsed)
    "Finished " + $Finish.ToString("HH:mm:ss.ffffzzz")
} 

