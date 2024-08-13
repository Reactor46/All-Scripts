<#
.SYNOPSIS 
    Check to see if the build uploaded to the storage acocunt is updated compared to the system record

.DESCRIPTION
    This runbook checks the last updated time of the cspkg and cscfg files stored in the designated storage
     account. It then compared with the previous recorded update time which is logged as automation variables. 
     If one of the build files is updated, then it triggers a new deployment.
     

.PARAMETER AzureConnectionName
    Name of the Azure connection setting that was created in the Automation service.
    This connection setting contains the subscription id and the name of the certificate setting that 
    holds the management certificate.

.PARAMETER StorageAccountName
    Name of the Azure storage account that is currently associated with the Azure connection. This is 
    the storage account where build package is uploaded to. A copy of updated cscfg is also uploaded.

.PARAMETER ContainerName
    Name of the Azure storage container for build packages.

.PARAMETER CscfgBlobName
    Name of the cscfg file.

.PARAMETER CspkgBlobName
    Name of the cspkg file.

.EXAMPLE
    Check-AzureCloudServiceBuildStatus -AzureConnectionName "Visual Studio Ultimate with MSDN" -StorageAccountName "MyStorageAccount" -ContainerName "MyContainer" `
    -CscfgBlobName "cloud.cscfg" -CspkgBlobName "cloud.cspkg"


.NOTES
    AUTHOR: Jing Chang
    LASTEDIT: April 30, 2014 
#>
workflow Check-AzureCloudServiceBuildStatus
{
     param(
            
        [parameter(Mandatory=$true)]
        [String]
        $AzureConnectionName,
            
        [parameter(Mandatory=$true)]
        [String]
        $StorageAccountName,

        [parameter(Mandatory=$true)]
        [String]
        $ContainerName,

        [parameter(Mandatory=$true)]
        [String]
        $CscfgBlobName,
        
        [parameter(Mandatory=$true)]
        [String]
        $CspkgBlobName
    )
    
    $IsBuildUpdated = $false
    $cscfgUpdated = $false
    $cspkgUpdated = $false
    
    $PrevCscfgBuildTime = Get-AutomationVariable -Name 'CscfgBuildTime'
    $PrevCspkgBuildTime = Get-AutomationVariable -Name 'CspkgBuildTime'
        
    # Call the Connect-Azure Runbook to set up the connection to Azure using the Automation connection asset
    Connect-Azure -AzureConnectionName $AzureConnectionName -StorageAccountName $StorageAccountName
    Select-AzureSubscription -SubscriptionName $AzureConnectionName

    # access the build location to get the latest timestamp
    $cscfgBlob = Get-AzureStorageBlob -Container $ContainerName -Blob $CscfgBlobName
    if ($cscfgBlob -ne $null)
    {
        $cscfgblobTime = [DateTime]$cscfgBlob.LastModified
        if ($cscfgblobTime -gt $PrevCscfgBuildTime)
        {
           $cscfgUpdated = $true
           Set-AutomationVariable -Name 'CscfgBuildTime' -Value $cscfgblobTime
        }
    }

    $cspkgBlob = Get-AzureStorageBlob -Container $ContainerName -Blob $CspkgBlobName
    if ($cspkgBlob -ne $null)
    {
        $cspkgblobTime = [DateTime]$cspkgBlob.LastModified
        if ($cspkgblobTime -gt $PrevCspkgBuildTime)
        {
           $cspkgUpdated = $true
           Set-AutomationVariable -Name 'CspkgBuildTime' -Value $cspkgblobTime
        }
    }

    $IsBuildUpdated = ($cscfgUpdated -or $cspkgUpdated)
    $cspkgAbsoluteUri = "https://" + $StorageAccountName + ".blob.core.windows.net/" + $ContainerName + "/" + $CspkgBlobName
    $cscfgAbsoluteUri = "https://" + $StorageAccountName + ".blob.core.windows.net/" + $ContainerName + "/" + $CscfgBlobName
   
    $BuildStatus = @{"IsBuildUpdated" = $IsBuildUpdated; "cspkgAbsoluteUri" = $cspkgAbsoluteUri; "cscfgAbsoluteUri" = $cscfgAbsoluteUri }
    
    # if one of the files is updated, set $isBuildUpdated to $true
    Write-Output $BuildStatus
}